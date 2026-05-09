import eel
import asyncio
import os
import sys
import re
import shutil
import tempfile
import subprocess
import base64
import edge_tts
import threading
import time
import win32com.client
import pandas as pd
import io

# --- 1. ПУТИ И ОКРУЖЕНИЕ ---
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
    root_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = base_dir

web_dir = os.path.join(base_dir, 'web')
audio_cache = os.path.join(web_dir, 'audio_cache')
output_dir = os.path.join(root_dir, 'Готовые аудио')

for d in [audio_cache, output_dir]: os.makedirs(d, exist_ok=True)

try:
    from bs4 import BeautifulSoup
    from ebooklib import epub
    import ebooklib
    import fitz
except ImportError:
    pass

# --- 2. ОБРАБОТКА ТЕКСТА ---
class TextProcessor:
    def __init__(self):
        self.yo_dict = {}
        self.phonetic_dict = {}
        self.load_dicts()

    def load_dicts(self):
        d_p = os.path.join(root_dir, 'dicts')
        if not os.path.exists(d_p): return
        for f, target in [('yo.dic', self.yo_dict), ('phonetic.txt', self.phonetic_dict)]:
            path = os.path.join(d_p, f)
            if os.path.exists(path):
                try:
                    with open(path, 'r', encoding='utf-8') as file:
                        for line in file:
                            if '=' in line:
                                k, v = line.strip().split('=', 1)
                                target[k.lower()] = v
                except: pass

    def process(self, text):
        if not text: return ""
        text = re.sub(r'\s+', ' ', text)
        for d in [self.yo_dict, self.phonetic_dict]:
            for k, v in d.items():
                text = re.compile(re.escape(k), re.IGNORECASE).sub(v, text)
        return text

text_processor = TextProcessor()

# --- 3. СИНТЕЗ (С АСИНХРОННОЙ ЗАЩИТОЙ И УПРАВЛЕНИЕМ ПРОЦЕССАМИ) ---
state = {"total": 0, "done": 0, "status": "idle", "chapter": ""}
stop_ev = threading.Event()
TIMEOUT_SECONDS = 45 
sapi_lock = threading.Lock() # Защита от зависаний драйверов Windows

async def run_subprocess_async(cmd, input_data=None):
    """Безопасный запуск процессов с гарантированным убийством при зависании"""
    proc = await asyncio.create_subprocess_exec(
        *cmd,
        stdin=asyncio.subprocess.PIPE if input_data else None,
        stdout=asyncio.subprocess.DEVNULL,
        stderr=asyncio.subprocess.DEVNULL,
        creationflags=0x08000000
    )
    try:
        if input_data:
            await proc.communicate(input=input_data)
        else:
            await proc.wait()
    except asyncio.CancelledError:
        try: proc.kill() # Жестко убиваем зомби-процесс
        except: pass
        raise

async def ensure_valid_mp3(filepath):
    if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
        ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
        cmd = [ffmpeg, "-y", "-f", "lavfi", "-i", "anullsrc=r=24000:cl=mono", "-t", "0.1", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", "-q:a", "4", filepath]
        await run_subprocess_async(cmd)

async def _synth_online_task(text, voice, rate, vol, pitch, out):
    tmp_mp3 = out.replace('.mp3', '_tmp.mp3')
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    try:
        await edge_tts.Communicate(text, voice, rate=rate, volume=vol, pitch=pitch).save(tmp_mp3)
        if os.path.exists(tmp_mp3):
            cmd = [ffmpeg, "-y", "-i", tmp_mp3, "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", "-q:a", "4", out]
            await run_subprocess_async(cmd)
    finally:
        if os.path.exists(tmp_mp3):
            try: os.unlink(tmp_mp3)
            except: pass

async def synth_online(text, voice, rate, vol, pitch, out, sem):
    async with sem:
        if stop_ev.is_set(): return
        try:
            await asyncio.wait_for(_synth_online_task(text, voice, rate, vol, pitch, out), timeout=TIMEOUT_SECONDS)
        except Exception as e:
            print(f"Ошибка (Online): {e}")
        finally:
            await ensure_valid_mp3(out)
            state["done"] += 1

async def _synth_offline_task(text, voice, rate, vol, out):
    tmp_wav = out.replace('.mp3', '.wav')
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    
    vol_val = float(str(vol).replace('%', '').replace('+', ''))
    v_float = max(0.1, 1.0 + (vol_val / 100.0))
    vol_filter = f"volume={v_float:.2f}"
    
    try:
        if voice.endswith('.onnx'):
            piper = os.path.join(root_dir, "piper.exe")
            r = float(str(rate).replace('%','').replace('+',''))
            scale = max(0.1, 1.0 / (1.0 + (r/100.0)))
            cmd = [piper, "-m", os.path.join(root_dir, voice), "--length_scale", f"{scale:.2f}", "-f", tmp_wav]
            await run_subprocess_async(cmd, input_data=text.encode('utf-8'))
        else:
            def run_sapi():
                with sapi_lock: # Защищаем COM-объекты Windows от мультипоточных сбоев
                    try:
                        import pythoncom; pythoncom.CoInitialize()
                        sp = win32com.client.Dispatch("SAPI.SpVoice")
                        vs = sp.GetVoices()
                        for i in range(vs.Count):
                            if voice == vs.Item(i).GetDescription(): sp.Voice = vs.Item(i); break
                        r_val = int(str(rate).replace('%','').replace('+',''))
                        sp.Rate = int(r_val/10)
                        fs = win32com.client.Dispatch("SAPI.SpFileStream")
                        fs.Open(tmp_wav, 3, False); sp.AudioOutputStream = fs
                        sp.Speak(text); fs.Close()
                    except Exception as e:
                        print(f"SAPI Inner Error: {e}")
                    finally:
                        try: pythoncom.CoUninitialize()
                        except: pass
            await asyncio.to_thread(run_sapi)
        
        if os.path.exists(tmp_wav):
            cmd = [ffmpeg, "-y", "-i", tmp_wav, "-filter:a", vol_filter, "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", "-q:a", "4", out]
            await run_subprocess_async(cmd)
    finally:
        if os.path.exists(tmp_wav):
            try: os.unlink(tmp_wav)
            except: pass

async def synth_offline(text, voice, rate, vol, out, sem):
    async with sem:
        if stop_ev.is_set(): return
        try:
            await asyncio.wait_for(_synth_offline_task(text, voice, rate, vol, out), timeout=TIMEOUT_SECONDS)
        except Exception as e:
            print(f"Ошибка (Offline): {e}")
        finally:
            await ensure_valid_mp3(out)
            state["done"] += 1

# --- 4. МЕДИА И СКЛЕЙКА ---
def generate_fx(fx_type):
    f_in = os.path.join(audio_cache, f"{fx_type}_in.mp3")
    f_out = os.path.join(audio_cache, f"{fx_type}_out.mp3")
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    
    if not os.path.exists(f_in) or os.path.getsize(f_in) < 100:
        if fx_type == "apollo":
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=2525:d=0.2", "-filter:a", "volume=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=2475:d=0.2", "-filter:a", "volume=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
        elif fx_type == "zvei":
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=1060:d=0.07", "-f", "lavfi", "-i", "sine=f=1270:d=0.07", "-f", "lavfi", "-i", "sine=f=1530:d=0.07", "-filter_complex", "concat=n=3:v=0:a=1,volume=0.15", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=1530:d=0.07", "-f", "lavfi", "-i", "sine=f=1270:d=0.07", "-f", "lavfi", "-i", "sine=f=1060:d=0.07", "-filter_complex", "concat=n=3:v=0:a=1,volume=0.15", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
        elif fx_type == "classic":
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=440:d=0.2", "-f", "lavfi", "-i", "sine=f=659:d=0.3", "-filter_complex", "concat=n=2:v=0:a=1,volume=0.15,afade=t=out:st=0.3:d=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
            subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", "sine=f=659:d=0.2", "-f", "lavfi", "-i", "sine=f=440:d=0.3", "-filter_complex", "concat=n=2:v=0:a=1,volume=0.15,afade=t=out:st=0.3:d=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
    
    if not os.path.exists(f_in): return None, None
    return f_in, f_out

def merge_audio(files, output):
    if not files: return
    list_p = output + "_list.txt"
    with open(list_p, "w", encoding="utf-8") as f:
        for p in files:
            safe_path = os.path.abspath(p).replace('\\', '/')
            f.write(f"file '{safe_path}'\n")
            
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    subprocess.run([ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", list_p, "-c:a", "libmp3lame", "-q:a", "4", "-ac", "2", "-ar", "24000", output], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=0x08000000)
    if os.path.exists(list_p): 
        try: os.unlink(list_p)
        except: pass

def create_video_for_chapter(audio_path, image_path, encoder):
    video_path = audio_path.replace(".mp3", ".mp4")
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    vf = "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2,format=yuv420p"
    cmd = [ffmpeg, "-y", "-loop", "1", "-i", image_path, "-i", audio_path, "-c:a", "copy", "-shortest", "-vf", vf]
    if encoder == "nvidia": cmd += ["-c:v", "h264_nvenc", "-b:v", "1500k"]
    elif encoder == "amd": cmd += ["-c:v", "h264_amf", "-b:v", "1500k"]
    else: cmd += ["-c:v", "libx264", "-preset", "ultrafast"]
    cmd.append(video_path); subprocess.run(cmd, stdout=subprocess.DEVNULL, creationflags=0x08000000)

# --- 5. EEL API ---
@eel.expose
def get_voices(): 
    try: return [{"name": v["ShortName"], "locale": v["Locale"]} for v in asyncio.run(edge_tts.list_voices())]
    except: return []

@eel.expose
def get_offline_voices():
    lv = [{"name": f, "locale": "Piper"} for f in os.listdir(root_dir) if f.endswith('.onnx')]
    try:
        import pythoncom; pythoncom.CoInitialize()
        vs = win32com.client.Dispatch("SAPI.SpVoice").GetVoices()
        for i in range(vs.Count): lv.append({"name": vs.Item(i).GetDescription(), "locale": "SAPI5"})
    except: pass
    return lv

@eel.expose
def get_progress():
    s = state.copy(); s["pct"] = int((s["done"] / max(s["total"], 1)) * 100)
    return s

@eel.expose
def stop_process(): stop_ev.set(); state["status"] = "stopped"

@eel.expose
def test_synthesis(text, voice, rate, vol, pitch, mode="online"):
    out = os.path.join(audio_cache, "test_check.mp3")
    async def run_test():
        sem = asyncio.Semaphore(1)
        proc = text_processor.process(text)
        if mode == 'online': await synth_online(proc, voice, rate, vol, pitch, out, sem)
        else: await synth_offline(proc, voice, rate, vol, out, sem)
    asyncio.run(run_test())
    return "/audio_cache/test_check.mp3"

@eel.expose
def get_book_info(filename, b64):
    if "," in b64: b64 = b64.split(",")[1]
    data = base64.b64decode(b64); ext = os.path.splitext(filename.lower())[1]; count = 0
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp: tmp.write(data); t_path = tmp.name
    try:
        if ext == ".epub":
            book = epub.read_epub(t_path); count = len([i for i in book.get_items_of_type(ebooklib.ITEM_DOCUMENT) if len(i.get_content()) > 200])
        elif ext == ".pdf": doc = fitz.open(t_path); count = len(doc); doc.close()
        else: count = 1
    except: count = 1
    finally:
        if os.path.exists(t_path): os.unlink(t_path)
    return {"count": count}

@eel.expose
def synthesize_epub(filename, b64, voice, v_diag, rate, vol, pitch, workers, book_name, mode="online", make_video=False, cover_b64=None, encoder="cpu", bgm_b64=None, bgm_vol="10", fx_type="none"):
    stop_ev.clear(); num_workers = int(workers); b64_data = b64.split(",")[1] if "," in b64 else b64
    data = base64.b64decode(b64_data); ext = os.path.splitext(filename.lower())[1]
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp: tmp.write(data); t_path = tmp.name
    chs = []
    try:
        if ext == ".epub":
            book = epub.read_epub(t_path)
            for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
                if any(x in item.get_name().lower() for x in ["nav", "toc", "cover"]): continue
                soup = BeautifulSoup(item.get_content(), "html.parser"); text = soup.get_text().strip()
                if len(text) > 200: chs.append({"title": f"Глава {len(chs)+1}", "text": text})
        elif ext == ".pdf":
            doc = fitz.open(t_path)
            for i, p in enumerate(doc):
                txt = p.get_text().strip()
                if len(txt) > 50: chs.append({"title": f"Стр {i+1}", "text": txt})
            doc.close()
        else: chs = [{"title": "Книга", "text": data.decode('utf-8', errors='ignore')}]
    except: pass
    finally:
        if os.path.exists(t_path): os.unlink(t_path)

    def task():
        cover_path = bgm_path = None
        try:
            safe_folder = re.sub(r'[\\/*?:"<>|]', "", book_name); f_dir = os.path.join(output_dir, safe_folder); os.makedirs(f_dir, exist_ok=True)
            if make_video and cover_b64:
                with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as c_tmp: c_tmp.write(base64.b64decode(cover_b64.split(",")[1])); cover_path = c_tmp.name
            if bgm_b64:
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as b_tmp: b_tmp.write(base64.b64decode(bgm_b64.split(",")[1])); bgm_path = b_tmp.name
            
            q_in = q_out = None
            if fx_type != "none": q_in, q_out = generate_fx(fx_type)

            all_data = []
            total = 0
            for ch in chs:
                clean = text_processor.process(ch["text"]); chunks = [s.strip() for s in re.split(r'(?<=[.!?…])\s+', clean) if s.strip()]
                if chunks: total += len(chunks); all_data.append({"title": ch["title"], "chunks": chunks})
            
            state.update({"total": total, "done": 0, "status": "running"})
            ffmpeg_exe = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
            
            for i, ch in enumerate(all_data):
                if stop_ev.is_set(): break
                state["chapter"] = ch["title"]; t_dir = tempfile.mkdtemp(); paths = []
                try:
                    async def run():
                        sem = asyncio.Semaphore(num_workers); tasks = []
                        for idx, txt in enumerate(ch["chunks"]):
                            
                            # ЧЕРЕДОВАНИЕ ЭФФЕКТОВ: Четные предложения получают In, нечетные - Out
                            if q_in and q_out:
                                paths.append(q_in if idx % 2 == 0 else q_out)
                                
                            p = os.path.join(t_dir, f"{idx:04d}.mp3"); paths.append(p)
                            v = v_diag if (v_diag and txt.startswith(('-', '—'))) else voice
                            if mode == "online": tasks.append(synth_online(txt, v, rate, vol, pitch, p, sem))
                            else: tasks.append(synth_offline(txt, v, rate, vol, p, sem))
                            
                        await asyncio.gather(*tasks)
                    asyncio.run(run())
                    
                    out = os.path.join(f_dir, f"{i+1:03d}.mp3"); merge_audio(paths, out)
                    
                    if bgm_path and os.path.exists(out):
                        mixed = out.replace(".mp3", "_m.mp3"); v_f = int(bgm_vol)/100.0
                        cmd = [ffmpeg_exe, "-y", "-i", out, "-stream_loop", "-1", "-i", bgm_path, "-filter_complex", f"[1:a]volume={v_f}[bg];[0:a][bg]amix=inputs=2:duration=first:dropout_transition=2,volume=1.8", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "2", mixed]
                        subprocess.run(cmd, stdout=subprocess.DEVNULL, creationflags=0x08000000)
                        if os.path.exists(mixed): shutil.move(mixed, out)
                    
                    if make_video and cover_path and os.path.exists(out): create_video_for_chapter(out, cover_path, encoder)
                except Exception as e:
                    print(f"Ошибка в главе {i}: {e}")
                finally:
                    try: shutil.rmtree(t_dir, ignore_errors=True)
                    except: pass
                    
            state["status"] = "done"
        except Exception as global_e: 
            state["status"] = "error"
            print(f"Global task error: {global_e}")
        finally:
            if cover_path and os.path.exists(cover_path): 
                try: os.unlink(cover_path)
                except: pass
            if bgm_path and os.path.exists(bgm_path): 
                try: os.unlink(bgm_path)
                except: pass
    threading.Thread(target=task, daemon=True).start()

@eel.expose
def process_dictionary(b64, filename, voice1, voice2, book_name, delay=1000, mode="online", vol="+0%"):
    stop_ev.clear(); b64_data = b64.split(",")[1] if "," in b64 else b64
    data = base64.b64decode(b64_data); ext = os.path.splitext(filename.lower())[1]
    try:
        df = pd.read_csv(io.BytesIO(data), sep=None, engine='python') if ext == '.csv' else pd.read_excel(io.BytesIO(data))
        pairs = df.iloc[:, :2].values.tolist(); f_dir = os.path.join(output_dir, f"Словарь_{book_name}"); os.makedirs(f_dir, exist_ok=True)
        state.update({"total": len(pairs), "done": 0, "status": "running"})
        def dict_task():
            for i, (w1, w2) in enumerate(pairs):
                if stop_ev.is_set(): break
                w1, w2 = str(w1), str(w2); state["chapter"] = f"Слово: {w1}"; t_dir = tempfile.mkdtemp()
                p1, p2 = os.path.join(t_dir, "1.mp3"), os.path.join(t_dir, "2.mp3")
                async def run_pair():
                    sem = asyncio.Semaphore(2)
                    await asyncio.gather(
                        synth_online(text_processor.process(w1), voice1, "+0%", vol, "+0%", p1, sem) if mode=="online" else synth_offline(text_processor.process(w1), voice1, "+0%", vol, p1, sem),
                        synth_online(text_processor.process(w2), voice2, "+0%", vol, "+0%", p2, sem) if mode=="online" else synth_offline(text_processor.process(w2), voice2, "+0%", vol, p2, sem)
                    )
                asyncio.run(run_pair())
                out_path = os.path.join(f_dir, f"{i+1:03d}_{re.sub(r'[^a-zA-Zа-яА-ЯёЁ0-9]', '_', w1)}.mp3")
                ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
                subprocess.run([ffmpeg, "-y", "-i", p1, "-i", p2, "-filter_complex", f"[1:a]adelay={delay}|{delay}[a2];[0:a][a2]concat=n=2:v=0:a=1", "-ac", "2", out_path], stdout=subprocess.DEVNULL, creationflags=0x08000000)
                try: shutil.rmtree(t_dir, ignore_errors=True)
                except: pass
                state["done"] = i + 1
            state["status"] = "done"
        threading.Thread(target=dict_task, daemon=True).start()
        return True
    except: return False

# --- 6. ЗАПУСК ---
def close_callback(route, websockets):
    if not websockets:
        print("Окно закрыто. Завершение процесса...")
        os._exit(0)

if __name__ == '__main__':
    eel.init(web_dir)
    ep = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    m = 'edge' if os.path.exists(ep) else 'default'
    try: eel.start('index.html', mode=m, size=(1100, 850), close_callback=close_callback)
    except: eel.start('index.html', mode=None, host='127.0.0.1', port=8888, close_callback=close_callback)