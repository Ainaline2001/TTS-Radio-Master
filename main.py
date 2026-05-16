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
        text = re.sub(r'\[\d+\]|\{\d+\}|\*+', '', text)
        text = re.sub(r'\s+', ' ', text)
        for d in [self.yo_dict, self.phonetic_dict]:
            for k, v in d.items():
                text = re.compile(re.escape(k), re.IGNORECASE).sub(v, text)
        return text

text_processor = TextProcessor()

SKIP_CHAPTERS = [
    'nav', 'toc', 'cover', 'примечани', 'приложени', 'сноски', 
    'библиографи', 'глоссарий', 'notes', 'appendix', 'содержание', 
    'об авторе', 'от автора', 'литература', 'указатель', 'reklama', 'реклама'
]

# --- 3. СИНТЕЗ ---
state = {"total": 0, "done": 0, "status": "idle", "chapter": ""}
stop_ev = threading.Event()
TIMEOUT_SECONDS = 60 
sapi_lock = threading.Lock() 

async def run_subprocess_async(cmd, input_data=None):
    proc = await asyncio.create_subprocess_exec(
        *cmd,
        stdin=asyncio.subprocess.PIPE if input_data else None,
        stdout=asyncio.subprocess.DEVNULL,
        stderr=asyncio.subprocess.DEVNULL,
        creationflags=0x08000000
    )
    try:
        if input_data: await proc.communicate(input=input_data)
        else: await proc.wait()
    except asyncio.CancelledError:
        try: proc.kill()
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
        except Exception as e: print(f"Online error: {e}")
        finally:
            await ensure_valid_mp3(out); state["done"] += 1

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
                with sapi_lock:
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
                    except: pass
                    finally: pythoncom.CoUninitialize()
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
        except Exception as e: print(f"Offline error: {e}")
        finally:
            await ensure_valid_mp3(out); state["done"] += 1

# --- 4. МЕДИА И СКЛЕЙКА (С ID3 ТЕГАМИ И ВИДЕО) ---
def merge_audio(files, output, title="Audio", album="Audiobook", track="1/1"):
    if not files: return
    list_p = output + "_list.txt"
    with open(list_p, "w", encoding="utf-8") as f:
        for p in files:
            safe_path = os.path.abspath(p).replace('\\', '/')
            f.write(f"file '{safe_path}'\n")
    
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    cmd = [
        ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", list_p, 
        "-c:a", "libmp3lame", "-q:a", "4", "-ac", "2", "-ar", "24000",
        "-metadata", f"title={title}",
        "-metadata", f"album={album}",
        "-metadata", f"track={track}",
        output
    ]
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=0x08000000)
    if os.path.exists(list_p): os.unlink(list_p)

def create_video_for_chapter(audio_path, image_path, encoder, is_shorts=False, shorts_duration=50):
    if not image_path or not os.path.exists(image_path):
        return
    
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    
    if is_shorts:
        video_path = audio_path.replace(".mp3", "_shorts_part%03d.mp4")
        filter_complex = (
            "color=c=black:s=1080x1920[bg];"
            "[0:v]scale=600:600,format=rgba,"
            "geq=r='r(X,Y)':g='g(X,Y)':b='b(X,Y)':a='if(lte(hypot(X-W/2,Y-H/2),W/2),255,0)',"
            "rotate=2*PI*t/12:c=black@0[disc];"
            "[1:a]showfreqs=s=900x400:mode=bar:ascale=log:colors=cyan[wave];"
            "[bg][disc]overlay=(W-w)/2:400[tmp];"
            "[tmp][wave]overlay=(W-w)/2:1200,format=yuv420p[outv]"
        )
    else:
        video_path = audio_path.replace(".mp3", ".mp4")
        filter_complex = (
            "color=c=black:s=1920x1080[bg];"
            "[0:v]scale=600:600,format=rgba,"
            "geq=r='r(X,Y)':g='g(X,Y)':b='b(X,Y)':a='if(lte(hypot(X-W/2,Y-H/2),W/2),255,0)',"
            "rotate=2*PI*t/12:c=black@0[disc];"
            "[1:a]showfreqs=s=1000x400:mode=bar:ascale=log:colors=cyan[wave];"
            "[bg][disc]overlay=150:H/2-h/2[tmp];"
            "[tmp][wave]overlay=W-w-150:H/2-h/2,format=yuv420p[outv]"
        )

    def run_render(current_encoder):
        cmd = [
            ffmpeg, "-y", "-loop", "1",
            "-i", image_path, "-i", audio_path,
            "-filter_complex", filter_complex,
            "-map", "[outv]", "-map", "1:a", "-c:a", "copy", "-shortest"
        ]
        if current_encoder == "amd": cmd += ["-c:v", "h264_amf", "-b:v", "1500k"]
        elif current_encoder == "nvidia": cmd += ["-c:v", "h264_nvenc", "-b:v", "1500k"]
        else: cmd += ["-c:v", "libx264", "-crf", "24", "-preset", "veryfast"]
            
        if is_shorts:
            cmd += ["-f", "segment", "-segment_time", str(shorts_duration), "-reset_timestamps", "1"]
            
        cmd.append(video_path)
        return subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, creationflags=0x08000000)

    print(f"Начинаем рендер видео: {'Shorts нарезка' if is_shorts else 'Полное видео'}")
    result = run_render(encoder)
    if result.returncode != 0: run_render("cpu")

def generate_dynamic_fx(fx_type, chapter_idx, output_dir):
    shift = (chapter_idx % 6) * 30 
    f_in = os.path.join(output_dir, f"{fx_type}_in.mp3")
    f_out = os.path.join(output_dir, f"{fx_type}_out.mp3")
    ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
    
    vol = 0.4
    if fx_type == "apollo":
        f1, f2 = 2525 + shift, 2475 + shift
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f1}:d=0.2", "-filter:a", f"volume={vol}", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f2}:d=0.2", "-filter:a", f"volume={vol}", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
    elif fx_type == "zvei":
        f1, f2, f3 = 1060 + shift, 1270 + shift, 1530 + shift
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f1}:d=0.07", "-f", "lavfi", "-i", f"sine=f={f2}:d=0.07", "-f", "lavfi", "-i", f"sine=f={f3}:d=0.07", "-filter_complex", f"concat=n=3:v=0:a=1,volume={vol-0.1}", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f3}:d=0.07", "-f", "lavfi", "-i", f"sine=f={f2}:d=0.07", "-f", "lavfi", "-i", f"sine=f={f1}:d=0.07", "-filter_complex", f"concat=n=3:v=0:a=1,volume={vol-0.1}", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
    elif fx_type == "classic":
        f1, f2 = 440 + shift, 659 + shift
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f1}:d=0.2", "-f", "lavfi", "-i", f"sine=f={f2}:d=0.3", "-filter_complex", f"concat=n=2:v=0:a=1,volume={vol-0.1},afade=t=out:st=0.3:d=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_in], stdout=subprocess.DEVNULL, creationflags=0x08000000)
        subprocess.run([ffmpeg, "-y", "-f", "lavfi", "-i", f"sine=f={f2}:d=0.2", "-f", "lavfi", "-i", f"sine=f={f1}:d=0.3", "-filter_complex", f"concat=n=2:v=0:a=1,volume={vol-0.1},afade=t=out:st=0.3:d=0.2", "-c:a", "libmp3lame", "-ar", "24000", "-ac", "1", f_out], stdout=subprocess.DEVNULL, creationflags=0x08000000)
    
    return f_in, f_out

# --- 5. EEL API (ПАРСИНГ, КНИГА, СЛОВАРЬ) ---
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
def get_book_info(filename, b64):
    if "," in b64: b64 = b64.split(",")[1]
    data = base64.b64decode(b64); ext = os.path.splitext(filename.lower())[1]; count = 0
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp: tmp.write(data); t_path = tmp.name
    try:
        if ext == ".epub":
            book = epub.read_epub(t_path)
            for item_id, linear in book.spine:
                item = book.get_item_with_id(item_id)
                if item and item.get_type() == ebooklib.ITEM_DOCUMENT:
                    name = item.get_name().lower()
                    if any(x in name for x in SKIP_CHAPTERS): continue
                    soup = BeautifulSoup(item.get_content(), "html.parser")
                    if len(soup.get_text().strip()) > 200: count += 1
        elif ext == ".pdf": 
            doc = fitz.open(t_path); count = len(doc); doc.close()
        else: count = 1
    except: count = 1
    finally:
        if os.path.exists(t_path): os.unlink(t_path)
    return {"count": count}

@eel.expose
def synthesize_epub(filename, b64, voice, v_diag, rate, vol, pitch, workers, book_name, mode="online", make_video=False, cover_b64=None, encoder="cpu", bgm_b64=None, bgm_vol="10", fx_type="none", is_shorts=False, shorts_duration=50, start_ch="1", end_ch=""):
    stop_ev.clear(); num_workers = int(workers)
    data = base64.b64decode(b64.split(",")[1] if "," in b64 else b64)
    ext = os.path.splitext(filename.lower())[1]
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp: tmp.write(data); t_path = tmp.name
    chs = []
    try:
        if ext == ".epub":
            book = epub.read_epub(t_path)
            for item_id, linear in book.spine:
                item = book.get_item_with_id(item_id)
                if item and item.get_type() == ebooklib.ITEM_DOCUMENT:
                    name = item.get_name().lower()
                    if any(x in name for x in SKIP_CHAPTERS): continue
                    soup = BeautifulSoup(item.get_content(), "html.parser")
                    h_tag = soup.find(['h1', 'h2', 'h3', 'title'])
                    title_text = h_tag.get_text().strip() if h_tag else ""
                    if any(x in title_text.lower() for x in SKIP_CHAPTERS): continue
                    text = soup.get_text().strip()
                    if len(text) > 200:
                        chs.append({"title": title_text[:50] or f"Глава {len(chs)+1}", "text": text})
        elif ext == ".pdf":
            doc = fitz.open(t_path)
            for i, p in enumerate(doc): chs.append({"title": f"Стр {i+1}", "text": p.get_text()})
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
                with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as c_tmp:
                    c_tmp.write(base64.b64decode(cover_b64.split(",")[1]))
                    cover_path = c_tmp.name

            if bgm_b64:
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as b_tmp:
                    b_tmp.write(base64.b64decode(bgm_b64.split(",")[1]))
                    bgm_path = b_tmp.name
            
            all_data = []
            for ch in chs:
                proc = text_processor.process(ch["text"])
                chunks = [s.strip() for s in re.split(r'(?<=[.!?…])\s+', proc) if s.strip()]
                if chunks: all_data.append({"title": ch["title"], "chunks": chunks})
            
            total_chapters = len(all_data)
            s_idx = max(0, int(start_ch) - 1) if start_ch else 0
            e_idx = int(end_ch) if end_ch else total_chapters
            selected_chapters = all_data[s_idx:e_idx]
            
            total_chunks = sum(len(c["chunks"]) for c in selected_chapters)
            state.update({"total": total_chunks, "done": 0, "status": "running"})
            ffmpeg_exe = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
            
            for i, ch in enumerate(selected_chapters):
                if stop_ev.is_set(): break
                real_idx = s_idx + i 
                state["chapter"] = ch["title"]; t_dir = tempfile.mkdtemp(); paths = []
                try:
                    q_in = q_out = None
                    if fx_type != "none": q_in, q_out = generate_dynamic_fx(fx_type, real_idx, t_dir)

                    async def run():
                        sem = asyncio.Semaphore(num_workers); tasks = []
                        for idx, txt in enumerate(ch["chunks"]):
                            if q_in and q_out: paths.append(q_in if idx % 2 == 0 else q_out)
                            p = os.path.join(t_dir, f"{idx:04d}.mp3"); paths.append(p)
                            v = v_diag if (v_diag and txt.startswith(('-', '—'))) else voice
                            if mode == "online": tasks.append(synth_online(txt, v, rate, vol, pitch, p, sem))
                            else: tasks.append(synth_offline(txt, v, rate, vol, p, sem))
                        await asyncio.gather(*tasks)
                    asyncio.run(run())
                    
                    out = os.path.join(f_dir, f"{real_idx+1:03d}.mp3")
                    clean_title = ch["title"].replace('"', '').replace("'", "")
                    merge_audio(paths, out, title=clean_title, album=book_name, track=f"{real_idx+1}/{total_chapters}")
                    
                    if make_video and cover_path:
                        create_video_for_chapter(out, cover_path, encoder, is_shorts, shorts_duration)
                finally:
                    try: shutil.rmtree(t_dir, ignore_errors=True)
                    except: pass
            state["status"] = "done"
        except Exception as e: print(f"Global error: {e}"); state["status"] = "error"
        finally:
            if cover_path: 
                try: os.unlink(cover_path)
                except: pass
            if bgm_path: 
                try: os.unlink(bgm_path)
                except: pass
    threading.Thread(target=task, daemon=True).start()

@eel.expose
def process_dictionary(b64, filename, voice1, voice2, book_name, delay=1000, mode="online", vol="+0%", merge_all=False, fx_type="none"):
    stop_ev.clear(); data = base64.b64decode(b64.split(",")[1] if "," in b64 else b64)
    ext = os.path.splitext(filename.lower())[1]
    try:
        df = pd.read_csv(io.BytesIO(data), sep=None, engine='python', header=None) if ext == '.csv' else pd.read_excel(io.BytesIO(data), header=None)
        pairs = df.iloc[:, :2].values.tolist(); f_dir = os.path.join(output_dir, f"Словарь_{book_name}"); os.makedirs(f_dir, exist_ok=True)
        state.update({"total": len(pairs), "done": 0, "status": "running"})
        
        def dict_task():
            generated_files = [] 
            for i, (w1, w2) in enumerate(pairs):
                if stop_ev.is_set(): break
                state["chapter"] = f"Слово: {str(w1)[:20]}"; t_dir = tempfile.mkdtemp()
                p1, p2 = os.path.join(t_dir, "1.mp3"), os.path.join(t_dir, "2.mp3")
                
                q_in = q_out = None
                if fx_type != "none":
                    q_in, q_out = generate_dynamic_fx(fx_type, i, t_dir)
                
                async def run_pair():
                    sem = asyncio.Semaphore(2)
                    v_v = vol if mode=="online" else vol
                    if mode=="online":
                        await asyncio.gather(
                            synth_online(str(w1), voice1, "+0%", v_v, "+0Hz", p1, sem),
                            synth_online(str(w2), voice2, "+0%", v_v, "+0Hz", p2, sem)
                        )
                    else:
                        await asyncio.gather(
                            synth_offline(str(w1), voice1, "+0%", v_v, p1, sem),
                            synth_offline(str(w2), voice2, "+0%", v_v, p2, sem)
                        )
                asyncio.run(run_pair())
                
                safe_w1 = re.sub(r'\W+', '_', str(w1))[:30]
                out = os.path.join(f_dir, f"{i+1:03d}_{safe_w1}.mp3")
                ffmpeg = os.path.join(root_dir, "ffmpeg.exe") if os.path.exists(os.path.join(root_dir, "ffmpeg.exe")) else "ffmpeg"
                
                delay_sec = float(delay) / 1000.0
                
                if fx_type != "none" and q_out and os.path.exists(q_out):
                    cmd = [ffmpeg, "-y", "-i", p1, "-i", p2, "-i", q_out, "-filter_complex", f"[0:a]apad=pad_dur={delay_sec}[a1];[a1][1:a][2:a]concat=n=3:v=0:a=1", "-ac", "2", out]
                else:
                    cmd = [ffmpeg, "-y", "-i", p1, "-i", p2, "-filter_complex", f"[0:a]apad=pad_dur={delay_sec}[a1];[a1][1:a]concat=n=2:v=0:a=1", "-ac", "2", out]
                
                subprocess.run(cmd, stdout=subprocess.DEVNULL, creationflags=0x08000000)
                
                if os.path.exists(out):
                    generated_files.append(out)
                    
                shutil.rmtree(t_dir, ignore_errors=True); state["done"] = i + 1
            
            if merge_all and not stop_ev.is_set() and generated_files:
                state["chapter"] = "Объединение словаря в один файл..."
                full_dict_path = os.path.join(f_dir, f"Словарь_ПОЛНЫЙ_{book_name}.mp3")
                merge_audio(generated_files, full_dict_path, title=f"Словарь {book_name}", album="Dictionary", track="1/1")
                
            state["status"] = "done"
            
        threading.Thread(target=dict_task, daemon=True).start(); return True
    except Exception as e: 
        print(f"Ошибка словаря: {e}")
        return False

def close_callback(route, websockets):
    if not websockets: os._exit(0)

if __name__ == '__main__':
    eel.init(web_dir)
    ep = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    m = 'edge' if os.path.exists(ep) else 'default'
    try: eel.start('index.html', mode=m, size=(1100, 850), close_callback=close_callback)
    except: eel.start('index.html', mode=None, host='127.0.0.1', port=8888, close_callback=close_callback)
