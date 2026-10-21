import streamlit as st
import datetime, os, sys, subprocess, re, json, csv, io, random, time
import sqlite3
from collections import Counter
from urllib.request import Request, urlopen
import xml.etree.ElementTree as ET

# ==================== 0. AUTO-INSTALL PACKAGES ====================
def _auto_pip(pkg, importer):
    try:
        return importer()
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pkg])
        return importer()

fitz = _auto_pip("PyMuPDF", lambda: __import__("fitz"))
gdown = _auto_pip("gdown", lambda: __import__("gdown"))
from google import genai
from google.genai import types
st_quill = _auto_pip("streamlit-quill", lambda: __import__("streamlit_quill"))

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "python-pptx"])
    from pptx import Presentation
    from pptx.util import Inches, Pt

# ==================== 1. CẤU HÌNH TRANG, GOOGLE DRIVE & BẢO MẬT ====================
st.set_page_config(
    page_title="Chấn Thương Chỉnh Hình",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="theme-color" content="#2B7A78">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
""", unsafe_allow_html=True)

# ---> DÁN LINK THƯ MỤC GOOGLE DRIVE CỦA BẠN VÀO ĐÂY <---
GDRIVE_FOLDER_URL = "https://drive.google.com/drive/u/0/folders/1tVqwvCq0u7aUrnLQx3XbWG4ezmy0Oto-Y"

LIBRARY_FILE = "dudu_library.json"
MISTAKES_FILE = "dudu_mistakes.json"
SETTINGS_FILE = "dudu_settings.json"
PROGRESS_FILE = "dudu_progress.json"
PDF_DIR = "kho_sach"
DB_FILE = os.path.join(PDF_DIR, "ortho_fts5.db")
SECRETS_DIR = ".streamlit"
SECRETS_PATH = os.path.join(SECRETS_DIR, "secrets.toml")
os.makedirs(PDF_DIR, exist_ok=True)

def sync_books_from_gdrive(force=False):
    """Tự động tải toàn bộ PDF từ thư mục Google Drive nếu kho_sach đang trống"""
    existing_pdfs = [f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")]
    if existing_pdfs and not force:
        return
    
    drive_url = st.secrets.get("GDRIVE_FOLDER_URL", GDRIVE_FOLDER_URL) if hasattr(st, "secrets") else GDRIVE_FOLDER_URL
    if not drive_url or "drive.google.com" not in drive_url:
        return
        
    with st.spinner("Đang tải đầy đủ kho sách gốc từ Google Drive (chỉ mất vài chục giây ở lần khởi động đầu tiên)..."):
        try:
            gdown.download_folder(url=drive_url, output=PDF_DIR, quiet=True, use_cookies=False)
        except Exception as e:
            st.warning(f"Chưa thể tải tự động từ Google Drive. Hãy kiểm tra quyền chia sẻ 'Bất kỳ ai có liên kết': {e}")

sync_books_from_gdrive(force=False)

def load_json_file(filepath, default_val):
    if not os.path.exists(filepath): return default_val
    try:
        with open(filepath, "r", encoding="utf-8") as f: return json.load(f)
    except Exception:
        bak = filepath + ".bak"
        if os.path.exists(bak):
            try:
                with open(bak, "r", encoding="utf-8") as f: return json.load(f)
            except Exception: pass
        return default_val

def save_json_file(filepath, data):
    tmp = filepath + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    if os.path.exists(filepath):
        try: os.replace(filepath, filepath + ".bak")
        except Exception: pass
    os.replace(tmp, filepath)

def repair_items_id(items_list, filepath):
    seen = set()
    changed = False
    for idx, item in enumerate(items_list):
        cid = item.get("id")
        if not cid or str(cid) == "None" or str(cid) in seen:
            item["id"] = f"item_{idx}_{datetime.datetime.now().strftime('%H%M%S')}_{random.randint(1000,9999)}"
            seen.add(item["id"])
            changed = True
        else: seen.add(str(cid))
    if changed: save_json_file(filepath, items_list)
    return items_list

def get_active_api_key():
    saved_settings = load_json_file(SETTINGS_FILE, {})
    k = saved_settings.get("api_key", "").strip().strip('"').strip("'")
    if k and k != "AQ.xxx": return k
    try:
        sk = st.secrets.get("GEMINI_API_KEY", "").strip().strip('"').strip("'")
        if sk and sk != "AQ.xxx": return sk
    except Exception: pass
    return ""

def save_api_key_everywhere(new_key: str):
    clean_k = new_key.strip().strip('"').strip("'")
    st.session_state.settings["api_key"] = clean_k
    save_json_file(SETTINGS_FILE, st.session_state.settings)
    try:
        os.makedirs(SECRETS_DIR, exist_ok=True)
        with open(SECRETS_PATH, "w", encoding="utf-8") as f: f.write(f'GEMINI_API_KEY = "{clean_k}"\n')
    except Exception: pass

# ==================== 2. CƠ SỞ DỮ LIỆU SQLITE FTS5 ====================
def get_db():
    conn = sqlite3.connect(DB_FILE, timeout=30, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db_tables():
    with get_db() as conn:
        conn.execute("CREATE VIRTUAL TABLE IF NOT EXISTS pages USING fts5(book, page UNINDEXED, text, tokenize='unicode61 remove_diacritics 2')")
        conn.execute("CREATE TABLE IF NOT EXISTS books (name TEXT PRIMARY KEY, pages INTEGER, indexed_at TEXT)")
        conn.commit()

init_db_tables()

def is_book_indexed(fname):
    with get_db() as conn:
        return bool(conn.execute("SELECT name FROM books WHERE name=?", (fname,)).fetchone())

def index_single_book_fast(fname):
    if is_book_indexed(fname): return False
    fpath = os.path.join(PDF_DIR, fname)
    if not os.path.exists(fpath): return False
    try:
        doc = fitz.open(fpath)
        rows = [(fname, idx + 1, doc.load_page(idx).get_text("text")) for idx in range(len(doc))]
        doc.close()
        with get_db() as conn:
            conn.executemany("INSERT INTO pages(book, page, text) VALUES (?,?,?)", rows)
            conn.execute("INSERT OR REPLACE INTO books(name, pages, indexed_at) VALUES (?,?,?)", (fname, len(rows), datetime.datetime.now().isoformat()))
            conn.commit()
        return True
    except Exception: return False

def get_page_data(fname, page_num):
    with get_db() as conn:
        row = conn.execute("SELECT text FROM pages WHERE book=? AND page=?", (fname, page_num)).fetchone()
    if not row: return None
    return {"text": row[0], "fpath": os.path.join(PDF_DIR, fname), "page_index": page_num - 1}

def get_total_indexed_pages(books):
    if not books: return 0
    ph = ",".join("?" * len(books))
    with get_db() as conn:
        row = conn.execute(f"SELECT SUM(pages) FROM books WHERE name IN ({ph})", books).fetchone()
    return (row[0] or 0) if row else 0

def smart_route_pages(query, books, max_pages=15):
    if not books: return "", [], []
    explicit = [int(p) for p in re.findall(r'(?:trang|tr\.?)\s*(\d+)', query.lower())]
    explicit_hits = []
    for bk in books:
        for p in explicit:
            d = get_page_data(bk, p)
            if d and d["text"].strip(): explicit_hits.append((bk, p, d["text"]))

    clean_q = re.sub(r'[^\w\s]', ' ', query.lower())
    stop = {"là", "gì", "và", "của", "các", "trong", "như", "thế", "nào", "cho", "tôi", "hỏi", "về", "bệnh", "cách", "tại", "sao", "vì", "khi"}
    terms = [t for t in clean_q.split() if len(t) > 1 and t not in stop and not t.isdigit()]
    fts_hits = []

    with get_db() as conn:
        if terms:
            fts_q = " OR ".join(f'"{t}"' for t in terms[:8])
            placeholders = ",".join("?" * len(books))
            try:
                cur = conn.execute(f"SELECT book, page, text, bm25(pages) AS rank FROM pages WHERE pages MATCH ? AND book IN ({placeholders}) ORDER BY rank LIMIT ?", (fts_q, *books, max_pages))
                fts_hits = [(r[0], r[1], r[2]) for r in cur.fetchall()]
            except Exception: fts_hits = []
        if not fts_hits and terms:
            placeholders = ",".join("?" * len(books))
            for t in terms[:3]:
                cur = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({placeholders}) AND text LIKE ? LIMIT ?", (*books, f"%{t}%", max_pages // 3))
                fts_hits.extend([(r[0], r[1], r[2]) for r in cur.fetchall()])

    seen = set()
    combined = []
    for bk, p, t in explicit_hits:
        if (bk, p) not in seen: combined.append((bk, p, t)); seen.add((bk, p))
    for bk, p, t in fts_hits:
        if (bk, p) not in seen and len(combined) < max_pages: combined.append((bk, p, t)); seen.add((bk, p))

    if not combined: return "", [], []
    routed = "=== MỤC LỤC ĐỊNH HƯỚNG ===\n"
    for bk, p, t in combined: routed += f"\n\n=== SÁCH: {bk} | TRANG {p} ===\n{t}"
    return routed, [(bk, p) for bk, p, _ in combined], [(bk, p) for bk, p, _ in combined][:2]

@st.cache_data(max_entries=120, show_spinner=False)
def get_page_image(fpath, page_index, zoom=1.4):
    doc = fitz.open(fpath)
    img_bytes = doc.load_page(page_index).get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")
    doc.close()
    return img_bytes

# ==================== 3. KHỞI TẠO STATE & HÀM BỔ TRỢ ====================
all_pdfs = sorted([f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")])
_defaults = {
    "library": lambda: repair_items_id(load_json_file(LIBRARY_FILE, []), LIBRARY_FILE),
    "mistakes": lambda: repair_items_id(load_json_file(MISTAKES_FILE, []), MISTAKES_FILE),
    "settings": lambda: load_json_file(SETTINGS_FILE, {"dark_mode": False, "api_key": ""}),
    "progress": lambda: load_json_file(PROGRESS_FILE, {"activity": {}, "quizzes": 0, "correct": 0}),
    "er_chat": list, "er_case_active": lambda: False, "qa_session": lambda: None,
    "news_headlines": list, "news_session": lambda: None, "anki_session": lambda: None,
    "xray_session": lambda: None, "interactive_quiz": lambda: None, "slide_session": lambda: None,
    "selected_model": lambda: "gemini-flash-lite-latest", "_books": lambda: list(all_pdfs),
    "current_page": lambda: "Hỏi đáp",
}
for k, factory in _defaults.items():
    if k not in st.session_state: st.session_state[k] = factory()

API_KEY = get_active_api_key()

def log_activity():
    today_str = datetime.date.today().isoformat()
    act = st.session_state.progress.setdefault("activity", {})
    act[today_str] = act.get(today_str, 0) + 1
    save_json_file(PROGRESS_FILE, st.session_state.progress)

def get_due_mistakes():
    today_str = datetime.date.today().isoformat()
    return [m for m in st.session_state.mistakes if m.get("next_review", "2000-01-01") <= today_str]

def save_to_library(item_type, title, content, tags=None):
    item_id = f"id_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{random.randint(1000,9999)}"
    st.session_state.library.append({"id": item_id, "time": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "type": item_type, "title": title, "content": content, "tags": tags or []})
    save_json_file(LIBRARY_FILE, st.session_state.library)
    st.toast("Đã lưu vào Kho bài viết")
    return item_id

def update_library_item(item_id, new_title, new_content, tags=None):
    for item in st.session_state.library:
        if str(item.get("id")) == str(item_id):
            item["title"] = new_title; item["content"] = new_content; item["time"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M") + " (Đã sửa)"
            if tags is not None: item["tags"] = tags
            break
    save_json_file(LIBRARY_FILE, st.session_state.library)

def delete_from_library(item_id):
    st.session_state.library = [i for i in st.session_state.library if str(i.get("id")) != str(item_id)]
    save_json_file(LIBRARY_FILE, st.session_state.library)

def sm2_update(mistake, quality: int):
    ease = mistake.get("ease", 2.5); interval = mistake.get("interval", 1); reps = mistake.get("repetitions", 0)
    if quality < 3: reps = 0; interval = 1
    else:
        interval = 1 if reps == 0 else 6 if reps == 1 else round(interval * ease)
        reps += 1
    ease += (0.1 - (5 - quality) * (0.08 + (5 - quality) * 0.02))
    mistake.update({"ease": max(1.3, ease), "interval": interval, "repetitions": reps, "next_review": (datetime.date.today() + datetime.timedelta(days=interval)).isoformat()})
    return mistake

def add_to_mistake_bank(q_obj, user_choice, topic):
    st.session_state.mistakes.append({"id": f"mis_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{random.randint(1000,9999)}", "time": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "topic": topic, "question": q_obj.get("question", ""), "options": q_obj.get("options", []), "answer": q_obj.get("answer", "A"), "user_wrong_choice": user_choice, "explanation": q_obj.get("explanation", ""), "ease": 2.5, "interval": 1, "repetitions": 0, "next_review": datetime.date.today().isoformat()})
    save_json_file(MISTAKES_FILE, st.session_state.mistakes)

@st.cache_data(show_spinner=False)
def text_to_speech(text: str, lang="vi"):
    try: from gtts import gTTS
    except ImportError: subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "gtts"]); from gtts import gTTS
    buf = io.BytesIO(); gTTS(text=re.sub(r'[#*`>\[\]_]', '', text)[:4000], lang=lang, slow=False).write_to_fp(buf); buf.seek(0)
    return buf.read()

def create_pptx_bytes(topic_title, raw_slide_text):
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[0]); s.shapes.title.text = topic_title.upper(); s.placeholders[1].text = f"Chuyên đề Chấn thương Chỉnh hình\nNgày {datetime.date.today().strftime('%d/%m/%Y')}"
    blocks = re.split(r'(?i)(?:###\s*SLIDE\s*\d+|SLIDE\s*\d+\s*:)', raw_slide_text)[1:] if len(re.split(r'(?i)(?:###\s*SLIDE\s*\d+|SLIDE\s*\d+\s*:)', raw_slide_text)) > 1 else [b for b in raw_slide_text.split("\n---\n") if b.strip()]
    for idx, blk in enumerate(blocks):
        blk = re.sub(r'```dot.*?```', '', blk, flags=re.DOTALL).strip()
        if not blk: continue
        parts = re.split(r'(?i)\*?\*?(?:Ghi chú thuyết trình|Speaker Notes)\*?\*?\s*:', blk)
        lines = [l.strip() for l in parts[0].strip().split('\n') if l.strip()]
        if not lines: continue
        slide = prs.slides.add_slide(prs.slide_layouts[1]); slide.shapes.title.text = re.sub(r'^[\*\#\-\:\s]+|[\*\#\s]+$', '', lines[0])[:120]
        tf = slide.placeholders[1].text_frame; tf.word_wrap = True; tf.clear()
        for b_i, b_l in enumerate(lines[1:9]):
            p = tf.paragraphs[0] if b_i == 0 else tf.add_paragraph()
            p.text = re.sub(r'^[\-\*\•\+]+\s*', '', b_l).replace('**', ''); p.font.size = Pt(20); p.space_after = Pt(10)
        if len(parts) > 1: slide.notes_slide.notes_text_frame.text = parts[1].strip().replace('**', '')
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ==================== 5. AI ENGINE & PROMPT SƯ PHẠM ĐỈNH CAO ====================
@st.cache_data(ttl=3600, show_spinner=False)
def scan_available_models(active_key: str):
    fallback = ["gemini-flash-lite-latest", "gemini-flash-latest", "gemini-pro-latest"]
    if not active_key: return fallback, False
    try:
        models = [m.name.replace("models/", "") for m in genai.Client(api_key=active_key).models.list() if "latest" in m.name.lower() and ("gemini" in m.name.lower() or "gemma" in m.name.lower()) and not any(ex in m.name.lower() for ex in ["2.5", "embedding", "imagen", "veo", "audio", "tts", "aqa", "learnlm", "robotics"])]
        if not models: return fallback, True
        models = sorted(list(set(models)), key=lambda n: (0 if n.lower() == "gemini-flash-lite-latest" else 1 if n.lower() == "gemini-flash-latest" else 2 if "flash" in n.lower() else 3, n))
        if "gemini-flash-lite-latest" not in models: models.insert(0, "gemini-flash-lite-latest")
        return models, True
    except Exception: return fallback, False

def generate_ai_response(prompt_text):
    client = genai.Client(api_key=get_active_api_key())
    return client.models.generate_content(
        model=st.session_state.get("selected_model", "gemini-flash-lite-latest"),
        contents=prompt_text,
        config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)
    )

def stream_ai_response(prompt_text, visual_page_keys=None):
    client = genai.Client(api_key=get_active_api_key())
    contents = []
    if visual_page_keys:
        for (fname, p_num) in visual_page_keys[:2]:
            try: contents.append(types.Part.from_bytes(data=get_page_image(os.path.join(PDF_DIR, fname), p_num - 1, zoom=1.1), mime_type="image/png"))
            except Exception: pass
    contents.append(prompt_text)

    placeholder = st.empty()
    full_text = ""
    try:
        for chunk in client.models.generate_content_stream(model=st.session_state.get("selected_model", "gemini-flash-lite-latest"), contents=contents, config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)):
            if chunk.text:
                full_text += chunk.text
                placeholder.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', full_text), unsafe_allow_html=True)
        placeholder.empty()
        return full_text
    except Exception as e:
        placeholder.empty()
        if "401" in str(e) or "UNAUTHENTICATED" in str(e): st.error("Lỗi xác thực API Key (401). Hãy dán khóa mới vào cột bên trái."); st.stop()
        return client.models.generate_content(model=st.session_state.get("selected_model", "gemini-flash-lite-latest"), contents=contents, config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)).text

PEDAGOGICAL_STYLE_RULES = """
NGUYÊN TẮC TRÌNH BÀY BẮT BUỘC ĐỂ TRÁNH KHỐI CHỮ DÀY ĐẶC:
1. KHÔNG VIẾT ĐOẠN VĂN DÀI HAY GỘP NHIỀU Ý VÀO NHAU. Mọi nội dung lâm sàng, X-quang, MRI, Điều trị phải được phân tách rõ ràng.
2. CẤU TRÚC PHÂN TẦNG VÀ TIÊU ĐỀ BẮT BUỘC PHẢI IN HOA + TÔ ĐẬM:
   - Các tiêu đề lớn phải dùng `1.`, `2.`, `3.` kèm IN HOA VÀ TÔ ĐẬM. (Ví dụ: **1. ĐẠI CƯƠNG VÀ CƠ CHẾ**, **2. ĐIỀU TRỊ BẢO TỒN VÀ PHẪU THUẬT**).
   - Dưới mỗi tiêu đề lớn, mọi phương pháp hay nguyên tắc bắt buộc phải xuống dòng và dùng Gạch đầu dòng (`-`).
   - Mở đầu mỗi gạch đầu dòng, từ khóa phải TÔ ĐẬM. 
   - Ví dụ chuẩn cho phần Điều trị:
     - **Nguyên tắc điều trị tối thượng:** Giải thích cơ chế... [Trang X]
     - **Chỉ định điều trị bảo tồn:** Tại sao lại chọn phương pháp này, cơ chế lành xương/mô là gì... [Trang X]
     - **Chỉ định phẫu thuật:** Lý do phẫu thuật, lựa chọn implant nào... [Trang X]
     - **X-quang:** Những góc đo nào, ý nghĩa cơ sinh học của chúng... [Trang X]
     - **Siêu âm / MRI:** ... [Trang X]
3. Trong mỗi gạch đầu dòng (`-`), diễn giải dài và sâu (5-8 câu), cung cấp đầy đủ các số liệu định lượng (độ, mm, tuần lễ) và cơ chế giải phẫu/sinh lý bệnh.
4. Lập **Bảng (Markdown Table)** khi cần so sánh. Cuối mỗi phần quan trọng đóng khung blockquote (`>`) đúc kết mẹo lâm sàng.
"""

def display_cited_pages(response_text, key_prefix="cite", candidate_pages=None):
    cited = sorted(set(int(p) for p in re.findall(r'[Tt]rang\s+(\d+)', response_text)))
    if not cited: return
    matched = [(bk, p, get_page_data(bk, p)) for bk, p in (candidate_pages or []) if p in cited and get_page_data(bk, p)]
    if not matched:
        for p in cited[:5]:
            for bk in st.session_state.get("_books", []):
                d = get_page_data(bk, p)
                if d: matched.append((bk, p, d)); break
    if not matched: return
    st.markdown("#### Hình ảnh minh họa cắt từ sách gốc:")
    for idx, (bk, p, d) in enumerate(matched[:6]):
        with st.expander(f"Xem gốc: {bk} - Trang {p}"):
            c1, c2 = st.columns([3, 2])
            with c1:
                try: st.image(get_page_image(d["fpath"], d["page_index"]), use_container_width=True)
                except Exception: st.write("Không render được ảnh")
            with c2: st.text_area("Văn bản gốc:", value=d["text"], height=280, disabled=True, key=f"{key_prefix}_{idx}_{p}")

def render_graphviz_if_present(text):
    dots = re.findall(r'```dot\n(.*?)\n```', text, re.DOTALL)
    clean = re.sub(r'```dot\n.*?\n```', '', text, flags=re.DOTALL)
    st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', clean), unsafe_allow_html=True)
    if dots:
        st.markdown("#### Sơ đồ thuật toán:")
        for d in dots:
            try: st.graphviz_chart(d)
            except Exception: st.code(d)

def compile_session_md(sess):
    md = sess["main_answer"]
    if sess.get("discussions"):
        md += "\n\n---\n### THẢO LUẬN BỔ SUNG"
        for it in sess["discussions"]:
            if it["type"] == "user_note": md += f"\n\n> Ý kiến ({it['time']}): {it['user_text']}\nĐối chiếu:\n{it.get('ai_reply', '')}"
            else: md += f"\n\n#### Phụ lục ({it['time']}): {it['user_text']}\n{it['ai_reply']}"
    return md

# ==================== 6. GIAO DIỆN & CSS ====================
dark_mode = st.session_state.settings.get("dark_mode", False)
bg_app, bg_sidebar, card, bg_subcard, border_col, text_main, text_sub, accent_col, input_bg = ("#0E141B", "#151F28", "#1B2733", "#131C26", "#2C3E50", "#ECF2F8", "#9FB3C8", "#38B2AC", "#111921") if dark_mode else ("#FAF7F2", "#F4EBE1", "#FFFFFF", "#F8F4EC", "#E6D5C3", "#2C1E16", "#6E503B", "#2B7A78", "#FFFFFF")

st.markdown(f"""
<style>
    /* Đồng bộ màu theo chế độ Sáng/Tối cho thanh chat input */
    [data-testid="stBottom"], [data-testid="stBottom"] > div {{
        background-color: {bg_app} !important;
    }}
    [data-testid="stChatInput"] {{
        background-color: {input_bg} !important;
        border: 1px solid {border_col} !important;
        border-radius: 12px !important;
    }}
    [data-testid="stChatInput"] textarea {{
        background-color: {input_bg} !important;
        color: {text_main} !important;
    }}
    [data-testid="stChatInput"] textarea::placeholder {{
        color: {text_sub} !important;
    }}
    [data-testid="stChatInput"] button {{
        color: {accent_col} !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==================== 7. SIDEBAR DỌC ====================
with st.sidebar:
    st.markdown(f"<h2 style='text-align:center;margin:0 0 12px 0;font-size:1.35rem;color:{accent_col};'>Chấn Thương Chỉnh Hình</h2>", unsafe_allow_html=True)
    new_dark = st.toggle("Chế độ ban đêm", value=dark_mode)
    if new_dark != dark_mode:
        st.session_state.settings["dark_mode"] = new_dark; save_json_file(SETTINGS_FILE, st.session_state.settings); st.rerun()

    avail_models, key_ok = scan_available_models(API_KEY)
    st.markdown("Mô hình AI:")
    current_m = st.session_state.get("selected_model", "gemini-flash-lite-latest")
    cur_idx = avail_models.index(current_m) if current_m in avail_models else 0
    selected_m = st.selectbox("Mô hình AI", avail_models, index=cur_idx, label_visibility="collapsed")
    if selected_m != st.session_state.selected_model: st.session_state.selected_model = selected_m; st.rerun()

    with st.expander("Cập nhật API Key", expanded=False):
        if not key_ok: st.caption("Khóa API lỗi 401 hoặc chưa nhập. Dán khóa API mới vào dưới đây:")
        new_key_in = st.text_input("GEMINI_API_KEY:", value="", type="password", placeholder="Dán API Key mới...")
        if st.button("Lưu API Key", use_container_width=True) and new_key_in.strip():
            save_api_key_everywhere(new_key_in); scan_available_models.clear(); st.success("Đã lưu!"); st.rerun()

    st.markdown("---")
    st.markdown(f"<p style='text-align:center;font-weight:700;color:{text_sub};margin-bottom:8px;'>CHUYÊN MỤC</p>", unsafe_allow_html=True)

    for label in ["Hỏi đáp", "Bảng tin", "Trực Cấp cứu", "Kho bài viết", "Tạo thẻ Anki", "X-quang", "USMLE & SRS", "Tạo Slide"]:
        is_active = st.session_state.current_page == label
        if st.button(label, key=f"nav_{label}", use_container_width=True, type="primary" if is_active else "secondary"):
            st.session_state.current_page = label; st.rerun()

    st.markdown("---")
    st.markdown(f"<p style='text-align:center;font-weight:700;color:{text_sub};margin-bottom:8px;'>KỆ SÁCH</p>", unsafe_allow_html=True)
    if st.button("Đồng bộ lại từ Google Drive", use_container_width=True):
        sync_books_from_gdrive(force=True)
        st.rerun()
        
    saved_pdfs = sorted([f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")])
    selected_local_books = []
    if saved_pdfs:
        for b_i, name in enumerate(saved_pdfs):
            if st.checkbox(name, value=True, key=f"book_{b_i}"): selected_local_books.append(name)
    else: st.info("Chưa có sách. Hãy dán link Google Drive vào code hoặc bấm Đồng bộ.")
    st.session_state["_books"] = selected_local_books

if st.session_state.get("_books", []):
    unindexed = [b for b in st.session_state["_books"] if not is_book_indexed(b)]
    if unindexed:
        pb = st.progress(0, text="Đang nạp sách mới...")
        for i, b in enumerate(unindexed):
            pb.progress(i / len(unindexed), text=f"Đang nạp: {b}"); index_single_book_fast(b)
        pb.empty()

# ==================== 8. NỘI DUNG TỪNG CHUYÊN MỤC ====================
page = st.session_state.current_page

if page == "Hỏi đáp":
    c_opt1, c_opt2, c_opt3 = st.columns([2.2, 1.8, 1])
    with c_opt1: ans_mode = st.radio("Chế độ:", ["Tự động", "Giải thích", "Bài soạn"], horizontal=True, key="qa_mode")
    with c_opt2: st.write(""); use_vision = st.toggle("Đính kèm ảnh cho AI", value=True, key="qa_vis_toggle")
    with c_opt3:
        st.write("")
        if st.button("Xóa phiên", key="qa_clr", use_container_width=True): st.session_state.qa_session = None; st.rerun()

    query = st.chat_input("Nhập câu hỏi (Nhấn Enter để gửi)...", key="qa_q")
    if query:
        books = st.session_state.get("_books", [])
        if not books: st.warning("Vui lòng chọn ít nhất 1 cuốn sách ở thanh bên trái.")
        else:
            ctx, matched, vis = smart_route_pages(query, books, max_pages=15)
            q_low = query.lower()
            is_specific = any(k in q_low for k in ["tại sao", "vì sao", "như thế nào", "khi nào", "bao lâu", "chỉ định", "biến chứng", "cơ chế", "là gì", "so sánh", "?"]) or len(q_low.split()) >= 6
            use_7 = (ans_mode == "Bài soạn") or (ans_mode == "Tự động" and not is_specific)

            if not use_7:
                prompt = f"""TÀI LIỆU GIÁO TRÌNH:\n{ctx}\n\nCÂU HỎI: "{query}"\n\n{PEDAGOGICAL_STYLE_RULES}
TRÌNH BÀY DẠNG LIST CHO TỪNG Ý LỚN DƯỚI ĐÂY (TIÊU ĐỀ LỚN PHẢI IN HOA VÀ TÔ ĐẬM):
**1. TRẢ LỜI TRỰC DIỆN & BẢN CHẤT CƠ CHẾ:** (Dùng gạch đầu dòng phân tích sâu)
**2. PHÂN TÍCH LÂM SÀNG & CẬN LÂM SÀNG:** (Dùng gạch đầu dòng riêng cho X-quang, Siêu âm... kèm `[Trang X]`)
**3. LƯU Ý LÂM SÀNG & RED FLAGS:**
**4. MẸO GHI NHỚ LÂU:** (Đặt trong blockquote `>`)
**5. SƠ ĐỒ THUẬT TOÁN** ```dot ... ```
"""
            else:
                prompt = f"""TÀI LIỆU GIÁO TRÌNH:\n{ctx}\n\nCHỦ ĐỀ: "{query}"\n\n{PEDAGOGICAL_STYLE_RULES}
TRÌNH BÀY BÀI SOẠN 7 PHẦN (CÁC TIÊU ĐỀ LỚN PHẢI IN HOA VÀ TÔ ĐẬM):
**1. TỔNG QUAN, ĐỊNH NGHĨA & GIẢI PHẪU - SINH CƠ HỌC**
**2. PHÂN LOẠI TỔN THƯƠNG & Ý NGHĨA LÂM SÀNG (KÈM BẢNG MARKDOWN)**
**3. TRIỆU CHỨNG LÂM SÀNG & CƠ CHẾ HÌNH THÀNH DẤU HIỆU**
**4. CẬN LÂM SÀNG** (Bắt buộc dùng gạch đầu dòng riêng rẽ: - **X-quang:**... \n - **Siêu âm:**... \n - **MRI:**...)
**5. NGUYÊN TẮC & PHÁC ĐỒ ĐIỀU TRỊ** (Bắt buộc dùng gạch đầu dòng riêng rẽ: - **Nguyên tắc tối thượng:**... \n - **Điều trị bảo tồn:**... \n - **Chỉ định phẫu thuật:**...)
**6. PHỤC HỒI CHỨC NĂNG & CHĂM SÓC SAU MỔ**
**7. TIÊN LƯỢNG, PHÒNG NGỪA & XỬ TRÍ BIẾN CHỨNG**
Cuối bài kèm sơ đồ ```dot ... ```.
"""
            st.session_state.qa_session = {"query": query, "routed_context": ctx, "matched_pages": matched, "visual_keys": vis, "main_answer": stream_ai_response(prompt, vis if use_vision else None), "discussions": [], "saved_id": None}
            log_activity()
            st.rerun()

    sess = st.session_state.qa_session
    if sess:
        st.markdown(f"**Chủ đề đang xem:** {sess['query']}")
        render_graphviz_if_present(sess["main_answer"])
        display_cited_pages(sess["main_answer"], candidate_pages=sess["matched_pages"])
        st.markdown("---")
        c1, c2, c3 = st.columns([1.5, 1.5, 1.5])
        with c1: tag_in = st.text_input("Tags (phẩy):", value="hỏi đáp", key="qa_tags", label_visibility="collapsed")
        with c2:
            if not sess.get("saved_id"):
                if st.button("Lưu bài viết", key="qa_save", type="primary", use_container_width=True):
                    sess["saved_id"] = save_to_library("Hỏi Đáp", sess["query"][:65], sess["main_answer"] + "\n\n---\n" + "".join([f"\n> Phụ lục ({d['time']}): {d['user_text']}\n{d['ai_reply']}" for d in sess.get("discussions", [])]), tags=[t.strip() for t in tag_in.split(",") if t.strip()])
                    st.rerun()
            else:
                if st.button("Cập nhật bài viết", key="qa_upd", type="primary", use_container_width=True):
                    update_library_item(sess["saved_id"], sess["query"][:65], sess["main_answer"] + "\n\n---\n" + "".join([f"\n> Phụ lục ({d['time']}): {d['user_text']}\n{d['ai_reply']}" for d in sess.get("discussions", [])]), tags=[t.strip() for t in tag_in.split(",") if t.strip()])
        with c3:
            if st.button("Nghe bài đọc (TTS)", key="qa_tts", use_container_width=True): st.audio(text_to_speech(sess["main_answer"]), format="audio/mp3")
        if sess.get("discussions"):
            st.markdown("---")
            for d in sess["discussions"]:
                st.markdown(f"**Phụ lục ({d['time']}):** {d['user_text']}")
                render_graphviz_if_present(d["ai_reply"])
        st.markdown("---")
        with st.form("qa_followup_form", clear_on_submit=True):
            mode = st.radio("Tương tác bổ sung:", ["Hỏi thêm", "Yêu cầu chỉnh sửa toàn bài"], horizontal=True)
            follow_q = st.text_input("Nhập nội dung hỏi thêm hoặc góp ý (Nhấn Enter để gửi):")
            if st.form_submit_button("Gửi yêu cầu") and follow_q.strip():
                extra_ctx, _, ex_vis = smart_route_pages(sess["query"] + " " + follow_q, st.session_state.get("_books", []), max_pages=10)
                if "chỉnh sửa" in mode.lower():
                    sess["main_answer"] = stream_ai_response(f"TÀI LIỆU:\n{sess['routed_context'][:18000]}\n{extra_ctx[:12000]}\nCHỦ ĐỀ: {sess['query']}\nBÀI CŨ: {sess['main_answer']}\nYÊU CẦU SỬA: {follow_q}\n{PEDAGOGICAL_STYLE_RULES}\nViết lại toàn bài chi tiết, xuống dòng bằng gạch đầu dòng (-) rõ ràng kèm [Trang X].", ex_vis if use_vision else None)
                else:
                    ans_f = stream_ai_response(f"CHỦ ĐỀ GỐC: {sess['query']}\nTÀI LIỆU:\n{sess['routed_context'][:18000]}\n{extra_ctx[:10000]}\nBÀI TRƯỚC:\n{sess['main_answer'][:4000]}\nYÊU CẦU: {follow_q}\n{PEDAGOGICAL_STYLE_RULES}\nGiải thích sâu sắc, chia gạch đầu dòng (-) rõ ràng kèm [Trang X].", ex_vis if use_vision else None)
                    sess["discussions"].append({"type": "feedback", "user_text": follow_q, "ai_reply": ans_f, "time": datetime.datetime.now().strftime("%H:%M")})
                if sess.get("saved_id"): update_library_item(sess["saved_id"], sess["query"][:65], sess["main_answer"] + "\n\n---\n" + "".join([f"\n> Phụ lục ({d['time']}): {d['user_text']}\n{d['ai_reply']}" for d in sess.get("discussions", [])]))
                st.rerun()

elif page == "Bảng tin":
    if not st.session_state.news_session:
        src = st.selectbox("Chọn nguồn:", ["Từ Giáo trình", "JOT (LWW)", "JBJS (LWW)", "Injury (Elsevier)"])
        if src == "Từ Giáo trình":
            if st.button("Bốc thăm 4 chuyên đề ngẫu nhiên từ sách", type="primary"):
                books = st.session_state.get("_books", [])
                if books:
                    with get_db() as conn: sample = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({','.join('?'*len(books))}) ORDER BY RANDOM() LIMIT 10", books).fetchall()
                    res = generate_ai_response(f"Dựa vào các trang sau:\n{chr(10).join([f'[Trang {p}]: {t[:500]}' for _, p, t in sample])}\nChọn 4 chủ đề Chấn thương Chỉnh hình KHÁC NHAU. Trả về đúng 4 dòng: Tên chủ đề | Mô tả ngắn.")
                    st.session_state.news_headlines = [(p[0].replace('*', '').strip(), p[1].replace('*', '').strip()) for p in [line.split('|') for line in res.text.strip().split('\n') if '|' in line]]
            if st.session_state.news_headlines:
                st.markdown("---")
                for i, (t, tz) in enumerate(st.session_state.news_headlines):
                    with st.container(border=True):
                        st.markdown(f"**{t}**"); st.caption(tz)
                        if st.button("Mở đọc chuyên đề này", key=f"read_{i}"): st.session_state.news_session = {"query": t, "main_answer": ""}; st.rerun()
        else:
            urls = {"JOT (LWW)": "https://journals.lww.com/jorthotrauma/_layouts/15/OAKS.Journals/feed.aspx?FeedType=CurrentIssue", "JBJS (LWW)": "https://journals.lww.com/jbjsjournal/_layouts/15/OAKS.Journals/feed.aspx?FeedType=CurrentIssue", "Injury (Elsevier)": "https://rss.sciencedirect.com/publication/science/00201383"}
            if st.button(f"Tải 5 bài báo mới nhất từ {src}", type="primary"):
                try:
                    with urlopen(Request(urls[src], headers={"User-Agent": "Mozilla/5.0"}), timeout=8) as r: root = ET.fromstring(r.read())
                    st.session_state.rss_items = [{"title": (it.findtext("title") or "").strip(), "link": (it.findtext("link") or "").strip()} for it in root.findall(".//item")[:5]]
                except Exception: st.warning("Không kết nối được nguồn RSS này.")
            for i, it in enumerate(st.session_state.get("rss_items", [])):
                with st.container(border=True):
                    st.markdown(f"**{i+1}. {it['title']}**"); st.markdown(f"[Xem bài gốc]({it['link']})")
                    if st.button("Phân tích chủ đề này", key=f"rss_an_{i}"): st.session_state.news_session = {"query": it["title"], "main_answer": ""}; st.rerun()
    else:
        ns = st.session_state.news_session
        if st.button("Quay lại Bảng tin"): st.session_state.news_session = None; st.rerun()
        st.markdown(f"## {ns['query']}")
        if not ns["main_answer"]:
            ctx, mp, _ = smart_route_pages(ns["query"], st.session_state.get("_books", []), max_pages=15)
            ans_news = stream_ai_response(f"TÀI LIỆU:\n{ctx}\n{PEDAGOGICAL_STYLE_RULES}\nViết bài 7 phần dạng list chia mục rõ ràng, CÁC TIÊU ĐỀ LỚN PHẢI IN HOA VÀ TÔ ĐẬM (chia X-quang, Siêu âm, Nguyên tắc, Bảo tồn, Phẫu thuật thành gạch đầu dòng riêng) về '{ns['query']}' kèm [Trang X] và sơ đồ ```dot ... ```.")
            ns.update({"routed_context": ctx, "matched_pages": mp, "main_answer": ans_news})
            st.rerun()
        render_graphviz_if_present(ns["main_answer"])
        display_cited_pages(ns["main_answer"], candidate_pages=ns.get("matched_pages", []))
        if st.button("Lưu bài viết này", type="primary"): save_to_library("Bảng tin", ns["query"][:65], ns["main_answer"], tags=["bảng tin"])

elif page == "Trực Cấp cứu":
    topic = st.text_input("Chủ đề ca bệnh (bỏ trống để ngẫu nhiên):", key="er_topic_in")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Bắt đầu ca trực mới", type="primary", use_container_width=True):
            ctx, _, _ = smart_route_pages(topic or "gãy xương cấp cứu", st.session_state.get("_books", []), max_pages=12)
            ans_er = stream_ai_response(f"Bạn là BS Trưởng tua trực Chấn thương Chỉnh hình. TÀI LIỆU:\n{ctx[:25000]}\nTạo BƯỚC 1/3: bệnh cảnh lâm sàng. CHƯA tiết lộ chẩn đoán và X-quang. Đặt 3 câu hỏi thử thách. Cần gạch đầu dòng rõ ràng. Chủ đề: {topic or 'Ca cấp cứu điển hình'}")
            st.session_state.er_chat = [{"role": "assistant", "content": ans_er}]; st.session_state.er_case_active = True; log_activity(); st.rerun()
    with c2:
        if st.button("Xóa ca trực", use_container_width=True): st.session_state.er_chat = []; st.session_state.er_case_active = False; st.rerun()

    if st.session_state.er_case_active:
        for i, m in enumerate(st.session_state.er_chat):
            with st.chat_message(m["role"]): render_graphviz_if_present(m["content"]); display_cited_pages(m["content"], key_prefix=f"er_{i}") if m["role"] == "assistant" else None

        u = st.chat_input("Nhập xử trí của bạn (Nhấn Enter để gửi)...")
        if u:
            st.session_state.er_chat.append({"role": "user", "content": u})
            with st.chat_message("user"): st.markdown(u)
            with st.chat_message("assistant"):
                ctx, _, _ = smart_route_pages((topic or "") + " " + u, st.session_state.get("_books", []), max_pages=12)
                reply_er = stream_ai_response(f"TÀI LIỆU:\n{ctx[:25000]}\nDIỄN BIẾN:\n{chr(10).join([f'[{m['role']}]: {m['content']}' for m in st.session_state.er_chat])}\n{PEDAGOGICAL_STYLE_RULES}\n1. Chấm điểm (thang 100), giải thích chi tiết đúng/sai kèm [Trang X] bằng dạng list rõ ràng, in hoa tiêu đề.\n2. Chuyển bước tiếp theo (Bước 1->2 cung cấp X-quang; Bước 2->3 hỏi kế hoạch mổ & PHCN; xong Bước 3 tổng kết + sơ đồ ```dot```).")
            st.session_state.er_chat.append({"role": "assistant", "content": reply_er}); st.rerun()

        ch, cs = st.columns(2)
        with ch:
            if st.button("Xin gợi ý", use_container_width=True): st.info(generate_ai_response(f"Đưa ra 3 gợi ý ngắn gọn gạch đầu dòng cho tình huống:\n{st.session_state.er_chat[-1]['content']}").text)
        with cs:
            if st.button("Lưu ca trực vào Kho", use_container_width=True): save_to_library("Trực Cấp Cứu", topic or "Ca lâm sàng cấp cứu", "\n\n---\n\n".join(f"**{m['role'].upper()}:**\n{m['content']}" for m in st.session_state.er_chat), tags=["case", "ER"])

elif page == "Kho bài viết":
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1: kw = st.text_input("Tìm kiếm bài viết:", key="lib_search")
    with c2: f_type = st.selectbox("Phân loại:", ["Tất cả"] + sorted(set(i.get("type", "Khác") for i in st.session_state.library)))
    with c3: f_tag = st.selectbox("Tag:", ["Tất cả"] + sorted({t for i in st.session_state.library for t in (i.get("tags") or [])}))

    if not st.session_state.library: st.info("Kho lưu trữ trống.")
    else:
        for idx, item in enumerate(reversed(st.session_state.library)):
            title = item.get("title", ""); content = item.get("content", ""); tags = item.get("tags") or []
            if kw.lower() and kw.lower() not in title.lower() and kw.lower() not in content.lower(): continue
            if f_type != "Tất cả" and item.get("type") != f_type: continue
            if f_tag != "Tất cả" and f_tag not in tags: continue

            iid = item["id"]
            with st.expander(f"[{item.get('type', 'Bài lưu')}] {title} ({item.get('time', '')}{' | Tags: ' + ', '.join(tags) if tags else ''})"):
                c_e, c_img, c_tts, c_d = st.columns([1, 1.2, 1.2, 1])
                with c_e: is_edit = st.toggle("Soạn thảo Rich Text", key=f"e_{iid}_{idx}")
                with c_img: show_img = st.toggle("Xem trang gốc", key=f"img_{iid}_{idx}")
                with c_tts: play_tts = st.button("Nghe bài (TTS)", key=f"tts_{iid}_{idx}", use_container_width=True)
                with c_d:
                    if st.button("Xóa", key=f"d_{iid}_{idx}", use_container_width=True): delete_from_library(iid); st.rerun()

                if play_tts: st.audio(text_to_speech(content), format="audio/mp3")

                if is_edit:
                    st.info("Bôi đen văn bản để sử dụng công cụ In đậm, In nghiêng, Highlight, chèn Link hoặc danh sách.")
                    new_t = st.text_input("Tiêu đề:", value=title, key=f"et_{iid}_{idx}")
                    new_tags = st.text_input("Tags (phẩy):", value=", ".join(tags), key=f"etg_{iid}_{idx}")
                    
                    try:
                        import markdown
                        html_content = markdown.markdown(content, extensions=['tables'])
                        new_c_html = st_quill(value=html_content, html=True, key=f"quill_{iid}_{idx}")
                        if st.button("Lưu thay đổi", key=f"sv_{iid}_{idx}", type="primary"):
                            update_library_item(iid, new_t, new_c_html, tags=[t.strip() for t in new_tags.split(",") if t.strip()])
                            st.rerun()
                    except Exception:
                        new_c = st.text_area("Nội dung Markdown:", value=content, height=320, key=f"ec_{iid}_{idx}")
                        if st.button("Lưu thay đổi", key=f"sv_{iid}_{idx}", type="primary"):
                            update_library_item(iid, new_t, new_c, tags=[t.strip() for t in new_tags.split(",") if t.strip()])
                            st.rerun()
                else:
                    if "<p>" in content or "<strong>" in content: st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', content), unsafe_allow_html=True)
                    else: render_graphviz_if_present(content)
                    if show_img: display_cited_pages(content, key_prefix=f"lib_{iid}_{idx}")

elif page == "Tạo thẻ Anki":
    topic = st.chat_input("Nhập chủ đề cần tạo 10 thẻ Flashcard (Nhấn Enter để tạo)...")
    if topic:
        ctx, _, _ = smart_route_pages(topic, st.session_state.get("_books", []), max_pages=14)
        raw_anki = stream_ai_response(f"Tạo 10 Flashcard (Hỏi - Đáp giải thích sâu cơ chế giải phẫu/sinh cơ học dạng gạch đầu dòng kèm [Trang X]) về: {topic}.\nPhần 1: Hiển thị đọc trên web.\nPhần 2: Đặt trong khối ```csv\nMặt trước ; Mặt sau\n``` để import vào Anki.\nNguồn:\n{ctx}")
        csv_m = re.search(r'```csv\n(.*?)\n```', raw_anki, re.DOTALL)
        st.session_state.anki_session = {"topic": topic, "display": re.sub(r'```csv\n.*?\n```', '', raw_anki, flags=re.DOTALL), "csv": csv_m.group(1) if csv_m else None}
        log_activity(); st.rerun()

    ak = st.session_state.anki_session
    if ak:
        c1, c2 = st.columns(2)
        with c1:
            if ak.get("csv"): st.download_button("Tải file .CSV Import vào Anki", data=ak["csv"].encode("utf-8-sig"), file_name=f"anki_{datetime.date.today()}.csv", mime="text/csv", use_container_width=True)
        with c2:
            if st.button("Lưu bộ thẻ vào Kho bài viết", use_container_width=True): save_to_library("Anki", ak["topic"], ak["display"], tags=["anki"])
        st.markdown("---"); st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', ak["display"]), unsafe_allow_html=True)

elif page == "X-quang":
    cl, cr = st.columns([1, 1])
    img_bytes = None
    with cl:
        src = st.radio("Nguồn hình ảnh:", ["Upload ảnh (PNG/JPG)", "Rút trang từ sách PDF"], horizontal=True)
        if "Upload" in src:
            f = st.file_uploader("Chọn ảnh X-quang / Giải phẫu:", type=["png", "jpg", "jpeg"])
            if f: img_bytes = f.read(); st.image(img_bytes, use_container_width=True)
        else:
            books = st.session_state.get("_books", [])
            if books:
                bk = st.selectbox("Chọn sách:", books)
                with get_db() as conn: row = conn.execute("SELECT pages FROM books WHERE name=?", (bk,)).fetchone()
                mx = row[0] if row else 1
                pg = st.number_input(f"Trang (1 - {mx}):", 1, mx, 1)
                d = get_page_data(bk, pg)
                if d: img_bytes = get_page_image(d["fpath"], d["page_index"]); st.image(img_bytes, caption=f"{bk} - Trang {pg}", use_container_width=True)
    with cr:
        mode = st.selectbox("Chế độ phân tích:", ["Đọc phim X-quang / CT-scan, Các góc đo & Phân loại", "Phân tích Đường mổ, Gian cơ & Thần kinh nguy cơ", "Đố vui nhận diện mốc Giải phẫu (Active Recall)"])
        note = st.text_input("Câu hỏi hoặc vùng cần chú ý trên hình:")
        if st.button("Phân tích hình ảnh", type="primary", use_container_width=True):
            if not img_bytes: st.warning("Vui lòng chọn hoặc tải lên 1 hình ảnh trước.")
            else:
                with st.spinner("Đang phân tích hình ảnh bằng AI Vision..."):
                    part = types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                    if "Đố vui" in mode: p = f"Chọn 5 mốc giải phẫu/dụng cụ trên hình ({note}), đặt 5 câu hỏi nhận diện gạch đầu dòng. Đặt đáp án chi tiết sau dòng '--- ĐÁP ÁN ---'."
                    elif "Đường mổ" in mode: p = f"Phân tích chi tiết đường mổ trên hình ({note}): liệt kê từng lớp giải phẫu, khoảng gian thần kinh và cấu trúc mạch máu/thần kinh nguy cơ cao bằng gạch đầu dòng."
                    else: p = f"Đọc phim X-quang/giải thích hình ảnh ({note}) thật chi tiết: gạch đầu dòng mô tả tổn thương, cơ chế di lệch, các góc đo giải phẫu và hướng xử trí."
                    r = genai.Client(api_key=get_active_api_key()).models.generate_content(model=st.session_state.selected_model, contents=[part, p])
                    st.session_state.xray_session = {"title": note or mode[:35], "content": r.text}; log_activity()

        xs = st.session_state.xray_session
        if xs:
            if "--- ĐÁP ÁN ---" in xs["content"]:
                a, b = xs["content"].split("--- ĐÁP ÁN ---", 1)
                st.markdown(a)
                with st.expander("Xem Đáp án"): st.markdown(b)
            else: st.markdown(xs["content"])
            if st.button("Lưu bài phân tích vào Kho"): save_to_library("X-quang", xs["title"], xs["content"], tags=["xray"])

elif page == "USMLE & SRS":
    def parse_usmle_questions(raw_text):
        qs = []
        blocks = re.split(r'===\s*CAU_HOI\s*===', raw_text, flags=re.IGNORECASE)
        for b in blocks:
            b = b.strip()
            if not b or "===DAP_AN===" not in b.upper(): continue
            try:
                q = re.split(r'===\s*LUA_CHON\s*===', b, flags=re.I)[0].strip()
                r1 = re.split(r'===\s*LUA_CHON\s*===', b, flags=re.I)[1]
                opts_raw = re.split(r'===\s*DAP_AN\s*===', r1, flags=re.I)[0].strip()
                r2 = re.split(r'===\s*DAP_AN\s*===', r1, flags=re.I)[1]
                ans = re.split(r'===\s*GIAI_THICH\s*===', r2, flags=re.I)[0].strip()
                exp = re.split(r'===\s*GIAI_THICH\s*===', r2, flags=re.I)[1].strip()
                exp = re.sub(r'===\s*KET_THUC\s*===.*$', '', exp, flags=re.I | re.DOTALL).strip()
                opts = [l.strip().lstrip("-*• ") for l in opts_raw.split('\n') if l.strip()]
                qs.append({"question": q, "options": opts[:5] if len(opts) >= 2 else ["A", "B", "C", "D"], "answer": (ans or "A")[0].upper(), "explanation": exp})
            except Exception: continue
        return qs

    @st.fragment
    def render_usmle():
        due_list = get_due_mistakes()
        t1, t2, t3 = st.tabs(["Làm Đề USMLE", f"Sổ Tay Câu Sai ({len(st.session_state.mistakes)})", f"Ôn SRS Hôm Nay ({len(due_list)})"])

        with t1:
            c1, c2 = st.columns([3, 1])
            with c1: topic = st.text_input("Chủ đề Quiz (Để trống = ngẫu nhiên):")
            with c2: n_q = st.number_input("Số câu:", 1, 10, 4)

            if st.button("Tạo đề trắc nghiệm", type="primary"):
                books = st.session_state.get("_books", [])
                if not books: st.warning("Vui lòng chọn sách ở thanh bên trái.")
                else:
                    with st.spinner("Đang thiết kế câu hỏi tình huống lâm sàng..."):
                        if not topic.strip():
                            with get_db() as conn: sample = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({','.join('?'*len(books))}) ORDER BY RANDOM() LIMIT 10", books).fetchall()
                            ctx = "\n\n".join(f"[SÁCH:{b} | TRANG {p}]\n{t}" for b, p, t in sample)
                            target = "Chủ đề ngẫu nhiên"
                        else: ctx, _, _ = smart_route_pages(topic, books, max_pages=14); target = topic.strip()

                        prompt = f"""TÀI LIỆU:\n{ctx[:30000]}
Tạo đúng {n_q} câu USMLE Step 2 CK về: {target}.
ĐỊNH DẠNG BẮT BUỘC CHO MỖI CÂU:
===CAU_HOI===
[Tình huống lâm sàng chi tiết + Câu hỏi]
===LUA_CHON===
A. [...]
B. [...]
C. [...]
D. [...]
E. [...]
===DAP_AN===
[A/B/C/D/E]
===GIAI_THICH===
- **Đáp án đúng:** [Giải thích sâu cơ chế giải phẫu, sinh cơ học dạng gạch đầu dòng kèm [Trang X]]
- **Phân tích phương án nhiễu:** [Gạch đầu dòng giải thích tại sao các đáp án khác sai]
===KET_THUC===
"""
                        res = generate_ai_response(prompt)
                        qs = parse_usmle_questions(res.text)
                        if qs: st.session_state.interactive_quiz = {"topic": target, "questions": qs, "submitted": False}; st.rerun()
                        else: st.error("Chưa tách được câu hỏi, vui lòng bấm Tạo đề lại.")

            iq = st.session_state.interactive_quiz
            if iq and "questions" in iq:
                user_sel = {}
                for qi, qit in enumerate(iq["questions"]):
                    with st.container(border=True):
                        st.markdown(f"**Câu {qi+1}:**\n\n{qit['question']}")
                        choice = st.radio(f"Chọn đáp án câu {qi+1}:", qit["options"], index=None, key=f"uq_{iq['topic']}_{qi}")
                        user_sel[qi] = choice
                        if iq["submitted"]:
                            if (choice or "")[0].upper() == qit["answer"][0].upper(): st.success(f"Chính xác! ({qit['answer'][0].upper()})")
                            else: st.error(f"Sai. Bạn chọn {(choice or '—')[0].upper()} | Đáp án đúng: {qit['answer'][0].upper()}")
                            st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', qit["explanation"]), unsafe_allow_html=True)

                if not iq["submitted"]:
                    if st.button("Nộp bài & Chấm điểm", type="primary", use_container_width=True):
                        iq["submitted"] = True; correct = sum(1 for qi, qit in enumerate(iq["questions"]) if (user_sel.get(qi) or "")[:1].upper() == qit["answer"][0].upper())
                        for qi, qit in enumerate(iq["questions"]):
                            if (user_sel.get(qi) or "")[:1].upper() != qit["answer"][0].upper(): add_to_mistake_bank(qit, user_sel.get(qi) or "Chưa chọn", iq["topic"])
                        st.session_state.progress["quizzes"] = st.session_state.progress.get("quizzes", 0) + 1
                        st.session_state.progress["correct"] = st.session_state.progress.get("correct", 0) + correct
                        save_json_file(PROGRESS_FILE, st.session_state.progress); log_activity(); st.rerun()

        with t2:
            if not st.session_state.mistakes: st.success("Sổ tay câu sai hiện đang trống.")
            else:
                for i, m in enumerate(reversed(st.session_state.mistakes)):
                    with st.expander(f"[{m.get('topic', '')}] {m['question'][:80]}..."):
                        st.markdown(m["question"])
                        for o in m.get("options", []): st.markdown(f"- {o}")
                        st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', f"**Đáp án đúng: {m['answer']}**\n\n{m.get('explanation', '')}"), unsafe_allow_html=True)
                        if st.button("Xóa khỏi sổ tay", key=f"dms_{i}"): st.session_state.mistakes = [x for x in st.session_state.mistakes if x.get("id") != m.get("id")]; save_json_file(MISTAKES_FILE, st.session_state.mistakes); st.rerun()

        with t3:
            if not due_list: st.success("Hôm nay không có câu hỏi nào đến hạn ôn tập SRS.")
            else:
                for idx, m in enumerate(due_list[:10]):
                    with st.container(border=True):
                        st.caption(f"Chủ đề: {m.get('topic', '')} | Lần lặp: {m.get('repetitions', 0)} | Ease: {m.get('ease', 2.5):.2f}")
                        st.markdown(m["question"])
                        for opt in m.get("options", []): st.markdown(f"- {opt}")
                        with st.expander("Xem đáp án & Đánh giá trí nhớ (SM-2)"):
                            st.markdown(re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', f"**Đáp án: {m['answer']}**\n\n{m.get('explanation', '')}"), unsafe_allow_html=True)
                            c1, c2, c3 = st.columns(3)
                            for col, (lbl, q_val) in zip([c1, c2, c3], [("Quên (0)", 0), ("Khó (3)", 3), ("Dễ (5)", 5)]):
                                with col:
                                    if st.button(lbl, key=f"srs_{m['id']}_{q_val}", use_container_width=True): sm2_update(m, q_val); save_json_file(MISTAKES_FILE, st.session_state.mistakes); st.rerun()

    render_usmle()

elif page == "Tạo Slide":
    topic = st.chat_input("Nhập chủ đề Slide báo cáo (Nhấn Enter để tạo)...")
    if topic:
        ctx, _, _ = smart_route_pages(topic, st.session_state.get("_books", []), max_pages=15)
        raw_slide = stream_ai_response(f"""Thiết kế 8 slide thuyết trình chi tiết về "{topic}". Mỗi slide đúng định dạng (dùng gạch đầu dòng rõ ràng):
### SLIDE [số]: [Tiêu đề]
- Ý 1 (Diễn giải chi tiết 3-4 câu)
- Ý 2 (Diễn giải chi tiết 3-4 câu)
- Ý 3 (Diễn giải chi tiết 3-4 câu)
**Ghi chú thuyết trình:** (Giải thích thật sâu cơ chế giải phẫu, sinh cơ học dạng list kèm [Trang X])
Nguồn: {ctx}""")
        st.session_state.slide_session = {"topic": topic, "text": raw_slide, "pptx": create_pptx_bytes(topic, raw_slide)}
        log_activity(); st.rerun()

    ss = st.session_state.slide_session
    if ss:
        st.markdown("### Chỉnh sửa Outline trước khi xuất file .PPTX")
        edited = st.text_area("Nội dung Outline (Có thể sửa trực tiếp):", value=ss["text"], height=320)
        if edited != ss["text"]: ss["text"] = edited; ss["pptx"] = create_pptx_bytes(ss["topic"], ss["text"])

        c1, c2 = st.columns(2)
        with c1:
            if ss.get("pptx"): st.download_button("Tải file PowerPoint (.PPTX)", data=ss["pptx"], file_name=f"Slide_{re.sub(r'[^A-Za-z0-9_]', '_', ss['topic'][:25])}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
        with c2:
            if st.button("Lưu Outline vào Kho bài viết", use_container_width=True): save_to_library("Slide", ss["topic"], ss["text"], tags=["slide"])
