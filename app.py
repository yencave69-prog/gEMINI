import streamlit as st
import streamlit.components.v1 as components
import datetime, os, sys, subprocess, re, json, csv, io, random, time
import sqlite3
import zipfile
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
from google import genai
from google.genai import types
from streamlit_quill import st_quill
gdown = _auto_pip("gdown", lambda: __import__("gdown"))

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "python-pptx"])
    from pptx import Presentation
    from pptx.util import Inches, Pt

# ==================== 1. CẤU HÌNH TRANG & DỮ LIỆU ĐÁM MÂY ====================
if "_sidebar_state" not in st.session_state:
    st.session_state["_sidebar_state"] = "expanded"

st.set_page_config(
    page_title="Chấn Thương Chỉnh Hình",
    layout="wide",
    initial_sidebar_state=st.session_state["_sidebar_state"]
)

st.markdown("""
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="theme-color" content="#2B7A78">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
""", unsafe_allow_html=True)

DRIVE_FILE_ID = "1g1FSfgKerUEFSM7K5mJnAlw2pH4u2uT2"
GDRIVE_FOLDER_URL = ""

LIBRARY_FILE = "dudu_library.json"
MISTAKES_FILE = "dudu_mistakes.json"
SETTINGS_FILE = "dudu_settings.json"
PROGRESS_FILE = "dudu_progress.json"
ACTIVE_SESSIONS_FILE = "dudu_active_sessions.json"
PDF_DIR = "kho_sach"
DB_FILE = os.path.join(PDF_DIR, "ortho_fts5.db")
SECRETS_DIR = ".streamlit"
SECRETS_PATH = os.path.join(SECRETS_DIR, "secrets.toml")
os.makedirs(PDF_DIR, exist_ok=True)

WORD_LIKE_TOOLBAR = [
    [{"header": [1, 2, 3, False]}],
    ["bold", "italic", "underline", "strike"],
    [{"color": []}, {"background": []}],
    [{"list": "ordered"}, {"list": "bullet"}],
    ["blockquote", "link", "image"],
    ["clean"]
]

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

def save_active_sessions():
    """Lưu bền vững phiên Hỏi đáp, Bảng tin, Cấp cứu xuống đĩa để chuyển tab hoặc bấm Lưu không bao giờ mất bài"""
    data = {
        "qa_session": st.session_state.get("qa_session"),
        "news_session": st.session_state.get("news_session"),
        "er_chat": st.session_state.get("er_chat", []),
        "er_case_active": st.session_state.get("er_case_active", False),
        "anki_session": st.session_state.get("anki_session"),
        "slide_session": st.session_state.get("slide_session", {}).get("text") if st.session_state.get("slide_session") else None,
        "slide_topic": st.session_state.get("slide_session", {}).get("topic") if st.session_state.get("slide_session") else None,
    }
    try:
        save_json_file(ACTIVE_SESSIONS_FILE, data)
    except Exception:
        pass

def repair_items_id(items_list, filepath):
    seen = set(); changed = False
    for idx, item in enumerate(items_list):
        cid = item.get("id")
        if not cid or str(cid) == "None" or str(cid) in seen:
            item["id"] = f"item_{idx}_{datetime.datetime.now().strftime('%H%M%S')}_{random.randint(1000,9999)}"
            seen.add(item["id"]); changed = True
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
    return "AQ.Ab8RN6IVfc-wewsRumOJ-eZF0Y0bedH16oIhp_o4KNHNbK05Cg"

def save_api_key_everywhere(new_key: str):
    clean_k = new_key.strip().strip('"').strip("'")
    st.session_state.settings["api_key"] = clean_k
    save_json_file(SETTINGS_FILE, st.session_state.settings)
    try:
        os.makedirs(SECRETS_DIR, exist_ok=True)
        with open(SECRETS_PATH, "w", encoding="utf-8") as f: f.write(f'GEMINI_API_KEY = "{clean_k}"\n')
    except Exception: pass

# ==================== 2. CƠ SỞ DỮ LIỆU SQLITE FTS5 & ĐỒNG BỘ SÁCH MỚI ====================
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
            conn.execute("DELETE FROM pages WHERE book=?", (fname,))
            conn.executemany("INSERT INTO pages(book, page, text) VALUES (?,?,?)", rows)
            conn.execute("INSERT OR REPLACE INTO books(name, pages, indexed_at) VALUES (?,?,?)", (fname, len(rows), datetime.datetime.now().isoformat()))
            conn.commit()
        return True
    except Exception: return False

def ensure_database_ready(force_reload=False, custom_folder_url=""):
    saved_cfg = load_json_file(SETTINGS_FILE, {})
    folder_url = custom_folder_url.strip() or saved_cfg.get("gdrive_folder_url", "").strip()
    if not folder_url and hasattr(st, "secrets"):
        try: folder_url = st.secrets.get("GDRIVE_FOLDER_URL", "").strip()
        except Exception: pass
    if not folder_url:
        folder_url = GDRIVE_FOLDER_URL.strip()

    if force_reload:
        if os.path.exists(DB_FILE):
            try: os.remove(DB_FILE)
            except Exception: pass

    if not os.path.exists(DB_FILE) or os.path.getsize(DB_FILE) < 100000:
        with st.spinner("⏳ Đang tải cơ sở dữ liệu sách từ Google Drive về Cloud..."):
            try:
                file_id = st.secrets.get("DRIVE_FILE_ID", DRIVE_FILE_ID) if hasattr(st, "secrets") else DRIVE_FILE_ID
                gdown.download(id=file_id, output=DB_FILE, quiet=False, fuzzy=True)
            except Exception as e:
                st.warning(f"Lưu ý khi tải file DB: {e}")

    init_db_tables()

    # Đồng bộ từ thư mục Google Drive nếu người dùng cung cấp link thư mục
    if folder_url and "drive.google.com" in folder_url and (force_reload or not any(f.lower().endswith(".pdf") for f in os.listdir(PDF_DIR))):
        with st.spinner("⏳ Đang tải các file sách mới (bao gồm thct) từ thư mục Google Drive..."):
            try:
                gdown.download_folder(url=folder_url, output=PDF_DIR, quiet=True, use_cookies=False)
                new_pdfs = sorted([f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")])
                if new_pdfs:
                    with get_db() as conn:
                        db_books = [r[0] for r in conn.execute("SELECT DISTINCT book FROM pages").fetchall()]
                        for old_b in db_books:
                            if old_b not in new_pdfs:
                                conn.execute("DELETE FROM pages WHERE book=?", (old_b,))
                                conn.execute("DELETE FROM books WHERE name=?", (old_b,))
                        conn.commit()
                    for pdf_f in new_pdfs:
                        index_single_book_fast(pdf_f)
            except Exception as e:
                st.warning(f"Không thể tải tự động thư mục Drive: {e}")

ensure_database_ready(force_reload=False)

local_pdf_files = sorted([f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")])
for pdf_f in local_pdf_files:
    if not is_book_indexed(pdf_f):
        index_single_book_fast(pdf_f)

def remove_books_from_db(book_names_to_delete):
    if not book_names_to_delete: return
    with get_db() as conn:
        for bname in book_names_to_delete:
            conn.execute("DELETE FROM pages WHERE book=?", (bname,))
            conn.execute("DELETE FROM books WHERE name=?", (bname,))
        conn.commit()

def deduplicate_legacy_books_auto():
    with get_db() as conn:
        rows = conn.execute("SELECT DISTINCT book FROM pages").fetchall()
    raw_books = sorted([r[0] for r in rows]) if rows else []

    known_old_duplicates = [
        "Dutton_s Orthopaedic-Examination, Evaluation and Intervention 4e.pdf",
        "Rockwood and Green_s Fractures in Adults 8e.pdf",
        "Chẩn đoán và điều trị trật khớp chi trên_downloadsachyhoc.com.pdf"
    ]
    to_auto_remove = []
    for old_b in known_old_duplicates:
        if old_b in raw_books:
            if ("Dutton" in old_b and "Dutton.pdf" in raw_books) or \
               ("Rockwood" in old_b and "Rockwood and Green.pdf" in raw_books) or \
               ("downloadsachyhoc" in old_b and ("chitren.pdf" in raw_books or len(raw_books) > 5)):
                to_auto_remove.append(old_b)

    if to_auto_remove:
        remove_books_from_db(to_auto_remove)
        raw_books = [b for b in raw_books if b not in to_auto_remove]

    saved_settings = load_json_file(SETTINGS_FILE, {})
    hidden_books = set(saved_settings.get("hidden_books", []))
    return [b for b in raw_books if b not in hidden_books]

all_books_in_db = deduplicate_legacy_books_auto()

def get_thct_emergency_books(all_books, selected_books):
    thct_matches = [b for b in all_books if "thct" in b.lower()]
    if thct_matches:
        return thct_matches
    return selected_books if selected_books else all_books

def get_page_data(fname, page_num):
    with get_db() as conn:
        row = conn.execute("SELECT text FROM pages WHERE book=? AND page=?", (fname, page_num)).fetchone()
    if not row: return None
    return {"text": row[0], "fpath": os.path.join(PDF_DIR, fname), "page_index": page_num - 1}

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

    seen = set(); combined = []
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
    if not os.path.exists(fpath): return None
    try:
        doc = fitz.open(fpath)
        img_bytes = doc.load_page(page_index).get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")
        doc.close()
        return img_bytes
    except Exception: return None

# ==================== 3. KHỞI TẠO STATE BỀN VỮNG ====================
_saved_active = load_json_file(ACTIVE_SESSIONS_FILE, {})

_defaults = {
    "library": lambda: repair_items_id(load_json_file(LIBRARY_FILE, []), LIBRARY_FILE),
    "mistakes": lambda: repair_items_id(load_json_file(MISTAKES_FILE, []), MISTAKES_FILE),
    "settings": lambda: load_json_file(SETTINGS_FILE, {"dark_mode": False, "api_key": "", "hidden_books": [], "gdrive_folder_url": ""}),
    "progress": lambda: load_json_file(PROGRESS_FILE, {"activity": {}, "quizzes": 0, "correct": 0}),
    "er_chat": lambda: _saved_active.get("er_chat", []),
    "er_case_active": lambda: _saved_active.get("er_case_active", False),
    "qa_session": lambda: _saved_active.get("qa_session", None),
    "news_headlines": list,
    "news_session": lambda: _saved_active.get("news_session", None),
    "anki_session": lambda: _saved_active.get("anki_session", None),
    "xray_session": lambda: None,
    "interactive_quiz": lambda: None,
    "slide_session": lambda: (
        {"topic": _saved_active["slide_topic"], "text": _saved_active["slide_session"], "pptx": create_pptx_bytes(_saved_active["slide_topic"], _saved_active["slide_session"])}
        if _saved_active.get("slide_session") and _saved_active.get("slide_topic") else None
    ),
    "selected_model": lambda: "gemini-flash-lite-latest",
    "_books": lambda: list(all_books_in_db),
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

def clean_cite_spans_for_storage(text: str):
    return re.sub(r'<span\s+class=["\']cite-tag["\']>(.*?)</span>', r'\1', text or "")

def save_to_library(item_type, title, content, tags=None):
    disk_lib = load_json_file(LIBRARY_FILE, [])
    existing_ids = {str(x.get("id")) for x in st.session_state.library}
    for d_item in disk_lib:
        if str(d_item.get("id")) not in existing_ids:
            st.session_state.library.append(d_item)

    item_id = f"id_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{random.randint(1000,9999)}"
    clean_content = clean_cite_spans_for_storage(content)
    new_entry = {
        "id": item_id,
        "time": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "type": item_type,
        "title": title,
        "content": clean_content,
        "tags": tags or []
    }
    st.session_state.library.append(new_entry)
    save_json_file(LIBRARY_FILE, st.session_state.library)
    save_active_sessions()
    st.toast("✅ Đã lưu thành công vào Kho bài viết!")
    return item_id

def update_library_item(item_id, new_title, new_content, tags=None):
    clean_content = clean_cite_spans_for_storage(new_content)
    found = False
    for item in st.session_state.library:
        if str(item.get("id")) == str(item_id):
            item["title"] = new_title
            item["content"] = clean_content
            item["time"] = datetime.datetime.now().strftime("%d/%m/%Y %H:%M") + " (Đã sửa)"
            if tags is not None: item["tags"] = tags
            found = True
            break
    if not found:
        return save_to_library("Hỏi Đáp", new_title, clean_content, tags=tags)
    save_json_file(LIBRARY_FILE, st.session_state.library)
    save_active_sessions()
    st.toast("✅ Đã cập nhật bài viết!")
    return item_id

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
    st.session_state.mistakes.append({
        "id": f"mis_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{random.randint(1000,9999)}",
        "time": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "topic": topic, "question": q_obj.get("question", ""), "options": q_obj.get("options", []),
        "answer": q_obj.get("answer", "A"), "user_wrong_choice": user_choice,
        "explanation": q_obj.get("explanation", ""), "ease": 2.5, "interval": 1,
        "repetitions": 0, "next_review": datetime.date.today().isoformat()
    })
    save_json_file(MISTAKES_FILE, st.session_state.mistakes)

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

# ==================== 4. AI ENGINE ====================
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

def _get_fallback_models():
    primary = st.session_state.get("selected_model", "gemini-flash-lite-latest")
    candidates = [primary, "gemini-flash-lite-latest", "gemini-flash-latest", "gemini-2.0-flash"]
    seen = set()
    return [m for m in candidates if not (m in seen or seen.add(m))]

def generate_ai_response(prompt_text):
    client = genai.Client(api_key=get_active_api_key())
    last_err = None
    with st.spinner("⏳ Đang xử lý dữ liệu với AI..."):
        for model_name in _get_fallback_models():
            for attempt in range(2):
                try:
                    return client.models.generate_content(
                        model=model_name,
                        contents=prompt_text,
                        config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)
                    )
                except Exception as e:
                    last_err = e
                    err_str = str(e)
                    if "401" in err_str or "UNAUTHENTICATED" in err_str or "API_KEY_INVALID" in err_str:
                        st.error("❌ Lỗi xác thực API Key. Hãy kiểm tra lại khóa API ở cột bên trái.")
                        st.stop()
                    time.sleep(1.2 * (attempt + 1))
    st.error(f"⚠️ Máy chủ AI đang quá tải tạm thời. Vui lòng thử lại sau vài giây. Chi tiết: {last_err}")
    st.stop()

def stream_ai_response(prompt_text, visual_page_keys=None):
    client = genai.Client(api_key=get_active_api_key())
    contents = []
    if visual_page_keys:
        for (fname, p_num) in visual_page_keys[:2]:
            img_b = get_page_image(os.path.join(PDF_DIR, fname), p_num - 1, zoom=1.1)
            if img_b: contents.append(types.Part.from_bytes(data=img_b, mime_type="image/png"))
    contents.append(prompt_text)

    placeholder = st.empty()
    placeholder.info("⏳ **Đang kết nối máy chủ AI và biên soạn câu trả lời... (Vui lòng chờ trong giây lát)**")
    
    full_text = ""
    last_err = None

    for model_name in _get_fallback_models():
        for attempt in range(2):
            try:
                full_text = ""
                for chunk in client.models.generate_content_stream(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)
                ):
                    if chunk.text:
                        full_text += chunk.text
                        placeholder.markdown(format_small_citations(full_text), unsafe_allow_html=True)
                if full_text.strip():
                    placeholder.empty()
                    return full_text
            except Exception as e:
                last_err = e
                err_str = str(e)
                if "401" in err_str or "UNAUTHENTICATED" in err_str or "API_KEY_INVALID" in err_str:
                    placeholder.empty()
                    st.error("❌ Lỗi xác thực API Key (401). Hãy kiểm tra lại khóa API ở cột bên trái.")
                    st.stop()
                if len(full_text) > 400:
                    placeholder.empty()
                    return full_text
                placeholder.info(f"⏳ **Máy chủ {model_name} đang bận, đang tự động chuyển luồng xử lý dự phòng (Lần thử {attempt+1})...**")
                time.sleep(1.5 * (attempt + 1))

        try:
            res = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=types.GenerateContentConfig(temperature=0.35, max_output_tokens=8192)
            )
            if res and res.text:
                placeholder.empty()
                return res.text
        except Exception as e:
            last_err = e
            continue

    placeholder.empty()
    st.error(f"⚠️ Máy chủ Google Gemini hiện đang quá tải (ServerError 503). Vui lòng bấm gửi lại câu hỏi sau vài giây! ({last_err})")
    st.stop()

@st.cache_data(max_entries=200, show_spinner=False)
def translate_to_vietnamese_cached(raw_text: str):
    if not raw_text or not raw_text.strip():
        return ""
    sample = raw_text[:1500]
    vn_chars = re.findall(r'[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]', sample.lower())
    if len(vn_chars) > 15:
        return raw_text
    try:
        client = genai.Client(api_key=get_active_api_key())
        prompt = (
            "Dịch toàn bộ đoạn văn bản y khoa Chấn thương Chỉnh hình sau đây sang Tiếng Việt chuẩn xác, "
            "rõ ràng, giữ nguyên các số hiệu hình ảnh (như Fig, Table) và cấu trúc đoạn. "
            "Chỉ trả về nội dung bản dịch Tiếng Việt, không thêm lời bình:\n\n" + raw_text[:4500]
        )
        for m in ["gemini-flash-lite-latest", "gemini-flash-latest"]:
            try:
                res = client.models.generate_content(
                    model=m,
                    contents=prompt,
                    config=types.GenerateContentConfig(temperature=0.2, max_output_tokens=4096)
                )
                if res and res.text:
                    return res.text.strip()
            except Exception:
                continue
        return raw_text
    except Exception:
        return raw_text

PEDAGOGICAL_STYLE_RULES = """
NGUYÊN TẮC TRÌNH BÀY BẮT BUỘC ĐỂ TRÁNH KHỐI CHỮ DÀY ĐẶC:
1. KHÔNG VIẾT ĐOẠN VĂN DÀI HAY GỘP NHIỀU Ý VÀO NHAU. Mọi nội dung lâm sàng, X-quang, MRI, Điều trị phải được phân tách rõ ràng.
2. CẤU TRÚC PHÂN TẦNG VÀ TIÊU ĐỀ BẮT BUỘC PHẢI IN HOA + TÔ ĐẬM:
   - Các tiêu đề lớn phải dùng `1.`, `2.`, `3.` kèm IN HOA VÀ TÔ ĐẬM. (Ví dụ: **1. ĐẠI CƯƠNG VÀ CƠ CHẾ**, **2. ĐIỀU TRỊ BẢO TỒN VÀ PHẪU THUẬT**).
   - Dưới mỗi tiêu đề lớn, mọi phương pháp hay nguyên tắc bắt buộc phải xuống dòng và dùng Gạch đầu dòng (`-`).
   - Mở đầu mỗi gạch đầu dòng, từ khóa phải TÔ ĐẬM.
3. Trong mỗi gạch đầu dòng (`-`), diễn giải dài và sâu (5-8 câu), cung cấp đầy đủ các số liệu định lượng (độ, mm, tuần lễ) và cơ chế giải phẫu/sinh lý bệnh kèm [Trang X].
4. Lập **Bảng (Markdown Table)** khi cần so sánh. Cuối mỗi phần quan trọng đóng khung blockquote (`>`) đúc kết mẹo lâm sàng.
"""

def format_small_citations(text: str):
    cleaned = clean_cite_spans_for_storage(text)
    return re.sub(r'(\[\s*(?:[^\]]*?[Tt]rang\s*\d+[^\]]*?)\s*\])', r'<span class="cite-tag">\1</span>', cleaned)

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
    st.markdown("#### 📖 Nội dung & Hình ảnh trích xuất từ sách gốc:")
    for idx, (bk, p, d) in enumerate(matched[:6]):
        with st.expander(f"Xem gốc: {bk} - Trang {p}"):
            with st.spinner("⏳ Đang tải trang sách và dịch nội dung sang Tiếng Việt..."):
                vn_text = translate_to_vietnamese_cached(d["text"])
                img_data = get_page_image(d["fpath"], d["page_index"])
                if img_data:
                    c1, c2 = st.columns([3, 2])
                    with c1: st.image(img_data, use_container_width=True)
                    with c2: st.text_area("Văn bản gốc (Tiếng Việt):", value=vn_text, height=280, disabled=True, key=f"{key_prefix}_{idx}_{p}")
                else:
                    st.text_area("Văn bản gốc (Tiếng Việt):", value=vn_text, height=220, disabled=True, key=f"{key_prefix}_{idx}_{p}")

def render_graphviz_if_present(text):
    dots = re.findall(r'```dot\n(.*?)\n```', text, re.DOTALL)
    clean = re.sub(r'```dot\n.*?\n```', '', text, flags=re.DOTALL)
    st.markdown(format_small_citations(clean), unsafe_allow_html=True)
    if dots:
        st.markdown("#### Sơ đồ thuật toán:")
        for d in dots:
            try: st.graphviz_chart(d)
            except Exception: st.code(d)

def force_open_sidebar():
    """Hàm mở lại thanh công cụ bên trái bằng cả State lẫn JavaScript để đảm bảo 100% hoạt động"""
    st.session_state["_sidebar_state"] = "expanded"
    components.html(
        """
        <script>
        try {
            const doc = window.parent.document;
            const selectors = [
                '[data-testid="stSidebarCollapsedControl"] button',
                '[data-testid="collapsedControl"] button',
                '[data-testid="stExpandSidebarButton"]',
                'button[aria-label*="sidebar" i]',
                'header[data-testid="stHeader"] button'
            ];
            for (const sel of selectors) {
                const btn = doc.querySelector(sel);
                if (btn) { btn.click(); break; }
            }
            const sb = doc.querySelector('[data-testid="stSidebar"]');
            if (sb && sb.getAttribute('aria-expanded') === 'false') {
                sb.setAttribute('aria-expanded', 'true');
            }
        } catch (e) {}
        </script>
        """,
        height=0,
    )

# ==================== 5. GIAO DIỆN, DARK MODE & NÚT MỞ SIDEBAR LUÔN HIỂN THỊ ====================
dark_mode = st.session_state.settings.get("dark_mode", False)
bg_app, bg_sidebar, card, bg_subcard, border_col, text_main, text_sub, accent_col, input_bg = (
    ("#0E141B", "#151F28", "#1B2733", "#131C26", "#2C3E50", "#ECF2F8", "#9FB3C8", "#38B2AC", "#111921")
    if dark_mode else
    ("#FAF7F2", "#F4EBE1", "#FFFFFF", "#F8F4EC", "#E6D5C3", "#2C1E16", "#6E503B", "#2B7A78", "#FFFFFF")
)

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, .stApp, .stMarkdown, p, li, h1, h2, h3, h4, h5, h6, label, input, textarea, button {{ font-family: 'Inter', sans-serif; }}
    [data-testid="stIconMaterial"], .material-icons, span[translate="no"] {{ font-family: 'Material Symbols Rounded' !important; font-weight: normal !important; font-style: normal !important; }}
    
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
        background-color: {bg_app} !important;
        color: {text_main} !important;
    }}
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div {{
        background-color: {bg_sidebar} !important;
        border-right: 1px solid {border_col} !important;
    }}

    /* SỬA LỖI 2: Hiển thị rõ ràng nút mở Sidebar ở góc trái trên cùng mọi lúc */
    header[data-testid="stHeader"] {{
        background: transparent !important;
        display: flex !important;
        visibility: visible !important;
        height: 3.2rem !important;
    }}
    [data-testid="stDecoration"], [data-testid="stStatusWidget"], .stAppDeployButton {{
        display: none !important;
    }}
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stExpandSidebarButton"] {{
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        position: fixed !important;
        top: 10px !important;
        left: 12px !important;
        z-index: 999999 !important;
        background-color: {accent_col} !important;
        border-radius: 8px !important;
        padding: 4px 8px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.35) !important;
    }}
    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="collapsedControl"] button,
    [data-testid="stSidebarHeader"] button {{
        background-color: {accent_col} !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
    }}
    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="collapsedControl"] svg,
    [data-testid="stSidebarHeader"] button svg {{
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }}

    .block-container {{ padding-top: 1.5rem !important; }}
    
    .stApp p, .stApp li, .stApp label, .stMarkdown {{ color: {text_main} !important; line-height: 1.76 !important; margin-bottom: 0.8rem !important; }}
    .stApp h1, .stApp h2, .stApp h3, .stApp h4 {{ color: {text_main} !important; margin-top: 1.5rem !important; margin-bottom: 0.8rem !important; }}
    .cite-tag {{ font-size: 0.72em !important; color: {text_sub} !important; opacity: 0.82; font-weight: 500; margin-left: 4px; background: {bg_subcard}; padding: 2px 6px; border-radius: 4px; }}
    
    .stButton > button {{ background-color: transparent !important; color: {text_main} !important; border: 1px solid {border_col} !important; border-radius: 8px; font-weight: 500; }}
    .stButton > button:hover {{ border-color: {accent_col} !important; color: {accent_col} !important; }}
    .stButton > button[kind="primary"] {{ background-color: {accent_col} !important; color: white !important; border: none !important; }}
    [data-testid="stSidebar"] .stButton > button {{ background: transparent !important; color: {text_main} !important; border: none !important; text-align: center !important; justify-content: center !important; display: flex !important; align-items: center !important; padding: 11px 14px !important; margin-bottom: 4px; width: 100% !important; }}
    [data-testid="stSidebar"] .stButton > button p {{ text-align: center !important; width: 100% !important; margin: 0 !important; }}
    [data-testid="stSidebar"] .stButton > button:hover {{ background-color: {bg_subcard} !important; border-radius: 10px; }}
    [data-testid="stSidebar"] .stButton > button[kind="primary"] {{ background-color: {accent_col} !important; color: white !important; font-weight: 600 !important; border-radius: 10px; }}
    [data-testid="stSidebar"] .stButton > button[kind="primary"] p {{ color: white !important; }}
    
    .stTextInput input, .stTextArea textarea, .stSelectbox [data-baseweb="select"] > div {{ background-color: {input_bg} !important; color: {text_main} !important; border: 1px solid {border_col} !important; border-radius: 8px !important; }}
    .stTextArea textarea:disabled {{ color: {text_main} !important; opacity: 0.92 !important; -webkit-text-fill-color: {text_main} !important; }}
    .stMarkdown table {{ width: 100%; border-collapse: collapse; margin: 16px 0; border: 1px solid {border_col}; }}
    .stMarkdown th {{ background-color: {bg_subcard}; color: {text_main}; padding: 12px; border-bottom: 1px solid {border_col}; text-align: left; }}
    .stMarkdown td {{ padding: 12px; border-bottom: 1px solid {border_col}; }}
    [data-testid="stExpander"], [data-testid="stMetric"] {{ background-color: transparent !important; border: 1px solid {border_col} !important; border-radius: 10px !important; margin-bottom: 10px !important; }}
    .element-container:has(> div:empty) {{ display: none !important; }}

    .ql-toolbar {{ background-color: #F8FAFC !important; border: 1px solid {border_col} !important; border-radius: 10px 10px 0 0 !important; padding: 8px !important; }}
    .ql-container {{ border: 1px solid {border_col} !important; border-radius: 0 0 10px 10px !important; font-family: 'Inter', sans-serif !important; font-size: 1.02rem !important; background-color: {input_bg} !important; }}
    .ql-editor {{ color: {text_main} !important; min-height: 400px !important; line-height: 1.75 !important; }}
    .ql-editor img, .stMarkdown img {{ max-width: 100% !important; height: auto !important; border-radius: 8px !important; margin: 12px 0 !important; box-shadow: 0 2px 8px rgba(0,0,0,0.12); }}

    [data-testid="stBottom"], [data-testid="stBottom"] > div, [data-testid="stBottomBlockContainer"] {{
        background-color: {bg_app} !important;
    }}
    [data-testid="stChatInput"], [data-testid="stChatInput"] > div {{
        background-color: {input_bg} !important;
        border: 1px solid {border_col} !important;
        border-radius: 12px !important;
    }}
    [data-testid="stChatInput"] textarea {{
        background-color: {input_bg} !important;
        color: {text_main} !important;
        caret-color: {text_main} !important;
    }}
    [data-testid="stChatInput"] textarea::placeholder {{
        color: {text_sub} !important;
    }}
    [data-testid="stChatInput"] button {{
        background-color: transparent !important;
        color: {accent_col} !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==================== 6. SIDEBAR DỌC ====================
with st.sidebar:
    st.markdown(f"<h2 style='text-align:center;margin:0 0 12px 0;font-size:1.35rem;color:{accent_col};'>Chấn Thương Chỉnh Hình</h2>", unsafe_allow_html=True)
    new_dark = st.toggle("Chế độ ban đêm", value=dark_mode)
    if new_dark != dark_mode: st.session_state.settings["dark_mode"] = new_dark; save_json_file(SETTINGS_FILE, st.session_state.settings); st.rerun()

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
            save_active_sessions()
            st.session_state.current_page = label
            st.rerun()

    st.markdown("---")
    st.markdown(f"<p style='text-align:center;font-weight:700;color:{text_sub};margin-bottom:8px;'>KỆ SÁCH ({len(all_books_in_db)})</p>", unsafe_allow_html=True)
    selected_local_books = []
    if all_books_in_db:
        for b_i, name in enumerate(all_books_in_db):
            if st.checkbox(name, value=True, key=f"book_{b_i}"): selected_local_books.append(name)
    else: st.error("Chưa nạp được dữ liệu sách từ Drive.")
    st.session_state["_books"] = selected_local_books

    with st.expander("🛠️ Quản lý & Đồng bộ sách mới (thct)", expanded=False):
        st.caption("1️⃣ **Thêm nhanh sách `thct.pdf` hoặc sách mới:**")
        uploaded_pdfs = st.file_uploader("Chọn file PDF từ máy (ví dụ thct.pdf):", type=["pdf"], accept_multiple_files=True, key="up_new_pdfs")
        if uploaded_pdfs:
            if st.button("📥 Nạp sách PDF vừa chọn vào Kệ sách", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang đọc và lập chỉ mục sách mới..."):
                    for uf in uploaded_pdfs:
                        save_path = os.path.join(PDF_DIR, uf.name)
                        with open(save_path, "wb") as out_f:
                            out_f.write(uf.read())
                        with get_db() as conn:
                            conn.execute("DELETE FROM books WHERE name=?", (uf.name,))
                            conn.commit()
                        index_single_book_fast(uf.name)
                st.success("✅ Đã thêm sách mới thành công!")
                st.rerun()

        st.markdown("---")
        st.caption("2️⃣ **Đồng bộ từ link thư mục Google Drive:**")
        cur_folder = st.session_state.settings.get("gdrive_folder_url", "")
        folder_in = st.text_input("Link thư mục Drive:", value=cur_folder, placeholder="https://drive.google.com/drive/folders/...")
        if st.button("🔄 Đồng bộ thư mục Drive & Làm mới Kệ sách", use_container_width=True):
            st.session_state.settings["gdrive_folder_url"] = folder_in.strip()
            st.session_state.settings["hidden_books"] = []
            save_json_file(SETTINGS_FILE, st.session_state.settings)
            ensure_database_ready(force_reload=True, custom_folder_url=folder_in.strip())
            st.rerun()

        st.markdown("---")
        st.caption("3️⃣ **Xóa tên sách cũ bị trùng:**")
        books_to_del = st.multiselect("Chọn sách cũ muốn xóa:", all_books_in_db, key="del_old_books_sel")
        if st.button("🧹 Xóa các tên sách đã chọn", use_container_width=True) and books_to_del:
            remove_books_from_db(books_to_del)
            hidden_list = set(st.session_state.settings.get("hidden_books", []))
            hidden_list.update(books_to_del)
            st.session_state.settings["hidden_books"] = list(hidden_list)
            save_json_file(SETTINGS_FILE, st.session_state.settings)
            st.success("Đã xóa tên sách cũ!")
            st.rerun()

# ==================== 7. THANH ĐIỀU HƯỚNG NHANH & NÚT MỞ MENU KHI SIDEBAR ĐÓNG ====================
page = st.session_state.current_page

# Thanh công cụ cố định trên đầu màn hình chính: giúp mở lại Sidebar hoặc chuyển tab ngay cả khi đã đóng Sidebar
top_c1, top_c2 = st.columns([1.1, 4.9])
with top_c1:
    if st.button("☰ Mở Menu / Kệ sách", key="btn_force_open_sb", use_container_width=True):
        force_open_sidebar()
        st.rerun()
with top_c2:
    nav_options = ["Hỏi đáp", "Bảng tin", "Trực Cấp cứu", "Kho bài viết", "Tạo thẻ Anki", "X-quang", "USMLE & SRS", "Tạo Slide"]
    quick_nav = st.selectbox(
        "Chuyển nhanh chuyên mục:",
        nav_options,
        index=nav_options.index(page) if page in nav_options else 0,
        label_visibility="collapsed",
        key="quick_top_nav_sel"
    )
    if quick_nav != page:
        save_active_sessions()
        st.session_state.current_page = quick_nav
        st.rerun()

# ==================== 8. NỘI DUNG TỪNG CHUYÊN MỤC ====================
if page == "Hỏi đáp":
    c_opt1, c_opt2, c_opt3 = st.columns([2.2, 1.8, 1])
    with c_opt1: ans_mode = st.radio("Chế độ:", ["Tự động", "Giải thích", "Bài soạn"], horizontal=True, key="qa_mode")
    with c_opt2: st.write(""); use_vision = st.toggle("Đính kèm ảnh cho AI", value=False, key="qa_vis_toggle")
    with c_opt3:
        st.write("")
        if st.button("Xóa phiên", key="qa_clr", use_container_width=True):
            st.session_state.qa_session = None
            save_active_sessions()
            st.rerun()

    query = st.chat_input("Nhập câu hỏi (Nhấn Enter để gửi)...", key="qa_q")
    if query:
        books = st.session_state.get("_books", all_books_in_db)
        if not books: st.warning("Vui lòng chọn ít nhất 1 cuốn sách ở thanh bên trái.")
        else:
            with st.spinner("⏳ Đang tra cứu giáo trình và biên soạn câu trả lời... (Vui lòng chờ trong giây lát)"):
                ctx, matched, vis = smart_route_pages(query, books, max_pages=15)
                q_low = query.lower()
                is_specific = any(k in q_low for k in ["tại sao", "vì sao", "như thế nào", "khi nào", "bao lâu", "chỉ định", "biến chứng", "cơ chế", "là gì", "so sánh", "?"]) or len(q_low.split()) >= 6
                use_7 = (ans_mode == "Bài soạn") or (ans_mode == "Tự động" and not is_specific)

                if not use_7:
                    prompt = f"""TÀI LIỆU GIÁO TRÌNH:\n{ctx}\n\nCÂU HỎI: "{query}"\n\n{PEDAGOGICAL_STYLE_RULES}
TRÌNH BÀY DẠNG LIST CHO TỪNG Ý LỚN DƯỚI ĐÂY (TIÊU ĐỀ LỚN PHẢI IN HOA VÀ TÔ ĐẬM):
**1. TRẢ LỜI TRỰC DIỆN & BẢN CHẤT CƠ CHẾ:**
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
**4. CẬN LÂM SÀNG** (- **X-quang:**... \n - **Siêu âm:**... \n - **MRI:**...)
**5. NGUYÊN TẮC & PHÁC ĐỒ ĐIỀU TRỊ** (- **Nguyên tắc tối thượng:**... \n - **Điều trị bảo tồn:**... \n - **Chỉ định phẫu thuật:**...)
**6. PHỤC HỒI CHỨC NĂNG & CHĂM SÓC SAU MỔ**
**7. TIÊN LƯỢNG, PHÒNG NGỪA & XỬ TRÍ BIẾN CHỨNG**
Cuối bài kèm sơ đồ ```dot ... ```.
"""
                ans_main = stream_ai_response(prompt, vis if use_vision else None)
                st.session_state.qa_session = {
                    "query": query,
                    "routed_context": ctx,
                    "matched_pages": matched,
                    "visual_keys": vis,
                    "main_answer": ans_main,
                    "discussions": [],
                    "saved_id": None
                }
                save_active_sessions()
                log_activity()
                st.rerun()

    sess = st.session_state.qa_session
    if sess:
        st.markdown(f"**Chủ đề đang xem:** {sess['query']}")
        render_graphviz_if_present(sess["main_answer"])
        display_cited_pages(sess["main_answer"], candidate_pages=sess["matched_pages"])
        st.markdown("---")
        c1, c2 = st.columns([2, 2])
        with c1: tag_in = st.text_input("Tags (cách nhau dấu phẩy):", value="hỏi đáp", key="qa_tags", label_visibility="collapsed")
        with c2:
            disc_list = sess.get("discussions", [])
            appendix_str = "\n\n---\n" + "".join([f"\n> Phụ lục ({d['time']}): {d['user_text']}\n{d['ai_reply']}" for d in disc_list]) if disc_list else ""
            full_save_body = sess["main_answer"] + appendix_str
            saved_exists = any(str(it.get("id")) == str(sess.get("saved_id")) for it in st.session_state.library)
            if not sess.get("saved_id") or not saved_exists:
                if st.button("💾 Lưu vào Kho bài viết", key="qa_save", type="primary", use_container_width=True):
                    sess["saved_id"] = save_to_library("Hỏi Đáp", sess["query"][:65], full_save_body, tags=[t.strip() for t in tag_in.split(",") if t.strip()])
                    save_active_sessions()
            else:
                if st.button("✅ Đã lưu vào Kho (Bấm để cập nhật)", key="qa_upd", type="primary", use_container_width=True):
                    sess["saved_id"] = update_library_item(sess["saved_id"], sess["query"][:65], full_save_body, tags=[t.strip() for t in tag_in.split(",") if t.strip()])
                    save_active_sessions()

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
                with st.spinner("⏳ Đang xử lý yêu cầu bổ sung... (Vui lòng chờ)"):
                    extra_ctx, _, ex_vis = smart_route_pages(sess["query"] + " " + follow_q, st.session_state.get("_books", []), max_pages=10)
                    if "chỉnh sửa" in mode.lower():
                        sess["main_answer"] = stream_ai_response(f"TÀI LIỆU:\n{sess.get('routed_context', '')[:18000]}\n{extra_ctx[:12000]}\nCHỦ ĐỀ: {sess['query']}\nBÀI CŨ: {sess['main_answer']}\nYÊU CẦU SỬA: {follow_q}\n{PEDAGOGICAL_STYLE_RULES}\nViết lại toàn bài chi tiết kèm [Trang X].", ex_vis if use_vision else None)
                    else:
                        ans_f = stream_ai_response(f"CHỦ ĐỀ GỐC: {sess['query']}\nTÀI LIỆU:\n{sess.get('routed_context', '')[:18000]}\n{extra_ctx[:10000]}\nBÀI TRƯỚC:\n{sess['main_answer'][:4000]}\nYÊU CẦU: {follow_q}\n{PEDAGOGICAL_STYLE_RULES}\nGiải thích sâu sắc kèm [Trang X].", ex_vis if use_vision else None)
                        sess["discussions"].append({"type": "feedback", "user_text": follow_q, "ai_reply": ans_f, "time": datetime.datetime.now().strftime("%H:%M")})
                    save_active_sessions()
                    st.rerun()

elif page == "Bảng tin":
    if not st.session_state.news_session:
        src = st.selectbox("Chọn nguồn:", ["Từ Giáo trình", "JOT (LWW)", "JBJS (LWW)", "Injury (Elsevier)"])
        if src == "Từ Giáo trình":
            if st.button("🎲 Bốc thăm 4 chuyên đề ngẫu nhiên từ sách", type="primary"):
                books = st.session_state.get("_books", all_books_in_db)
                if books:
                    with st.spinner("⏳ Đang rút ngẫu nhiên chuyên đề từ giáo trình..."):
                        ph_books = ",".join("?" * len(books))
                        with get_db() as conn: sample = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({ph_books}) ORDER BY RANDOM() LIMIT 10", books).fetchall()
                        sample_text = "\n".join([f"[Trang {p}]: {t[:500]}" for _, p, t in sample])
                        res = generate_ai_response(f"Dựa vào các trang sau:\n{sample_text}\nChọn 4 chủ đề Chấn thương Chỉnh hình KHÁC NHAU. Trả về đúng 4 dòng: Tên chủ đề | Mô tả ngắn.")
                        st.session_state.news_headlines = [(p[0].replace('*', '').strip(), p[1].replace('*', '').strip()) for p in [line.split('|') for line in res.text.strip().split('\n') if '|' in line]]
            if st.session_state.news_headlines:
                st.markdown("---")
                for i, (t, tz) in enumerate(st.session_state.news_headlines):
                    with st.container(border=True):
                        st.markdown(f"**{t}**"); st.caption(tz)
                        if st.button("Mở đọc chuyên đề này", key=f"read_{i}"):
                            st.session_state.news_session = {"query": t, "main_answer": "", "saved_id": None}
                            save_active_sessions()
                            st.rerun()
        else:
            urls = {"JOT (LWW)": "https://journals.lww.com/jorthotrauma/_layouts/15/OAKS.Journals/feed.aspx?FeedType=CurrentIssue", "JBJS (LWW)": "https://journals.lww.com/jbjsjournal/_layouts/15/OAKS.Journals/feed.aspx?FeedType=CurrentIssue", "Injury (Elsevier)": "https://rss.sciencedirect.com/publication/science/00201383"}
            if st.button(f"Tải 5 bài báo mới nhất từ {src}", type="primary"):
                with st.spinner(f"⏳ Đang tải bài báo từ {src}..."):
                    try:
                        with urlopen(Request(urls[src], headers={"User-Agent": "Mozilla/5.0"}), timeout=8) as r: root = ET.fromstring(r.read())
                        st.session_state.rss_items = [{"title": (it.findtext("title") or "").strip(), "link": (it.findtext("link") or "").strip()} for it in root.findall(".//item")[:5]]
                    except Exception: st.warning("Không kết nối được nguồn RSS này.")
            for i, it in enumerate(st.session_state.get("rss_items", [])):
                with st.container(border=True):
                    st.markdown(f"**{i+1}. {it['title']}**"); st.markdown(f"[Xem bài gốc]({it['link']})")
                    if st.button("Phân tích chủ đề này", key=f"rss_an_{i}"):
                        st.session_state.news_session = {"query": it["title"], "main_answer": "", "saved_id": None}
                        save_active_sessions()
                        st.rerun()
    else:
        ns = st.session_state.news_session
        if st.button("⬅️ Quay lại Bảng tin", key="back_news_btn"):
            st.session_state.news_session = None
            save_active_sessions()
            st.rerun()
        st.markdown(f"## {ns['query']}")
        if not ns["main_answer"]:
            with st.spinner("⏳ Đang biên soạn bài chuyên đề từ giáo trình... (Vui lòng chờ)"):
                ctx, mp, _ = smart_route_pages(ns["query"], st.session_state.get("_books", all_books_in_db), max_pages=15)
                ans_news = stream_ai_response(f"TÀI LIỆU:\n{ctx}\n{PEDAGOGICAL_STYLE_RULES}\nViết bài 7 phần dạng list chia mục rõ ràng, CÁC TIÊU ĐỀ LỚN PHẢI IN HOA VÀ TÔ ĐẬM về '{ns['query']}' kèm [Trang X] và sơ đồ ```dot ... ```.")
                ns.update({"routed_context": ctx, "matched_pages": mp, "main_answer": ans_news})
                save_active_sessions()
                st.rerun()

        render_graphviz_if_present(ns["main_answer"])
        display_cited_pages(ns["main_answer"], key_prefix="news_cite", candidate_pages=ns.get("matched_pages", []))

        st.markdown("---")
        cn1, cn2 = st.columns([2, 2])
        with cn1:
            news_tags = st.text_input("Tags:", value="bảng tin, chuyên đề", key="news_tag_in", label_visibility="collapsed")
        with cn2:
            saved_exists = any(str(it.get("id")) == str(ns.get("saved_id")) for it in st.session_state.library)
            if not ns.get("saved_id") or not saved_exists:
                if st.button("💾 Lưu bài viết này vào Kho", key="news_save_btn", type="primary", use_container_width=True):
                    ns["saved_id"] = save_to_library("Bảng tin", ns["query"][:65], ns["main_answer"], tags=[t.strip() for t in news_tags.split(",") if t.strip()])
                    save_active_sessions()
            else:
                if st.button("✅ Đã lưu (Bấm để cập nhật lại)", key="news_upd_btn", type="primary", use_container_width=True):
                    ns["saved_id"] = update_library_item(ns["saved_id"], ns["query"][:65], ns["main_answer"], tags=[t.strip() for t in news_tags.split(",") if t.strip()])
                    save_active_sessions()

elif page == "Kho bài viết":
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1.2])
    with c1: kw = st.text_input("🔍 Tìm kiếm bài viết:", key="lib_search")
    with c2: f_type = st.selectbox("Phân loại:", ["Tất cả"] + sorted(set(i.get("type", "Khác") for i in st.session_state.library)))
    with c3: f_tag = st.selectbox("Tag:", ["Tất cả"] + sorted({t for i in st.session_state.library for t in (i.get("tags") or [])}))
    with c4:
        st.write("")
        if st.button("➕ Tạo bài viết trống mới", type="primary", use_container_width=True):
            save_to_library("Tự soạn", "Bài ghi chép mới", "<p>Nhập nội dung ghi chép tại đây...</p>", tags=["ghi chép"])
            st.rerun()

    if not st.session_state.library:
        st.info("Kho lưu trữ hiện đang trống. Hãy lưu bài từ mục Hỏi đáp, Bảng tin hoặc bấm 'Tạo bài viết trống mới'.")
    else:
        for idx, item in enumerate(reversed(st.session_state.library)):
            title = item.get("title", ""); content = item.get("content", ""); tags = item.get("tags") or []
            if kw.lower() and kw.lower() not in title.lower() and kw.lower() not in content.lower(): continue
            if f_type != "Tất cả" and item.get("type") != f_type: continue
            if f_tag != "Tất cả" and f_tag not in tags: continue

            iid = item["id"]
            tag_label = " | Tags: " + ", ".join(tags) if tags else ""
            with st.expander(f"[{item.get('type', 'Bài lưu')}] {title} ({item.get('time', '')}{tag_label})"):
                c_e, c_img, c_d = st.columns([1.4, 1.4, 1])
                with c_e: is_edit = st.toggle("✏️ Chỉnh sửa (Word Editor)", key=f"e_{iid}_{idx}")
                with c_img: show_img = st.toggle("📖 Xem trang gốc", key=f"img_{iid}_{idx}")
                with c_d:
                    if st.button("🗑️ Xóa", key=f"d_{iid}_{idx}", use_container_width=True): delete_from_library(iid); st.rerun()

                if is_edit:
                    st.caption("💡 **Hướng dẫn:** Bôi đen chữ để **In đậm (B)**, **Tô màu Highlight** (biểu tượng chữ A có nền màu). Để **chèn hình ngay tại vị trí con trỏ**, hãy đặt con trỏ chuột vào dòng muốn chèn rồi bấm vào **biểu tượng Bức tranh (Image)** trên thanh công cụ.")
                    new_t = st.text_input("Tiêu đề bài viết:", value=title, key=f"et_{iid}_{idx}")
                    new_tags = st.text_input("Tags (phân cách bằng dấu phẩy):", value=", ".join(tags), key=f"etg_{iid}_{idx}")
                    
                    try:
                        import markdown
                        clean_raw = clean_cite_spans_for_storage(content)
                        is_already_html = any(tag in clean_raw for tag in ["<p>", "<img", "<strong>", "<h1>", "<h2>", "<h3>", "<ul"])
                        html_content = clean_raw if is_already_html else markdown.markdown(clean_raw, extensions=['tables', 'fenced_code'])
                        
                        new_c_html = st_quill(
                            value=html_content,
                            html=True,
                            toolbar=WORD_LIKE_TOOLBAR,
                            key=f"quill_{iid}_{idx}"
                        )
                        if st.button("💾 Lưu thay đổi bài viết", key=f"sv_{iid}_{idx}", type="primary"):
                            saved_body = new_c_html if (new_c_html and len(new_c_html.strip()) > 0) else html_content
                            update_library_item(iid, new_t, saved_body, tags=[t.strip() for t in new_tags.split(",") if t.strip()])
                            st.rerun()
                    except Exception as e:
                        st.warning(f"Đang dùng trình sửa văn bản dự phòng: {e}")
                        new_c = st.text_area("Nội dung:", value=clean_cite_spans_for_storage(content), height=350, key=f"ec_{iid}_{idx}")
                        if st.button("💾 Lưu thay đổi", key=f"sv_fb_{iid}_{idx}", type="primary"):
                            update_library_item(iid, new_t, new_c, tags=[t.strip() for t in new_tags.split(",") if t.strip()])
                            st.rerun()
                else:
                    clean_view = clean_cite_spans_for_storage(content)
                    if any(tag in clean_view for tag in ["<p>", "<img", "<strong>"]):
                        st.markdown(format_small_citations(clean_view), unsafe_allow_html=True)
                    else:
                        render_graphviz_if_present(clean_view)
                    if show_img: display_cited_pages(clean_view, key_prefix=f"lib_{iid}_{idx}")

elif page == "Trực Cấp cứu":
    er_books = get_thct_emergency_books(all_books_in_db, st.session_state.get("_books", all_books_in_db))
    thct_detected = [b for b in all_books_in_db if "thct" in b.lower()]
    if thct_detected:
        st.success(f"📚 **Nguồn sách cấp cứu đang kích hoạt:** `{', '.join(thct_detected)}`")
    else:
        with st.expander("⚠️ Chưa thấy file sách 'thct' trong Kệ sách — Bấm vào đây để nạp nhanh file thct.pdf", expanded=True):
            er_up = st.file_uploader("Chọn file thct.pdf từ máy tính/iPad của bạn:", type=["pdf"], key="er_direct_upload")
            if er_up and st.button("📥 Nạp ngay sách thct vào hệ thống", type="primary"):
                with st.spinner("⏳ Đang nạp sách cấp cứu thct..."):
                    sp = os.path.join(PDF_DIR, er_up.name)
                    with open(sp, "wb") as f_out: f_out.write(er_up.read())
                    index_single_book_fast(er_up.name)
                st.success("✅ Đã nạp xong sách thct!")
                st.rerun()

    topic = st.text_input("Chủ đề ca bệnh (bỏ trống để bốc thăm ngẫu nhiên từ sách thct):", key="er_topic_in")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🚑 Bắt đầu ca trực mới", type="primary", use_container_width=True):
            with st.spinner("⏳ Đang rút tình huống cấp cứu lâm sàng từ sách thct..."):
                if topic.strip():
                    ctx, _, _ = smart_route_pages(topic, er_books, max_pages=14)
                else:
                    ph_b = ",".join("?" * len(er_books))
                    with get_db() as conn:
                        sample = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({ph_b}) ORDER BY RANDOM() LIMIT 12", er_books).fetchall()
                    ctx = "\n\n".join(f"=== SÁCH: {b} | TRANG {p} ===\n{t}" for b, p, t in sample)
                ans_er = stream_ai_response(f"Bạn là BS Trưởng tua trực Chấn thương Chỉnh hình. Dựa sát vào TÀI LIỆU SÁCH CẤP CỨU (thct) dưới đây:\n{ctx[:25000]}\nTạo BƯỚC 1/3: bệnh cảnh lâm sàng cấp cứu thực tế. CHƯA tiết lộ chẩn đoán xác định và X-quang. Đặt 3 câu hỏi thử thách xử trí ban đầu. Cần gạch đầu dòng rõ ràng. Chủ đề: {topic or 'Tình huống cấp cứu từ sách thct'}")
                st.session_state.er_chat = [{"role": "assistant", "content": ans_er}]
                st.session_state.er_case_active = True
                save_active_sessions()
                log_activity()
                st.rerun()
    with c2:
        if st.button("Xóa ca trực", use_container_width=True):
            st.session_state.er_chat = []
            st.session_state.er_case_active = False
            save_active_sessions()
            st.rerun()

    if st.session_state.er_case_active:
        for i, m in enumerate(st.session_state.er_chat):
            with st.chat_message(m["role"]):
                render_graphviz_if_present(m["content"])
                if m["role"] == "assistant": display_cited_pages(m["content"], key_prefix=f"er_{i}")

        u = st.chat_input("Nhập xử trí của bạn (Nhấn Enter để gửi)...")
        if u:
            st.session_state.er_chat.append({"role": "user", "content": u})
            with st.chat_message("user"): st.markdown(u)
            with st.chat_message("assistant"):
                with st.spinner("⏳ BS Trưởng tua đang chấm điểm và đối chiếu sách thct..."):
                    ctx, _, _ = smart_route_pages((topic or "") + " " + u, er_books, max_pages=14)
                    history_log = "\n".join([f"[{m['role']}]: {m['content']}" for m in st.session_state.er_chat])
                    er_prompt = (
                        f"TÀI LIỆU SÁCH CẤP CỨU (thct):\n{ctx[:25000]}\n"
                        f"DIỄN BIẾN:\n{history_log}\n"
                        f"{PEDAGOGICAL_STYLE_RULES}\n"
                        "1. Chấm điểm (thang 100), giải thích chi tiết đúng/sai theo sát phác đồ sách thct kèm [Trang X] bằng dạng list rõ ràng, in hoa tiêu đề.\n"
                        "2. Chuyển bước tiếp theo (Bước 1->2 cung cấp X-quang/CLS; Bước 2->3 hỏi kế hoạch phẫu thuật/bảo tồn & PHCN; xong Bước 3 tổng kết + sơ đồ ```dot```)."
                    )
                    reply_er = stream_ai_response(er_prompt)
            st.session_state.er_chat.append({"role": "assistant", "content": reply_er})
            save_active_sessions()
            st.rerun()

        ch, cs = st.columns(2)
        with ch:
            if st.button("💡 Xin gợi ý", use_container_width=True):
                last_msg = st.session_state.er_chat[-1]["content"]
                st.info(generate_ai_response(f"Đưa ra 3 gợi ý ngắn gọn gạch đầu dòng cho tình huống:\n{last_msg}").text)
        with cs:
            if st.button("💾 Lưu ca trực vào Kho", use_container_width=True):
                er_full_text = "\n\n---\n\n".join([f"**{m['role'].upper()}:**\n{m['content']}" for m in st.session_state.er_chat])
                save_to_library("Trực Cấp Cứu", topic or "Ca lâm sàng cấp cứu (thct)", er_full_text, tags=["case", "ER", "thct"])

elif page == "Tạo thẻ Anki":
    topic = st.chat_input("Nhập chủ đề cần tạo 10 thẻ Flashcard (Nhấn Enter để tạo)...")
    if topic:
        with st.spinner("⏳ Đang biên soạn 10 thẻ Flashcard Anki... (Vui lòng chờ)"):
            ctx, _, _ = smart_route_pages(topic, st.session_state.get("_books", all_books_in_db), max_pages=14)
            raw_anki = stream_ai_response(f"Tạo 10 Flashcard (Hỏi - Đáp giải thích sâu cơ chế giải phẫu/sinh cơ học dạng gạch đầu dòng kèm [Trang X]) về: {topic}.\nPhần 1: Hiển thị đọc trên web.\nPhần 2: Đặt trong khối ```csv\nMặt trước ; Mặt sau\n``` để import vào Anki.\nNguồn:\n{ctx}")
            csv_m = re.search(r'```csv\n(.*?)\n```', raw_anki, re.DOTALL)
            st.session_state.anki_session = {"topic": topic, "display": re.sub(r'```csv\n.*?\n```', '', raw_anki, flags=re.DOTALL), "csv": csv_m.group(1) if csv_m else None}
            save_active_sessions()
            log_activity()
            st.rerun()

    ak = st.session_state.anki_session
    if ak:
        c1, c2 = st.columns(2)
        with c1:
            if ak.get("csv"): st.download_button("📥 Tải file .CSV Import vào Anki", data=ak["csv"].encode("utf-8-sig"), file_name=f"anki_{datetime.date.today()}.csv", mime="text/csv", use_container_width=True)
        with c2:
            if st.button("💾 Lưu bộ thẻ vào Kho bài viết", use_container_width=True): save_to_library("Anki", ak["topic"], ak["display"], tags=["anki"])
        st.markdown("---"); st.markdown(format_small_citations(ak["display"]), unsafe_allow_html=True)

elif page == "X-quang":
    cl, cr = st.columns([1, 1])
    img_bytes = None
    with cl:
        f = st.file_uploader("Chọn ảnh X-quang / CT / MRI / Giải phẫu (PNG, JPG):", type=["png", "jpg", "jpeg"])
        if f: img_bytes = f.read(); st.image(img_bytes, use_container_width=True)
    with cr:
        mode = st.selectbox("Chế độ phân tích:", ["Đọc phim X-quang / CT-scan, Các góc đo & Phân loại", "Phân tích Đường mổ, Gian cơ & Thần kinh nguy cơ", "Đố vui nhận diện mốc Giải phẫu (Active Recall)"])
        note = st.text_input("Câu hỏi hoặc vùng cần chú ý trên hình:")
        if st.button("🔍 Phân tích hình ảnh", type="primary", use_container_width=True):
            if not img_bytes: st.warning("Vui lòng tải lên 1 hình ảnh trước.")
            else:
                with st.spinner("⏳ Đang phân tích hình ảnh bằng AI Vision..."):
                    part = types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                    if "Đố vui" in mode: p = f"Chọn 5 mốc giải phẫu/dụng cụ trên hình ({note}), đặt 5 câu hỏi nhận diện gạch đầu dòng. Đặt đáp án chi tiết sau dòng '--- ĐÁP ÁN ---'."
                    elif "Đường mổ" in mode: p = f"Phân tích chi tiết đường mổ trên hình ({note}): liệt kê từng lớp giải phẫu, khoảng gian thần kinh và cấu trúc mạch máu/thần kinh nguy cơ cao bằng gạch đầu dòng."
                    else: p = f"Đọc phim X-quang/giải thích hình ảnh ({note}) thật chi tiết: gạch đầu dòng mô tả tổn thương, cơ chế di lệch, các góc đo giải phẫu và hướng xử trí."
                    res = generate_ai_response([part, p])
                    st.session_state.xray_session = {"title": note or mode[:35], "content": res.text}; log_activity()

        xs = st.session_state.xray_session
        if xs:
            if "--- ĐÁP ÁN ---" in xs["content"]:
                a, b = xs["content"].split("--- ĐÁP ÁN ---", 1)
                st.markdown(a)
                with st.expander("Xem Đáp án"): st.markdown(b)
            else: st.markdown(xs["content"])
            if st.button("💾 Lưu bài phân tích vào Kho", type="primary"): save_to_library("X-quang", xs["title"], xs["content"], tags=["xray"])

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

    due_list = get_due_mistakes()
    t1, t2, t3 = st.tabs(["Làm Đề USMLE", f"Sổ Tay Câu Sai ({len(st.session_state.mistakes)})", f"Ôn SRS Hôm Nay ({len(due_list)})"])

    with t1:
        c1, c2 = st.columns([3, 1])
        with c1: topic = st.text_input("Chủ đề Quiz (Để trống = ngẫu nhiên):")
        with c2: n_q = st.number_input("Số câu:", 1, 10, 4)

        if st.button("📝 Tạo đề trắc nghiệm", type="primary"):
            books = st.session_state.get("_books", all_books_in_db)
            if not books: st.warning("Vui lòng chọn sách ở thanh bên trái.")
            else:
                with st.spinner("⏳ Đang thiết kế câu hỏi tình huống lâm sàng USMLE..."):
                    if not topic.strip():
                        ph_books = ",".join("?" * len(books))
                        with get_db() as conn: sample = conn.execute(f"SELECT book, page, text FROM pages WHERE book IN ({ph_books}) ORDER BY RANDOM() LIMIT 10", books).fetchall()
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
                        if (choice or "")[:1].upper() == qit["answer"][0].upper(): st.success(f"Chính xác! ({qit['answer'][0].upper()})")
                        else: st.error(f"Sai. Bạn chọn {(choice or '—')[:1].upper()} | Đáp án đúng: {qit['answer'][0].upper()}")
                        st.markdown(format_small_citations(qit["explanation"]), unsafe_allow_html=True)

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
                    st.markdown(format_small_citations(f"**Đáp án đúng: {m['answer']}**\n\n{m.get('explanation', '')}"), unsafe_allow_html=True)
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
                        st.markdown(format_small_citations(f"**Đáp án: {m['answer']}**\n\n{m.get('explanation', '')}"), unsafe_allow_html=True)
                        c1, c2, c3 = st.columns(3)
                        for col, (lbl, q_val) in zip([c1, c2, c3], [("Quên (0)", 0), ("Khó (3)", 3), ("Dễ (5)", 5)]):
                            with col:
                                if st.button(lbl, key=f"srs_{m['id']}_{q_val}", use_container_width=True): sm2_update(m, q_val); save_json_file(MISTAKES_FILE, st.session_state.mistakes); st.rerun()

elif page == "Tạo Slide":
    topic = st.chat_input("Nhập chủ đề Slide báo cáo (Nhấn Enter để tạo)...")
    if topic:
        with st.spinner("⏳ Đang thiết kế cấu trúc 8 Slide thuyết trình và xuất file PowerPoint..."):
            ctx, _, _ = smart_route_pages(topic, st.session_state.get("_books", all_books_in_db), max_pages=15)
            raw_slide = stream_ai_response(f"""Thiết kế 8 slide thuyết trình chi tiết về "{topic}". Mỗi slide đúng định dạng:
### SLIDE [số]: [Tiêu đề]
- Ý 1 (Diễn giải chi tiết 3-4 câu)
- Ý 2 (Diễn giải chi tiết 3-4 câu)
- Ý 3 (Diễn giải chi tiết 3-4 câu)
**Ghi chú thuyết trình:** (Giải thích thật sâu cơ chế giải phẫu, sinh cơ học dạng list kèm [Trang X])
Nguồn: {ctx}""")
            st.session_state.slide_session = {"topic": topic, "text": raw_slide, "pptx": create_pptx_bytes(topic, raw_slide)}
            save_active_sessions()
            log_activity()
            st.rerun()

    ss = st.session_state.slide_session
    if ss:
        st.markdown("### Chỉnh sửa Outline trước khi xuất file .PPTX")
        edited = st.text_area("Nội dung Outline (Có thể sửa trực tiếp):", value=ss["text"], height=320)
        if edited != ss["text"]:
            ss["text"] = edited
            ss["pptx"] = create_pptx_bytes(ss["topic"], ss["text"])
            save_active_sessions()

        c1, c2 = st.columns(2)
        with c1:
            safe_fname = re.sub(r'[^A-Za-z0-9_]', '_', ss["topic"][:25])
            if ss.get("pptx"): st.download_button("📥 Tải file PowerPoint (.PPTX)", data=ss["pptx"], file_name=f"Slide_{safe_fname}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
        with c2:
            if st.button("💾 Lưu Outline vào Kho bài viết", use_container_width=True): save_to_library("Slide", ss["topic"], ss["text"], tags=["slide"])
