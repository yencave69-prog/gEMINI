import os, sqlite3, datetime
import fitz  # PyMuPDF

PDF_DIR = "kho_sach"
DB_FILE = os.path.join(PDF_DIR, "ortho_fts5.db")

# 1. Xóa file db cũ nếu còn sót
for ext in ["", "-wal", "-shm"]:
    p = DB_FILE + ext
    if os.path.exists(p):
        os.remove(p)

# 2. Tạo bảng FTS5 mới tinh
conn = sqlite3.connect(DB_FILE)
conn.execute("CREATE VIRTUAL TABLE pages USING fts5(book, page UNINDEXED, text, tokenize='unicode61 remove_diacritics 2')")
conn.execute("CREATE TABLE books (name TEXT PRIMARY KEY, pages INTEGER, indexed_at TEXT)")

# 3. Quét tất cả các file PDF hiện có trong kho_sach
pdfs = sorted([f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")])
print(f"Tìm thấy {len(pdfs)} cuốn sách PDF. Bắt đầu tạo ortho_fts5.db mới...\n")

for idx_b, fname in enumerate(pdfs, 1):
    fpath = os.path.join(PDF_DIR, fname)
    print(f"[{idx_b}/{len(pdfs)}] Đang nạp sách: {fname} ...")
    try:
        doc = fitz.open(fpath)
        rows = [(fname, i + 1, doc.load_page(i).get_text("text")) for i in range(len(doc))]
        doc.close()
        conn.executemany("INSERT INTO pages(book, page, text) VALUES (?,?,?)", rows)
        conn.execute("INSERT INTO books(name, pages, indexed_at) VALUES (?,?,?)", (fname, len(rows), datetime.datetime.now().isoformat()))
        conn.commit()
        print(f"   -> Hoàn tất ({len(rows)} trang)")
    except Exception as e:
        print(f"   -> Lỗi đọc file {fname}: {e}")

# 4. Nén tối ưu dung lượng file db
print("\nĐang tối ưu hóa dung lượng file DB...")
conn.execute("VACUUM")
conn.close()
print(f"\n=== ĐÃ TẠO XONG FILE MỚI TẠI: {DB_FILE} ===")
input("Nhấn Enter để thoát...")
