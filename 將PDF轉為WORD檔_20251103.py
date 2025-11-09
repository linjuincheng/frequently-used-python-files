import os
import sys
import shutil
import tempfile
from pathlib import Path

INPUT_DIR = r"C:\_python 2024\2024final"
OUTPUT_SUBDIR = "converted"
DPI = 300  # 影像保底法解析度

# ---- 依賴檢查 ----
MISSING = []
try:
    import win32com.client as win32
    from win32com.client import constants as wd_const
except Exception:
    win32 = None
    wd_const = None
    MISSING.append("pywin32")

try:
    from pdf2docx import Converter as PDF2DOCX_Converter
except Exception:
    PDF2DOCX_Converter = None
    MISSING.append("pdf2docx")

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
    MISSING.append("pymupdf")

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
except Exception:
    Document = None
    WD_ORIENT = None
    MISSING.append("python-docx")

if MISSING:
    print("缺少套件：", ", ".join(MISSING))
    print("請先執行：pip install " + " ".join(MISSING))
    sys.exit(1)

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def convert_with_word(in_pdf: Path, out_docx: Path) -> bool:
    """用 Word 的 PDF Reflow 轉 DOCX。成功回 True。"""
    if win32 is None:
        return False
    word = None
    doc = None
    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # No alerts
        # 若 PDF 被保護視圖擋住，可用 ProtectedViewOpen，再 .Edit 進行
        try:
            doc = word.Documents.Open(str(in_pdf))  # 直接開
        except Exception:
            pv = word.ProtectedViewWindows.Open(str(in_pdf))
            doc = pv.Edit()

        doc.SaveAs(str(out_docx), FileFormat=wd_const.wdFormatXMLDocument)  # 12
        return out_docx.exists() and out_docx.stat().st_size > 0
    except Exception as e:
        # 失敗時清理殘檔
        if out_docx.exists():
            try: out_docx.unlink()
            except Exception: pass
        return False
    finally:
        try:
            if doc: doc.Close(False)
        except Exception:
            pass
        try:
            if word: word.Quit()
        except Exception:
            pass

def convert_with_pdf2docx(in_pdf: Path, out_docx: Path) -> bool:
    if PDF2DOCX_Converter is None:
        return False
    try:
        cv = PDF2DOCX_Converter(str(in_pdf))
        cv.convert(str(out_docx))
        cv.close()
        return out_docx.exists() and out_docx.stat().st_size > 0
    except Exception:
        if out_docx.exists():
            try: out_docx.unlink()
            except Exception: pass
        return False

def image_docx_fallback(in_pdf: Path, out_docx: Path, dpi=DPI) -> bool:
    """保底：每頁渲染成圖片貼進 Word，幾乎 1:1 不跑版。"""
    if fitz is None or Document is None:
        return False
    try:
        docx = Document()
        section = docx.sections[0]
        if WD_ORIENT:
            section.orientation = WD_ORIENT.PORTRAIT
        page_width = section.page_width - section.left_margin - section.right_margin

        with fitz.open(str(in_pdf)) as pdf:
            if pdf.page_count == 0:
                return False
            for i in range(pdf.page_count):
                if i > 0:
                    docx.add_page_break()
                page = pdf.load_page(i)
                zoom = dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                    tf.write(pix.tobytes("png"))
                    tmp = Path(tf.name)
                docx.add_picture(str(tmp), width=page_width)
                try: tmp.unlink()
                except Exception: pass

        docx.save(str(out_docx))
        return True
    except Exception:
        if out_docx.exists():
            try: out_docx.unlink()
            except Exception: pass
        return False

def convert_one(pdf_path: Path, out_dir: Path):
    out_docx = out_dir / (pdf_path.stem + ".docx")

    # (1) 先試 Word（最適合你當前環境）
    if convert_with_word(pdf_path, out_docx):
        print(f"[OK-WRD] {pdf_path.name} → {out_docx.name}")
        return

    # (2) 再試 pdf2docx（可編輯）
    if convert_with_pdf2docx(pdf_path, out_docx):
        print(f"[OK-P2D] {pdf_path.name} → {out_docx.name}")
        return

    # (3) 最後保底：整頁影像貼進 Word（不會跑版、圖片不會消失）
    if image_docx_fallback(pdf_path, out_docx):
        print(f"[OK-IMG] {pdf_path.name} → {out_docx.name}（影像保底）")
        return

    print(f"[FAIL ] {pdf_path.name}（所有方法失敗）")

def main():
    in_dir = Path(INPUT_DIR)
    out_dir = in_dir / OUTPUT_SUBDIR
    ensure_dir(out_dir)

    pdfs = [p for p in in_dir.iterdir() if p.suffix.lower() == ".pdf"]
    if not pdfs:
        print(f"找不到 PDF：{in_dir}")
        return

    print(f"輸入：{in_dir}\n輸出：{out_dir}\n共 {len(pdfs)} 個 PDF\n")
    for p in pdfs:
        convert_one(p, out_dir)

if __name__ == "__main__":
    main()

