import hashlib
import io
import os
import sqlite3
from datetime import datetime
from typing import List, Tuple

import streamlit as st
from groq import Groq
from PIL import Image, ImageEnhance, ImageOps
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from docx import Document

# ---------------------------------------------------------------------------
# Page Configuration & Styling
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="DocuFree Studio — Scanner, PDF & Word Maker",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .stApp {
        background-color: #0b0f14;
        color: #f1f5f9;
    }
    .hero-banner {
        background: linear-gradient(135deg, #111827 0%, #0f2926 100%);
        border: 1px solid #1f2937;
        border-left: 4px solid #10b981;
        border-radius: 14px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1.2rem;
    }
    .tool-card {
        background-color: #131a24;
        border: 1px solid #1e293b;
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 0.75rem;
    }
    .badge-free {
        display: inline-block;
        background-color: rgba(16, 185, 129, 0.15);
        color: #34d399;
        border: 1px solid rgba(16, 185, 129, 0.4);
        padding: 0.15rem 0.6rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

DB_PATH = "docufree_users.db"
FREE_GROQ_MODEL = "openai/gpt-oss-20b"


# ---------------------------------------------------------------------------
# Database & Authentication Helpers
# ---------------------------------------------------------------------------
def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def hash_password(password: str, salt: str = "docufree_salt_v1") -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000
    )
    return dk.hex()


def register_user(username: str, password: str) -> Tuple[bool, str]:
    clean_user = username.strip().lower()
    if len(clean_user) < 3:
        return False, "Username must be at least 3 characters long."
    if len(password) < 4:
        return False, "Password must be at least 4 characters long."

    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
            (clean_user, hash_password(password), datetime.utcnow().isoformat()),
        )
        conn.commit()
        conn.close()
        return True, f"Account created for '{clean_user}'. You can log in now."
    except sqlite3.IntegrityError:
        return False, "That username already exists. Please choose another."
    except Exception as exc:
        return False, f"Database error: {exc}"


def authenticate_user(username: str, password: str) -> Tuple[bool, str]:
    clean_user = username.strip().lower()
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute(
            "SELECT password_hash FROM users WHERE username = ?", (clean_user,)
        )
        row = cur.fetchone()
        conn.close()
        if row and row[0] == hash_password(password):
            return True, clean_user
        return False, ""
    except Exception:
        return False, ""


# ---------------------------------------------------------------------------
# Image Processing, PDF & Word Builders
# ---------------------------------------------------------------------------
def enhance_scan_image(
    img: Image.Image,
    mode: str = "Document Scan (Crisp)",
    rotate_deg: int = 0,
) -> Image.Image:
    work = img.convert("RGB")
    if rotate_deg in (90, 180, 270):
        work = work.rotate(-rotate_deg, expand=True)

    if mode == "Document Scan (Crisp)":
        gray = ImageOps.grayscale(work)
        contrasted = ImageEnhance.Contrast(gray).enhance(1.55)
        sharpened = ImageEnhance.Sharpness(contrasted).enhance(1.4)
        return sharpened.convert("RGB")
    if mode == "High Contrast B&W":
        gray = ImageOps.grayscale(work)
        contrasted = ImageEnhance.Contrast(gray).enhance(2.2)
        bw = contrasted.point(lambda x: 255 if x > 145 else 0, mode="1")
        return bw.convert("RGB")
    if mode == "Bright & Clean":
        bright = ImageEnhance.Brightness(work).enhance(1.12)
        return ImageEnhance.Contrast(bright).enhance(1.25)

    return work


def images_to_pdf_bytes(images: List[Image.Image], quality: int = 85) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w, page_h = A4
    margin = 24
    usable_w = page_w - 2 * margin
    usable_h = page_h - 2 * margin

    for img in images:
        rgb = img.convert("RGB")
        img_buf = io.BytesIO()
        rgb.save(img_buf, format="JPEG", quality=quality, optimize=True)
        img_buf.seek(0)

        iw, ih = rgb.size
        scale = min(usable_w / max(iw, 1), usable_h / max(ih, 1))
        draw_w = iw * scale
        draw_h = ih * scale
        x = (page_w - draw_w) / 2
        y = (page_h - draw_h) / 2

        c.drawImage(ImageReader(img_buf), x, y, width=draw_w, height=draw_h)
        c.showPage()

    c.save()
    buf.seek(0)
    return buf.getvalue()


def content_to_docx_bytes(
    title: str,
    images: List[Image.Image],
    text_body: str = "",
) -> bytes:
    doc = Document()
    doc.add_heading(title.strip() or "Scanned Document", level=1)

    if text_body.strip():
        doc.add_heading("Document Text", level=2)
        for line in text_body.splitlines():
            doc.add_paragraph(line)

    if images:
        doc.add_heading("Pages / Scans", level=2)
        for idx, img in enumerate(images, start=1):
            rgb = img.convert("RGB")
            # Resize wide images to fit comfortably on a Word page
            max_px = 1400
            if max(rgb.size) > max_px:
                rgb.thumbnail((max_px, max_px))
            img_stream = io.BytesIO()
            rgb.save(img_stream, format="PNG")
            img_stream.seek(0)

            doc.add_paragraph(f"Page {idx}")
            from docx.shared import Inches

            doc.add_picture(img_stream, width=Inches(6.0))

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


def merge_pdfs_bytes(pdf_files: List[bytes]) -> bytes:
    writer = PdfWriter()
    for raw in pdf_files:
        reader = PdfReader(io.BytesIO(raw))
        for page in reader.pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()


def watermark_pdf_bytes(pdf_raw: bytes, watermark_text: str) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_raw))
    writer = PdfWriter()

    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        wm_buf = io.BytesIO()
        c = canvas.Canvas(wm_buf, pagesize=(w, h))
        c.saveState()
        c.setFont("Helvetica-Bold", 28)
        c.setFillGray(0.75, alpha=0.35)
        c.translate(w / 2, h / 2)
        c.rotate(30)
        c.drawCentredString(0, 0, watermark_text[:50])
        c.restoreState()
        c.save()
        wm_buf.seek(0)

        wm_reader = PdfReader(wm_buf)
        page.merge_page(wm_reader.pages[0])
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Groq Free-Tier Helper
# ---------------------------------------------------------------------------
def get_groq_api_key() -> str:
    try:
        if "GROQ_API_KEY" in st.secrets and st.secrets["GROQ_API_KEY"]:
            return str(st.secrets["GROQ_API_KEY"]).strip()
    except Exception:
        pass
    return os.getenv("GROQ_API_KEY", "").strip()


def run_groq_text_assistant(prompt_text: str, task: str, target_lang: str = "English") -> str:
    api_key = get_groq_api_key()
    if not api_key:
        return (
            "Groq API key not found. Add `GROQ_API_KEY` in Streamlit Secrets or "
            "environment variables to enable AI text cleanup, summarization, and translation."
        )

    system_map = {
        "Clean & Format Document Text": (
            "You are a document formatting assistant. Clean up spacing, punctuation, "
            "and line breaks while preserving every fact and number accurately."
        ),
        "Summarize Key Points": (
            "Summarize the provided document text clearly with concise bullet points "
            "and key takeaways."
        ),
        "Translate Text": (
            f"Translate the following text accurately into {target_lang}. "
            "Keep formatting clean."
        ),
    }

    try:
        client = Groq(api_key=api_key)
        completion = client.chat.completions.create(
            model=FREE_GROQ_MODEL,
            messages=[
                {"role": "system", "content": system_map.get(task, system_map["Clean & Format Document Text"])},
                {"role": "user", "content": prompt_text},
            ],
            temperature=0.2,
            max_tokens=1500,
        )
        return completion.choices[0].message.content or "No response returned."
    except Exception as exc:
        return (
            f"Could not complete AI request right now (free-tier rate limit or network issue): {exc}"
        )


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
init_db()

if "user" not in st.session_state:
    st.session_state.user = "Guest"
if "scanned_pages" not in st.session_state:
    st.session_state.scanned_pages = []
if "ai_output" not in st.session_state:
    st.session_state.ai_output = ""

# Sidebar: Account & Workspace Controls
with st.sidebar:
    st.markdown("### 👤 Account & Access")
    st.markdown('<span class="badge-free">100% FREE • NO LOCKS</span>', unsafe_allow_html=True)
    st.write("")

    if st.session_state.user != "Guest":
        st.success(f"Signed in as **{st.session_state.user}**")
        if st.button("Log Out", use_container_width=True):
            st.session_state.user = "Guest"
            st.rerun()
    else:
        st.info("Browsing as **Guest** (all tools unlocked)")
        auth_tab1, auth_tab2 = st.tabs(["Log In", "Sign Up"])
        with auth_tab1:
            login_u = st.text_input("Username", key="login_user")
            login_p = st.text_input("Password", type="password", key="login_pass")
            if st.button("Log In", use_container_width=True):
                ok, uname = authenticate_user(login_u, login_p)
                if ok:
                    st.session_state.user = uname
                    st.rerun()
                else:
                    st.error("Invalid username or password.")
        with auth_tab2:
            reg_u = st.text_input("Choose Username", key="reg_user")
            reg_p = st.text_input("Choose Password", type="password", key="reg_pass")
            if st.button("Create Account", use_container_width=True):
                ok, msg = register_user(reg_u, reg_p)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

    st.divider()
    st.markdown("### 🗂️ Active Scan Queue")
    st.caption(f"Pages in queue: **{len(st.session_state.scanned_pages)}**")
    if st.session_state.scanned_pages and st.button("Clear Queue", use_container_width=True):
        st.session_state.scanned_pages = []
        st.rerun()

# Hero Header
st.markdown(
    """
    <div class="hero-banner">
        <h2 style="margin:0;">📄 DocuFree Studio</h2>
        <p style="margin:0.35rem 0 0 0; color:#cbd5e1;">
            Free camera scanner, image-to-PDF & Word (.docx) maker, PDF merger/watermarker,
            and Groq AI text assistant — with zero paywalls.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_scan, tab_export, tab_pdf_tools, tab_ai = st.tabs(
    [
        "📷 Scan & Upload",
        "📥 Export PDF & Word",
        "🛠️ PDF Tools",
        "✨ AI Text Assistant",
    ]
)

# ---------------------------------------------------------------------------
# Tab 1: Camera Scanner & Image Uploader
# ---------------------------------------------------------------------------
with tab_scan:
    st.subheader("Scan Documents or Upload Photos")
    col_left, col_right = st.columns([1.15, 0.85], gap="large")

    with col_left:
        scan_mode = st.selectbox(
            "Scan Enhancement Filter",
            ["Document Scan (Crisp)", "High Contrast B&W", "Bright & Clean", "Original"],
        )
        rotate_deg = st.selectbox("Rotate Image", [0, 90, 180, 270])

        source_tab1, source_tab2 = st.tabs(["📸 Live Camera", "🖼️ Upload Photos"])

        with source_tab1:
            cam_photo = st.camera_input("Take a photo of your document or page")
            if cam_photo is not None:
                raw_img = Image.open(cam_photo)
                processed = enhance_scan_image(raw_img, scan_mode, rotate_deg)
                st.image(processed, caption="Preview Scan", use_container_width=True)
                if st.button("➕ Add Captured Page to Queue", type="primary"):
                    st.session_state.scanned_pages.append(processed)
                    st.success("Page added to your document queue!")

        with source_tab2:
            uploaded_imgs = st.file_uploader(
                "Upload one or more photos (JPG, PNG, WEBP)",
                type=["jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
            )
            if uploaded_imgs:
                preview_list = [
                    enhance_scan_image(Image.open(f), scan_mode, rotate_deg)
                    for f in uploaded_imgs
                ]
                st.image(
                    preview_list[0],
                    caption=f"Previewing Page 1 of {len(preview_list)}",
                    use_container_width=True,
                )
                if st.button(f"➕ Add All {len(preview_list)} Photo(s) to Queue", type="primary"):
                    st.session_state.scanned_pages.extend(preview_list)
                    st.success(f"Added {len(preview_list)} page(s) to queue!")

    with col_right:
        st.markdown("#### Queued Pages")
        pages = st.session_state.scanned_pages
        if not pages:
            st.info("No pages in queue yet. Capture a photo or upload images on the left.")
        else:
            for idx, pg in enumerate(pages):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.image(pg, caption=f"Page {idx + 1}", use_container_width=True)
                with c2:
                    if st.button("Remove", key=f"del_pg_{idx}"):
                        st.session_state.scanned_pages.pop(idx)
                        st.rerun()

# ---------------------------------------------------------------------------
# Tab 2: Convert to PDF & Word (.docx)
# ---------------------------------------------------------------------------
with tab_export:
    st.subheader("Convert Queued Scans & Text into PDF or Word")
    doc_title = st.text_input("Document Title", value="My_Scanned_Document")
    include_text = st.text_area(
        "Optional Text / Notes to Include in Word (.docx)",
        value=st.session_state.ai_output,
        height=140,
        placeholder="Paste text, notes, or AI-cleaned text here...",
    )

    c_pdf, c_docx = st.columns(2, gap="large")

    with c_pdf:
        st.markdown('<div class="tool-card">', unsafe_allow_html=True)
        st.markdown("#### 📕 Export as PDF")
        pdf_quality = st.slider("JPEG Compression Quality", min_value=50, max_value=100, value=85)
        if st.session_state.scanned_pages:
            pdf_bytes = images_to_pdf_bytes(st.session_state.scanned_pages, quality=pdf_quality)
            st.download_button(
                label="⬇️ Download PDF",
                data=pdf_bytes,
                file_name=f"{doc_title.strip() or 'document'}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
            )
        else:
            st.warning("Add at least one scanned/uploaded page in **Scan & Upload** to create a PDF.")
        st.markdown("</div>", unsafe_allow_html=True)

    with c_docx:
        st.markdown('<div class="tool-card">', unsafe_allow_html=True)
        st.markdown("#### 📘 Export as Word (.docx)")
        if st.session_state.scanned_pages or include_text.strip():
            docx_bytes = content_to_docx_bytes(
                title=doc_title,
                images=st.session_state.scanned_pages,
                text_body=include_text,
            )
            st.download_button(
                label="⬇️ Download Word (.docx)",
                data=docx_bytes,
                file_name=f"{doc_title.strip() or 'document'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
            )
        else:
            st.warning("Add scanned pages or enter text above to generate a Word document.")
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Tab 3: PDF Utilities (Merge & Watermark)
# ---------------------------------------------------------------------------
with tab_pdf_tools:
    st.subheader("PDF Utilities")
    u1, u2 = st.columns(2, gap="large")

    with u1:
        st.markdown('<div class="tool-card">', unsafe_allow_html=True)
        st.markdown("#### 🔗 Merge Multiple PDFs")
        pdfs_to_merge = st.file_uploader(
            "Select 2 or more PDF files",
            type=["pdf"],
            accept_multiple_files=True,
            key="merge_uploader",
        )
        if pdfs_to_merge and len(pdfs_to_merge) >= 2:
            merged_data = merge_pdfs_bytes([f.getvalue() for f in pdfs_to_merge])
            st.download_button(
                label="⬇️ Download Merged PDF",
                data=merged_data,
                file_name="merged_document.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        elif pdfs_to_merge:
            st.caption("Upload at least 2 PDFs to merge.")
        st.markdown("</div>", unsafe_allow_html=True)

    with u2:
        st.markdown('<div class="tool-card">', unsafe_allow_html=True)
        st.markdown("#### 💧 Add Watermark to PDF")
        wm_pdf = st.file_uploader("Upload a PDF file", type=["pdf"], key="wm_uploader")
        wm_text = st.text_input("Watermark Text", value="CONFIDENTIAL / DRAFT")
        if wm_pdf and wm_text.strip():
            watermarked_data = watermark_pdf_bytes(wm_pdf.getvalue(), wm_text.strip())
            st.download_button(
                label="⬇️ Download Watermarked PDF",
                data=watermarked_data,
                file_name="watermarked_document.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Tab 4: Groq Free-Tier AI Text Assistant
# ---------------------------------------------------------------------------
with tab_ai:
    st.subheader("AI Document Text Assistant (Groq Free Tier)")
    st.caption(f"Powered by Groq model: `{FREE_GROQ_MODEL}`")

    ai_task = st.selectbox(
        "Choose AI Task",
        ["Clean & Format Document Text", "Summarize Key Points", "Translate Text"],
    )
    target_lang = "English"
    if ai_task == "Translate Text":
        target_lang = st.text_input("Target Language", value="Urdu")

    source_text = st.text_area(
        "Paste Document Text or Notes",
        height=180,
        placeholder="Paste typed text, lecture notes, or extracted text here...",
    )

    if st.button("✨ Run AI Assistant", type="primary"):
        if not source_text.strip():
            st.warning("Please enter some text first.")
        else:
            with st.spinner("Processing with Groq..."):
                result = run_groq_text_assistant(source_text, ai_task, target_lang)
                st.session_state.ai_output = result

    if st.session_state.ai_output:
        st.markdown("#### AI Result")
        st.text_area("Output (also pre-filled in Export Word tab)", value=st.session_state.ai_output, height=180)
