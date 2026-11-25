# DocuFree Studio

A free, open-access document scanner, image-to-PDF and Word (`.docx`) converter, PDF utility tool, and Groq-powered text assistant built with **Python** and **Streamlit**.Open-access CamScanner-style web app built with Python, Streamlit, Pillow, python-docx, pypdf, and Groq API—featuring user accounts, camera/photo scanning, PDF &amp; Word conversion, PDF tools, and AI text cleanup/summarization/translation.

## Features
- **Account System:** Create an account, sign in, or use Guest mode with zero feature locks or paywalls.
- **Scan & Upload:** Capture photos via camera or upload multiple images (`JPG`, `PNG`, `WEBP`) with scan enhancement filters and rotation.
- **PDF & Word Export:** Convert scanned/uploaded pages into a multi-page `.pdf` or a `.docx` Word document (with optional text/notes).
- **PDF Tools:** Merge multiple PDFs or add a diagonal watermark to a PDF.
- **Groq AI Text Assistant:** Clean up text, summarize key points, or translate text using Groq's free-tier `openai/gpt-oss-20b` model.

## Run Locally
```bash
pip install -r requirements.txt
streamlit run app.py
