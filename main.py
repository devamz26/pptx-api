from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
import os, uuid, io, requests

# Image conversion helpers
from PIL import Image

# Optional SVG support
try:
    import cairosvg
    HAS_CAIROSVG = True
except Exception:
    HAS_CAIROSVG = False

app = FastAPI()

# -------------------------
# Health route
# -------------------------
@app.get("/")
def health():
    return {"status": "ok", "endpoints": ["/docs", "/pptx/create"]}

# -------------------------
# Models
# -------------------------
class ImageItem(BaseModel):
    url: HttpUrl
    width_inch: Optional[float] = None
    height_inch: Optional[float] = None
    caption: Optional[str] = None

class SlideItem(BaseModel):
    heading: str
    bullets: Optional[List[str]] = []
    images: Optional[List[ImageItem]] = []

class CreatePptxInput(BaseModel):
    title: str
    subtitle: Optional[str] = None
    slides: List[SlideItem]
    footer: Optional[str] = None

# -------------------------
# Storage
# -------------------------
FILES_DIR = "generated"
os.makedirs(FILES_DIR, exist_ok=True)

# -------------------------
# Robust image fetcher + converter
# -------------------------
def fetch_image_bytes(url: str) -> io.BytesIO:
    """
    Download the image at `url` and return an in-memory file-like object
    that python-pptx can embed (PNG/JPEG/GIF).
    - Converts WebP -> PNG using Pillow.
    - Converts SVG -> PNG using cairosvg if available.
    - Follows redirects and uses a browser User-Agent.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; PPTX-Generator/1.2)",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    }
    r = requests.get(url, headers=headers, timeout=25, allow_redirects=True)
    r.raise_for_status()

    ct = (r.headers.get("Content-Type") or "").lower()
    url_l = url.lower()

    # SVG handling
    if ("image/svg" in ct) or url_l.endswith(".svg"):
        if not HAS_CAIROSVG:
            raise ValueError("SVG image found but cairosvg not installed on server.")
        # Convert SVG bytes to PNG bytes
        png_bytes = cairosvg.svg2png(bytestring=r.content)
        return io.BytesIO(png_bytes)

    # WebP handling -> convert to PNG
    if ("image/webp" in ct) or url_l.endswith(".webp"):
        img = Image.open(io.BytesIO(r.content)).convert("RGBA")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

    # PNG/JPEG/GIF passthrough
    if any(t in ct for t in ("image/png", "image/jpeg", "image/jpg", "image/gif")):
        return io.BytesIO(r.content)

    # Fallback by extension
    if url_l.endswith((".png", ".jpg", ".jpeg", ".gif")):
        return io.BytesIO(r.content)

    raise ValueError(f"Unsupported image type: {ct or 'unknown'}")

# -------------------------
# PPTX builder
# -------------------------
def build_pptx(payload: CreatePptxInput, output_path: str):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]

    # Title slide
    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = payload.title
    if payload.subtitle:
        try:
            slide.placeholders[1].text = payload.subtitle
        except Exception:
            pass

    # Content slides
    for s in payload.slides:
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = s.heading[:255]

        # Bullets
        body = slide.shapes.placeholders[1].text_frame
        body.clear()
        if s.bullets:
            body.text = s.bullets[0][:1000]
            for b in s.bullets[1:]:
                p = body.add_paragraph()
                p.text = b[:1000]
                p.level = 0

        # Images
        if s.images:
            top = Inches(2.8)
            max_width = Inches(6.5)
            for img in s.images:
                try:
                    stream = fetch_image_bytes(str(img.url))

                    # Add picture, auto-fit to width
                    if img.width_inch and img.height_inch:
                        pic = slide.shapes.add_picture(
                            stream, Inches(0.5), top,
                            width=Inches(img.width_inch), height=Inches(img.height_inch)
                        )
                    elif img.width_inch:
                        pic = slide.shapes.add_picture(
                            stream, Inches(0.5), top, width=Inches(img.width_inch)
                        )
                    elif img.height_inch:
                        pic = slide.shapes.add_picture(
                            stream, Inches(0.5), top, height=Inches(img.height_inch)
                        )
                    else:
                        pic = slide.shapes.add_picture(stream, Inches(1), top, width=max_width)

                    # center horizontally
                    pic.left = int((prs.slide_width - pic.width) / 2)

                    # caption
                    if img.caption:
                        cap = slide.shapes.add_textbox(pic.left, pic.top + pic.height + Inches(0.08), pic.width, Inches(0.36))
                        cap_tf = cap.text_frame
                        cap_tf.text = img.caption[:200]
                        try:
                            cap_tf.paragraphs[0].runs[0].font.size = Pt(12)
                        except Exception:
                            pass

                    top = pic.top + pic.height + Inches(0.24)

                except Exception as e:
                    # Put the error message into the slide to see why it failed
                    p = body.add_paragraph()
                    p.text = f"[Image failed: {img.url} — {str(e)}]"
                    p.level = 0

    # Footer
    if payload.footer:
        for slide in prs.slides:
            tx = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.3))
            run = tx.text_frame.paragraphs[0].add_run()
            run.text = payload.footer[:120]
            run.font.size = Pt(10)

    prs.save(output_path)

# -------------------------
# Routes
# -------------------------
@app.post("/pptx/create")
async def create_pptx(payload: CreatePptxInput):
    file_id = uuid.uuid4().hex
    filename = f"{file_id}.pptx"
    path = os.path.join(FILES_DIR, filename)

    build_pptx(payload, path)

    base_url = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")
    return JSONResponse({"download_url": f"{base_url}/files/{filename}", "file_name": filename})

@app.get("/files/{filename}")
async def serve_file(filename: str):
    path = os.path.join(FILES_DIR, filename)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=filename)
