from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
import os, uuid, io, requests

app = FastAPI()

# ---------------------------------
# Root Health Check Route
# ---------------------------------
@app.get("/")
def health():
    return {"status": "ok", "endpoints": ["/docs", "/pptx/create"]}


# ---------------------------------
# Data Models
# ---------------------------------
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


FILES_DIR = "generated"
os.makedirs(FILES_DIR, exist_ok=True)


# ---------------------------------
# Robust Image Fetcher
# ---------------------------------
def fetch_image_bytes(url: str) -> io.BytesIO:
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; PPTX-Generator/1.1; +https://pptx-api)",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    }

    r = requests.get(url, headers=headers, timeout=20, allow_redirects=True)
    r.raise_for_status()

    content_type = (r.headers.get("Content-Type") or "").lower()

    supported = ("image/png", "image/jpeg", "image/jpg", "image/gif")

    # Accept PNG/JPG/GIF, reject SVG/WebP because python-pptx cannot embed them
    if not any(t in content_type for t in supported):
        if not url.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
            raise ValueError(f"Unsupported image type: {content_type}")

    return io.BytesIO(r.content)


# ---------------------------------
# PPTX Builder
# ---------------------------------
def build_pptx(payload: CreatePptxInput, output_path: str):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]

    # Title Slide
    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = payload.title
    if payload.subtitle:
        slide.placeholders[1].text = payload.subtitle

    # Content Slides
    for s in payload.slides:
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = s.heading[:255]

        body = slide.shapes.placeholders[1].text_frame
        body.clear()

        # Bullets
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

                    # Auto-fitting image
                    pic = slide.shapes.add_picture(
                        stream, Inches(1), top, width=max_width
                    )

                    # Center image
                    pic.left = int((prs.slide_width - pic.width) / 2)

                    # Caption
                    if img.caption:
                        cap = slide.shapes.add_textbox(
                            pic.left, pic.top + pic.height + Inches(0.1),
                            pic.width, Inches(0.4)
                        )
                        cap.text_frame.text = img.caption[:140]
                        cap.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

                    top = pic.top + pic.height + Inches(0.4)

                except Exception as e:
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


# ---------------------------------
# API Routes
# ---------------------------------
@app.post("/pptx/create")
async def create_pptx(payload: CreatePptxInput):
    file_id = uuid.uuid4().hex
    filename = f"{file_id}.pptx"
    path = os.path.join(FILES_DIR, filename)

    build_pptx(payload, path)

    base_url = os.getenv("PUBLIC_BASE_URL", "https://pptx-api-8eqj.onrender.com")

    return JSONResponse({
        "download_url": f"{base_url}/files/{filename}",
        "file_name": filename
    })


@app.get("/files/{filename}")
async def serve_file(filename: str):
    path = os.path.join(FILES_DIR, filename)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename
    )
