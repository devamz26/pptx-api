from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
import os, uuid, io, requests

app = FastAPI()

# -----------------------
# ROOT HEALTH CHECK ROUTE
# -----------------------
@app.get("/")
def health():
    return {"status": "ok", "endpoints": ["/docs", "/pptx/create"]}

# -----------------------
# DATA MODELS
# -----------------------
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

# Directory to store generated PPTX files
FILES_DIR = "generated"
os.makedirs(FILES_DIR, exist_ok=True)

# -----------------------
# IMAGE FETCH HELPER
# -----------------------
def fetch_image_bytes(url: str) -> io.BytesIO:
    headers = {"User-Agent": "Mozilla/5.0 (pptx-generator/1.0)"}
    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()

    ctype = r.headers.get("Content-Type", "").lower()
    if not any(x in ctype for x in ["image/jpeg", "image/png", "image/gif", "application/octet-stream"]):
        raise ValueError(f"Unsupported image content-type: {ctype}")

    return io.BytesIO(r.content)

# -----------------------
# FOOTER HELPER
# -----------------------
def add_footer(prs: Presentation, text: str):
    for slide in prs.slides:
        tx = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.3))
        run = tx.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(10)

# -----------------------
# PPTX BUILDER
# -----------------------
def build_pptx(payload: CreatePptxInput, output_path: str):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    bullet_layout = prs.slide_layouts[1]

    # Title Slide
    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = payload.title
    if payload.subtitle:
        slide.placeholders[1].text = payload.subtitle

    # Content slides
    for s in payload.slides:
        slide = prs.slides.add_slide(bullet_layout)
        slide.shapes.title.text = s.heading[:255]

        # Bullets
        body = slide.shapes.placeholders[1].text_frame
        body.clear()
        if s.bullets:
            body.text = s.bullets[0]
            for bullet in s.bullets[1:]:
                p = body.add_paragraph()
                p.text = bullet
                p.level = 0

               # Images
        if s.images:
            top = Inches(2.8)
            max_width = Inches(6.5)

            for img in s.images:
                try:
                    stream = fetch_image_bytes(str(img.url))

                    # Choose picture sizing
                    if img.width_inch and img.height_inch:
                        pic = slide.shapes.add_picture(
                            stream,
                            Inches(0.5),
                            top,
                            width=Inches(img.width_inch),
                            height=Inches(img.height_inch)
                        )
                    elif img.width_inch:
                        pic = slide.shapes.add_picture(
                            stream,
                            Inches(0.5),
                            top,
                            width=Inches(img.width_inch)
                        )
                    elif img.height_inch:
                        pic = slide.shapes.add_picture(
                            stream,
                            Inches(0.5),
                            top,
                            height=Inches(img.height_inch)
                        )
                    else:
                        # Default — scale to width 6.5"
                        pic = slide.shapes.add_picture(
                            stream,
                            Inches(1),
                            top,
                            width=max_width
                        )

                    # Center the image horizontally
                    pic.left = int((prs.slide_width - pic.width) / 2)

                    # Optional caption
                    if img.caption:
                        cap = slide.shapes.add_textbox(
                            pic.left,
                            pic.top + pic.height + Inches(0.1),
                            pic.width,
                            Inches(0.4)
                        )
                        cap_tf = cap.text_frame
                        cap_tf.text = img.caption
                        cap_tf.paragraphs[0].runs[0].font.size = Pt(12)

                    # Move top down for next image
                    top = pic.top + pic.height + Inches(0.4)

                except Exception as e:
                    p = body.add_paragraph()
                    p.text = f"[Image failed: {img.url}]"
                    p.level = 0

