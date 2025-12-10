from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, HttpUrl
from typing import List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
import os, uuid, io, requests

app = FastAPI()
FILES_DIR = "generated"
os.makedirs(FILES_DIR, exist_ok=True)

# ---- Models ----
class ImageItem(BaseModel):
    url: HttpUrl
    width_inch: Optional[float] = None   # optional width override
    height_inch: Optional[float] = None  # optional height override
    caption: Optional[str] = None

class SlideItem(BaseModel):
    heading: str
    bullets: Optional[List[str]] = []
    images: Optional[List[ImageItem]] = []  # <-- NEW

class CreatePptxInput(BaseModel):
    title: str
    subtitle: Optional[str] = None
    slides: List[SlideItem]
    theme: Optional[str] = "simple"
    footer: Optional[str] = None

# ---- Helpers ----
def fetch_image_bytes(url: str) -> io.BytesIO:
    # Some hosts block default python UA; set a browsery UA and short timeout
    headers = {"User-Agent": "Mozilla/5.0 (pptx-generator/1.0)"}
    r = requests.get(url, headers=headers, timeout=15)
    r.raise_for_status()
    # Basic content-type check
    ct = r.headers.get("Content-Type", "").lower()
    if not any(x in ct for x in ["image/jpeg", "image/png", "image/gif", "application/octet-stream"]):
        # allow octet-stream (common for direct file hosts)
        raise ValueError(f"Unsupported image content-type: {ct}")
    return io.BytesIO(r.content)

def add_footer(prs: Presentation, text: str):
    for slide in prs.slides:
        tx = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(9), Inches(0.3))
        run = tx.text_frame.paragraphs[0].add_run()
        run.text = text[:120]
        run.font.size = Pt(10)

# ---- Core builder ----
def build_pptx(payload: CreatePptxInput, output_path: str):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]   # Title slide
    bullet_layout = prs.slide_layouts[1]  # Title + Content

    # Title slide
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
            body.text = s.bullets[0][:1000]
            for b in s.bullets[1:]:
                p = body.add_paragraph()
                p.text = b[:1000]
                p.level = 0

        # Images (embedded)
        # Default placement: below the text placeholder, centered, width ~6.5"
        # Adjust if width/height provided.
        if s.images:
            top = Inches(2.8)  # roughly below text area
            left_default = Inches(1.0)
            max_width = Inches(6.5)
            for img in s.images:
                try:
                    stream = fetch_image_bytes(str(img.url))
                    # Compute size rules
                    if img.width_inch and img.height_inch:
                        pic = slide.shapes.add_picture(stream, Inches(0.5), top,
                                                       width=Inches(img.width_inch),
                                                       height=Inches(img.height_inch))
                    elif img.width_inch:
                        pic = slide.shapes.add_picture(stream, Inches(0.5), top,
                                                       width=Inches(img.width_inch))
                    elif img.height_inch:
                        pic = slide.shapes.add_picture(stream, Inches(0.5), top,
                                                       height=Inches(img.height_inch))
                    else:
                        # Fit to max_width, keep aspect by specifying width only
                        pic = slide.shapes.add_picture(stream, left_default, top, width=max_width)
                    # Center image horizontally if not using explicit left
                    pic.left = int((prs.slide_width - pic.width) / 2)

                    # Optional caption
                    if img.caption:
                        cap_box = slide.shapes.add_textbox(pic.left, pic.top + pic.height + Inches(0.1),
                                                           pic.width, Inches(0.4))
                        cap_tf = cap_box.text_frame
                        cap_tf.text = img.caption[:140]
                        cap_tf.paragraphs[0].runs[0].font.size = Pt(12)

                    # Stack subsequent images under previous one
                    top = pic.top + pic.height + Inches(0.4)

                except Exception as e:
                    # Add a note if an image fails
                    p = body.add_paragraph()
                    p.text = f"[Image failed to embed: {img.url}]"
                    p.level = 0

    if payload.footer:
        add_footer(prs, payload.footer)

    prs.save(output_path)

# ---- Routes ----
@app.post("/pptx/create")
async def create_pptx(payload: CreatePptxInput):
    file_id = uuid.uuid4().hex
    file_name = f"{file_id}.pptx"
    local_path = os.path.join(FILES_DIR, file_name)
    build_pptx(payload, local_path)
    base = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")
    return JSONResponse({"download_url": f"{base}/files/{file_name}", "file_name": file_name})

@app.get("/files/{file_name}")
async def serve_file(file_name: str):
    path = os.path.join(FILES_DIR, file_name)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=file_name
    )
