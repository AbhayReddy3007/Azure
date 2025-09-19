import os, re, tempfile, base64, datetime
import fitz, docx, requests
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
from PIL import Image
import streamlit as st
from vertexai.generative_models import GenerativeModel
import vertexai

# ---------------- CONFIG ----------------
PROJECT_ID = "drl-zenai-prod"  
REGION = "us-central1"
vertexai.init(project=PROJECT_ID, location=REGION)

TEXT_MODEL_NAME = "gemini-2.0-flash"
TEXT_MODEL = GenerativeModel(TEXT_MODEL_NAME)


# ---------------- HELPERS ----------------
def call_vertex(prompt: str) -> str:
    response = TEXT_MODEL.generate_content(prompt)
    return response.text.strip()

def extract_text(path: str, filename: str) -> str:
    name = filename.lower()
    if name.endswith(".pdf"):
        text_parts = []
        doc = fitz.open(path)
        for page in doc:
            text_parts.append(page.get_text("text"))
        doc.close()
        return "\n".join(text_parts)
    if name.endswith(".docx"):
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    if name.endswith(".txt"):
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return ""

def split_text(text: str, chunk_size: int = 8000, overlap: int = 300):
    if not text:
        return []
    chunks, start, n = [], 0, len(text)
    while start < n:
        end = min(start + chunk_size, n)
        chunks.append(text[start:end])
        if end == n:
            break
        start = max(0, end - overlap)
    return chunks

def summarize_long_text(full_text: str) -> str:
    chunks = split_text(full_text)
    if len(chunks) <= 1:
        return call_vertex(f"Summarize:\n\n{full_text}")
    partial_summaries = []
    for idx, ch in enumerate(chunks, start=1):
        mapped = call_vertex(f"Summarize this part:\n\n{ch}")
        partial_summaries.append(f"Chunk {idx}:\n{mapped.strip()}")
    combined = "\n\n".join(partial_summaries)
    return call_vertex(f"Combine these summaries:\n\n{combined}")

def generate_title(summary: str) -> str:
    prompt = f"""Read the summary and create a clean, short title (<10 words):
Summary:
{summary}"""
    return call_vertex(prompt).strip()

def parse_points(points_text: str):
    points = []
    current_title, current_content = None, []
    lines = [re.sub(r"[#*>`]", "", ln).rstrip() for ln in points_text.splitlines()]
    for line in lines:
        if not line:
            continue
        m = re.match(r"^\s*(Slide|Section)\s*(\d+)\s*:\s*(.+)$", line, re.IGNORECASE)
        if m:
            if current_title:
                points.append({"title": current_title, "description": "\n".join(current_content)})
            current_title, current_content = m.group(3).strip(), []
            continue
        if line.strip().startswith(("-", "•", "*")):
            text = line.lstrip("-•*").strip()
            if text:
                current_content.append(f"• {text}")
        else:
            if line.strip():
                current_content.append(line.strip())
    if current_title:
        points.append({"title": current_title, "description": "\n".join(current_content)})
    return points

def generate_outline_from_desc(description: str):
    prompt = f"""Create a PowerPoint outline on: {description}.
Each slide should have:
- Title
- 3–4 bullet points
No title slide.
Format:
Slide 1: <Title>
- Bullet
- Bullet
"""
    return parse_points(call_vertex(prompt))

# ---------------- PPT GENERATOR ----------------
def clean_title_text(title: str) -> str:
    return re.sub(r"\s+", " ", title.strip()) if title else "Presentation"

def resize_image(image_path, max_width=800, max_height=600):
    try:
        img = Image.open(image_path)
        img.thumbnail((max_width, max_height))
        resized_path = image_path.replace(".png", "_resized.png")
        img.save(resized_path, "PNG")
        return resized_path
    except Exception:
        return image_path

def create_ppt(title, points, filename="output.pptx", images=None):
    prs = Presentation()
    PRIMARY_PURPLE = RGBColor(94, 42, 132)
    SECONDARY_TEAL = RGBColor(0, 185, 163)
    TEXT_DARK = RGBColor(40, 40, 40)
    BG_LIGHT = RGBColor(244, 244, 244)
    title = clean_title_text(title)

    # Title slide
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = PRIMARY_PURPLE
    textbox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
    tf = textbox.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

    # Content slides
    for idx, item in enumerate(points, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        bg_color = BG_LIGHT if idx % 2 == 0 else RGBColor(255, 255, 255)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        key_point = clean_title_text(item.get("title", ""))
        description = item.get("description", "")

        textbox = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8), Inches(1.5))
        tf = textbox.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        p = tf.add_paragraph()
        p.text = key_point
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = PRIMARY_PURPLE
        p.alignment = PP_ALIGN.LEFT

        # Accent underline
        shape = slide.shapes.add_shape(1, Inches(0.8), Inches(1.6), Inches(3), Inches(0.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = SECONDARY_TEAL

        if description:
            textbox = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(5), Inches(4))
            tf = textbox.text_frame
            tf.word_wrap = True
            for line in description.split("\n"):
                if line.strip():
                    bullet = tf.add_paragraph()
                    bullet.text = line.strip()
                    bullet.font.size = Pt(22)
                    bullet.font.color.rgb = TEXT_DARK

        textbox = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(8), Inches(0.3))
        tf = textbox.text_frame
        p = tf.add_paragraph()
        p.text = "Generated with AI"
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(150, 150, 150)
        p.alignment = PP_ALIGN.RIGHT

    prs.save(filename)
    return filename


# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="AI Productivity Suite", layout="wide")
st.title("AI Productivity Suite")

uploaded_file = st.file_uploader("📂 Upload a document", type=["pdf", "docx", "txt"])
if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    text = extract_text(tmp_path, uploaded_file.name)
    os.remove(tmp_path)

    if text:
        summary = summarize_long_text(text)
        title = generate_title(summary)
        st.success(f"✅ Document uploaded! Suggested title: **{title}**")

        if st.button("📑 Generate PPT Outline"):
            outline = generate_outline_from_desc(summary)
            st.subheader(f"Outline: {title}")
            for i, slide in enumerate(outline, start=1):
                st.markdown(f"**Slide {i}: {slide['title']}**")
                st.markdown(slide["description"].replace("\n", "\n\n"))

            if st.button("🎯 Generate PPT"):
                filename = f"{re.sub(r'[^A-Za-z0-9_.-]', '_', title)}.pptx"
                ppt_path = create_ppt(title, outline, filename)
                with open(ppt_path, "rb") as f:
                    st.download_button("⬇️ Download PPT", f, file_name=filename)
    else:
        st.error("❌ Could not read file content.")
