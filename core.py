import os
import uuid
import tempfile
import subprocess
import base64
import requests
import json
from io import BytesIO
from datetime import datetime
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx_formatter import apply_guidelines
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
import urllib3
import time
from tenacity import retry, stop_after_attempt, wait_exponential, before_sleep_log
import logging
import anthropic

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app")

urllib3.disable_warnings()

class SlideData(BaseModel):
    title: str = Field(description="The main title for the slide")
    narrative: str = Field(description="1-2 line explanatory narrative under the title, setting up the slide's argument.", default="")
    punchline: str = Field(description="One takeaway line; unique per slide, placed at the bottom.")
    key_takeaway: str = Field(description="A single powerful sentence summarizing the strategic impact or core takeaway of the slide.", default="Strategic growth driver.")
    layout_type: str = Field(description="One of: title_and_content, two_column, diagram")
    slide_archetype: str = Field(description="Must be one of: title, agenda, divider, standard, table, deep_dive, roadmap, options", default="standard")
    bullet_points: list[str] = Field(description="Maximum 3-5 bullet points. The main content extracted and summarized.")
    table_data: list[list[str]] = Field(description="2D array of strings for table/matrix slides. First row is headers.", default=[])
    icon_keyword: str = Field(description="A single keyword to search for an icon representing the slide's intent")
    keep_original_image: bool = Field(description="Set to true if the original image contains important visual data like a chart, diagram, or photo that should be kept on the slide.")

class PresentationData(BaseModel):
    slides: list[SlideData] = Field(description="List of generated slides")

class SectionData(BaseModel):
    heading: str = Field(description="Section heading")
    content: str = Field(description="Multiple paragraphs of text for this section")
    bullet_points: list[str] = Field(description="Optional bullet points for this section")

class DocumentData(BaseModel):
    title: str = Field(description="Document title")
    sections: list[SectionData] = Field(description="Document sections")

THEMES = {
    "dark corporate": {
        "bg": (28, 30, 38),          # Deep slate grey background
        "accent": (0, 161, 241),    # Vivid blue accent
        "text": (245, 245, 245),    # Off-white text
        "subtext": (170, 175, 185)  # Dimmed text
    },
    "modern light": {
        "bg": (250, 252, 255),      # Crisp very-light blue/white
        "accent": (230, 57, 70),    # Strong red accent
        "text": (33, 37, 41),       # Almost black text
        "subtext": (108, 117, 125)  # Grey subtext
    },
    "pastel": {
        "bg": (245, 241, 237),      # Soft cream
        "accent": (148, 201, 169),  # Mint green
        "text": (73, 80, 87),       # Soft dark grey text
        "subtext": (140, 145, 150)  # Lighter grey
    },
    "blue accent": {
        "bg": (255, 255, 255),      # Pure white
        "accent": (0, 80, 158),     # Navy blue
        "text": (10, 25, 47),       # Very dark blue text
        "subtext": (80, 90, 110)    # Mid blue-grey
    }
}

def _get_theme_colors(theme_str: str):
    t = theme_str.lower()
    if "dark" in t: return THEMES["dark corporate"]
    elif "pastel" in t: return THEMES["pastel"]
    elif "blue" in t: return THEMES["blue accent"]
    return THEMES["modern light"]

def _create_themed_presentation(theme_str: str):
    """
    Creates a new Presentation() and injects a robust, stylish layout template 
    into the master slide so every slide inherits a consistent, beautiful design.
    """
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    
    colors = _get_theme_colors(theme_str)
    bg_color = RGBColor(*colors["bg"])
    accent_color = RGBColor(*colors["accent"])
    
    # Apply to Slide Master so it inherits everywhere
    master = prs.slide_master
    
    # 1. Set solid background
    master.background.fill.solid()
    master.background.fill.fore_color.rgb = bg_color
    
    return prs, colors

def _apply_theme_ribbons(slide, prs, colors):
    """Adds the thematic ribbons to an individual slide since python-pptx doesn't support adding shapes to masters."""
    from pptx.enum.shapes import MSO_SHAPE
    accent_color = RGBColor(*colors["accent"])
    
    top_ribbon = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        0, 0, prs.slide_width, Inches(0.15)
    )
    top_ribbon.fill.solid()
    top_ribbon.fill.fore_color.rgb = accent_color
    top_ribbon.line.fill.background()
    
    bottom_ribbon = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        0, prs.slide_height - Inches(0.3), prs.slide_width, Inches(0.3)
    )
    bottom_ribbon.fill.solid()
    bottom_ribbon.fill.fore_color.rgb = accent_color
    bottom_ribbon.line.fill.background()

def _apply_aptos_narrow(shape, font_color=None):
    if not hasattr(shape, 'text_frame'):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Aptos Narrow'
            if font_color:
                run.font.color.rgb = font_color

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True
)
def _call_genai_with_retry(client, pil_img, prompt_text):
    try:
        return client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[pil_img, prompt_text],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=SlideData,
                temperature=0.2
            ),
        )
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower() or "rate" in str(e).lower():
            logger.warning(f"GenAI rate limit hit (429/Quota): {e}")
        else:
            logger.warning(f"GenAI API error: {e}")
        raise e

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True
)
def _call_anthropic_with_retry(client, b64_img, prompt_text):
    try:
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=1024,
            temperature=0.2,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": b64_img
                            }
                        },
                        {
                            "type": "text",
                            "text": prompt_text + "\n\nRespond ONLY with a valid JSON object matching the requested schema."
                        }
                    ]
                }
            ]
        )
        return response.content[0].text
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower() or "rate" in str(e).lower():
            logger.warning(f"Anthropic rate limit hit (429/Quota): {e}")
        else:
            logger.warning(f"Anthropic API error: {e}")
        raise e

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True
)
def _call_genai_text_with_retry(client, prompt_text, schema):
    try:
        return client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt_text],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=schema,
                temperature=0.4
            ),
        )
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower() or "rate" in str(e).lower():
            logger.warning(f"GenAI text rate limit hit: {e}")
        raise e

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    reraise=True
)
def _call_anthropic_text_with_retry(client, prompt_text, schema):
    try:
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=4000,
            temperature=0.4,
            messages=[
                {
                    "role": "user",
                    "content": prompt_text + f"\n\nRespond ONLY with a valid JSON object matching the requested schema:\n{schema.schema_json()}"
                }
            ]
        )
        return response.content[0].text
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower() or "rate" in str(e).lower():
            logger.warning(f"Anthropic text rate limit hit: {e}")
        raise e

def _fit_image_to_slide(slide, img_path, slide_width, slide_height, margin):
    img = Image.open(img_path)
    img_width, img_height = img.size
    page_aspect = img_width / img_height
    slide_aspect = (slide_width - 2 * margin) / (slide_height - 2 * margin)
    
    if page_aspect > slide_aspect:
        width = slide_width - 2 * margin
        height = width / page_aspect
    else:
        height = slide_height - 2 * margin
        width = height * page_aspect
        
    left = (slide_width - width) / 2 + margin
    top = (slide_height - height) / 2 + margin
    
    slide.shapes.add_picture(img_path, left, top, width, height)

def _trigger_webhook(webhook_url: str, payload: dict):
    """
    Push the generated artifact info to the specified webhook URL.
    """
    if not webhook_url:
        return
    try:
        # Avoid hanging the generation response by using a short timeout for webhook
        requests.post(webhook_url, json=payload, timeout=10)
    except Exception as e:
        print(f"Failed to trigger webhook at {webhook_url}: {e}")

def format_document(doc_source: str, is_url: bool = True, webhook_url: str = None) -> dict:
    """
    Downloads/Reads a DOCX file, applies corporate guidelines, and returns the formatted DOCX URL.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    input_filename = "input.docx"
    input_path = os.path.join(run_dir, input_filename)
    output_filename = "formatted_document.docx"
    output_path = os.path.join(run_dir, output_filename)
    
    try:
        if is_url:
            if doc_source.startswith(('http://', 'https://')):
                response = requests.get(doc_source, verify=False)
                response.raise_for_status()
                with open(input_path, 'wb') as f:
                    f.write(response.content)
            else:
                if ";base64," in doc_source:
                    _, b64_data = doc_source.split(";base64,")
                    with open(input_path, 'wb') as f:
                        f.write(base64.b64decode(b64_data))
                else:
                    import shutil
                    shutil.copy2(doc_source, input_path)
        else:
            with open(input_path, 'wb') as f:
                f.write(base64.b64decode(doc_source))
                
        # Apply formatting guidelines
        apply_guidelines(input_path, output_path)
        
        file_url = _get_file_url(execution_id, output_filename)
        _record_success(file_url, output_filename)
        return {
            "success": True,
            "message": "Document formatted successfully.",
            "file_url": file_url,
            "execution_id": execution_id,
            "filename": output_filename
        }
        
    except Exception as e:
        stats["failed_creations"] += 1
        return {
            "success": False,
            "message": f"Error formatting document: {str(e)}"
        }

def _send_progress(webhook_url, message, status="in_progress"):
    if not webhook_url:
        return
    try:
        requests.post(webhook_url, json={"status": status, "message": message}, verify=False, timeout=5)
    except:
        pass

def process_pdf_to_artifacts(
    pdf_source: str, 
    is_url: bool = True, 
    instructions: str = "", 
    layout_theme: str = "", 
    visual_iconography: str = "", 
    slide_content_rules: str = "",
    target_format: str = "pptx",
    webhook_url: str = None,
    api_key: str = ""
) -> dict:
    """
    Converts a PDF into a PPTX or DOCX, incorporating custom guidelines.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    input_filename = "input.pdf"
    input_path = os.path.join(run_dir, input_filename)
    
    try:
        # Load PDF
        if is_url:
            if pdf_source.startswith(('http://', 'https://')):
                response = requests.get(pdf_source, verify=False)
                response.raise_for_status()
                with open(input_path, 'wb') as f:
                    f.write(response.content)
            else:
                if ";base64," in pdf_source:
                    _, b64_data = pdf_source.split(";base64,")
                    with open(input_path, 'wb') as f:
                        f.write(base64.b64decode(b64_data))
                else:
                    import shutil
                    shutil.copy2(pdf_source, input_path)
        else:
            with open(input_path, 'wb') as f:
                f.write(base64.b64decode(pdf_source))
                
        doc = fitz.open(input_path)
        try:
            if target_format.lower() == "pptx":
                output_filename = "converted_presentation.pptx"
                output_path = os.path.join(run_dir, output_filename)
                
                prs, theme_colors = _create_themed_presentation(layout_theme)
                
                # Add an instructions slide to pass metadata or indicate rules
                if instructions or layout_theme or visual_iconography or slide_content_rules:
                    slide = prs.slides.add_slide(prs.slide_layouts[1]) # Title and Content
                    slide.shapes.title.text = "Generated PPTX Guidelines applied"
                    tf = slide.placeholders[1].text_frame
                    tf.text = "The following guidelines were requested for this presentation:\n"
                    if layout_theme: tf.add_paragraph().text = f"- Theme: {layout_theme}"
                    if visual_iconography: tf.add_paragraph().text = f"- Iconography: {visual_iconography}"
                    if slide_content_rules: tf.add_paragraph().text = f"- Content Rules: {slide_content_rules}"
                    if instructions: tf.add_paragraph().text = f"- Instructions: {instructions}"
                    
                blank_layout = prs.slide_layouts[6]
                
                # Determine which AI client to use
                has_genai = False
                has_anthropic = False
                client = None
                
                # Check for explicit key or environment key
                use_anthropic = False
                if api_key.startswith("sk-ant") or (not api_key and os.environ.get("ANTHROPIC_API_KEY") and not os.environ.get("GEMINI_API_KEY") and not os.environ.get("GOOGLE_API_KEY")):
                    use_anthropic = True
                    
                if use_anthropic:
                    try:
                        k = api_key if api_key else os.environ.get("ANTHROPIC_API_KEY")
                        proxy_url = os.environ.get("GCP_PROXY_FOR_CLAUD")
                        if proxy_url:
                            client = anthropic.Anthropic(api_key=k, base_url=proxy_url, max_retries=0)
                        else:
                            client = anthropic.Anthropic(api_key=k, max_retries=0)
                        has_anthropic = True
                    except:
                        pass
                else:
                    try:
                        if api_key:
                            client = genai.Client(api_key=api_key)
                        elif os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"):
                            client = genai.Client()
                        if client:
                            has_genai = True
                    except:
                        pass
                        
                ai_rate_limit_fallback_count = 0

                for page_num in range(len(doc)):
                    _send_progress(webhook_url, f"Processing page {page_num + 1} of {len(doc)}...")
                    page = doc[page_num]
                    # Generate a high-res image for AI to analyze
                    mat = fitz.Matrix(3.0, 3.0)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    img_path = os.path.join(run_dir, f"page_{page_num}.png")
                    pix.save(img_path)
                    
                    if has_genai or has_anthropic:
                        try:
                            prompt_text = f"""Analyze this presentation slide image thoroughly. 
                            1. Extract all text content and structure it logically.
                            2. Suggest a new layout type (choose strictly from: title_and_content, two_column, diagram).
                            3. Provide an overarching punchline.
                            4. Provide a single keyword for a visual icon that represents the core idea.
                            5. Determine if the original image contains complex diagrams, charts, or essential visual data that MUST be kept on the slide (set keep_original_image to true if so).
                            
                            User Instructions: {instructions}
                            Layout Theme: {layout_theme}
                            Content Rules: {slide_content_rules}"""
                            
                            try:
                                if has_anthropic:
                                    with open(img_path, "rb") as image_file:
                                        b64_img = base64.b64encode(image_file.read()).decode('utf-8')
                                    prompt_text += f"\n\nJSON SCHEMA:\n{SlideData.schema_json()}"
                                    raw_text = _call_anthropic_with_retry(client, b64_img, prompt_text)
                                else:
                                    pil_img = Image.open(img_path)
                                    response = _call_genai_with_retry(client, pil_img, prompt_text)
                                    raw_text = response.text.strip()
                            except Exception as api_err:
                                print(f"AI API rate limit or other error after retries: {api_err}")
                                raise api_err

                            try:
                                # Clean response string just in case it has markdown code block formatting
                                if raw_text.startswith("```json"):
                                    raw_text = raw_text[7:]
                                elif raw_text.startswith("```"):
                                    raw_text = raw_text[3:]
                                if raw_text.endswith("```"):
                                    raw_text = raw_text[:-3]
                                slide_data = json.loads(raw_text.strip())
                            except Exception as parse_e:
                                print(f"JSON Parse error for page {page_num}: {parse_e}\nRaw output was: {raw_text}")
                                raise parse_e
                            
                            # Build editable slide
                            l_type = slide_data.get('layout_type', 'title_and_content')
                            if l_type == 'diagram':
                                slide_layout = prs.slide_layouts[6] # Blank
                                slide = prs.slides.add_slide(slide_layout)
                                
                                # Apply theme ribbons
                                _apply_theme_ribbons(slide, prs, theme_colors)
                                
                                # We keep the original image for diagrams
                                _fit_image_to_slide(slide, img_path, SLIDE_WIDTH, SLIDE_HEIGHT, MARGIN)
                                
                                # Add a title text box over the image
                                left = Inches(0.5)
                                top = Inches(0.2)
                                width = Inches(12.0)
                                height = Inches(0.5)
                                txBox = slide.shapes.add_textbox(left, top, width, height)
                                tf = txBox.text_frame
                                p = tf.add_paragraph()
                                p.text = slide_data.get('title', f"Slide {page_num + 1}")
                                p.font.bold = True
                                p.font.size = Pt(28)
                                _apply_aptos_narrow(txBox)
                            else:
                                if l_type == 'two_column':
                                    slide_layout = prs.slide_layouts[3] # Two Content
                                else:
                                    slide_layout = prs.slide_layouts[1] # Title and Content
                                    
                                slide = prs.slides.add_slide(slide_layout)
                                
                                # Set Title
                                title_shape = slide.shapes.title
                                title_shape.left = Inches(0.5)
                                title_shape.top = Inches(0.25)
                                title_shape.width = SLIDE_WIDTH - Inches(1.0)
                                title_shape.height = Inches(0.8)
                                _apply_aptos_narrow(title_shape, font_color=RGBColor(*theme_colors["text"]))
                                
                                # Set Narrative
                                left = Inches(0.5)
                                top = Inches(0.95)
                                width = SLIDE_WIDTH - Inches(1.0)
                                height = Inches(0.5)
                                txBox = slide.shapes.add_textbox(left, top, width, height)
                                tf = txBox.text_frame
                                p = tf.add_paragraph()
                                p.text = slide_data.get('narrative', '')
                                p.font.size = Pt(16)
                                p.font.color.rgb = RGBColor(*theme_colors["text"])
                                _apply_aptos_narrow(txBox)
                                
                                # Set Punchline at bottom
                                left = Inches(0.5)
                                top = SLIDE_HEIGHT - Inches(0.8)
                                width = SLIDE_WIDTH - Inches(1.0)
                                height = Inches(0.4)
                                txBox_punch = slide.shapes.add_textbox(left, top, width, height)
                                tf_punch = txBox_punch.text_frame
                                p = tf_punch.add_paragraph()
                                p.text = slide_data.get('punchline', '')
                                p.font.italic = True
                                p.font.size = Pt(14)
                                p.font.color.rgb = RGBColor(*theme_colors["subtext"])
                                _apply_aptos_narrow(txBox_punch)
                                
                                # Set Bullets
                                body_shape = slide.placeholders[1]
                                body_shape.left = Inches(0.5)
                                body_shape.top = Inches(1.6)
                                body_shape.width = SLIDE_WIDTH - Inches(1.0)
                                body_shape.height = Inches(4.8)
                                tf = body_shape.text_frame
                                tf.word_wrap = True
                                tf.text = "" # clear default
                                for bullet in slide_data.get('bullet_points', []):
                                    p = tf.add_paragraph()
                                    p.text = bullet
                                    p.level = 0
                                _apply_aptos_narrow(body_shape, font_color=RGBColor(*theme_colors["text"]))
                                
                                # Add AI Generated Icon
                                icon_keyword = slide_data.get('icon_keyword', 'presentation')
                                icon_url = f"https://api.dicebear.com/9.x/icons/png?seed={icon_keyword}&backgroundColor=ffffff"
                                try:
                                    icon_resp = requests.get(icon_url, verify=False, timeout=10)
                                    if icon_resp.status_code == 200:
                                        icon_path = os.path.join(run_dir, f"icon_{page_num}.png")
                                        with open(icon_path, "wb") as f:
                                            f.write(icon_resp.content)
                                        slide.shapes.add_picture(icon_path, Inches(11.5), Inches(0.5), Inches(1), Inches(1))
                                except:
                                    pass
                                    
                                # If two column, or if AI indicated we should keep the image, put it in the right placeholder
                                keep_image = slide_data.get('keep_original_image', False)
                                if (l_type == 'two_column' or keep_image) and len(slide.placeholders) > 2:
                                    # Adjust left body shape to be half width
                                    body_shape.width = (SLIDE_WIDTH / 2) - Inches(0.75)
                                    
                                    right_body_shape = slide.placeholders[2]
                                    right_body_shape.left = (SLIDE_WIDTH / 2) + Inches(0.25)
                                    right_body_shape.top = Inches(1.8)
                                    right_body_shape.width = (SLIDE_WIDTH / 2) - Inches(0.75)
                                    right_body_shape.height = Inches(5.0)
                                    tf_right = right_body_shape.text_frame
                                    tf_right.word_wrap = True
                                    tf_right.text = "Original Context"
                                    _apply_aptos_narrow(right_body_shape, font_color=RGBColor(*theme_colors["text"]))
                                    
                                    # Insert the original image overlapping the placeholder slightly
                                    slide.shapes.add_picture(
                                        img_path, 
                                        right_body_shape.left, 
                                        right_body_shape.top + Inches(0.5), 
                                        right_body_shape.width, 
                                        right_body_shape.height - Inches(0.5)
                                    )
                                elif keep_image:
                                    # Just add it to the bottom right
                                    slide.shapes.add_picture(img_path, Inches(8.0), Inches(4.0), Inches(4.5), Inches(3.0))
                        except Exception as e:
                            print(f"GenAI failed for page {page_num}: {e}")
                            ai_rate_limit_fallback_count += 1
                            
                            # Fallback if AI completely fails
                            slide = prs.slides.add_slide(blank_layout)
                            _fit_image_to_slide(slide, img_path, SLIDE_WIDTH, SLIDE_HEIGHT, MARGIN)
                    else:
                        slide = prs.slides.add_slide(blank_layout)
                        _fit_image_to_slide(slide, img_path, SLIDE_WIDTH, SLIDE_HEIGHT, MARGIN)
                        
                    os.remove(img_path)
                    
                prs.save(output_path)
            else:
                # DOCX
                output_filename = "converted_document.docx"
                output_path = os.path.join(run_dir, output_filename)
                
                docx_doc = DocxDocument()
                docx_doc.add_heading('Generated Document from PDF', 0)
                
                if instructions or layout_theme or visual_iconography or slide_content_rules:
                    docx_doc.add_heading('Generation Guidelines', level=1)
                    if layout_theme: docx_doc.add_paragraph(f"Theme: {layout_theme}", style='List Bullet')
                    if visual_iconography: docx_doc.add_paragraph(f"Iconography: {visual_iconography}", style='List Bullet')
                    if slide_content_rules: docx_doc.add_paragraph(f"Content Rules: {slide_content_rules}", style='List Bullet')
                    if instructions: docx_doc.add_paragraph(f"Instructions: {instructions}", style='List Bullet')
                    
                docx_doc.add_heading('Extracted Content', level=1)
                
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text("text")
                    if text.strip():
                        docx_doc.add_paragraph(text)
                
                docx_doc.save(output_path)
                
                # Apply corporate guidelines to the generated docx
                formatted_output_filename = "final_formatted_document.docx"
                formatted_output_path = os.path.join(run_dir, formatted_output_filename)
                apply_guidelines(output_path, formatted_output_path)
                output_filename = formatted_output_filename
        finally:
            doc.close()
            
        file_url = _get_file_url(execution_id, output_filename)
        stats["successful_creations"] += 1
        _add_to_history(execution_id, output_filename, file_url, "process_pdf")
        
        msg = f"Successfully generated {target_format.upper()} from PDF."
        if target_format.lower() == "pptx" and not (has_genai or has_anthropic):
            msg += " Note: No valid API key found. Fell back to generating static image slides."
        elif target_format.lower() == "pptx" and ai_rate_limit_fallback_count > 0:
            msg += f" Note: {ai_rate_limit_fallback_count} slides fell back to original images due to AI API rate limits or errors. Retries were attempted."

        response_payload = {
            "success": True,
            "message": msg,
            "file_url": file_url,
            "download_path": f"/downloads/{execution_id}/{output_filename}",
            "execution_id": execution_id,
            "filename": output_filename
        }
        _trigger_webhook(webhook_url, response_payload)
        return response_payload
        
    except Exception as e:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": f"Error converting PDF: {str(e)}"
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload
    """
    Downloads/Reads a DOCX file, applies corporate guidelines, and returns the formatted DOCX URL.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    input_filename = "input.docx"
    input_path = os.path.join(run_dir, input_filename)
    output_filename = "formatted_document.docx"
    output_path = os.path.join(run_dir, output_filename)
    
    try:
        if is_url:
            if doc_source.startswith(('http://', 'https://')):
                response = requests.get(doc_source, verify=False)
                response.raise_for_status()
                with open(input_path, 'wb') as f:
                    f.write(response.content)
            else:
                if ";base64," in doc_source:
                    _, b64_data = doc_source.split(";base64,")
                    with open(input_path, 'wb') as f:
                        f.write(base64.b64decode(b64_data))
                else:
                    import shutil
                    shutil.copy2(doc_source, input_path)
        else:
            with open(input_path, 'wb') as f:
                f.write(base64.b64decode(doc_source))
                
        # Apply formatting guidelines
        apply_guidelines(input_path, output_path)
        
        file_url = _get_file_url(execution_id, output_filename)
        
        stats["successful_creations"] += 1
        _add_to_history(execution_id, output_filename, file_url, "format_docx")
        
        response_payload = {
            "success": True,
            "message": "Document formatted successfully.",
            "file_url": file_url,
            "download_path": f"/downloads/{execution_id}/{output_filename}",
            "execution_id": execution_id,
            "filename": output_filename
        }
        _trigger_webhook(webhook_url, response_payload)
        return response_payload
        
    except Exception as e:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": f"Error formatting document: {str(e)}"
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload

# Global stats
stats = {
    "requests_received": 0,
    "successful_creations": 0,
    "failed_creations": 0,
    "last_request_time": None,
    "last_success_file_url": None,
    "last_success_filename": None
}

# Global history for last X generated artifacts
MAX_HISTORY_ITEMS = int(os.environ.get("MAX_HISTORY_ITEMS", "10"))
generation_history = []

def _add_to_history(execution_id: str, filename: str, file_url: str, artifact_type: str):
    """Add a successful generation to the history list."""
    history_item = {
        "execution_id": execution_id,
        "filename": filename,
        "file_url": file_url,
        "type": artifact_type,
        "timestamp": datetime.now().isoformat()
    }
    generation_history.insert(0, history_item) # Add to front (newest first)
    
    # Keep only the last X items
    while len(generation_history) > MAX_HISTORY_ITEMS:
        generation_history.pop()

OUTPUT_DIR = os.environ.get("PPTX_OUTPUT_DIR", tempfile.gettempdir())
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Slide dimensions (16:9 widescreen)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
MARGIN = Inches(0.25)

def _get_file_url(execution_id: str, filename: str) -> str:
    file_path = os.path.abspath(os.path.join(OUTPUT_DIR, execution_id, filename))
    base_url = os.environ.get("BASE_URL", "")
    
    if base_url == "file://":
        return f"file:///{file_path.replace(chr(92), '/')}"
    elif base_url:
        prefix = base_url.rstrip('/')
        return f"{prefix}/downloads/{execution_id}/{filename}"
    else:
        # Fallback absolute path if no BASE_URL is set
        return f"http://localhost:8000/downloads/{execution_id}/{filename}"

def generate_presentation(python_code: str, webhook_url: str = None) -> dict:
    """
    Executes Python code to generate a PPTX file.
    Returns a dict with 'success', 'message', and optionally 'file_url' or 'execution_id'.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    script_path = os.path.join(run_dir, "script.py")
    
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(python_code)
        
    try:
        result = subprocess.run(
            ["python", "script.py"],
            cwd=run_dir, 
            capture_output=True, 
            text=True, 
            timeout=60
        )
        
        if result.returncode != 0:
            stats["failed_creations"] += 1
            return {
                "success": False,
                "message": f"Error executing code:\n{result.stderr}\n\nStdout:\n{result.stdout}"
            }
            
        pptx_files = [f for f in os.listdir(run_dir) if f.endswith(".pptx")]
        
        if not pptx_files:
            stats["failed_creations"] += 1
            return {
                "success": False,
                "message": "Execution succeeded but no .pptx file was found. Ensure your code calls presentation.save('output.pptx').\n\nStdout:\n" + result.stdout
            }
            
        file_url = _get_file_url(execution_id, pptx_files[0])
            
        stats["successful_creations"] += 1
        _add_to_history(execution_id, pptx_files[0], file_url, "generate_pptx")
        
        response_payload = {
            "success": True,
            "message": "Presentation generated successfully.",
            "file_url": file_url,
            "download_path": f"/downloads/{execution_id}/{pptx_files[0]}",
            "execution_id": execution_id,
            "filename": pptx_files[0]
        }
        _trigger_webhook(webhook_url, response_payload)
        return response_payload
        
    except subprocess.TimeoutExpired:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": "Error: Python code execution timed out after 60 seconds."
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload
    except Exception as e:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": f"Error: {str(e)}"
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload

def image_to_presentation(image_source: str, is_url: bool = True, webhook_url: str = None) -> dict:
    """
    Converts an image into a PPTX presentation with a single slide containing the image perfectly fitted.
    image_source: URL, file path, or base64 string.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    try:
        # Load image
        if is_url:
            if image_source.startswith(('http://', 'https://')):
                response = requests.get(image_source, verify=False)
                response.raise_for_status()
                img = Image.open(BytesIO(response.content))
            else:
                # Assume local file path or base64
                if ";base64," in image_source:
                    _, b64_data = image_source.split(";base64,")
                    img = Image.open(BytesIO(base64.b64decode(b64_data)))
                else:
                    img = Image.open(image_source)
        else:
            # Direct base64 string without data URI scheme
            img = Image.open(BytesIO(base64.b64decode(image_source)))
            
        # Save image locally to temp file to be used by python-pptx
        img_ext = img.format.lower() if img.format else "png"
        if img_ext == "jpeg":
            img_ext = "jpg"
        img_path = os.path.join(run_dir, f"source_image.{img_ext}")
        
        # Convert RGBA to RGB for JPEG
        if img.mode == 'RGBA' and img_ext in ['jpg', 'jpeg']:
            img = img.convert('RGB')
            
        img.save(img_path)
        
        # Create presentation
        prs = Presentation()
        prs.slide_width = SLIDE_WIDTH
        prs.slide_height = SLIDE_HEIGHT
        
        blank_layout = prs.slide_layouts[6] # Blank layout
        slide = prs.slides.add_slide(blank_layout)
        
        # Fit image to slide while preserving aspect ratio
        img_width, img_height = img.size
        page_aspect = img_width / img_height
        slide_aspect = (SLIDE_WIDTH - 2 * MARGIN) / (SLIDE_HEIGHT - 2 * MARGIN)
        
        if page_aspect > slide_aspect:
            width = SLIDE_WIDTH - 2 * MARGIN
            height = width / page_aspect
        else:
            height = SLIDE_HEIGHT - 2 * MARGIN
            width = height * page_aspect
            
        left = (SLIDE_WIDTH - width) / 2 + MARGIN
        top = (SLIDE_HEIGHT - height) / 2 + MARGIN
        
        slide.shapes.add_picture(img_path, left, top, width, height)
        
        output_filename = "image_presentation.pptx"
        output_path = os.path.join(run_dir, output_filename)
        prs.save(output_path)
        
        file_url = _get_file_url(execution_id, output_filename)
        
        stats["successful_creations"] += 1
        _add_to_history(execution_id, output_filename, file_url, "image_to_pptx")
        
        response_payload = {
            "success": True,
            "message": "Image presentation generated successfully.",
            "file_url": file_url,
            "download_path": f"/downloads/{execution_id}/{output_filename}",
            "execution_id": execution_id,
            "filename": output_filename
        }
        _trigger_webhook(webhook_url, response_payload)
        return response_payload
        
    except Exception as e:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": f"Error converting image to presentation: {str(e)}"
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload
def generate_artifacts_from_prompt(
    prompt: str,
    target_format: str = "pptx",
    presentation_style: str = "Detailed",
    layout_theme: str = "Modern Light",
    num_slides: int = 5,
    webhook_url: str = None,
    api_key: str = ""
) -> dict:
    """
    Dynamically generates a full presentation or document strictly from a text prompt.
    Takes into account requested themes and styles.
    """
    stats["requests_received"] += 1
    stats["last_request_time"] = datetime.now().isoformat()
    
    execution_id = str(uuid.uuid4())
    run_dir = os.path.join(OUTPUT_DIR, execution_id)
    os.makedirs(run_dir, exist_ok=True)
    
    try:
        has_genai = False
        has_anthropic = False
        client = None
        
        use_anthropic = False
        if api_key.startswith("sk-ant") or (not api_key and os.environ.get("ANTHROPIC_API_KEY") and not os.environ.get("GEMINI_API_KEY") and not os.environ.get("GOOGLE_API_KEY")):
            use_anthropic = True
            
        if use_anthropic:
            try:
                k = api_key if api_key else os.environ.get("ANTHROPIC_API_KEY")
                proxy_url = os.environ.get("GCP_PROXY_FOR_CLAUD")
                if proxy_url:
                    client = anthropic.Anthropic(api_key=k, base_url=proxy_url, max_retries=0)
                else:
                    client = anthropic.Anthropic(api_key=k, max_retries=0)
                has_anthropic = True
            except:
                pass
        else:
            try:
                if api_key:
                    client = genai.Client(api_key=api_key)
                elif os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"):
                    client = genai.Client()
                if client:
                    has_genai = True
            except:
                pass
                
        if not (has_genai or has_anthropic):
            raise Exception("No valid GenAI or Anthropic API key configured.")
            
    if target_format.lower() == "pptx":
        # AI prompt strictly enforcing the Presentation Creation Kit anatomy
        _send_progress(webhook_url, "Generating presentation outline with AI...")
        system_prompt = f"""You are an expert presentation designer and strategic consultant.
Create a {num_slides}-slide presentation outline on the following topic: {prompt}

Presentation Style Constraint: {presentation_style}
Theme Concept: {layout_theme}

STRICT SLIDE ANATOMY (Generic Presentation Kit):
DO NOT brand this presentation with 'JPL', 'JEMP', or specific corporate tags unless explicitly requested by the user. Use neutral terms like 'The Organization' or 'The Platform'.

Each slide MUST be structured according to this strict contract:
1. Title: A concise, impactful title.
2. Narrative: A 1-2 line explanatory narrative setting up the slide's argument.
3. Content/Visual: Provide 3-5 max bullet points with parallel grammar. Do not exceed 5 bullets.
4. Archetypes: Assign a `slide_archetype` (title, agenda, divider, standard, table, deep_dive, roadmap, options).
5. For 'table', 'roadmap', or 'options', provide `table_data` as a 2D array of strings where the first row is headers.
6. Punchline: A single takeaway line summarizing the strategic impact. Unique per slide.

DECK STRUCTURE & STORYTELLING:
- Storyline Flow: Context/Vision -> Execution Model (Tracks/Phases) -> Options -> Architecture -> Roadmap -> Risks -> Recommendation.
- Define terms clearly (e.g., Track = enduring workstream; Project = deliverable).
- No filler slides: every slide answers a question. Density over page-count chasing.
- For comparisons/options, use the 'table' archetype to matrix the options against criteria.

Write the output in the JSON format matching this schema:
"""
            if use_anthropic:
                raw_text = _call_anthropic_text_with_retry(client, system_prompt, PresentationData)
            else:
                response = _call_genai_text_with_retry(client, system_prompt, PresentationData)
                raw_text = response.text
                
            try:
                raw_text = raw_text.strip()
                if raw_text.startswith("```json"): raw_text = raw_text[7:]
                elif raw_text.startswith("```"): raw_text = raw_text[3:]
                if raw_text.endswith("```"): raw_text = raw_text[:-3]
                data = json.loads(raw_text.strip())
            except Exception as e:
                raise Exception(f"Failed to parse LLM JSON: {e}")
                
            output_filename = "generated_presentation.pptx"
            output_path = os.path.join(run_dir, output_filename)
            
            prs, theme_colors = _create_themed_presentation(layout_theme)
            
            _send_progress(webhook_url, "Generating presentation slides...")
            
            slides_data = data.get("slides", [])
            for i, s_data in enumerate(slides_data):
                l_type = s_data.get('layout_type', 'title_and_content')
                archetype = s_data.get('slide_archetype', 'standard')
                
                # Title Slide Archetype
                if archetype == 'title':
                    slide_layout = prs.slide_layouts[0] # Title layout
                    slide = prs.slides.add_slide(slide_layout)
                    slide.shapes.title.text = s_data.get('title', 'Presentation')
                    _apply_aptos_narrow(slide.shapes.title, font_color=RGBColor(*theme_colors["text"]))
                    if len(slide.placeholders) > 1:
                        subtitle = slide.placeholders[1]
                        subtitle.text = s_data.get('narrative', '') + "\n" + s_data.get('punchline', '')
                        _apply_aptos_narrow(subtitle, font_color=RGBColor(*theme_colors["subtext"]))
                    continue
                    
                # Section Divider Archetype
                if archetype == 'divider':
                    slide_layout = prs.slide_layouts[6] # Blank
                    slide = prs.slides.add_slide(slide_layout)
                    # Center huge text
                    txBox = slide.shapes.add_textbox(Inches(1), Inches(3), SLIDE_WIDTH - Inches(2), Inches(1.5))
                    tf = txBox.text_frame
                    p = tf.add_paragraph()
                    p.text = s_data.get('title', 'Section')
                    p.font.size = Pt(44)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(*theme_colors["text"])
                    p.alignment = PP_ALIGN.CENTER
                    _apply_aptos_narrow(txBox)
                    continue

                if l_type == 'two_column':
                    slide_layout = prs.slide_layouts[3]
                else:
                    slide_layout = prs.slide_layouts[1]
                    
                slide = prs.slides.add_slide(slide_layout)
                
                # Set Title
                title_shape = slide.shapes.title
                title_shape.text = s_data.get('title', f"Slide {i + 1}")
                title_shape.left = Inches(0.5)
                title_shape.top = Inches(0.25)
                title_shape.width = SLIDE_WIDTH - Inches(1.0)
                title_shape.height = Inches(0.8)
                _apply_aptos_narrow(title_shape, font_color=RGBColor(*theme_colors["text"]))
                
                # Set Narrative
                left = Inches(0.5)
                top = Inches(0.95)
                width = SLIDE_WIDTH - Inches(1.0)
                height = Inches(0.5)
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.add_paragraph()
                p.text = s_data.get('narrative', '')
                p.font.size = Pt(16)
                p.font.color.rgb = RGBColor(*theme_colors["text"])
                _apply_aptos_narrow(txBox)
                
                # Set Punchline at bottom
                left = Inches(0.5)
                top = SLIDE_HEIGHT - Inches(0.8)
                width = SLIDE_WIDTH - Inches(1.0)
                height = Inches(0.4)
                txBox_punch = slide.shapes.add_textbox(left, top, width, height)
                tf_punch = txBox_punch.text_frame
                p = tf_punch.add_paragraph()
                p.text = s_data.get('punchline', '')
                p.font.italic = True
                p.font.size = Pt(14)
                p.font.color.rgb = RGBColor(*theme_colors["subtext"])
                _apply_aptos_narrow(txBox_punch)
                
                # Render Table if archetype matches and data exists
                if archetype in ['table', 'roadmap', 'options'] and s_data.get('table_data'):
                    table_data = s_data.get('table_data')
                    rows = len(table_data)
                    cols = len(table_data[0]) if rows > 0 else 0
                    if rows > 0 and cols > 0:
                        # Clear default text box
                        slide.placeholders[1].text_frame.text = ""
                        # Add Table
                        table_shape = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.6), SLIDE_WIDTH - Inches(1.0), Inches(4.5))
                        table = table_shape.table
                        for r_idx, row_data in enumerate(table_data):
                            for c_idx, cell_value in enumerate(row_data):
                                if c_idx < cols:
                                    cell = table.cell(r_idx, c_idx)
                                    cell.text = str(cell_value)
                                    _apply_aptos_narrow(cell.text_frame, font_color=RGBColor(*theme_colors["text"]))
                                    if r_idx == 0: # Header
                                        cell.fill.solid()
                                        cell.fill.fore_color.rgb = RGBColor(*theme_colors["accent"])
                                        for paragraph in cell.text_frame.paragraphs:
                                            paragraph.font.color.rgb = RGBColor(255,255,255)
                                            paragraph.font.bold = True
                    continue
                    
                # Otherwise, Standard Bullets
                body_shape = slide.placeholders[1]
                body_shape.left = Inches(0.5)
                body_shape.top = Inches(1.6)
                body_shape.width = SLIDE_WIDTH - Inches(1.0)
                body_shape.height = Inches(4.8)
                tf = body_shape.text_frame
                tf.word_wrap = True
                tf.text = "" # clear default
                for bullet in s_data.get('bullet_points', []):
                    p = tf.add_paragraph()
                    p.text = bullet
                    p.level = 0
                _apply_aptos_narrow(body_shape, font_color=RGBColor(*theme_colors["text"]))
                
                # Add AI Generated Icon
                icon_keyword = s_data.get('icon_keyword', 'presentation')
                icon_url = f"https://api.dicebear.com/9.x/icons/png?seed={icon_keyword}&backgroundColor=ffffff"
                try:
                    icon_resp = requests.get(icon_url, verify=False, timeout=10)
                    if icon_resp.status_code == 200:
                        icon_path = os.path.join(run_dir, f"icon_{i}.png")
                        with open(icon_path, "wb") as f:
                            f.write(icon_resp.content)
                        slide.shapes.add_picture(icon_path, Inches(11.5), Inches(0.5), Inches(1), Inches(1))
                except:
                    pass
                
                # Two Column adjustment
                if l_type == 'two_column' and len(slide.placeholders) > 2:
                    body_shape.width = (SLIDE_WIDTH / 2) - Inches(0.75)
                    
                    right_body_shape = slide.placeholders[2]
                    right_body_shape.left = (SLIDE_WIDTH / 2) + Inches(0.25)
                    right_body_shape.top = Inches(1.6)
                    right_body_shape.width = (SLIDE_WIDTH / 2) - Inches(0.75)
                    right_body_shape.height = Inches(4.8)
                    tf_right = right_body_shape.text_frame
                    tf_right.word_wrap = True
                    tf_right.text = "Additional Context / Visuals"
                    _apply_aptos_narrow(right_body_shape, font_color=text_color)
                
                # Strategic Impact Box (Bottom Ribbon / Takeaway)
                takeaway_text = s_data.get('key_takeaway', '')
                if takeaway_text:
                    from pptx.enum.shapes import MSO_SHAPE
                    rect = slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE, 
                        Inches(0.5), Inches(6.2), SLIDE_WIDTH - Inches(1.0), Inches(0.9)
                    )
                    rect.fill.solid()
                    if "dark" in theme_low:
                        rect.fill.fore_color.rgb = RGBColor(40, 45, 55)
                        rect.line.color.rgb = RGBColor(100, 150, 255)
                    else:
                        rect.fill.fore_color.rgb = RGBColor(240, 245, 250)
                        rect.line.color.rgb = header_bg_color
                    
                    tf_rect = rect.text_frame
                    tf_rect.word_wrap = True
                    
                    # Bold Header
                    p = tf_rect.paragraphs[0]
                    p.text = "Strategic Takeaway"
                    p.font.bold = True
                    p.font.size = Pt(14)
                    if "dark" in theme_low: p.font.color.rgb = RGBColor(255, 255, 255)
                    else: p.font.color.rgb = header_bg_color
                    
                    # Takeaway Content
                    p2 = tf_rect.add_paragraph()
                    p2.text = takeaway_text
                    p2.font.size = Pt(13)
                    if "dark" in theme_low: p2.font.color.rgb = RGBColor(200, 200, 200)
                    else: p2.font.color.rgb = RGBColor(80, 80, 80)
                    
                    _apply_aptos_narrow(rect)
                
                # Footer text
                footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.2), SLIDE_WIDTH - Inches(1.0), Inches(0.3))
                tf_footer = footer.text_frame
                p_footer = tf_footer.paragraphs[0]
                p_footer.text = f"Slide {i + 1} of {len(slides_data)} | Generated via {layout_theme} Theme"
                p_footer.font.size = Pt(10)
                if "dark" in theme_low: p_footer.font.color.rgb = RGBColor(150, 150, 150)
                else: p_footer.font.color.rgb = RGBColor(120, 120, 120)
                p_footer.alignment = PP_ALIGN.RIGHT
                _apply_aptos_narrow(footer)
            
            prs.save(output_path)
            
        else: # DOCX
            _send_progress(webhook_url, "Generating document content with AI...")
            system_prompt = f"""You are an expert document author and strategic consultant.
Create a detailed, well-structured document on the following topic: {prompt}

Document Style Constraint: {presentation_style}
Theme/Tone: {layout_theme}

STRICT DOCUMENT ORGANIZATION (The Formatting Kit):
1. Introduction: Must explain what the document is and how sections are organized. Outline the roadmap at a section level.
2. Structure: Follow a logical order (e.g., Vision -> Execution/Tracks -> Architecture/Options -> Details).
3. Clarity: Separate Projects (what you build) from Tracks (how you execute). Include comparative analysis where options are discussed.
4. Formatting: The document must be well organized into headings, content paragraphs, and bullet points where useful for scanning.
5. Depth: Do not drop content, only add. Avoid padding just to hit a page count.
"""
            if use_anthropic:
                raw_text = _call_anthropic_text_with_retry(client, system_prompt, DocumentData)
            else:
                response = _call_genai_text_with_retry(client, system_prompt, DocumentData)
                raw_text = response.text
                
            try:
                raw_text = raw_text.strip()
                if raw_text.startswith("```json"): raw_text = raw_text[7:]
                elif raw_text.startswith("```"): raw_text = raw_text[3:]
                if raw_text.endswith("```"): raw_text = raw_text[:-3]
                data = json.loads(raw_text.strip())
            except Exception as e:
                raise Exception(f"Failed to parse LLM JSON: {e}")
                
            output_filename = "generated_document.docx"
            output_path = os.path.join(run_dir, output_filename)
            
            docx_doc = DocxDocument()
            docx_doc.add_heading(data.get("title", "Generated Document"), 0)
            
            for section in data.get("sections", []):
                docx_doc.add_heading(section.get("heading", "Section"), level=1)
                for paragraph in section.get("content", "").split("\n\n"):
                    if paragraph.strip():
                        docx_doc.add_paragraph(paragraph.strip())
                for bullet in section.get("bullet_points", []):
                    docx_doc.add_paragraph(bullet, style='List Bullet')
                    
            docx_doc.save(output_path)
            
            # Apply formatting guidelines
            formatted_output_filename = "final_formatted_document.docx"
            formatted_output_path = os.path.join(run_dir, formatted_output_filename)
            apply_guidelines(output_path, formatted_output_path)
            output_filename = formatted_output_filename
            
        file_url = _get_file_url(execution_id, output_filename)
        stats["successful_creations"] += 1
        _add_to_history(execution_id, output_filename, file_url, "generate_from_prompt")
        
        response_payload = {
            "success": True,
            "message": f"Successfully generated {target_format.upper()} from prompt.",
            "file_url": file_url,
            "download_path": f"/downloads/{execution_id}/{output_filename}",
            "execution_id": execution_id,
            "filename": output_filename
        }
        _trigger_webhook(webhook_url, response_payload)
        return response_payload

    except Exception as e:
        stats["failed_creations"] += 1
        error_payload = {
            "success": False,
            "message": f"Error generating from prompt: {str(e)}"
        }
        _trigger_webhook(webhook_url, error_payload)
        return error_payload
