import os
import uuid
import tempfile
import subprocess
import base64
import requests
from io import BytesIO
from datetime import datetime
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx_formatter import apply_guidelines

def format_document(doc_source: str, is_url: bool = True) -> dict:
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

def process_pdf_to_artifacts(
    pdf_source: str, 
    is_url: bool = True, 
    instructions: str = "", 
    layout_theme: str = "", 
    visual_iconography: str = "", 
    slide_content_rules: str = "",
    target_format: str = "pptx"
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
        
        if target_format.lower() == "pptx":
            output_filename = "converted_presentation.pptx"
            output_path = os.path.join(run_dir, output_filename)
            
            prs = Presentation()
            prs.slide_width = SLIDE_WIDTH
            prs.slide_height = SLIDE_HEIGHT
            
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
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_path = os.path.join(run_dir, f"page_{page_num}.png")
                pix.save(img_path)
                
                slide = prs.slides.add_slide(blank_layout)
                
                img_width = pix.width
                img_height = pix.height
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
                os.remove(img_path)
                
            doc.close()
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
            
            doc.close()
            docx_doc.save(output_path)
            
            # Apply corporate guidelines to the generated docx
            formatted_output_filename = "final_formatted_document.docx"
            formatted_output_path = os.path.join(run_dir, formatted_output_filename)
            apply_guidelines(output_path, formatted_output_path)
            output_filename = formatted_output_filename
            
        file_url = _get_file_url(execution_id, output_filename)
        stats["successful_creations"] += 1
        return {
            "success": True,
            "message": f"Successfully generated {target_format.upper()} from PDF.",
            "file_url": file_url,
            "execution_id": execution_id,
            "filename": output_filename
        }
        
    except Exception as e:
        stats["failed_creations"] += 1
        return {
            "success": False,
            "message": f"Error converting PDF: {str(e)}"
        }
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

# Global stats
stats = {
    "requests_received": 0,
    "successful_creations": 0,
    "failed_creations": 0,
    "last_request_time": None
}

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
    else:
        prefix = base_url.rstrip('/') if base_url else ""
        return f"{prefix}/downloads/{execution_id}/{filename}"

def generate_presentation(python_code: str) -> dict:
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
        return {
            "success": True,
            "message": "Presentation generated successfully.",
            "file_url": file_url,
            "execution_id": execution_id,
            "filename": pptx_files[0]
        }
        
    except subprocess.TimeoutExpired:
        stats["failed_creations"] += 1
        return {
            "success": False,
            "message": "Error: Python code execution timed out after 60 seconds."
        }
    except Exception as e:
        stats["failed_creations"] += 1
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }

def image_to_presentation(image_source: str, is_url: bool = True) -> dict:
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
        return {
            "success": True,
            "message": "Image presentation generated successfully.",
            "file_url": file_url,
            "execution_id": execution_id,
            "filename": output_filename
        }
        
    except Exception as e:
        stats["failed_creations"] += 1
        return {
            "success": False,
            "message": f"Error converting image to presentation: {str(e)}"
        }
