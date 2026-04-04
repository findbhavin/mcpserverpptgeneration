import os
import uuid
import tempfile
import subprocess
import base64
import requests
from io import BytesIO
from datetime import datetime
from PIL import Image
from docx_formatter import apply_guidelines
from pptx import Presentation
from pptx.util import Inches

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
