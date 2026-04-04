from mcp.server.fastmcp import FastMCP
from core import generate_presentation, image_to_presentation, format_document, process_pdf_to_artifacts

mcp = FastMCP("pptx_generator")

@mcp.tool()
def process_pdf(
    pdf_source: str, 
    is_url: bool = True, 
    instructions: str = "", 
    layout_theme: str = "", 
    visual_iconography: str = "", 
    slide_content_rules: str = "",
    target_format: str = "pptx"
) -> str:
    """
    Takes a fresh PDF file and instructions, converting it into a styled PPTX or formatted DOCX.
    Args:
        pdf_source: URL, file path, or base64 string of the PDF.
        is_url: True if pdf_source is a URL/file path, False if raw base64 string.
        instructions: Abstract requests/instructions for how the content should be treated.
        layout_theme: Requested theme (e.g., 'Modern Corporate', 'Dark Mode').
        visual_iconography: Rules regarding icons and imagery.
        slide_content_rules: Guidelines on how to split text across slides.
        target_format: 'pptx' or 'docx'.
    Returns the URL/path to the generated file.
    """
    result = process_pdf_to_artifacts(
        pdf_source, is_url, instructions, layout_theme, 
        visual_iconography, slide_content_rules, target_format
    )
    if result["success"]:
        return result["file_url"]
    else:
        return result["message"]

@mcp.tool()
def generate_pptx(python_code: str) -> str:
    """
    Generate a PPTX file using python-pptx by executing the provided Python code.
    The code should save the presentation to the current working directory.
    Returns the URL/path to the generated PPTX file.
    """
    result = generate_presentation(python_code)
    if result["success"]:
        return result["file_url"]
    else:
        return result["message"]

@mcp.tool()
def image_to_pptx(image_source: str, is_url: bool = True) -> str:
    """
    Converts an image into a PPTX presentation with a single slide containing the image perfectly fitted.
    Args:
        image_source: URL, file path, or base64 string of the image.
        is_url: True if image_source is a URL or file path or data URI, False if it's a raw base64 string.
    Returns the URL/path to the generated PPTX file.
    """
    result = image_to_presentation(image_source, is_url)
    if result["success"]:
        return result["file_url"]
    else:
        return result["message"]

@mcp.tool()
def apply_docx_template(doc_source: str, is_url: bool = True) -> str:
    """
    Takes an existing DOCX document and reformats it to strictly follow corporate guidelines:
    - Global font: Aptos Narrow
    - Auto-formats Tables, Table of Contents, Figures, and Headings
    Args:
        doc_source: URL, file path, or base64 string of the docx.
        is_url: True if doc_source is a URL/file path, False if raw base64 string.
    Returns the URL/path to the newly formatted DOCX file.
    """
    result = format_document(doc_source, is_url)
    if result["success"]:
        return result["file_url"]
    else:
        return result["message"]

if __name__ == "__main__":
    mcp.run()
