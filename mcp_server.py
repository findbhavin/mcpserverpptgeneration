from mcp.server.fastmcp import FastMCP
from core import generate_presentation, image_to_presentation, format_document

mcp = FastMCP("pptx_generator")

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
