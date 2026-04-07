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
    target_format: str = "pptx",
    webhook_url: str = None,
    api_key: str = ""
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
        webhook_url: Optional webhook URL to POST the JSON result to.
        api_key: Optional Gemini or Google GenAI API key. If omitted, falls back to server env vars.
    Returns:
        A JSON string containing the 'success' boolean, 'message', and 'file_url'.
    """
    import json
    result = process_pdf_to_artifacts(
        pdf_source, is_url, instructions, layout_theme, 
        visual_iconography, slide_content_rules, target_format,
        webhook_url=webhook_url, api_key=api_key
    )
    return json.dumps(result, indent=2)

@mcp.tool()
def generate_pptx(python_code: str, webhook_url: str = None) -> str:
    """
    Generate a PPTX file using python-pptx by executing the provided Python code.
    The code should save the presentation to the current working directory.
    Args:
        python_code: The Python script to generate the pptx. Must end with `prs.save("output.pptx")`.
        webhook_url: Optional webhook URL to POST the JSON result to.
    Returns:
        A JSON string containing the 'success' boolean, 'message', and 'file_url'.
    """
    import json
    result = generate_presentation(python_code, webhook_url=webhook_url)
    return json.dumps(result, indent=2)

@mcp.tool()
def image_to_pptx(image_source: str, is_url: bool = True, webhook_url: str = None) -> str:
    """
    Converts an image into a PPTX presentation with a single slide containing the image perfectly fitted.
    Args:
        image_source: URL, file path, or base64 string of the image.
        is_url: True if image_source is a URL or file path or data URI, False if it's a raw base64 string.
        webhook_url: Optional webhook URL to POST the JSON result to.
    Returns:
        A JSON string containing the 'success' boolean, 'message', and 'file_url'.
    """
    import json
    result = image_to_presentation(image_source, is_url, webhook_url=webhook_url)
    return json.dumps(result, indent=2)

@mcp.tool()
def apply_docx_template(doc_source: str, is_url: bool = True, webhook_url: str = None) -> str:
    """
    Takes an existing DOCX document and reformats it to strictly follow corporate guidelines.
    Args:
        doc_source: URL, file path, or base64 string of the docx.
        is_url: True if doc_source is a URL/file path, False if raw base64 string.
        webhook_url: Optional webhook URL to POST the JSON result to.
    Returns:
        A JSON string containing the 'success' boolean, 'message', and 'file_url'.
    """
    import json
    result = format_document(doc_source, is_url, webhook_url=webhook_url)
    return json.dumps(result, indent=2)

@mcp.tool()
def generate_from_prompt(
    prompt: str,
    target_format: str = "pptx",
    presentation_style: str = "Detailed",
    layout_theme: str = "Modern Light",
    num_slides: int = 5,
    webhook_url: str = None,
    api_key: str = ""
) -> str:
    """
    Dynamically generates a full presentation or document strictly from a text prompt.
    Args:
        prompt: The main topic or prompt to generate the presentation/document from.
        target_format: 'pptx' or 'docx'.
        presentation_style: E.g., "Detailed", "Abstract", "Executive", "Minimalist".
        layout_theme: E.g., "Dark Corporate", "Light Modern", "Pastel".
        num_slides: Number of slides to generate (if pptx).
        webhook_url: Optional webhook URL to POST the JSON result to.
        api_key: AI API Key to use for content generation.
    Returns:
        A JSON string containing the 'success' boolean, 'message', and 'file_url'.
    """
    import json
    from core import generate_artifacts_from_prompt
    result = generate_artifacts_from_prompt(
        prompt=prompt,
        target_format=target_format,
        presentation_style=presentation_style,
        layout_theme=layout_theme,
        num_slides=num_slides,
        webhook_url=webhook_url,
        api_key=api_key
    )
    return json.dumps(result, indent=2)

if __name__ == "__main__":
    mcp.run()
