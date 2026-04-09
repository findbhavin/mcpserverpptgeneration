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

@mcp.tool()
def get_capabilities() -> str:
    """
    Returns a comprehensive list of capabilities and accepted parameters for this MCP server.
    External RAG agents should call this first to understand what the server can do.
    """
    import json
    capabilities = {
        "server_name": "ppt-doc-generator",
        "description": "Advanced server for generating, parsing, and formatting PPTX and DOCX documents.",
        "capabilities": [
            {
                "tool_name": "generate_from_prompt",
                "description": "Dynamically generate full presentations (PPTX) or documents (DOCX) from abstract text prompts.",
                "supported_formats": ["pptx", "docx"],
                "supported_themes": ["Modern Light", "Dark Corporate", "Pastel", "Blue Accent"],
                "supported_styles": ["Detailed", "Abstract", "Executive", "Minimalist"],
                "ai_integration": "Uses Gemini or Anthropic to create content and DiceBear for iconography."
            },
            {
                "tool_name": "process_pdf",
                "description": "Convert PDF documents into editable PPTX presentations or formatted DOCX documents.",
                "features": ["AI-driven text extraction", "Layout preservation", "Iconography generation", "Summarization"],
                "supported_formats": ["pptx", "docx"]
            },
            {
                "tool_name": "convert_image_to_pptx",
                "description": "Convert a single image (URL, base64, or file path) into a one-slide presentation."
            },
            {
                "tool_name": "apply_docx_template",
                "description": "Reformat an existing DOCX document to adhere to strict corporate branding and layout guidelines (Aptos Narrow font, specific heading sizes, table formatting, etc.)."
            },
            {
                "tool_name": "generate_pptx",
                "description": "Directly execute python-pptx code in a sandboxed environment to construct a presentation programmatically."
            }
        ],
        "interaction_guidelines": [
            "Use webhook_url for long-running generation tasks to get asynchronous results.",
            "Base64 or public URLs are supported for file inputs.",
            "When using process_pdf or generate_from_prompt, explicitly passing api_key is recommended if the server is not pre-configured with them.",
            "CRITICAL: To download files, the RAG agent MUST append the 'download_path' to the known MCP server URL rather than relying solely on 'file_url', as LLMs may hallucinate local paths like 'sandbox:/mnt/data/'."
        ],
        "presentation_standardization": {
            "protocol": "STRICT",
            "description": "All presentations generated by this MCP server follow a strict protocol for standardization.",
            "features": [
                "Uses a robust, themed Master Slide instead of blank slides.",
                "Consistent slide background color with elegant colored ribbons at the header and footer.",
                "Aptos Narrow font is set uniformly across all elements.",
                "Four distinct professional themes supported: 'Dark Corporate', 'Modern Light', 'Pastel', and 'Blue Accent'.",
                "Slide elements (titles, punchlines, bullets, context sidebars) automatically adjust their colors (dark/light text) to ensure readability based on the selected theme."
            ]
        }
    }
    return json.dumps(capabilities, indent=2)

if __name__ == "__main__":
    mcp.run()
