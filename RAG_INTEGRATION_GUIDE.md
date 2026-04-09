# Integration Guide: RAG Framework to Presentation/Document MCP Server

This guide outlines how your Retrieval-Augmented Generation (RAG) framework should integrate with and utilize the `ppt-doc-generator` MCP Server to fulfill user requests for PowerPoint presentations (PPTX) and Word documents (DOCX).

## 0. Strict Protocol Setup: Presentation Standardization (Templates/Themes)

**Before you start generating a presentation, the RAG Agent MUST understand and enforce the following standards:**

We do not just create blank slides. The MCP server utilizes a `_create_themed_presentation()` function that generates a robust, themed Master Slide.

**The Strict Standards:**
1. **Themed Master Slides:** Every presentation uses a consistent slide background color and elegant colored ribbons at the header and footer.
2. **Typography:** The **Aptos Narrow** font is set uniformly across all elements.
3. **Available Themes:** We support four distinct professional themes:
   * `"Dark Corporate"`
   * `"Modern Light"`
   * `"Pastel"`
   * `"Blue Accent"`
4. **Dynamic Readability:** The slide elements (titles, punchlines, bullets, context sidebars) automatically adjust their colors (dark/light text) to ensure readability based on the selected theme.

**Agent Requirement:** If the user does not specify a theme, the RAG Agent should **alternatively use some good default themes** (like "Modern Light" or "Dark Corporate") when calling the MCP tools.

## 1. Initial Handshake: Capability Discovery

Before attempting to generate artifacts, the RAG agent should query the MCP Server's capabilities to understand the available tools, accepted formats, themes, and styles.

**Tool to Call:** `get_capabilities`
**Parameters:** None

**How to use the response:** 
The server will return a JSON manifest detailing the supported `layout_theme` options (e.g., "Modern Light", "Dark Corporate") and `presentation_style` options (e.g., "Detailed", "Executive"). The RAG agent should dynamically inject these constraints into its internal prompt formulation logic so it only requests styles the server natively understands.

## 2. Core Generation Workflows

Based on the user's request, the RAG framework should map its intent to one of the following MCP tools:

### Scenario A: Generating from Scratch (Prompt-Based)
When the user asks to "create a presentation about X" or "write a document summarizing Y", and the RAG has compiled the research text.

**Tool to Call:** `generate_from_prompt`
**Parameters:**
*   `prompt` (string, required): The core content or topic. **Best Practice:** Pass the finalized, researched text from the RAG directly into this parameter.
*   `target_format` (string): `"pptx"` or `"docx"`
*   `presentation_style` (string): E.g., `"Detailed"`, `"Executive"`, `"Minimalist"`.
*   `layout_theme` (string): E.g., `"Modern Light"`, `"Dark Corporate"`.
*   `num_slides` (integer): Desired length (default 5).
*   `api_key` (string, optional): Pass the AI provider API key if the MCP server isn't globally configured with one.
*   `webhook_url` (string, optional): Recommended for async notification upon completion.

### Scenario B: Converting Existing Documents (PDF to PPTX/DOCX)
When the user provides a source PDF and wants it transformed into an editable presentation or a styled document.

**Tool to Call:** `process_pdf`
**Parameters:**
*   `pdf_source` (string, required): A public URL to the PDF, a local file path, or a raw Base64 string.
*   `is_url` (boolean): Set to `true` if `pdf_source` is a URL/path, `false` if Base64.
*   `instructions` (string): Specific rules for the conversion (e.g., "Focus only on financial tables").
*   `target_format` (string): `"pptx"` or `"docx"`
*   `layout_theme` (string), `visual_iconography` (string), `slide_content_rules` (string): Styling parameters.
*   `webhook_url` (string, optional)

### Scenario C: Branding / Formatting Existing DOCX
When a user has an unformatted DOCX and needs it aligned to strict corporate standards (Aptos Narrow, specific table/heading styles).

**Tool to Call:** `apply_docx_template`
**Parameters:**
*   `doc_source` (string, required): URL, path, or Base64 of the DOCX.
*   `is_url` (boolean).
*   `webhook_url` (string, optional).

## 3. Best Practices for RAG Agents

1.  **Asynchronous Handling (Webhooks):** Generation tasks (especially PDF processing or long presentations) can take 10-60 seconds due to AI image processing and slide generation. The RAG framework should provide a `webhook_url` whenever possible. The server will immediately accept the request and later POST a JSON payload `{"success": true, "file_url": "...", ...}` to the webhook. The RAG should tell the user "I am generating your presentation, it will be ready shortly" while waiting for the webhook.
2.  **API Key Management:** If the MCP Server is deployed in a stateless/secure environment without hardcoded AI keys, the RAG agent must pass the user's (or system's) `api_key` (Gemini or Anthropic) in the tool call.
3.  **Handling Fallbacks:** The MCP server gracefully handles rate limits by falling back to static image generation if AI parsing fails. The returned JSON message will indicate if a fallback occurred. The RAG should inform the user if the resulting slides are less editable due to rate limits.