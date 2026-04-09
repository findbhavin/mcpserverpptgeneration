# RAG to MCP Server Interface Contract

This document defines the strict API contract and data schemas for communication between external Retrieval-Augmented Generation (RAG) agents and the `ppt-doc-generator` MCP Server. 

To ensure the RAG agents and the MCP server are perfectly in sync, any RAG agent calling this server must adhere to the parameters and enumerations defined below.

---

## 1. Allowed Enumerations (Standardization)

Whenever specifying themes, styles, or formats, the RAG agent MUST use one of the following exact string values.

### `target_format`
* `"pptx"` (PowerPoint Presentation)
* `"docx"` (Word Document)

### `layout_theme`
* `"Dark Corporate"` (Slate grey backgrounds, white text, vivid blue accents)
* `"Modern Light"` (Crisp light blue/white backgrounds, dark text, red accents)
* `"Pastel"` (Soft cream backgrounds, dark grey text, mint green accents)
* `"Blue Accent"` (Pure white backgrounds, navy blue text and accents)

### `presentation_style`
* `"Detailed"` (Comprehensive bullet points, deep context)
* `"Executive"` (High-level summaries, metric-focused)
* `"Abstract"` (Conceptual, minimal text)
* `"Minimalist"` (Very sparse text, high reliance on iconography/layout)

---

## 2. MCP Tool Specifications

### Tool: `generate_from_prompt`
**Description:** Dynamically generates a full presentation (PPTX) or document (DOCX) from an abstract text prompt.
**Type:** Synchronous (returns JSON) & Asynchronous (posts to webhook).

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `prompt` | `string` | **Yes** | - | The main topic, content, or finalized RAG research text. |
| `target_format` | `string` | No | `"pptx"` | Must be from the `target_format` enum. |
| `presentation_style`| `string` | No | `"Detailed"` | Must be from the `presentation_style` enum. |
| `layout_theme` | `string` | No | `"Modern Light"` | Must be from the `layout_theme` enum. |
| `num_slides` | `integer`| No | `5` | Number of slides to generate (if `target_format` is pptx). |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload upon completion. |
| `api_key` | `string` | No | `""` | Optional AI provider API key (Gemini/Anthropic). |

### Tool: `process_pdf`
**Description:** Converts a PDF document into an editable PPTX presentation or formatted DOCX.

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `pdf_source` | `string` | **Yes** | - | URL, local file path, or raw Base64 string of the PDF. |
| `is_url` | `boolean`| No | `true` | Set `true` if `pdf_source` is URL/path, `false` if Base64. |
| `instructions` | `string` | No | `null` | Specific rules for conversion (e.g., "Summarize tables"). |
| `target_format` | `string` | No | `"pptx"` | Must be from the `target_format` enum. |
| `layout_theme` | `string` | No | `null` | Overrides the PDF's implicit theme. Use enum values. |
| `visual_iconography`| `string` | No | `null` | Instructions for AI-generated icons (e.g., "flat vector"). |
| `slide_content_rules`| `string` | No | `null` | Rules for text volume per slide. |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload. |
| `api_key` | `string` | No | `""` | Optional AI provider API key. |

### Tool: `apply_docx_template`
**Description:** Reformats an existing DOCX document to adhere to strict corporate branding (Aptos Narrow, standardized headings).

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `doc_source` | `string` | **Yes** | - | URL, local file path, or raw Base64 string of the DOCX. |
| `is_url` | `boolean`| No | `true` | Set `true` if URL/path, `false` if Base64. |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload. |

### Tool: `get_capabilities`
**Description:** Returns the live JSON manifest of the server's capabilities, enums, and rules.
**Parameters:** None.

---

## 3. Response Schema Contract

Whether returned synchronously by the MCP Tool execution, or asynchronously via the `webhook_url` POST request, the JSON payload follows a strict schema.

### Success Response
```json
{
  "success": true,
  "message": "Successfully generated PPTX from prompt.",
  "file_url": "https://mcpserver-url.app/downloads/123e4567/generated_presentation.pptx",
  "download_path": "/downloads/123e4567/generated_presentation.pptx",
  "execution_id": "123e4567-e89b-12d3-a456-426614174000",
  "filename": "generated_presentation.pptx"
}
```
* **CRITICAL REQUIREMENT FOR RAG FRONTEND**: The RAG Agent must construct the download link for the user by combining its known base server URL with the `download_path` field. It should **not** directly render `file_url` if hallucination (e.g., `sandbox:/...`) is occurring within the LLM's output.

### Error Response
```json
{
  "success": false,
  "message": "Error generating from prompt: [Exception details]"
}
```

### Progress Webhook Schema (For long-running tasks)
If `webhook_url` is provided, the server will periodically send progress updates before sending the final success/error payload.
```json
{
  "status": "in_progress",
  "message": "Generating presentation outline with AI..."
}
```
*RAG agents can use this to update user-facing loading states.*