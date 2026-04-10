# RAG to MCP Server Interface Contract

This document defines the strict API contract and data schemas for communication between external Retrieval-Augmented Generation (RAG) agents and the `ppt-doc-generator` MCP Server.

To ensure the RAG agents and the MCP server are perfectly in sync, any RAG agent calling this server must adhere to the parameters and enumerations defined below.

---

## 1. Allowed Enumerations (Standardization)

Whenever specifying themes, styles, or formats, the RAG agent MUST use one of the following exact string values (unless noted, theme matching is **substring-based** on the lowercase string — e.g. `"My Deck — Studio Dark"` still resolves to Studio Dark colors).

### `target_format`

* `"pptx"` (PowerPoint Presentation)
* `"docx"` (Word Document)

### `layout_theme`

**Recommended default (API/MCP):** `"Studio Light"` — neutral light canvas (cool grey-white background, indigo header/footer ribbons, strong title and punchline contrast). This is the canonical default when the caller omits `layout_theme` where a default applies.

**Default dark pair:** `"Studio Dark"` (same palette as `"Presentation Dark"`) — slate background, amber accents, warm titles, green punchlines.

| Value | Resolved palette (summary) |
| :--- | :--- |
| `"Studio Light"` | Same as `"Presentation Light"`: cool neutral light background, indigo ribbons, slate titles, green punchline accent |
| `"Presentation Light"` | Same as Studio Light |
| `"Studio Dark"` | Same as `"Presentation Dark"`: dark slate canvas, amber/green punchline styling |
| `"Presentation Dark"` | Same as Studio Dark |
| `"Modern Light"` | Crisp very-light blue/white, red accent ribbons, dark text |
| `"Dark Corporate"` | Deep slate grey, vivid blue accent, light text |
| `"Pastel"` | Soft cream, mint accent, soft dark grey text |
| `"Blue Accent"` | Pure white, navy accents (substring `"blue"` also maps here) |

**Legacy alias:** if the theme string contains `voiceqa` (case-insensitive), it resolves to the **dark** Studio/Presentation Dark palette. Prefer `"Studio Dark"` or `"Presentation Dark"` for new integrations.

**Optional two-column visual layout:** the default for prompt-generated decks is **row infographic** (one AI icon per bullet plus a larger hero icon). To request the **split** layout (text left, icon grid + hero right), include a phrase in `layout_theme` such as: `split layout`, `split-panel`, `two-panel`, `two-column visual`, or `split visual` (see server implementation for the full keyword set).

**Fallback:** unknown or empty theme strings fall back to **Presentation Light** colors (same as Studio Light).

### `presentation_style`

* `"Detailed"` (Comprehensive bullet points, deep context)
* `"Executive"` (High-level summaries, metric-focused)
* `"Abstract"` (Conceptual, minimal text)
* `"Minimalist"` (Very sparse text, high reliance on iconography/layout)

---

## 2. MCP Tool Specifications

### Tool: `generate_from_prompt`

**Description:** Dynamically generates a full presentation (PPTX) or document (DOCX) from an abstract text prompt.  
**Type:** Synchronous (returns JSON) & asynchronous (posts to `webhook_url`).

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `prompt` | `string` | **Yes** | — | The main topic, content, or finalized RAG research text. |
| `target_format` | `string` | No | `"pptx"` | Must be from the `target_format` enum. |
| `presentation_style` | `string` | No | `"Detailed"` | Must be from the `presentation_style` enum. |
| `layout_theme` | `string` | No | `"Studio Light"` | Theme string; see `layout_theme` enum and substring rules above. |
| `num_slides` | `integer` | No | `5` | Number of slides to generate (if `target_format` is `pptx`). |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload upon completion. |
| `api_key` | `string` | No | `""` | Optional AI provider API key (Gemini/Anthropic). |

### Tool: `process_pdf`

**Description:** Converts a PDF document into an editable PPTX presentation or formatted DOCX.

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `pdf_source` | `string` | **Yes** | — | URL, local file path, or raw Base64 string of the PDF. |
| `is_url` | `boolean` | No | `true` | Set `true` if `pdf_source` is URL/path, `false` if Base64. |
| `instructions` | `string` | No | `null` | Specific rules for conversion (e.g. "Summarize tables"). |
| `target_format` | `string` | No | `"pptx"` | Must be from the `target_format` enum. |
| `layout_theme` | `string` | No | `""` | Theme string; empty string uses server fallback (Presentation Light / Studio Light palette). |
| `visual_iconography` | `string` | No | `null` | Instructions for AI-generated icons (e.g. "flat vector"). |
| `slide_content_rules` | `string` | No | `null` | Rules for text volume per slide. |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload. |
| `api_key` | `string` | No | `""` | Optional AI provider API key. |

### Tool: `image_to_pptx`

**Description:** Converts an image into an editable PPTX (vision extraction plus reconstruction).  
**Parameters:** `image_source`, `is_url`, optional `webhook_url`, `api_key`, `layout_theme` (default **`"Studio Light"`** if omitted when the tool supplies a default).

### Tool: `apply_docx_template`

**Description:** Reformats an existing DOCX document to adhere to strict corporate branding (Aptos Narrow, standardized headings).

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `doc_source` | `string` | **Yes** | — | URL, local file path, or raw Base64 string of the DOCX. |
| `is_url` | `boolean` | No | `true` | Set `true` if URL/path, `false` if Base64. |
| `webhook_url` | `string` | No | `null` | URL to receive the final JSON payload. |

### Tool: `get_capabilities`

**Description:** Returns the live JSON manifest of the server's capabilities, enums, and rules (including an expanded `supported_themes` list).  
**Parameters:** None.

---

## 3. Strict Presentation Standardization Protocol

The MCP Server implements a highly restrictive visual engine for all generated PPTX files. The RAG agent **must be aware** that the output will be strictly formatted according to these rules:

1. **Slide anatomy**
   - Every content slide includes:
     - **Title** (top; single-line style where enforced)
     - **Narrative** (1–2 lines below the title)
     - **Body** (bullets/table as generated)
     - **Punchline** (bottom)
2. **AI-generated icons (strict prompt-generated decks)**
   - Content slides require **`icon_keyword`** (hero icon seed) and **`bullet_icon_seeds`** (one seed per bullet). Icons are rendered via **DiceBear** (deterministic PNGs from seeds), with **one icon per bullet row** and a **larger hero icon** unless the theme requests **split layout** (see `layout_theme` above).
3. **Slide flow**
   - Decks follow an approved layout plan (title, sections, index, content types as generated).
4. **Typography**
   - The engine uses **Aptos Narrow** broadly across deck elements.
5. **Theming**
   - Supported palettes include **Studio Light / Studio Dark** (default pair), **Presentation Light / Presentation Dark** (equivalent colors), plus **Modern Light**, **Dark Corporate**, **Pastel**, and **Blue Accent**. Backgrounds, ribbons, and text colors adjust per resolved theme.
   - If the user does not specify a theme, agents should prefer **`"Studio Light"`** (or rely on server defaults).
6. **No corporate branding**
   - The presentation engine is generic. It does **not** inject "JPL", "JEMP", or other specific corporate branding unless the user asked. It remains white-labeled.

---

## 4. Response Schema Contract

Whether returned synchronously by the MCP tool execution, or asynchronously via the `webhook_url` POST request, the JSON payload follows a strict schema.

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

* **CRITICAL REQUIREMENT FOR RAG FRONTEND**: The RAG agent must construct the download link for the user by combining its known base server URL with the `download_path` field. It should **not** directly render `file_url` if hallucination (e.g. `sandbox:/...`) is occurring within the LLM's output.

### Error Response

```json
{
  "success": false,
  "message": "Error generating from prompt: [Exception details]"
}
```

### Progress Webhook Schema (for long-running tasks)

If `webhook_url` is provided, the server will periodically send progress updates before sending the final success/error payload.

```json
{
  "status": "in_progress",
  "message": "Generating presentation outline with AI..."
}
```

*RAG agents can use this to update user-facing loading states.*
