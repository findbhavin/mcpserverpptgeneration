# Agentic Pipeline & RAG Integration Guidelines

This document provides best practices and architectural guidelines for integrating this MCP Server into Agentic Workflows and Retrieval-Augmented Generation (RAG) pipelines.

## 1. Overview of the MCP Server Capabilities

This server exposes three powerful MCP tools designed to be seamlessly called by LLM-driven agents:

1. `generate_pptx(python_code: str)`: Dynamically generates a PowerPoint presentation by executing `python-pptx` code in an isolated sandbox.
2. `image_to_pptx(image_source: str, is_url: bool)`: Takes an image (e.g., an architecture diagram, chart, or generated graphic) and perfectly fits it into a 16:9 presentation slide.
3. `apply_docx_template(doc_source: str, is_url: bool)`: Ingests an existing `.docx` file and automatically reformats it to strictly comply with corporate guidelines (Aptos Narrow font, specific table margins, heading sizes, caption rules, etc.).

## 2. Recommended RAG Pipeline Architecture

When building an agentic pipeline (using frameworks like LangChain, AutoGen, CrewAI, or Claude Desktop) that utilizes this server, follow this standard flow:

### Phase 1: Retrieval & Context Gathering
* **Action**: The agent receives a user prompt (e.g., "Create a presentation on our Q3 Financial Results").
* **RAG Step**: The agent queries the vector database or knowledge graph to retrieve the necessary unstructured text, financial tables, and figures.

### Phase 2: Content Synthesis & Structuring
* **Action**: The agent processes the retrieved context and outlines the artifact.
* **Best Practice**: The agent should explicitly decide *which* artifact to generate:
  * If highly structured multi-slide content is needed -> **Use `generate_pptx`**
  * If a single high-quality chart/diagram was retrieved/generated -> **Use `image_to_pptx`**
  * If a long-form textual report was drafted -> **Use `apply_docx_template`**

### Phase 3: Tool Execution & Self-Correction
* **Action**: The agent writes the payload and invokes the relevant MCP Tool.
* **Self-Correction Loop**: The MCP server returns a detailed `message` containing `stdout` and `stderr` if execution fails. The agent MUST be instructed to read these errors, fix its payload (e.g., fixing a syntax error in the Python code), and re-invoke the tool.

### Phase 4: Delivery
* **Action**: The tool returns a `file_url` upon success.
* **Delivery**: The agent presents this URL to the user as a downloadable artifact.

---

## 3. Best Practices for Agents Using `generate_pptx`

Instruct your agents with the following prompt guidelines to ensure high success rates when generating Python code for PPTX creation:

1. **Always Use `python-pptx`**: The environment only supports the `python-pptx` library for presentation generation. Do not try to use `win32com` or other OS-dependent libraries.
2. **Mandatory Save Step**: The python script **MUST** end with saving the file to the current working directory.
   ```python
   # Correct
   prs.save("output.pptx")
   ```
3. **Handle Text Overflows**: Agents should be prompted to calculate or estimate text length to avoid overflowing slide boundaries. Use multiple slides if the RAG context is too long.
4. **Layout Usage**: Rely on standard layouts (0 = Title, 1 = Title and Content, 5 = Title Only, 6 = Blank).
5. **Data Visualization**: If the RAG context contains tabular data, the agent should use the `python-pptx` table shapes (`slide.shapes.add_table`) rather than trying to draw text boxes manually.

---

## 4. Best Practices for Agents Using `image_to_pptx`

1. **Source Reliability**: Ensure the `image_source` provided is publicly accessible if `is_url` is `true`. If the image is generated locally in the agent's memory (e.g., via a matplotlib Python execution), pass it as a Base64 string and set `is_url=False`.
2. **Aspect Ratios**: The MCP server automatically scales and centers the image for a 16:9 slide, but agents should ideally generate or crop images to 16:9 (e.g., 1920x1080) for the best visual result without letterboxing.

---

## 5. Best Practices for Agents Using `apply_docx_template`

1. **Decouple Content from Styling**: During the RAG generation phase, the agent should only focus on generating raw `.docx` content (e.g., generating markdown, then using `pypandoc` or standard libraries to create a basic DOCX). 
2. **Rely on the Server for Branding**: The agent should NOT spend tokens writing complex python code to style the document. Instead, it should pass the raw, unstyled `.docx` to `apply_docx_template`.
3. **Triggering Formats**: To ensure the MCP server formats items correctly, the agent should use standard Word constructs when creating the initial draft:
   * Use standard Headings (`Heading 1`, `Heading 2`, `Heading 3`).
   * Group images with their captions logically.
   * Insert standard Tables rather than tab-separated text.

## 6. Example Agent System Prompt Addition

If you are using Claude or an OpenAI model, inject this into the system prompt:

> "You have access to a Presentation and Document Generation MCP Server. 
> When asked to create a PowerPoint, synthesize your knowledge into a coherent outline, then write a `python-pptx` script to generate it. The script must save the file to 'output.pptx'. Pass this script to the `generate_pptx` tool.
> If the tool returns a Python syntax error or traceback, analyze the error, correct your Python code, and call the tool again.
> When asked to create a formal report, first create a basic DOCX file, then pass it to `apply_docx_template` to automatically enforce our corporate branding standards."
