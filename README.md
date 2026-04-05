# MCP Server - PPTX Generation

This repository provides an MCP (Model Context Protocol) server for generating PowerPoint (.pptx) presentations using Python code.

It provides three ways to interact:
1. **MCP Protocol**: Can be connected as an MCP tool for agents (Claude, cursor, etc.) to use natively via the `/mcp/sse` endpoint.
2. **REST API**: Standard endpoints (`/api/generate`, `/api/stats`) for integration with other frameworks.
3. **Web UI**: A simple, intuitive web interface to track statistics and generate presentations manually.

## Features

- Built with FastAPI & FastMCP.
- Highly isolated environment for executing python-pptx generation code.
- GCP Cloud Run ready.

## Local Development

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the server:**
   ```bash
   uvicorn app:app --reload
   ```

   The web UI will be available at [http://localhost:8000](http://localhost:8000).
   The MCP SSE endpoint will be available at [http://localhost:8000/mcp/sse](http://localhost:8000/mcp/sse).

3. **Using Docker Compose:**
   ```bash
   docker-compose up --build
   ```

## Deploying to GCP Cloud Run

The easiest way to deploy is using the provided Dockerfile and Cloud Run:

```bash
gcloud builds submit --config cloudbuild.yaml
```

Or deploying manually:

```bash
gcloud run deploy pptx-generator \
  --source . \
  --allow-unauthenticated \
  --memory 2Gi \
  --region us-central1
```

## API Usage

### `POST /api/generate`
Generates a PPTX file given the provided python-pptx code.

**Request body:**
```json
{
  "python_code": "from pptx import Presentation\nprs = Presentation()\nslide = prs.slides.add_slide(prs.slide_layouts[0])\nslide.shapes.title.text = 'Hello World'\nprs.save('output.pptx')"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Presentation generated successfully.",
  "file_url": "/downloads/abc-123/output.pptx",
  "execution_id": "abc-123",
  "filename": "output.pptx"
}
```

### `POST /api/image-to-pptx`
Converts an image into a presentation with a single slide perfectly fitting the image. Useful for taking generated charts or diagrams and immediately wrapping them into slides.

**Request body:**
```json
{
  "image_source": "https://example.com/image.png",
  "is_url": true
}
```
*(You can also pass a base64 string directly with `is_url: false` or a base64 data URI with `is_url: true`)*

### `POST /api/process-pdf`
Processes a fresh PDF file (or Base64 content) and converts it into a `pptx` or `docx` format, optionally taking into account instructions around layout themes, visual iconography, and slide content rules.

**Request body:**
```json
{
  "pdf_source": "https://example.com/report.pdf",
  "is_url": true,
  "instructions": "Extract financial tables only",
  "layout_theme": "Modern Corporate",
  "visual_iconography": "Flat design, tech icons",
  "slide_content_rules": "Max 5 bullets per slide",
  "target_format": "pptx"
}
```
### `GET /api/stats`
Returns the current request and generation statistics.
