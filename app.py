from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import os
import uvicorn

from core import generate_presentation, image_to_presentation, format_document, process_pdf_to_artifacts, stats, OUTPUT_DIR
from mcp_server import mcp

app = FastAPI(title="PPTX Generator API", description="API and UI for generating PowerPoint presentations")

# Mount the MCP SSE application
mcp_starlette = mcp.sse_app()
app.mount("/mcp", mcp_starlette)

# Pydantic models for API
class GenerateRequest(BaseModel):
    python_code: str

class ImageRequest(BaseModel):
    image_source: str
    is_url: bool = True

class DocxRequest(BaseModel):
    doc_source: str
    is_url: bool = True

class PdfRequest(BaseModel):
    pdf_source: str
    is_url: bool = True
    instructions: str = ""
    layout_theme: str = ""
    visual_iconography: str = ""
    slide_content_rules: str = ""
    target_format: str = "pptx"

@app.get("/", response_class=HTMLResponse)
async def index():
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>PPTX Generator Dashboard</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; background-color: #f5f5f5; }}
            .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            h1 {{ color: #333; }}
            .stats {{ display: flex; gap: 20px; margin-bottom: 30px; }}
            .stat-box {{ flex: 1; background: #e9ecef; padding: 15px; border-radius: 6px; text-align: center; }}
            .stat-box h3 {{ margin: 0 0 10px 0; font-size: 14px; color: #666; }}
            .stat-box p {{ margin: 0; font-size: 24px; font-weight: bold; color: #333; }}
            textarea {{ width: 100%; height: 200px; margin-bottom: 15px; padding: 10px; font-family: monospace; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }}
            button {{ background: #007bff; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 16px; }}
            button:hover {{ background: #0056b3; }}
            .result {{ margin-top: 20px; padding: 15px; border-radius: 4px; display: none; }}
            .success {{ background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }}
            .error {{ background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>PPTX Generator Dashboard</h1>
            
            <div class="stats">
                <div class="stat-box">
                    <h3>Requests Received</h3>
                    <p>{stats['requests_received']}</p>
                </div>
                <div class="stat-box">
                    <h3>Successful Creations</h3>
                    <p>{stats['successful_creations']}</p>
                </div>
                <div class="stat-box">
                    <h3>Failed Creations</h3>
                    <p>{stats['failed_creations']}</p>
                </div>
            </div>

            <div style="display: flex; gap: 10px; margin-bottom: 20px;">
                <button id="tabCode" style="flex: 1; background: #007bff;" onclick="switchTab('code')">Generate from Code</button>
                <button id="tabImage" style="flex: 1; background: #6c757d;" onclick="switchTab('image')">Image to PPTX</button>
                <button id="tabDocx" style="flex: 1; background: #6c757d;" onclick="switchTab('docx')">Format DOCX Template</button>
                <button id="tabPdf" style="flex: 1; background: #6c757d;" onclick="switchTab('pdf')">Process PDF</button>
            </div>

            <div id="sectionCode">
                <h2>Trigger Presentation Creation</h2>
                <form id="generateForm">
                    <textarea id="pythonCode" placeholder="Enter python-pptx code here... Example:
from pptx import Presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = 'Hello World'
prs.save('output.pptx')"></textarea>
                    <button type="submit" id="submitBtn">Generate PPTX</button>
                </form>
            </div>

            <div id="sectionImage" style="display: none;">
                <h2>Convert Image to PPTX</h2>
                <form id="imageForm">
                    <input type="text" id="imageSource" placeholder="Enter Image URL or Data URI..." style="width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;" required>
                    <button type="submit" id="submitImageBtn">Convert to PPTX</button>
                </form>
            </div>

            <div id="sectionDocx" style="display: none;">
                <h2>Format Document (Apply Template)</h2>
                <form id="docxForm">
                    <input type="text" id="docxSource" placeholder="Enter DOCX File URL or Base64..." style="width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;" required>
                    <button type="submit" id="submitDocxBtn">Format Document</button>
                </form>
            </div>

            <div id="sectionPdf" style="display: none;">
                <h2>Process PDF to Presentation/Document</h2>
                <form id="pdfForm">
                    <input type="text" id="pdfSource" placeholder="Enter PDF File URL or Base64..." style="width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;" required>
                    
                    <textarea id="pdfInstructions" placeholder="Abstract Instructions (e.g., 'Extract financial tables only')..." style="height: 60px;"></textarea>
                    <input type="text" id="pdfLayoutTheme" placeholder="Layout Theme (e.g., 'Modern Corporate')" style="width: 100%; padding: 10px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;">
                    <input type="text" id="pdfIconography" placeholder="Visual Iconography (e.g., 'Flat design, tech icons')" style="width: 100%; padding: 10px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box;">
                    <textarea id="pdfContentRules" placeholder="Slide Content Rules (e.g., 'Max 5 bullets per slide')..." style="height: 60px;"></textarea>
                    
                    <select id="pdfTargetFormat" style="width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px;">
                        <option value="pptx">Generate PPTX</option>
                        <option value="docx">Generate DOCX</option>
                    </select>

                    <button type="submit" id="submitPdfBtn">Process PDF</button>
                </form>
            </div>

            <div id="resultBox" class="result"></div>
        </div>

        <script>
            function switchTab(tab) {{
                document.getElementById('sectionCode').style.display = tab === 'code' ? 'block' : 'none';
                document.getElementById('sectionImage').style.display = tab === 'image' ? 'block' : 'none';
                document.getElementById('sectionDocx').style.display = tab === 'docx' ? 'block' : 'none';
                document.getElementById('sectionPdf').style.display = tab === 'pdf' ? 'block' : 'none';
                document.getElementById('tabCode').style.background = tab === 'code' ? '#007bff' : '#6c757d';
                document.getElementById('tabImage').style.background = tab === 'image' ? '#007bff' : '#6c757d';
                document.getElementById('tabDocx').style.background = tab === 'docx' ? '#007bff' : '#6c757d';
                document.getElementById('tabPdf').style.background = tab === 'pdf' ? '#007bff' : '#6c757d';
                document.getElementById('resultBox').style.display = 'none';
            }}

            document.getElementById('generateForm').addEventListener('submit', async (e) => {{
                e.preventDefault();
                const btn = document.getElementById('submitBtn');
                const resultBox = document.getElementById('resultBox');
                const code = document.getElementById('pythonCode').value;
                
                if (!code) return;
                
                btn.disabled = true;
                btn.textContent = 'Generating...';
                resultBox.style.display = 'none';
                
                try {{
                    const response = await fetch('/api/generate', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ python_code: code }})
                    }});
                    
                    const data = await response.json();
                    
                    resultBox.style.display = 'block';
                    if (data.success) {{
                        resultBox.className = 'result success';
                        resultBox.innerHTML = `<strong>Success!</strong> Presentation generated. <br><a href="${{data.file_url}}" target="_blank">Download ${{data.filename || 'File'}}</a>`;
                        
                        setTimeout(() => window.location.reload(), 2000);
                    }} else {{
                        resultBox.className = 'result error';
                        resultBox.innerHTML = `<strong>Error!</strong><br><pre>${{data.message}}</pre>`;
                    }}
                }} catch (err) {{
                    resultBox.style.display = 'block';
                    resultBox.className = 'result error';
                    resultBox.textContent = 'Network error occurred.';
                }} finally {{
                    btn.disabled = false;
                    btn.textContent = 'Generate PPTX';
                }}
            }});

            document.getElementById('imageForm').addEventListener('submit', async (e) => {{
                e.preventDefault();
                const btn = document.getElementById('submitImageBtn');
                const resultBox = document.getElementById('resultBox');
                const source = document.getElementById('imageSource').value;
                
                if (!source) return;
                
                btn.disabled = true;
                btn.textContent = 'Converting...';
                resultBox.style.display = 'none';
                
                try {{
                    const response = await fetch('/api/image-to-pptx', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ image_source: source, is_url: true }})
                    }});
                    
                    const data = await response.json();
                    
                    resultBox.style.display = 'block';
                    if (data.success) {{
                        resultBox.className = 'result success';
                        resultBox.innerHTML = `<strong>Success!</strong> Presentation generated. <br><a href="${{data.file_url}}" target="_blank">Download ${{data.filename || 'File'}}</a>`;
                        
                        setTimeout(() => window.location.reload(), 2000);
                    }} else {{
                        resultBox.className = 'result error';
                        resultBox.innerHTML = `<strong>Error!</strong><br><pre>${{data.message}}</pre>`;
                    }}
                }} catch (err) {{
                    resultBox.style.display = 'block';
                    resultBox.className = 'result error';
                    resultBox.textContent = 'Network error occurred.';
                }} finally {{
                    btn.disabled = false;
                    btn.textContent = 'Convert to PPTX';
                }}
            }});
            document.getElementById('docxForm').addEventListener('submit', async (e) => {{
                e.preventDefault();
                const btn = document.getElementById('submitDocxBtn');
                const resultBox = document.getElementById('resultBox');
                const source = document.getElementById('docxSource').value;
                
                if (!source) return;
                
                btn.disabled = true;
                btn.textContent = 'Formatting...';
                resultBox.style.display = 'none';
                
                try {{
                    const response = await fetch('/api/format-docx', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ doc_source: source, is_url: true }})
                    }});
                    
                    const data = await response.json();
                    
                    resultBox.style.display = 'block';
                    if (data.success) {{
                        resultBox.className = 'result success';
                        resultBox.innerHTML = `<strong>Success!</strong> Document formatted. <br><a href="${{data.file_url}}" target="_blank">Download ${{data.filename || 'File'}}</a>`;
                        
                        setTimeout(() => window.location.reload(), 2000);
                    }} else {{
                        resultBox.className = 'result error';
                        resultBox.innerHTML = `<strong>Error!</strong><br><pre>${{data.message}}</pre>`;
                    }}
                }} catch (err) {{
                    resultBox.style.display = 'block';
                    resultBox.className = 'result error';
                    resultBox.textContent = 'Network error occurred.';
                }} finally {{
                    btn.disabled = false;
                    btn.textContent = 'Format Document';
                }}
            }});
            document.getElementById('pdfForm').addEventListener('submit', async (e) => {{
                e.preventDefault();
                const btn = document.getElementById('submitPdfBtn');
                const resultBox = document.getElementById('resultBox');
                
                const source = document.getElementById('pdfSource').value;
                const instructions = document.getElementById('pdfInstructions').value;
                const theme = document.getElementById('pdfLayoutTheme').value;
                const iconography = document.getElementById('pdfIconography').value;
                const rules = document.getElementById('pdfContentRules').value;
                const format = document.getElementById('pdfTargetFormat').value;
                
                if (!source) return;
                
                btn.disabled = true;
                btn.textContent = 'Processing...';
                resultBox.style.display = 'none';
                
                try {{
                    const response = await fetch('/api/process-pdf', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ 
                            pdf_source: source, 
                            is_url: true,
                            instructions: instructions,
                            layout_theme: theme,
                            visual_iconography: iconography,
                            slide_content_rules: rules,
                            target_format: format
                        }})
                    }});
                    
                    const data = await response.json();
                    
                    resultBox.style.display = 'block';
                    if (data.success) {{
                        resultBox.className = 'result success';
                        resultBox.innerHTML = `<strong>Success!</strong> File generated. <br><a href="${{data.file_url}}" target="_blank">Download ${{data.filename || 'File'}}</a>`;
                        setTimeout(() => window.location.reload(), 2000);
                    }} else {{
                        resultBox.className = 'result error';
                        resultBox.innerHTML = `<strong>Error!</strong><br><pre>${{data.message}}</pre>`;
                    }}
                }} catch (err) {{
                    resultBox.style.display = 'block';
                    resultBox.className = 'result error';
                    resultBox.textContent = 'Network error occurred.';
                }} finally {{
                    btn.disabled = false;
                    btn.textContent = 'Process PDF';
                }}
            }});
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.get("/api/stats")
async def get_stats():
    return stats

@app.post("/api/generate")
async def api_generate(request: GenerateRequest):
    return generate_presentation(request.python_code)

@app.post("/api/image-to-pptx")
async def api_image_to_pptx(request: ImageRequest):
    return image_to_presentation(request.image_source, request.is_url)

@app.post("/api/format-docx")
async def api_format_docx(request: DocxRequest):
    return format_document(request.doc_source, request.is_url)

@app.post("/api/process-pdf")
async def api_process_pdf(request: PdfRequest):
    return process_pdf_to_artifacts(
        request.pdf_source, 
        request.is_url, 
        request.instructions, 
        request.layout_theme, 
        request.visual_iconography, 
        request.slide_content_rules, 
        request.target_format
    )

@app.get("/downloads/{execution_id}/{filename}")
async def download_file(execution_id: str, filename: str):
    file_path = os.path.join(OUTPUT_DIR, execution_id, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=filename)
    return {"error": "File not found"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
