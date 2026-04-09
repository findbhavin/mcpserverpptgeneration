# Fixing the Sandbox Links in the RAG Frontend

When the RAG framework calls our MCP Server, the MCP server responds with a JSON payload that looks exactly like this:

```json
{
  "success": true,
  "message": "Successfully generated PPTX from prompt.",
  "file_url": "https://mcpserverpptgeneration-707676665280.europe-west1.run.app/downloads/a1b2c3d4/generated_presentation.pptx",
  "execution_id": "a1b2c3d4",
  "filename": "generated_presentation.pptx"
}
```

The error where the user sees a broken link like `sandbox:/mnt/data/AI_Adoption_in_2026.pptx` is happening because **Claude is hallucinating a local file path**, and the RAG frontend is rendering exactly what Claude outputs.

Here is the exact code modification needed in the RAG framework (likely inside the `execute_tool` or `send_tool_result_to_claude` functions) to fix this:

## The Problem
Right now, the RAG framework calls the MCP Server, gets the JSON result, and immediately feeds the raw JSON back to Claude as a `tool_result`. Claude sees the URL but decides to output a Markdown response like: *"I have generated the presentation. You can download it here: [sandbox:/mnt/data/...]"*.

## The Solution
You must **intercept** the `file_url` from the MCP response, force Claude to use it, OR bypass Claude entirely and render a direct download button in the chat UI.

### Option 1: Provide the download_path back to the frontend

The MCP Server now responds with an explicit `download_path` field (e.g., `/downloads/execution_id/filename.pptx`).
Since your RAG agent already knows the `MCP_URL` it uses to connect to the MCP server, you can dynamically construct the absolute, reliable download URL directly in your RAG code, entirely bypassing Claude's output.

### Python Example:

```python
# Assuming you called an MCP tool and got this response string:
# '{"success": true, "file_url": "...", "download_path": "/downloads/abc-123/file.pptx", "execution_id": "..."}'

import json
mcp_response = json.loads(tool_result_string)

# The base URL your RAG agent uses to connect to the MCP server
MCP_BASE_URL = "https://mcpserverpptgeneration-707676665280.europe-west1.run.app"

if mcp_response.get("success") and "download_path" in mcp_response:
    reliable_url = f"{MCP_BASE_URL}{mcp_response['download_path']}"
    
    # Render this link directly in the chat UI
    st.markdown(f"**Success!** Download your presentation here: [Download PPTX]({reliable_url})")
```

## Option 2: Force Claude to use the exact URL via prompt injection
When you send the `tool_result` back to Claude, append explicit instructions telling it exactly how to format the Markdown link using the real `file_url`.

**Modify the RAG Python code where it handles the MCP HTTP response:**
```python
# Example RAG Frontend Code:
response = requests.post(MCP_URL + "/api/generate_from_prompt", json=block.input)
mcp_data = response.json()

if mcp_data.get("success") and mcp_data.get("file_url"):
    real_url = mcp_data["file_url"]
    filename = mcp_data.get("filename", "Presentation.pptx")
    
    # INJECT STRICT INSTRUCTIONS TO CLAUDE
    tool_result_content = (
        f"Tool execution succeeded. The file is available at the URL: {real_url}. "
        f"CRITICAL INSTRUCTION: When you reply to the user, you MUST provide exactly this Markdown link: "
        f"[{filename}]({real_url}) "
        f"Do NOT use 'sandbox:/mnt/data/' paths."
    )
else:
    tool_result_content = str(mcp_data)

send_tool_result_to_claude(tool_result_content)
```

### Option 2: Bypass Claude and Render it in the Chat UI (Better UX)
Instead of relying on Claude to print the link, intercept the success response and explicitly inject a system message or a clickable HTML button directly into the user's chat stream.

**Modify the RAG Python code:**
```python
# Example RAG Frontend Code:
response = requests.post(MCP_URL + "/api/generate_from_prompt", json=block.input)
mcp_data = response.json()

# 1. Send generic success back to Claude so it knows the task is done
send_tool_result_to_claude("Success. The presentation was generated and the download link has been provided to the user.")

# 2. Append a direct, un-hallucinated download card to the chat history
if mcp_data.get("success") and mcp_data.get("file_url"):
    real_url = mcp_data["file_url"]
    
    # Assuming your RAG uses standard chat history rendering:
    chat_history.append({
        "role": "system", 
        "content": f"✅ Generation Complete: <a href='{real_url}' target='_blank'>Download Presentation</a>"
    })
```

By explicitly grabbing the `"file_url"` property out of the JSON response and either hard-coding it into the chat UI (Option 2) or forcing Claude to use it verbatim (Option 1), the hallucinated `sandbox:/` links will disappear.