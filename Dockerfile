FROM python:3.12-slim

# Create a non-root user to run the service
RUN useradd -m mcpuser

WORKDIR /app

# Create a shared directory for outputs
RUN mkdir -p /app/outputs && chown mcpuser:mcpuser /app/outputs
ENV PPTX_OUTPUT_DIR=/app/outputs
ENV HOME=/tmp

# Install dependencies
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy the server code
COPY core.py mcp_server.py app.py docx_formatter.py /app/

# Give ownership to the non-root user
RUN chown -R mcpuser:mcpuser /app

# Drop to non-root user for security
USER mcpuser

# In a real environment, you might restrict network access using Docker networks 
# or run with gVisor via the Docker runtime flag: --runtime=runsc

# Run the FastAPI server which also mounts the MCP SSE endpoints
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1"]
