FROM python:3.12-slim

WORKDIR /app

RUN pip install --no-cache-dir "mcp[cli]>=1.27.0" "httpx>=0.27.0" "msal>=1.28.0" "uvicorn>=0.30.0" "pydantic-settings>=2.0.0"

COPY server.py .

EXPOSE 8000

CMD ["python", "server.py"]
