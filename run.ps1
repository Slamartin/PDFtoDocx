# Starts the converter on http://127.0.0.1:8765 and opens it in the browser.
Set-Location $PSScriptRoot
uv sync
Start-Process "http://127.0.0.1:8765"
uv run uvicorn app.main:app --host 127.0.0.1 --port 8765
