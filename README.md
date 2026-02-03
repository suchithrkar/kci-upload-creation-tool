# KCI Upload Creation Tool

Minimal FastAPI scaffold for an internal single-user web app.

## Requirements
- Python 3.11+

## Setup
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run
```bash
uvicorn app.main:app --reload
```

## Endpoints
- `GET /health`
- `GET /upload/status`
