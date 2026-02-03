from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.api.health import router as health_router
from app.api.upload import router as upload_router
from app.api.process import router as process_router

app = FastAPI(
    title="KCI Upload Creation Tool",
    version="0.1.0",
)

# Serve static UI
app.mount("/static", StaticFiles(directory="app/static"), name="static")

@app.get("/")
def index() -> FileResponse:
    return FileResponse(Path("app/static/index.html"))

# API routes
app.include_router(health_router)
app.include_router(upload_router)
app.include_router(process_router)