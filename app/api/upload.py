from fastapi import APIRouter

router = APIRouter(prefix="/upload", tags=["upload"])


@router.get("/status")
def upload_status() -> dict[str, str]:
    return {"status": "not_implemented"}
