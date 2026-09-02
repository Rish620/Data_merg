from io import BytesIO
from fastapi import FastAPI, File, UploadFile, HTTPException
import pandas as pd

app = FastAPI(title="Reltio Configuration Automation Platform API", version="0.1.0")

@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "profiling-api"}

@app.post("/api/profile")
async def profile(file: UploadFile = File(...)) -> dict:
    if not file.filename or not file.filename.lower().endswith((".csv", ".xlsx", ".json")):
        raise HTTPException(status_code=400, detail="Upload a CSV, XLSX, or JSON file.")
    content = await file.read()
    try:
        suffix = file.filename.lower().rsplit(".", 1)[-1]
        if suffix == "csv":
            frame = pd.read_csv(BytesIO(content))
        elif suffix == "xlsx":
            frame = pd.read_excel(BytesIO(content))
        else:
            frame = pd.read_json(BytesIO(content))
    except Exception as exc:
        raise HTTPException(status_code=422, detail=f"Could not profile file: {exc}") from exc
    return {"filename": file.filename, "rows": len(frame), "columns": len(frame.columns), "attributes": list(frame.columns)}
