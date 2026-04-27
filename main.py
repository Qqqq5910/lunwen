import shutil
import uuid
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from analyzer.pipeline import analyze_docx
from analyzer.security import cleanup_old_files, max_upload_bytes, read_token
import os

app = FastAPI(title="论文格式终检助手 final")
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
UPLOAD_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    html_path = BASE_DIR / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


def save_upload(file, target):
    limit = max_upload_bytes()
    size = 0
    with target.open("wb") as buffer:
        while True:
            chunk = file.file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                raise HTTPException(status_code=413, detail="文件过大")
            buffer.write(chunk)


@app.post("/api/analyze")
async def analyze(thesis_file: UploadFile = File(...), school_requirement_file: UploadFile | None = File(default=None), fix_superscript: bool = Query(default=True), fix_citation_ranges: bool = Query(default=True), fix_school_format: bool = Query(default=False)):
    ttl_hours = int(os.getenv("FILE_TTL_HOURS", "24"))
    cleanup_old_files([UPLOAD_DIR, REPORT_DIR], ttl_hours)
    if not thesis_file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="论文目前只支持 .docx 文件")
    if school_requirement_file and not school_requirement_file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="学校格式要求目前只支持 .docx 文件")
    job_id = uuid.uuid4().hex
    job_upload_dir = UPLOAD_DIR / job_id
    job_upload_dir.mkdir(parents=True, exist_ok=True)
    thesis_path = job_upload_dir / Path(thesis_file.filename).name
    save_upload(thesis_file, thesis_path)
    school_path = None
    if school_requirement_file:
        school_path = job_upload_dir / Path(school_requirement_file.filename).name
        save_upload(school_requirement_file, school_path)
    return analyze_docx(str(thesis_path), school_requirement_path=str(school_path) if school_path else None, fix_superscript=fix_superscript, fix_citation_ranges=fix_citation_ranges, fix_school=fix_school_format, job_id=job_id, report_base_dir=REPORT_DIR)


def check_token(job_id, token):
    expected = read_token(REPORT_DIR / job_id)
    if not expected or not token or token != expected:
        raise HTTPException(status_code=403, detail="下载链接无效或已过期")


@app.get("/download/{job_id}/fixed/{filename}")
def download_fixed(job_id: str, filename: str, token: str):
    check_token(job_id, token)
    target = REPORT_DIR / job_id / "fixed" / filename
    if not target.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(str(target), filename=filename)


@app.get("/download/{job_id}/report/{filename}")
def download_report(job_id: str, filename: str, token: str):
    check_token(job_id, token)
    target = REPORT_DIR / job_id / filename
    if not target.exists():
        raise HTTPException(status_code=404, detail="报告不存在")
    return FileResponse(str(target), filename=filename)


@app.get("/health")
def health():
    return {"status": "ok"}
