import json
import os
import secrets
import time
from pathlib import Path


def generate_token():
    return secrets.token_urlsafe(32)


def write_token(job_dir, token):
    path = Path(job_dir) / "token.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"token": token}, ensure_ascii=False), encoding="utf-8")


def read_token(job_dir):
    path = Path(job_dir) / "token.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8")).get("token")
    except Exception:
        return None


def cleanup_old_files(paths, ttl_hours=24):
    now = time.time()
    ttl = ttl_hours * 3600
    for raw in paths:
        base = Path(raw)
        if not base.exists():
            continue
        for child in base.iterdir():
            try:
                if now - child.stat().st_mtime > ttl:
                    if child.is_dir():
                        shutil_rmtree(child)
                    else:
                        child.unlink(missing_ok=True)
            except Exception:
                continue


def shutil_rmtree(path):
    for sub in sorted(path.rglob("*"), reverse=True):
        if sub.is_file():
            sub.unlink(missing_ok=True)
        elif sub.is_dir():
            sub.rmdir()
    path.rmdir()


def max_upload_bytes():
    mb = int(os.getenv("MAX_UPLOAD_MB", "30"))
    return mb * 1024 * 1024
