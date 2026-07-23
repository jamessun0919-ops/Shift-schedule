import json
import os
import re
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded
from slowapi.util import get_remote_address
from starlette.background import BackgroundTask

from backend.generator import generate_xlsx
from backend.heuristics import guess_axis_range, guess_template
from backend.parser import extract_top_preview, get_hidden_columns, parse_schedule

app = FastAPI(title="班表格式轉換工具 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 只限流實際做解析/產生檔案這類有運算成本的端點（analyze/preview/convert），
# 範本 CRUD 與 health check 是輕量 JSON 讀寫，不設限，避免擋到正常操作流程。
# 門檻設定為「濫用斷路器」而非一般使用配額：單一使用者/單一店家正常操作
# （分析+反覆調整欄位對照+存範本+下載）一次session實測約 4~8 次呼叫，
# 20 次/分鐘遠高於正常使用量，只會擋到明顯異常的重複呼叫。
limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20MB
ALLOWED_EXTENSIONS = {".xlsx", ".ods"}
_SAFE_ID_RE = re.compile(r"^[A-Za-z0-9_\-一-鿿]+$")


def _save_upload_to_temp(file: UploadFile) -> str:
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(400, f"不支援的檔案格式：{ext or '未知'}，僅支援 .xlsx / .ods")

    contents = file.file.read(MAX_UPLOAD_SIZE + 1)
    if not contents:
        raise HTTPException(400, "上傳檔案為空")
    if len(contents) > MAX_UPLOAD_SIZE:
        raise HTTPException(400, "檔案過大，上限 20MB")

    fd, tmp_path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(contents)
    return tmp_path


def _parse_template_json(template_json: str) -> dict:
    try:
        return json.loads(template_json)
    except json.JSONDecodeError as e:
        raise HTTPException(400, f"template JSON 格式錯誤：{e}")


def _validate_ids(scope_id: str, template_id: str = None):
    if not _SAFE_ID_RE.match(scope_id):
        raise HTTPException(400, "scope_id 含不合法字元")
    if template_id is not None and not _SAFE_ID_RE.match(template_id):
        raise HTTPException(400, "template_id 含不合法字元")


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.post("/api/analyze")
@limiter.limit("20/minute")
async def analyze(request: Request, file: UploadFile = File(...)):
    tmp_path = _save_upload_to_temp(file)
    try:
        template = guess_template(tmp_path)
        preview_rows = extract_top_preview(tmp_path, template["sheet_name"])
        hidden_cols = sorted(get_hidden_columns(tmp_path, template["sheet_name"]))
        try:
            # 座標軸範圍猜測是錦上添花的顯示用預設值，用猜出的範本試解析失敗
            # 不該讓整個 /api/analyze 跟著失敗，靜默退回預設 8~24 即可。
            parsed = parse_schedule(tmp_path, template)
            template["display"] = guess_axis_range(parsed["employees"])
        except Exception:
            template["display"] = {"axis_start": 8, "axis_end": 24}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(422, f"分析失敗：{e}")
    finally:
        os.remove(tmp_path)

    return {
        "suggested_template": template,
        "preview_rows": preview_rows,
        "hidden_cols": hidden_cols,
    }


@app.post("/api/preview")
@limiter.limit("20/minute")
async def preview(request: Request, file: UploadFile = File(...), template: str = Form(...)):
    tmp_path = _save_upload_to_temp(file)
    template_dict = _parse_template_json(template)
    try:
        result = parse_schedule(tmp_path, template_dict)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(422, f"解析失敗，請確認範本欄位對照：{e}")
    finally:
        os.remove(tmp_path)

    employees = result["employees"]

    return {
        "employees_count": len(employees),
        "anomalies": result["anomalies"],
        "is_healthy": result["is_healthy"],
        "employees": employees,
        "month": result["month"],
    }


@app.post("/api/convert")
@limiter.limit("20/minute")
async def convert(
    request: Request,
    file: UploadFile = File(...),
    template: str = Form(...),
):
    tmp_path = _save_upload_to_temp(file)
    template_dict = _parse_template_json(template)

    out_path = None
    try:
        result = parse_schedule(tmp_path, template_dict)

        out_fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(out_fd)

        generate_xlsx(result["employees"], out_path, template_dict, month=result["month"])
    except HTTPException:
        raise
    except Exception as e:
        if out_path and os.path.exists(out_path):
            os.remove(out_path)
        raise HTTPException(422, f"轉換失敗：{e}")
    finally:
        os.remove(tmp_path)

    return FileResponse(
        out_path,
        filename="converted.xlsx",
        background=BackgroundTask(os.remove, out_path),
    )


@app.get("/api/templates")
async def list_templates(scope_id: str):
    _validate_ids(scope_id)
    scope_dir = TEMPLATES_DIR / scope_id
    if not scope_dir.exists():
        return []

    results = []
    for f in sorted(scope_dir.glob("*.json")):
        data = json.loads(f.read_text(encoding="utf-8"))
        results.append({
            "template_id": data.get("template_id", f.stem),
            "template_name": data.get("template_name", f.stem),
            "updated_at": data.get("updated_at"),
        })
    return results


@app.get("/api/templates/{scope_id}/{template_id}")
async def get_template(scope_id: str, template_id: str):
    _validate_ids(scope_id, template_id)
    path = TEMPLATES_DIR / scope_id / f"{template_id}.json"
    if not path.exists():
        raise HTTPException(404, "找不到範本")
    return json.loads(path.read_text(encoding="utf-8"))


@app.post("/api/templates/{scope_id}")
async def create_template(scope_id: str, template: dict):
    _validate_ids(scope_id)
    scope_dir = TEMPLATES_DIR / scope_id
    scope_dir.mkdir(parents=True, exist_ok=True)

    template_id = uuid.uuid4().hex[:12]
    now = datetime.now(timezone.utc).isoformat()
    record = {
        **template,
        "template_id": template_id,
        "scope_id": scope_id,
        "created_at": now,
        "updated_at": now,
    }
    (scope_dir / f"{template_id}.json").write_text(
        json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {"template_id": template_id}


@app.put("/api/templates/{scope_id}/{template_id}")
async def update_template(scope_id: str, template_id: str, template: dict):
    _validate_ids(scope_id, template_id)
    path = TEMPLATES_DIR / scope_id / f"{template_id}.json"
    if not path.exists():
        raise HTTPException(404, "找不到範本")

    existing = json.loads(path.read_text(encoding="utf-8"))
    now = datetime.now(timezone.utc).isoformat()
    record = {
        **template,
        "template_id": template_id,
        "scope_id": scope_id,
        "created_at": existing.get("created_at", now),
        "updated_at": now,
    }
    path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"status": "ok"}


@app.delete("/api/templates/{scope_id}/{template_id}")
async def delete_template(scope_id: str, template_id: str):
    _validate_ids(scope_id, template_id)
    path = TEMPLATES_DIR / scope_id / f"{template_id}.json"
    if not path.exists():
        raise HTTPException(404, "找不到範本")
    path.unlink()
    return {"status": "ok"}


FRONTEND_DIR = BASE_DIR / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
