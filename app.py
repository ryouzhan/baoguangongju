import os
import shutil
import uuid
import zipfile
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from core.config import (
    load_columns_config, save_columns_config,
    load_exchange_rate, save_exchange_rate,
    DEFAULT_CONFIG
)
from core.airscript import fetch_purchase_from_airscript
from core.customs_doc import generate_customs_document
from core.fba_merger import FbaCustomsMerger
from core.declaration_generator import process_all_fba_declarations

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMP_DIR = os.path.join(BASE_DIR, "temp_uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI(title="报关数据整合系统 Web 版", version="2.1.0")

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

FBA_SESSIONS = {}

@app.get("/", response_class=HTMLResponse)
async def read_index():
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Web Customs Tool - Loading...</h1>"

@app.get("/api/config")
async def get_config():
    cfg = load_columns_config()
    rate = load_exchange_rate()
    return {"status": "ok", "config": cfg, "exchange_rate": rate}

class ConfigUpdateRequest(BaseModel):
    config: dict
    exchange_rate: Optional[float] = 7.2

@app.post("/api/config")
async def update_config(data: ConfigUpdateRequest):
    try:
        saved_cfg = save_columns_config(data.config)
        if data.exchange_rate is not None:
            save_exchange_rate(float(data.exchange_rate))
        return {"status": "ok", "message": "配置已成功保存！", "config": saved_cfg}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/airscript/test")
async def test_airscript():
    try:
        df = fetch_purchase_from_airscript()
        return {
            "status": "ok",
            "message": f"连接成功！从 AirScript 云端获取到 {len(df)} 条采购单记录",
            "count": len(df),
            "columns": df.columns.tolist()
        }
    except Exception as e:
        return {"status": "error", "message": f"连接 AirScript 失败: {str(e)}"}

# ----------------- 功能 1: 报关资料在线生成 -----------------
@app.post("/api/customs_doc/generate")
async def api_generate_customs_doc(
    delivery_file: UploadFile = File(...),
    exchange_rate: float = Form(7.2)
):
    session_id = str(uuid.uuid4())[:8]
    work_dir = os.path.join(OUTPUT_DIR, f"customs_{session_id}")
    os.makedirs(work_dir, exist_ok=True)

    input_delivery_path = os.path.join(work_dir, delivery_file.filename)
    with open(input_delivery_path, "wb") as f:
        shutil.copyfileobj(delivery_file.file, f)

    logs = []
    def log_cb(msg):
        logs.append(msg)

    try:
        out_filepath, stats = generate_customs_document(
            delivery_file_path=input_delivery_path,
            output_dir=work_dir,
            exchange_rate=exchange_rate,
            status_logger=log_cb
        )
        return {
            "status": "ok",
            "stats": stats,
            "logs": logs,
            "download_url": f"/api/download?path={out_filepath}&name={stats['filename']}"
        }
    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
            "logs": logs
        }

# ----------------- 功能 2: FBA 数据合并 -----------------
@app.post("/api/fba_merger/load")
async def api_fba_merger_load(source_file: UploadFile = File(...)):
    session_id = str(uuid.uuid4())[:8]
    sess_dir = os.path.join(TEMP_DIR, f"fba_{session_id}")
    os.makedirs(sess_dir, exist_ok=True)

    file_path = os.path.join(sess_dir, source_file.filename)
    with open(file_path, "wb") as f:
        shutil.copyfileobj(source_file.file, f)

    merger = FbaCustomsMerger()
    ok, res = merger.load_file(file_path)
    if not ok:
        return {"status": "error", "message": res}

    FBA_SESSIONS[session_id] = merger

    return {
        "status": "ok",
        "session_id": session_id,
        "row_count": res["row_count"],
        "address_count": res["address_count"],
        "address_list": res["address_list"],
        "fba_list": res["fba_list"]
    }

class FbaMergeExportRequest(BaseModel):
    session_id: str
    groups: List[List[str]]

@app.post("/api/fba_merger/export")
async def api_fba_merger_export(req: FbaMergeExportRequest):
    merger = FBA_SESSIONS.get(req.session_id)
    if not merger:
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传文件！")
    
    try:
        out_dir = os.path.join(OUTPUT_DIR, f"fba_merged_{req.session_id}")
        os.makedirs(out_dir, exist_ok=True)
        out_path, stats = merger.export_merged_excel(req.groups, out_dir)
        return {
            "status": "ok",
            "stats": stats,
            "download_url": f"/api/download?path={out_path}&name={stats['filename']}"
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

# ----------------- 功能 3: FBA 报关单套版生成与智能清理（合并原第四功能） -----------------
@app.post("/api/declaration/generate")
async def api_declaration_generate(
    template_file: UploadFile = File(...),
    data_file: UploadFile = File(...),
    auto_clean: Optional[bool] = Form(True)
):
    session_id = str(uuid.uuid4())[:8]
    work_dir = os.path.join(OUTPUT_DIR, f"decl_{session_id}")
    save_folder = os.path.join(work_dir, "declarations")
    os.makedirs(save_folder, exist_ok=True)

    tpl_path = os.path.join(work_dir, template_file.filename)
    with open(tpl_path, "wb") as f:
        shutil.copyfileobj(template_file.file, f)

    data_path = os.path.join(work_dir, data_file.filename)
    with open(data_path, "wb") as f:
        shutil.copyfileobj(data_file.file, f)

    logs = []
    def log_cb(msg):
        logs.append(msg)

    # 直接执行套版填充 + 空白行清理 + 合同号规范命名，一步到位
    ok, res = process_all_fba_declarations(
        template_path=tpl_path,
        data_path=data_path,
        save_folder=save_folder,
        status_callback=log_cb
    )

    if not ok:
        return {"status": "error", "message": str(res), "logs": logs}

    files = res["files"]
    if not files:
        return {"status": "error", "message": "未能生成任何报关单文件！", "logs": logs}

    if len(files) == 1:
        single_path = files[0]
        single_name = os.path.basename(single_path)
        return {
            "status": "ok",
            "total_fba": res["total_fba"],
            "success_count": res["success_count"],
            "file_names": [single_name],
            "logs": logs,
            "download_url": f"/api/download?path={single_path}&name={single_name}"
        }

    # 多文件自动打包为 ZIP 供一键下载
    zip_filename = f"报关资料_{session_id}.zip"
    zip_path = os.path.join(work_dir, zip_filename)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in files:
            arcname = os.path.basename(fpath)
            zf.write(fpath, arcname=arcname)

    return {
        "status": "ok",
        "total_fba": res["total_fba"],
        "success_count": res["success_count"],
        "file_names": [os.path.basename(f) for f in files],
        "logs": logs,
        "download_url": f"/api/download?path={zip_path}&name={zip_filename}"
    }

@app.get("/api/download")
async def download_file(path: str, name: str):
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="文件不存在或已被清理！")
    return FileResponse(
        path=path,
        filename=name,
        media_type="application/octet-stream"
    )

if __name__ == "__main__":
    import uvicorn
    print("🚀 启动报关数据整合系统 Web 版: http://127.0.0.1:8000")
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=False)
