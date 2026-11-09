import os, io, json, datetime
from fastapi import FastAPI, Body
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

APP_VERSION = "v5.0.0 (Полная сборка: Предпросмотр А4, Сквозной ГСМ, Профили, Роли, Карты DWG)"

app = FastAPI(title="ПУТЁВКА v5.0.0 — Редактор путевых листов (РПС 0345001)")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "data", "putevka_db.json")
INIT_FILE = os.path.join(BASE_DIR, "data", "app_data.json")

def load_valid_state():
    for path in (DB_FILE, INIT_FILE):
        if os.path.exists(path) and os.path.getsize(path) > 50:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and data.get("profiles"):
                    return data
            except Exception:
                pass
    return {}

@app.get("/")
def index():
    return FileResponse(
        os.path.join(BASE_DIR, "static", "index.html"),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"}
    )

@app.get("/api/version")
def get_version():
    return {
        "version": APP_VERSION,
        "release_date": "2026-11-08",
        "has_saved_db": os.path.exists(DB_FILE)
    }

@app.get("/api/state")
def get_state():
    return load_valid_state()

@app.post("/api/state")
def save_state(payload: dict = Body(...)):
    os.makedirs(os.path.dirname(DB_FILE), exist_ok=True)
    # Сохраняем существующие xls_schema и dwg_previews, если фронтенд не прислал их в облегченном пакете
    existing = load_valid_state()
    if "xls_schema" not in payload and "xls_schema" in existing:
        payload["xls_schema"] = existing["xls_schema"]
    if "dwg_previews" not in payload and "dwg_previews" in existing:
        payload["dwg_previews"] = existing["dwg_previews"]
    payload["appVersion"] = "v5.0.0"
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return {
        "status": "ok",
        "saved_at": datetime.datetime.now().isoformat(),
        "version": APP_VERSION
    }

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
