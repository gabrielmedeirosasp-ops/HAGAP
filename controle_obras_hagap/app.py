import json
import threading
import time
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, send_from_directory

import db
import gmail_sync

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
CONFIG_PATH = BASE / "config.json"

app = Flask(__name__, template_folder="templates", static_folder="static")
sync_lock = threading.Lock()
runtime = {
    "running": False,
    "last_error": "",
    "last_result": None,
}


def load_config():
    cfg = {
        "port": 5055,
        "sync_interval_minutes": 10,
        "history_start": "2025-01-01",
    }
    if CONFIG_PATH.exists():
        try:
            cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def sync_once():
    if not sync_lock.acquire(blocking=False):
        return {"ok": False, "status": "em_execucao"}

    runtime["running"] = True
    runtime["last_error"] = ""

    try:
        result = gmail_sync.run_sync(interactive=False)
        runtime["last_result"] = result
        return result
    except Exception as exc:
        runtime["last_error"] = str(exc)
        return {"ok": False, "status": "erro", "erro": str(exc)}
    finally:
        runtime["running"] = False
        sync_lock.release()


def sync_worker():
    # Dá tempo para o servidor iniciar antes da primeira sincronização.
    time.sleep(3)

    while True:
        cfg = load_config()
        interval = max(2, int(cfg.get("sync_interval_minutes", 10)))

        if gmail_sync.gmail_configured():
            result = sync_once()
            print(f"[ROTINA AUTOMÁTICA] {result}", flush=True)
        else:
            runtime["last_error"] = (
                "Gmail ainda não conectado. Execute conectar_gmail.bat uma única vez."
            )

        time.sleep(interval * 60)


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/logo-hagap")
def logo_hagap():
    # Reutiliza a logo do projeto atual sem alterar nem duplicar o site antigo.
    return send_from_directory(ROOT / "static", "logo.png")


@app.get("/api/resumo")
def api_resumo():
    return jsonify(
        {
            "resumo": db.summary(),
            "ultima_sync": db.latest_sync(),
            "runtime": {
                "running": runtime["running"],
                "last_error": runtime["last_error"],
                "gmail_configured": gmail_sync.gmail_configured(),
            },
        }
    )


@app.get("/api/projetos")
def api_projetos():
    return jsonify(db.dashboard_rows())


@app.get("/api/projeto/<projeto>")
def api_projeto(projeto):
    projeto = "".join(c for c in projeto if c.isdigit())[:7]
    if len(projeto) != 7:
        return jsonify({"erro": "Projeto inválido"}), 400
    return jsonify(
        {
            "projeto": projeto,
            "eventos": db.project_events(projeto),
        }
    )


@app.post("/api/sincronizar")
def api_sincronizar():
    if runtime["running"]:
        return jsonify({"ok": False, "status": "em_execucao"}), 409

    # Roda em thread para não travar a página durante a primeira carga histórica.
    def runner():
        result = sync_once()
        print(f"[SINCRONIZAÇÃO MANUAL] {result}", flush=True)

    threading.Thread(target=runner, daemon=True).start()
    return jsonify({"ok": True, "status": "iniciada"})


@app.get("/api/health")
def api_health():
    return jsonify(
        {
            "ok": True,
            "hora": datetime.now().isoformat(timespec="seconds"),
            "gmail_configured": gmail_sync.gmail_configured(),
            "sync_running": runtime["running"],
        }
    )


def main():
    db.init_db()
    cfg = load_config()

    threading.Thread(target=sync_worker, daemon=True, name="gmail-sync").start()

    print("=" * 72)
    print("HAGAP — CONTROLE DE OBRAS")
    print("Projeto independente. O site HAGAP antigo não é alterado.")
    print(f"Painel: http://127.0.0.1:{cfg['port']}")
    print(f"Sincronização automática: a cada {cfg['sync_interval_minutes']} min")
    print("=" * 72)

    app.run(
        host="127.0.0.1",
        port=int(cfg["port"]),
        debug=False,
        use_reloader=False,
    )


if __name__ == "__main__":
    main()
