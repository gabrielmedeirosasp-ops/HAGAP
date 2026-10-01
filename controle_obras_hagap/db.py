import json
import sqlite3
from contextlib import contextmanager
from pathlib import Path

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "controle_obras.db"


@contextmanager
def conn():
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    c.execute("PRAGMA foreign_keys=ON")
    try:
        yield c
        c.commit()
    finally:
        c.close()


def init_db():
    with conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS emails (
                message_id TEXT PRIMARY KEY,
                thread_id TEXT,
                subject TEXT NOT NULL,
                sender TEXT,
                event_ts TEXT,
                gmail_url TEXT,
                kind TEXT,
                processed_at TEXT NOT NULL,
                raw_json TEXT
            );

            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                message_id TEXT NOT NULL,
                kind TEXT NOT NULL,
                project TEXT,
                project_base TEXT,
                ref_number TEXT,
                related_ref TEXT,
                status TEXT,
                municipality TEXT,
                event_ts TEXT,
                subject TEXT,
                gmail_url TEXT,
                attachment_name TEXT,
                details_json TEXT,
                UNIQUE(message_id, kind, project, ref_number, attachment_name),
                FOREIGN KEY(message_id) REFERENCES emails(message_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_events_project ON events(project_base);
            CREATE INDEX IF NOT EXISTS idx_events_kind ON events(kind);
            CREATE INDEX IF NOT EXISTS idx_events_ref ON events(ref_number);
            CREATE INDEX IF NOT EXISTS idx_events_related_ref ON events(related_ref);

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );

            CREATE TABLE IF NOT EXISTS sync_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL,
                scanned INTEGER DEFAULT 0,
                inserted INTEGER DEFAULT 0,
                errors INTEGER DEFAULT 0,
                note TEXT
            );
            """
        )


def get_setting(key, default=None):
    with conn() as c:
        row = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default


def set_setting(key, value):
    with conn() as c:
        c.execute(
            """
            INSERT INTO settings(key, value) VALUES(?, ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value
            """,
            (key, str(value)),
        )


def email_exists(message_id):
    with conn() as c:
        return c.execute(
            "SELECT 1 FROM emails WHERE message_id=?", (message_id,)
        ).fetchone() is not None


def save_email(meta, raw_json):
    with conn() as c:
        c.execute(
            """
            INSERT OR IGNORE INTO emails(
                message_id, thread_id, subject, sender, event_ts,
                gmail_url, kind, processed_at, raw_json
            ) VALUES(?,?,?,?,?,?,?,?,?)
            """,
            (
                meta.get("message_id"),
                meta.get("thread_id"),
                meta.get("subject", ""),
                meta.get("sender", ""),
                meta.get("event_ts", ""),
                meta.get("gmail_url", ""),
                meta.get("kind", ""),
                meta.get("processed_at", ""),
                json.dumps(raw_json, ensure_ascii=False),
            ),
        )


def save_event(event):
    project = event.get("project")
    project_base = project[:7] if project and len(project) >= 7 else project

    with conn() as c:
        c.execute(
            """
            INSERT OR IGNORE INTO events(
                message_id, kind, project, project_base, ref_number, related_ref,
                status, municipality, event_ts, subject, gmail_url,
                attachment_name, details_json
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                event.get("message_id"),
                event.get("kind"),
                project,
                project_base,
                event.get("ref_number"),
                event.get("related_ref"),
                event.get("status"),
                event.get("municipality"),
                event.get("event_ts"),
                event.get("subject"),
                event.get("gmail_url"),
                event.get("attachment_name"),
                json.dumps(event.get("details", {}), ensure_ascii=False),
            ),
        )


def resolve_omb_projects():
    with conn() as c:
        ombs = c.execute(
            """
            SELECT id, related_ref FROM events
            WHERE kind='OMB' AND (project IS NULL OR project='')
              AND related_ref IS NOT NULL AND related_ref<>''
            """
        ).fetchall()

        for omb in ombs:
            pde = c.execute(
                """
                SELECT project, project_base
                FROM events
                WHERE kind='PDE' AND ref_number=?
                  AND project IS NOT NULL AND project<>''
                ORDER BY event_ts DESC
                LIMIT 1
                """,
                (omb["related_ref"],),
            ).fetchone()
            if pde:
                c.execute(
                    "UPDATE events SET project=?, project_base=? WHERE id=?",
                    (pde["project"], pde["project_base"], omb["id"]),
                )


def start_sync(started_at):
    with conn() as c:
        cur = c.execute(
            "INSERT INTO sync_log(started_at,status) VALUES(?, 'running')",
            (started_at,),
        )
        return cur.lastrowid


def finish_sync(sync_id, finished_at, status, scanned, inserted, errors, note=""):
    with conn() as c:
        c.execute(
            """
            UPDATE sync_log
            SET finished_at=?, status=?, scanned=?, inserted=?, errors=?, note=?
            WHERE id=?
            """,
            (finished_at, status, scanned, inserted, errors, note, sync_id),
        )


def latest_sync():
    with conn() as c:
        row = c.execute(
            "SELECT * FROM sync_log ORDER BY id DESC LIMIT 1"
        ).fetchone()
        return dict(row) if row else None


def dashboard_rows():
    with conn() as c:
        rows = c.execute(
            """
            WITH projetos AS (
                SELECT DISTINCT project_base AS projeto
                FROM events
                WHERE project_base IS NOT NULL
                  AND length(project_base)=7
            )
            SELECT
                p.projeto,
                MAX(CASE WHEN e.kind='DOC_FINAL' THEN e.event_ts END) AS documentos_data,
                MAX(CASE WHEN e.kind='DOC_PARCIAL' THEN e.event_ts END) AS parcial_data,
                COUNT(DISTINCT CASE WHEN e.kind='PDE' THEN e.ref_number END) AS qtd_pde,
                COUNT(DISTINCT CASE WHEN e.kind='PLV' THEN e.ref_number END) AS qtd_plv,
                COUNT(DISTINCT CASE WHEN e.kind='OMB' THEN e.ref_number END) AS qtd_omb,
                COUNT(CASE WHEN e.kind='BMD' THEN 1 END) AS qtd_bmd,
                COUNT(CASE WHEN e.kind='FFO' THEN 1 END) AS qtd_ffo,
                MAX(CASE WHEN e.kind='PDE' THEN e.event_ts END) AS pde_data,
                MAX(CASE WHEN e.kind='PLV' THEN e.event_ts END) AS plv_data,
                MAX(CASE WHEN e.kind='OMB' THEN e.event_ts END) AS omb_data,
                MAX(CASE WHEN e.kind='BMD' THEN e.event_ts END) AS bmd_data,
                MAX(CASE WHEN e.kind='FFO' THEN e.event_ts END) AS ffo_data,
                MAX(CASE WHEN e.municipality IS NOT NULL AND e.municipality<>'' THEN e.municipality END) AS municipio
            FROM projetos p
            LEFT JOIN events e ON e.project_base=p.projeto
            GROUP BY p.projeto
            ORDER BY COALESCE(documentos_data, parcial_data, pde_data, plv_data, omb_data, bmd_data, ffo_data) DESC,
                     p.projeto DESC
            """
        ).fetchall()

    result = []
    for r in rows:
        d = dict(r)
        d["documentos"] = bool(d["documentos_data"])
        d["parcial"] = bool(d["parcial_data"])
        d["programada"] = bool(d["qtd_pde"] or d["qtd_plv"] or d["qtd_omb"])
        d["medida"] = bool(d["qtd_bmd"])
        d["fechada"] = bool(d["qtd_ffo"])
        result.append(d)
    return result


def project_events(project_base):
    with conn() as c:
        rows = c.execute(
            """
            SELECT *
            FROM events
            WHERE project_base=?
            ORDER BY event_ts DESC, id DESC
            """,
            (project_base,),
        ).fetchall()
    out = []
    for row in rows:
        d = dict(row)
        try:
            d["details"] = json.loads(d.pop("details_json") or "{}")
        except Exception:
            d["details"] = {}
            d.pop("details_json", None)
        out.append(d)
    return out


def summary():
    rows = dashboard_rows()
    return {
        "total": len(rows),
        "concluidas": sum(1 for x in rows if x["documentos"]),
        "parciais": sum(1 for x in rows if x["parcial"] and not x["documentos"]),
        "programadas": sum(1 for x in rows if x["programada"]),
        "com_bmd": sum(1 for x in rows if x["medida"]),
        "com_ffo": sum(1 for x in rows if x["fechada"]),
    }
