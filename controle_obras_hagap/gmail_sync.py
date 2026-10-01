import argparse
import base64
import io
import json
import os
import re
import time
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path

import fitz
import pdfplumber
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import db
import parsers

BASE = Path(__file__).resolve().parent
CONFIG_PATH = BASE / "config.json"
CREDENTIALS_PATH = BASE / "credentials.json"
TOKEN_PATH = BASE / "token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

DEFAULT_CONFIG = {
    "history_start": "2025-01-01",
    "sync_interval_minutes": 10,
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        try:
            cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def gmail_configured():
    return CREDENTIALS_PATH.exists() and TOKEN_PATH.exists()


def get_credentials(interactive=False):
    creds = None

    if TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except Exception:
            creds = None

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")

    if creds and creds.valid:
        return creds

    if not interactive:
        raise RuntimeError(
            "Gmail ainda não conectado. Execute conectar_gmail.bat uma única vez."
        )

    if not CREDENTIALS_PATH.exists():
        raise RuntimeError(
            "Arquivo credentials.json não encontrado na pasta controle_obras_hagap."
        )

    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return creds


def get_service(interactive=False):
    return build("gmail", "v1", credentials=get_credentials(interactive), cache_discovery=False)


def decode_b64url(data):
    if not data:
        return b""
    padding = "=" * (-len(data) % 4)
    return base64.urlsafe_b64decode(data + padding)


def headers_dict(payload):
    out = {}
    for h in payload.get("headers", []):
        out[h.get("name", "").lower()] = h.get("value", "")
    return out


def walk_parts(part):
    yield part
    for child in part.get("parts", []) or []:
        yield from walk_parts(child)


def html_to_text(value):
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value or "")
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = value.replace("&nbsp;", " ").replace("&amp;", "&")
    return " ".join(value.split())


def extract_body(payload):
    plain = []
    html = []

    for part in walk_parts(payload):
        mime = part.get("mimeType", "")
        data = (part.get("body") or {}).get("data")
        if not data:
            continue
        try:
            text = decode_b64url(data).decode("utf-8", errors="replace")
        except Exception:
            continue
        if mime == "text/plain":
            plain.append(text)
        elif mime == "text/html":
            html.append(html_to_text(text))

    if plain:
        return "\n".join(plain)
    if html:
        return "\n".join(html)
    return ""


def list_attachments(payload):
    found = []
    for part in walk_parts(payload):
        filename = part.get("filename") or ""
        body = part.get("body") or {}
        attachment_id = body.get("attachmentId")
        if filename and attachment_id:
            found.append(
                {
                    "filename": filename,
                    "mime_type": part.get("mimeType", ""),
                    "attachment_id": attachment_id,
                }
            )
    return found


def pdf_text(data):
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        text = ""

    if text.strip():
        return text

    try:
        doc = fitz.open(stream=data, filetype="pdf")
        return "\n".join(page.get_text("text") or "" for page in doc)
    except Exception:
        return ""


def list_message_ids(service, query):
    token = None
    while True:
        resp = (
            service.users()
            .messages()
            .list(userId="me", q=query, maxResults=500, pageToken=token)
            .execute()
        )
        for item in resp.get("messages", []) or []:
            yield item["id"]
        token = resp.get("nextPageToken")
        if not token:
            break


def incremental_after_date():
    cfg = load_config()
    history_start = datetime.strptime(cfg["history_start"], "%Y-%m-%d")

    last = db.get_setting("last_successful_sync")
    if not last:
        return history_start

    try:
        dt = datetime.fromisoformat(last)
    except Exception:
        return history_start

    # Reconsulta 3 dias para capturar reenvios/atrasos.
    dt = dt - timedelta(days=3)
    return max(history_start, dt)


def query_set(after_date):
    d = after_date.strftime("%Y/%m/%d")
    return [
        ("DOC", f'after:{d} in:sent subject:DOCUMENTOS'),
        ("PDE", f'after:{d} subject:"PDE número:"'),
        ("PLV", f'after:{d} subject:"PLV número:"'),
        ("OMB", f'after:{d} subject:"OMB Enviado Integração"'),
        ("MED", f'after:{d} {{subject:BMD subject:FFO}}'),
    ]


def parse_email_ts(headers, internal_date):
    if internal_date:
        try:
            return datetime.fromtimestamp(int(internal_date) / 1000).isoformat(timespec="seconds")
        except Exception:
            pass

    raw = headers.get("date")
    if raw:
        try:
            return parsedate_to_datetime(raw).replace(tzinfo=None).isoformat(timespec="seconds")
        except Exception:
            pass

    return datetime.now().isoformat(timespec="seconds")


def attachment_bytes(service, message_id, attachment_id):
    obj = (
        service.users()
        .messages()
        .attachments()
        .get(userId="me", messageId=message_id, id=attachment_id)
        .execute()
    )
    return decode_b64url(obj.get("data", ""))


def save_parsed_events(message_meta, parsed, attachment_name=""):
    projects = parsed.get("projects") or [None]

    for project in projects:
        event = {
            "message_id": message_meta["message_id"],
            "kind": parsed["kind"],
            "project": project,
            "ref_number": parsed.get("ref_number", ""),
            "related_ref": parsed.get("related_ref", ""),
            "status": parsed.get("status", ""),
            "municipality": parsed.get("municipality", ""),
            "event_ts": message_meta["event_ts"],
            "subject": message_meta["subject"],
            "gmail_url": message_meta["gmail_url"],
            "attachment_name": attachment_name,
            "details": parsed.get("details", {}),
        }
        db.save_event(event)


def process_message(service, message_id):
    if db.email_exists(message_id):
        return 0

    msg = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    payload = msg.get("payload") or {}
    headers = headers_dict(payload)
    subject = headers.get("subject", "")
    kind = parsers.classify_subject(subject)

    if not kind:
        return 0

    body = extract_body(payload)
    attachments = list_attachments(payload)
    attachment_names = [x["filename"] for x in attachments]

    event_ts = parse_email_ts(headers, msg.get("internalDate"))
    meta = {
        "message_id": message_id,
        "thread_id": msg.get("threadId", ""),
        "subject": subject,
        "sender": headers.get("from", ""),
        "event_ts": event_ts,
        "gmail_url": f"https://mail.google.com/mail/u/0/#all/{message_id}",
        "kind": kind,
        "processed_at": datetime.now().isoformat(timespec="seconds"),
    }

    db.save_email(
        meta,
        {
            "labelIds": msg.get("labelIds", []),
            "snippet": msg.get("snippet", ""),
            "attachment_names": attachment_names,
        },
    )

    inserted = 0

    if kind in ("DOC_FINAL", "DOC_PARCIAL"):
        parsed = parsers.parse_documents(kind, subject, body, attachment_names)
        save_parsed_events(meta, parsed)
        return max(1, len(parsed.get("projects") or []))

    pdfs = [
        a for a in attachments
        if a["filename"].lower().endswith(".pdf")
        or a["mime_type"].lower() == "application/pdf"
    ]

    if kind in ("PDE", "PLV", "OMB"):
        texts = []
        used_names = []
        for a in pdfs:
            data = attachment_bytes(service, message_id, a["attachment_id"])
            texts.append(pdf_text(data))
            used_names.append(a["filename"])
        joined = "\n\n".join(texts)

        if kind == "PDE":
            parsed = parsers.parse_pde(subject, joined)
        elif kind == "PLV":
            parsed = parsers.parse_plv(subject, joined)
        else:
            parsed = parsers.parse_omb(subject, joined)

        save_parsed_events(meta, parsed, " | ".join(used_names))
        return max(1, len(parsed.get("projects") or []))

    # BMD / FFO: um mesmo e-mail pode conter os dois tipos.
    if kind in ("BMD", "FFO", "BMD_FFO"):
        if not pdfs:
            parsed = parsers.parse_measurement(
                "FFO" if kind == "FFO" else "BMD",
                subject,
                "",
                "",
            )
            save_parsed_events(meta, parsed)
            return max(1, len(parsed.get("projects") or []))

        for a in pdfs:
            name_upper = parsers.norm(a["filename"])
            if "FFO" in name_upper or "FF0" in name_upper:
                attachment_kind = "FFO"
            elif "BMD" in name_upper:
                attachment_kind = "BMD"
            elif kind == "FFO":
                attachment_kind = "FFO"
            else:
                attachment_kind = "BMD"

            data = attachment_bytes(service, message_id, a["attachment_id"])
            text = pdf_text(data)
            parsed = parsers.parse_measurement(
                attachment_kind,
                subject,
                text,
                a["filename"],
            )
            save_parsed_events(meta, parsed, a["filename"])
            inserted += max(1, len(parsed.get("projects") or []))

    return inserted


def run_sync(interactive=False):
    db.init_db()
    started = datetime.now()
    sync_id = db.start_sync(started.isoformat(timespec="seconds"))

    scanned = 0
    inserted = 0
    errors = 0
    error_notes = []

    try:
        service = get_service(interactive=interactive)
        after_date = incremental_after_date()

        ids = []
        seen = set()
        for _, query in query_set(after_date):
            for message_id in list_message_ids(service, query):
                if message_id not in seen:
                    seen.add(message_id)
                    ids.append(message_id)

        print(f"[GMAIL] Mensagens candidatas: {len(ids)}")
        print(f"[GMAIL] Janela de busca a partir de: {after_date:%d/%m/%Y}")

        for idx, message_id in enumerate(ids, 1):
            scanned += 1
            try:
                n = process_message(service, message_id)
                inserted += n
            except Exception as exc:
                errors += 1
                error_notes.append(f"{message_id}: {exc}")
                print(f"[ERRO] {message_id}: {exc}")

            if idx % 25 == 0 or idx == len(ids):
                print(
                    f"[GMAIL] {idx}/{len(ids)} | eventos novos: {inserted} | erros: {errors}",
                    flush=True,
                )

        db.resolve_omb_projects()

        now = datetime.now()
        # Só avança o marcador quando a rotina terminou; emails já processados
        # são deduplicados por message_id.
        db.set_setting("last_successful_sync", now.isoformat(timespec="seconds"))

        status = "ok" if errors == 0 else "ok_com_erros"
        db.finish_sync(
            sync_id,
            now.isoformat(timespec="seconds"),
            status,
            scanned,
            inserted,
            errors,
            "\n".join(error_notes[:50]),
        )

        return {
            "ok": True,
            "status": status,
            "scanned": scanned,
            "inserted": inserted,
            "errors": errors,
        }

    except Exception as exc:
        now = datetime.now()
        db.finish_sync(
            sync_id,
            now.isoformat(timespec="seconds"),
            "erro",
            scanned,
            inserted,
            errors + 1,
            str(exc),
        )
        raise


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--auth-only", action="store_true")
    parser.add_argument("--sync", action="store_true")
    args = parser.parse_args()

    db.init_db()

    if args.auth_only:
        get_service(interactive=True)
        print("[CONFIRMADO] Gmail conectado com permissão somente de leitura.")
        return

    result = run_sync(interactive=False)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
