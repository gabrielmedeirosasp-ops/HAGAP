from __future__ import annotations

import importlib.util
import io
import json
import mimetypes
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
ROOT_REPO = BASE.parent

SITE_URL = "https://copel0.sharepoint.com/sites/VORUMU-OBRAS-HAGAP-"
PASTA_TEAMS = "/sites/VORUMU-OBRAS-HAGAP-/Documentos Compartilhados/HAGAP -/3- Obras para Execução"
URL_INICIAL = (
    "https://copel0.sharepoint.com/sites/VORUMU-OBRAS-HAGAP-/"
    "Shared%20Documents/Forms/AllItems.aspx"
)
PERFIL_EDGE = Path(r"C:\HAGAP\SHAREPOINT_BROWSER_PROFILE")
DRIVE_DESTINO_ID = "1TTdT09m-WOVeEiii2hed9y4JhsNw8uh9"
RELATORIO = BASE / "RELATORIO_ENVIO_PROGRAMADOS_DRIVE.json"

RE_PROJ = re.compile(r"(?<!\d)(\d{7}[A-Z]?)(?!\d)", re.I)


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def escolher_planilha() -> Path:
    candidatos = [
        Path.cwd() / "PROGRAMAÇÃO UMU .xlsx",
        BASE / "PROGRAMAÇÃO UMU .xlsx",
        Path.home() / "Downloads" / "PROGRAMAÇÃO UMU .xlsx",
        Path.home() / "Desktop" / "PROGRAMAÇÃO UMU .xlsx",
        Path.home() / "Documents" / "PROGRAMAÇÃO UMU .xlsx",
    ]
    for p in candidatos:
        if p.exists():
            return p

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        arq = filedialog.askopenfilename(
            title="Selecione PROGRAMAÇÃO UMU .xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        root.destroy()
        if arq:
            return Path(arq)
    except Exception:
        pass

    raise RuntimeError("PROGRAMAÇÃO UMU .xlsx não localizada.")


def parse_data(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v

    s = str(v or "").strip()

    m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            pass

    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b", s)
    if m:
        d, mo, y = map(int, m.groups())
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            pass

    return None


def projeto_da_celula(v):
    if v is None or isinstance(v, bool):
        return ""

    if isinstance(v, int):
        s = str(v)
    elif isinstance(v, float) and v.is_integer():
        s = str(int(v))
    else:
        s = str(v).strip().upper()

    return s if re.fullmatch(r"\d{7}[A-Z]?", s) else ""


def projetos_mes_atual(xlsx: Path) -> list[str]:
    """
    Leitura corrigida para a PROGRAMAÇÃO UMU:
    DATA e PROJETO podem estar deslocados por colunas vazias/mescladas.
    """
    wb = load_workbook(xlsx, data_only=True, read_only=True)
    if "MÊS ATUAL" not in wb.sheetnames:
        raise RuntimeError(
            "A aba 'MÊS ATUAL' não foi localizada. Abas: " + ", ".join(wb.sheetnames)
        )

    ws = wb["MÊS ATUAL"]
    hoje = date.today()
    data_corrente = None
    projetos = []

    for row in ws.iter_rows(values_only=True):
        if not row:
            continue

        data_linha = None
        for v in row:
            d = parse_data(v)
            if d:
                data_linha = d
                break

        if data_linha:
            data_corrente = data_linha

        if not data_corrente:
            continue

        if (data_corrente.year, data_corrente.month) != (hoje.year, hoje.month):
            continue

        for v in row:
            p = projeto_da_celula(v)
            if p and p not in projetos:
                projetos.append(p)

    if not projetos:
        raise RuntimeError(
            f"Nenhum projeto de 7 dígitos localizado em "
            f"{hoje.month:02d}/{hoje.year} na aba MÊS ATUAL."
        )

    return projetos


def odata_path(path):
    return str(path).replace("'", "''")


def url_api(endpoint):
    return quote(SITE_URL.rstrip("/") + endpoint, safe=":/?&=$(),'%-")


def endpoint_files(folder):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(folder)
        + "')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,UniqueId"
    )


def endpoint_folders(folder):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(folder)
        + "')/Folders?$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount"
    )


def endpoint_download(path):
    return (
        "/_api/web/GetFileByServerRelativePath(decodedUrl='"
        + odata_path(path)
        + "')/$value"
    )


def request_get_retry(context, url, headers=None, tentativas=6, timeout=120000):
    last = None
    for n in range(1, tentativas + 1):
        try:
            r = context.request.get(url, headers=headers or {}, timeout=timeout)
            if r.ok:
                return r
            last = RuntimeError(f"SharePoint HTTP {r.status}: {r.text()[:250]}")
        except Exception as exc:
            last = exc
        if n < tentativas:
            time.sleep(min(30, 2 ** (n - 1)))
    raise RuntimeError(f"Falha SharePoint após {tentativas} tentativas: {last}")


def get_json(context, endpoint):
    return request_get_retry(
        context,
        url_api(endpoint),
        headers={"Accept": "application/json;odata=nometadata"},
    ).json()


def aguardar_login(context, page, limite=300):
    page.goto(URL_INICIAL, wait_until="domcontentloaded", timeout=120000)
    inicio = time.time()
    last = ""
    while time.time() - inicio < limite:
        try:
            get_json(context, endpoint_folders(PASTA_TEAMS))
            return
        except Exception as exc:
            last = str(exc)
            print(
                "\rAguardando autenticação COPEL/Teams no Edge... faça login/MFA se solicitado.      ",
                end="",
                flush=True,
            )
            time.sleep(3)
    print()
    raise RuntimeError("Não foi possível validar o Teams. Último erro: " + last)


def listar_arvore(context, folder):
    files = get_json(context, endpoint_files(folder)).get("value", [])
    folders = get_json(context, endpoint_folders(folder)).get("value", [])
    folders = [x for x in folders if str(x.get("Name", "")).lower() != "forms"]
    return files, folders


def scan_folder(context, folder):
    all_files = []
    stack = [folder]
    while stack:
        current = stack.pop()
        files, folders = listar_arvore(context, current)
        all_files.extend(files)
        for sub in folders:
            u = sub.get("ServerRelativeUrl")
            if u:
                stack.append(u)
    return all_files


def project_ids_from_name(s):
    return {m.upper() for m in RE_PROJ.findall(str(s or ""))}


def get_drive_service():
    """Reutiliza somente a autenticação Drive já existente do HAGAP."""
    candidatos = [
        Path(r"C:\HAGAP\gerador.py"),
        BASE / "gerador.py",
        BASE.parent / "gerador.py",
        BASE.parent.parent / "gerador.py",
    ]

    gerador_py = next((p for p in candidatos if p.exists()), None)
    if not gerador_py:
        raise RuntimeError(
            "Não encontrei gerador.py com a autenticação do Google Drive. "
            "Mantenha C:\\HAGAP\\gerador.py ou coloque este programa ao lado dele."
        )

    try:
        spec = importlib.util.spec_from_file_location("hagap_gerador_drive", gerador_py)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    except Exception as exc:
        raise RuntimeError(f"Não consegui abrir a autenticação Drive do HAGAP: {exc}")

    if not hasattr(mod, "_get_drive_service_local"):
        raise RuntimeError("O gerador.py localizado não possui a integração Google Drive esperada.")

    svc = mod._get_drive_service_local()
    if not svc:
        raise RuntimeError("Autenticação Google Drive não disponível no HAGAP.")

    return svc


def drive_children(service, parent_id):
    out = []
    token = None
    while True:
        r = service.files().list(
            q=f"'{parent_id}' in parents and trashed=false",
            fields="nextPageToken,files(id,name,mimeType,webViewLink)",
            pageSize=1000,
            pageToken=token,
        ).execute()
        out.extend(r.get("files", []))
        token = r.get("nextPageToken")
        if not token:
            break
    return out


def drive_find_exact(service, parent_id, name, mime=None):
    esc = name.replace("'", "\\'")
    q = f"'{parent_id}' in parents and name='{esc}' and trashed=false"
    if mime:
        q += f" and mimeType='{mime}'"
    r = service.files().list(q=q, fields="files(id,name,mimeType,webViewLink)", pageSize=10).execute()
    return (r.get("files") or [None])[0]


def drive_ensure_folder(service, parent_id, name):
    mime = "application/vnd.google-apps.folder"
    found = drive_find_exact(service, parent_id, name, mime)
    if found:
        return found["id"], False
    meta = {"name": name, "mimeType": mime, "parents": [parent_id]}
    x = service.files().create(body=meta, fields="id,name,webViewLink").execute()
    return x["id"], True


def drive_upload_bytes(service, parent_id, name, data):
    from googleapiclient.http import MediaIoBaseUpload

    found = drive_find_exact(service, parent_id, name)
    if found:
        return found["id"], False

    mime = mimetypes.guess_type(name)[0] or "application/octet-stream"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    meta = {"name": name, "parents": [parent_id]}
    x = service.files().create(body=meta, media_body=media, fields="id,name,webViewLink").execute()
    return x["id"], True


def relative_under(root, full):
    root = root.rstrip("/")
    if full.startswith(root + "/"):
        return full[len(root) + 1:]
    return full.rsplit("/", 1)[-1]


def upload_source_folder(context, drive, source_folder, target_parent_id):
    source_name = source_folder.rstrip("/").rsplit("/", 1)[-1]
    root_drive_id, created = drive_ensure_folder(drive, target_parent_id, source_name)
    files = scan_folder(context, source_folder)

    folders_cache = {"": root_drive_id}
    uploaded = 0
    skipped = 0

    for item in files:
        remote = item.get("ServerRelativeUrl", "")
        if not remote:
            continue
        rel = relative_under(source_folder, remote)
        parts = [x for x in rel.split("/") if x]
        if not parts:
            continue

        parent_rel = ""
        parent_id = root_drive_id
        for part in parts[:-1]:
            next_rel = parent_rel + ("/" if parent_rel else "") + part
            if next_rel not in folders_cache:
                fid, _ = drive_ensure_folder(drive, parent_id, part)
                folders_cache[next_rel] = fid
            parent_id = folders_cache[next_rel]
            parent_rel = next_rel

        fname = parts[-1]
        existing = drive_find_exact(drive, parent_id, fname)
        if existing:
            skipped += 1
            continue

        data = request_get_retry(
            context, url_api(endpoint_download(remote)), timeout=180000
        ).body()
        _, made = drive_upload_bytes(drive, parent_id, fname, data)
        uploaded += 1 if made else 0
        skipped += 0 if made else 1

    return {"folder": source_name, "folder_created": created, "uploaded": uploaded, "skipped": skipped}


def upload_root_file(context, drive, item, target_parent_id):
    name = item.get("Name") or item.get("ServerRelativeUrl", "").rsplit("/", 1)[-1]
    existing = drive_find_exact(drive, target_parent_id, name)
    if existing:
        return {"file": name, "uploaded": 0, "skipped": 1}
    data = request_get_retry(
        context, url_api(endpoint_download(item["ServerRelativeUrl"])), timeout=180000
    ).body()
    _, made = drive_upload_bytes(drive, target_parent_id, name, data)
    return {"file": name, "uploaded": 1 if made else 0, "skipped": 0 if made else 1}


def main():
    planilha = escolher_planilha()
    projetos = projetos_mes_atual(planilha)
    alvos = set(projetos)

    print("=" * 78)
    print("HAGAP — ENVIAR PROJETOS PROGRAMADOS PARA O DRIVE")
    print(f"Planilha: {planilha}")
    print(f"Mês: {date.today().month:02d}/{date.today().year}")
    print(f"Projetos únicos: {len(projetos)}")
    print("[CONFIRMADO] Fonte dos arquivos: Teams/SharePoint — 3- Obras para Execução")
    print("[CONFIRMADO] O Teams é somente leitura; nada será removido.")
    print("=" * 78)

    PERFIL_EDGE.mkdir(parents=True, exist_ok=True)
    drive = get_drive_service()

    try:
        meta = drive.files().get(
            fileId=DRIVE_DESTINO_ID,
            fields="id,name,mimeType,capabilities(canAddChildren)",
        ).execute()
        if meta.get("mimeType") != "application/vnd.google-apps.folder":
            raise RuntimeError("O destino informado no Drive não é uma pasta.")
        if not (meta.get("capabilities") or {}).get("canAddChildren", False):
            raise RuntimeError("A autenticação HAGAP não possui permissão para adicionar arquivos nessa pasta do Drive.")
        print(f"[CONFIRMADO] Drive destino: {meta.get('name')} ({DRIVE_DESTINO_ID})")
    except Exception as exc:
        raise RuntimeError(f"Não consegui validar a pasta destino no Drive: {exc}")

    root_items = drive_children(drive, DRIVE_DESTINO_ID)
    already = set()
    for x in root_items:
        already |= project_ids_from_name(x.get("name", ""))

    faltavam = alvos - already
    print(f"[CONFIRMADO] Projetos já representados no Drive: {len(alvos & already)}")
    print(f"[PENDENTE] Projetos a procurar no Teams: {len(faltavam)}")

    report = {
        "planilha": str(planilha),
        "mes": f"{date.today().month:02d}/{date.today().year}",
        "projetos_programacao": projetos,
        "ja_no_drive": sorted(alvos & already),
        "localizados_teams": {},
        "nao_localizados_teams": [],
        "uploads": [],
        "divergencias": [],
    }

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=False,
        )

        try:
            page = context.pages[0] if context.pages else context.new_page()
            aguardar_login(context, page)
            print("\n[CONFIRMADO] Sessão Teams/SharePoint validada.")

            root_files, root_folders = listar_arvore(context, PASTA_TEAMS)

            # Indexa pastas de primeiro nível pelo número que aparece no nome
            # e, quando necessário, pelos nomes dos arquivos internos.
            folder_projects = {}
            for idx, folder in enumerate(root_folders, start=1):
                url = folder.get("ServerRelativeUrl")
                name = folder.get("Name", "")
                ids = set(project_ids_from_name(name)) & alvos

                if not ids and url:
                    files = scan_folder(context, url)
                    for item in files:
                        ids |= project_ids_from_name(item.get("Name", "")) & alvos
                        ids |= project_ids_from_name(item.get("ServerRelativeUrl", "")) & alvos

                if ids:
                    folder_projects[url] = ids
                    for proj in ids:
                        report["localizados_teams"].setdefault(proj, []).append(name)

                print(
                    f"\r[VARREDURA] Pastas Teams: {idx}/{len(root_folders)} | projetos localizados: "
                    f"{len(report['localizados_teams'])}",
                    end="",
                    flush=True,
                )

            print()

            root_file_projects = {}
            for item in root_files:
                ids = (
                    project_ids_from_name(item.get("Name", ""))
                    | project_ids_from_name(item.get("ServerRelativeUrl", ""))
                ) & alvos
                if ids:
                    root_file_projects[item["ServerRelativeUrl"]] = ids
                    for proj in ids:
                        report["localizados_teams"].setdefault(proj, []).append(item.get("Name", ""))

            # Detecta projetos com mais de uma origem de primeiro nível.
            for proj, origins in report["localizados_teams"].items():
                uniq = sorted(set(origins))
                if len(uniq) > 1:
                    report["divergencias"].append(
                        {
                            "projeto": proj,
                            "tipo": "MULTIPLAS_ORIGENS_TEAMS",
                            "origens": uniq,
                        }
                    )

            # Upload de pastas inteiras somente quando contêm pelo menos um projeto
            # que ainda não está representado no destino.
            for folder_url, ids in folder_projects.items():
                needed = ids & faltavam
                if not needed:
                    continue
                print(f"[UPLOAD] Pasta Teams: {folder_url.rsplit('/',1)[-1]} | projetos: {', '.join(sorted(needed))}")
                result = upload_source_folder(context, drive, folder_url, DRIVE_DESTINO_ID)
                result["projetos"] = sorted(needed)
                report["uploads"].append(result)

            for remote, ids in root_file_projects.items():
                needed = ids & faltavam
                if not needed:
                    continue
                item = next((x for x in root_files if x.get("ServerRelativeUrl") == remote), None)
                if not item:
                    continue
                print(f"[UPLOAD] Arquivo Teams: {item.get('Name')} | projetos: {', '.join(sorted(needed))}")
                result = upload_root_file(context, drive, item, DRIVE_DESTINO_ID)
                result["projetos"] = sorted(needed)
                report["uploads"].append(result)

            localizados = set(report["localizados_teams"])
            report["nao_localizados_teams"] = sorted(faltavam - localizados)

        finally:
            context.close()

    RELATORIO.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print()
    print("=" * 78)
    print("RESUMO")
    print(f"Programação: {len(projetos)} projeto(s)")
    print(f"Já no Drive: {len(report['ja_no_drive'])}")
    print(f"Localizados no Teams: {len(report['localizados_teams'])}")
    print(f"Não localizados no Teams: {len(report['nao_localizados_teams'])}")
    print(f"Pastas/arquivos enviados: {len(report['uploads'])}")
    print(f"Divergências: {len(report['divergencias'])}")
    print(f"Relatório: {RELATORIO}")
    print("=" * 78)

    if report["nao_localizados_teams"]:
        print("\n[NÃO LOCALIZADO] Projetos não encontrados por número em pasta/arquivo do Teams:")
        for p in report["nao_localizados_teams"]:
            print(" -", p)

    if report["divergencias"]:
        print("\n[DIVERGÊNCIA] Projetos encontrados em mais de uma origem do Teams:")
        for d in report["divergencias"]:
            print(" -", d["projeto"], "=>", " | ".join(d["origens"]))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelado.")
    except Exception as exc:
        print(f"\n[ERRO] {exc}")
        sys.exit(1)
