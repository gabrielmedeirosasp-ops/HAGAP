from __future__ import annotations

import json
import os
import re
import shutil
import sys
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


BASE = Path(__file__).resolve().parent

SITE_URL = "https://copel0.sharepoint.com/sites/VORUMU-OBRAS-HAGAP-"
PASTA_TEAMS = "/sites/VORUMU-OBRAS-HAGAP-/Documentos Compartilhados/HAGAP -/3- Obras para Execução"
URL_TEAMS = "https://copel0.sharepoint.com/sites/VORUMU-OBRAS-HAGAP-/Shared%20Documents/Forms/AllItems.aspx"

DRIVE_FOLDER_ID = "1TTdT09m-WOVeEiii2hed9y4JhsNw8uh9"
URL_DRIVE = f"https://drive.google.com/drive/u/0/folders/{DRIVE_FOLDER_ID}"

PERFIL_EDGE = Path(r"C:\HAGAP\ENVIO_PROGRAMADOS_BROWSER")
TEMP_ROOT = Path(r"C:\HAGAP\TEMP_ENVIO_PROGRAMADOS")
RELATORIO = BASE / "RELATORIO_ENVIO_PROGRAMADOS_DRIVE.json"
ESTADO = BASE / "ESTADO_ENVIO_PROGRAMADOS.json"

RE_PROJ = re.compile(r"(?<!\d)(\d{7}[A-Z]?)(?!\d)", re.I)

# [CONFIRMADO] Estes projetos já estavam representados na pasta destino
# quando o programa foi preparado. O estado local complementa esta lista
# após cada envio bem-sucedido.
JA_NO_DRIVE_INICIAL = {
    "1513221","1706284","1711120","1718931","1718955",
    "1726018","1726914","1727315","1727889","1731349",
    "1731354","1732558","1737088","1737089","1739821",
}


def carregar_estado():
    if not ESTADO.exists():
        return {"enviados": []}
    try:
        data = json.loads(ESTADO.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"enviados": []}
        data.setdefault("enviados", [])
        return data
    except Exception:
        return {"enviados": []}


def salvar_enviado(projetos):
    estado = carregar_estado()
    atual = {str(x).upper() for x in estado.get("enviados", [])}
    atual.update(str(x).upper() for x in projetos)
    estado["enviados"] = sorted(atual)
    estado["atualizado_em"] = datetime.now().isoformat(timespec="seconds")
    ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


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


def escolher_planilha():
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


def projetos_mes_atual(xlsx):
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


def project_ids_from_name(s):
    return {m.upper() for m in RE_PROJ.findall(str(s or ""))}


# ============================================================
# TEAMS / SHAREPOINT - SOMENTE LEITURA
# ============================================================

def odata_path(path):
    return str(path).replace("'", "''")


def url_api(endpoint):
    return quote(SITE_URL.rstrip("/") + endpoint, safe=":/?&=$(),'%-")


def endpoint_files(folder):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(folder)
        + "')/Files?$select=Name,ServerRelativeUrl,Length"
    )


def endpoint_folders(folder):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(folder)
        + "')/Folders?$select=Name,ServerRelativeUrl,ItemCount"
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


def aguardar_login_teams(context, page, limite=300):
    page.goto(URL_TEAMS, wait_until="domcontentloaded", timeout=120000)
    inicio = time.time()
    ultimo = ""

    while time.time() - inicio < limite:
        try:
            get_json(context, endpoint_folders(PASTA_TEAMS))
            print("\n[CONFIRMADO] Teams/SharePoint autenticado.")
            return
        except Exception as exc:
            ultimo = str(exc)
            print(
                "\rAguardando login COPEL/Teams no Edge... faça login/MFA se solicitado.      ",
                end="",
                flush=True,
            )
            time.sleep(3)

    print()
    raise RuntimeError("Não foi possível validar o Teams. Último erro: " + ultimo)


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


def safe_windows_name(name):
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    name = name.rstrip(". ")
    return name or "_"


def baixar_pasta_teams(context, source_folder):
    if TEMP_ROOT.exists():
        shutil.rmtree(TEMP_ROOT, ignore_errors=True)
    TEMP_ROOT.mkdir(parents=True, exist_ok=True)

    source_name = safe_windows_name(source_folder.rstrip("/").rsplit("/", 1)[-1])
    local_root = TEMP_ROOT / source_name
    local_root.mkdir(parents=True, exist_ok=True)

    files = scan_folder(context, source_folder)

    for idx, item in enumerate(files, start=1):
        remote = item.get("ServerRelativeUrl", "")
        if not remote:
            continue

        rel = remote[len(source_folder.rstrip("/") + "/"):] if remote.startswith(source_folder.rstrip("/") + "/") else remote.rsplit("/", 1)[-1]
        parts = [safe_windows_name(x) for x in rel.split("/") if x]
        if not parts:
            continue

        destino = local_root.joinpath(*parts)
        destino.parent.mkdir(parents=True, exist_ok=True)

        print(f"\r[DOWNLOAD TEAMS] {idx}/{len(files)} - {parts[-1][:55]:55}", end="", flush=True)
        data = request_get_retry(
            context,
            url_api(endpoint_download(remote)),
            timeout=180000,
        ).body()
        destino.write_bytes(data)

    print()
    return local_root


def baixar_arquivo_teams(context, item):
    TEMP_ROOT.mkdir(parents=True, exist_ok=True)
    name = safe_windows_name(item.get("Name") or item.get("ServerRelativeUrl", "").rsplit("/", 1)[-1])
    destino = TEMP_ROOT / name
    data = request_get_retry(
        context,
        url_api(endpoint_download(item["ServerRelativeUrl"])),
        timeout=180000,
    ).body()
    destino.write_bytes(data)
    return destino


# ============================================================
# GOOGLE DRIVE - LOGIN NORMAL PELO EDGE
# ============================================================

def aguardar_login_drive(page, limite=300):
    page.goto(URL_DRIVE, wait_until="domcontentloaded", timeout=120000)
    inicio = time.time()

    while time.time() - inicio < limite:
        url = page.url.lower()

        if "drive.google.com/drive" in url:
            print("\n[CONFIRMADO] Google Drive aberto.")
            return

        print(
            "\rAguardando login do Google no Edge... entre na sua conta se solicitado.       ",
            end="",
            flush=True,
        )
        time.sleep(3)

    print()
    raise RuntimeError("Não foi possível abrir a pasta do Google Drive.")


def _clicar_novo(page):
    tentativas = [
        lambda: page.get_by_role("button", name=re.compile(r"^(Novo|New)$", re.I)).first,
        lambda: page.locator('[aria-label="Novo"]').first,
        lambda: page.locator('[aria-label="New"]').first,
        lambda: page.get_by_text(re.compile(r"^(Novo|New)$", re.I), exact=True).first,
    ]

    for getloc in tentativas:
        try:
            loc = getloc()
            if loc.count() and loc.is_visible(timeout=1500):
                loc.click()
                return
        except Exception:
            pass

    raise RuntimeError("Não encontrei o botão Novo/New no Google Drive.")


def _menu_upload(page, pasta):
    regex = (
        re.compile(r"^(Upload de pasta|Folder upload)$", re.I)
        if pasta
        else re.compile(r"^(Upload de arquivo|File upload)$", re.I)
    )

    loc = page.get_by_text(regex, exact=True)
    if not loc.count():
        # Alguns layouts expõem como menuitem.
        loc = page.get_by_role("menuitem", name=regex)

    if not loc.count():
        raise RuntimeError(
            "Não encontrei a opção "
            + ("Upload de pasta" if pasta else "Upload de arquivo")
            + " no Google Drive."
        )

    return loc.first


def _esperar_upload(page, nome, timeout_s):
    inicio = time.time()
    texto_ok = re.compile(
        r"(Upload conclu[ií]do|Upload complete|Conclu[ií]do|Complete)",
        re.I,
    )

    while time.time() - inicio < timeout_s:
        try:
            body = page.locator("body").inner_text(timeout=3000)
            if texto_ok.search(body):
                return True
            if nome.lower() in body.lower() and time.time() - inicio > 12:
                # O item já apareceu na pasta. Dá uma margem para finalizar.
                time.sleep(3)
                return True
        except Exception:
            pass

        time.sleep(2)

    return False


def upload_drive_browser(page, caminho):
    caminho = Path(caminho)
    if not caminho.exists():
        raise RuntimeError(f"Arquivo/pasta local não existe: {caminho}")

    page.goto(URL_DRIVE, wait_until="domcontentloaded", timeout=120000)
    time.sleep(2)

    _clicar_novo(page)
    menu = _menu_upload(page, caminho.is_dir())

    try:
        with page.expect_file_chooser(timeout=15000) as fc_info:
            menu.click()
        chooser = fc_info.value
        chooser.set_files(str(caminho))
    except Exception as exc:
        raise RuntimeError(f"Falha ao selecionar {caminho.name} para upload: {exc}")

    total = 0
    if caminho.is_dir():
        for p in caminho.rglob("*"):
            if p.is_file():
                try:
                    total += p.stat().st_size
                except Exception:
                    pass
    else:
        total = caminho.stat().st_size

    timeout_s = max(120, min(900, 90 + int(total / (1024 * 1024)) * 8))

    print(f"[UPLOAD DRIVE] {caminho.name} ({total/1024/1024:.1f} MB)")
    ok = _esperar_upload(page, caminho.name, timeout_s)

    if not ok:
        raise RuntimeError(
            f"Não consegui confirmar o término do upload de {caminho.name}. "
            "Confira o Google Drive antes de repetir."
        )


# ============================================================
# FLUXO PRINCIPAL
# ============================================================

def main():
    planilha = escolher_planilha()
    projetos = projetos_mes_atual(planilha)
    alvos = set(projetos)

    estado = carregar_estado()
    conhecidos = set(JA_NO_DRIVE_INICIAL)
    conhecidos.update(str(x).upper() for x in estado.get("enviados", []))

    faltantes = alvos - conhecidos

    print("=" * 78)
    print("HAGAP — ENVIAR PROJETOS PROGRAMADOS PARA O DRIVE")
    print(f"Planilha: {planilha}")
    print(f"Mês: {date.today().month:02d}/{date.today().year}")
    print(f"Projetos únicos: {len(projetos)}")
    print(f"[CONFIRMADO] Já conhecidos no Drive/estado local: {len(alvos & conhecidos)}")
    print(f"[PENDENTE] A procurar no Teams: {len(faltantes)}")
    print("[CONFIRMADO] Teams é somente leitura; nada será removido.")
    print("[CONFIRMADO] Google Drive será acessado pelo login normal do Edge.")
    print("=" * 78)

    relatorio = {
        "planilha": str(planilha),
        "mes": f"{date.today().month:02d}/{date.today().year}",
        "projetos_programacao": projetos,
        "ja_no_drive_ou_estado": sorted(alvos & conhecidos),
        "localizados_teams": {},
        "nao_localizados_teams": [],
        "divergencias": [],
        "uploads": [],
    }

    PERFIL_EDGE.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=False,
        )

        try:
            teams_page = context.pages[0] if context.pages else context.new_page()
            aguardar_login_teams(context, teams_page)

            drive_page = context.new_page()
            aguardar_login_drive(drive_page)

            root_files, root_folders = listar_arvore(context, PASTA_TEAMS)

            folder_projects = {}
            project_sources = {}

            for idx, folder in enumerate(root_folders, start=1):
                url = folder.get("ServerRelativeUrl")
                name = folder.get("Name", "")
                ids = project_ids_from_name(name) & faltantes

                if not ids and url:
                    files = scan_folder(context, url)
                    for item in files:
                        ids |= project_ids_from_name(item.get("Name", "")) & faltantes
                        ids |= project_ids_from_name(item.get("ServerRelativeUrl", "")) & faltantes

                if ids:
                    folder_projects[url] = ids
                    for proj in ids:
                        project_sources.setdefault(proj, set()).add(url)
                        relatorio["localizados_teams"].setdefault(proj, []).append(name)

                print(
                    f"\r[VARREDURA] Pastas Teams: {idx}/{len(root_folders)} | "
                    f"projetos localizados: {len(project_sources)}      ",
                    end="",
                    flush=True,
                )

            print()

            root_file_projects = {}
            for item in root_files:
                ids = (
                    project_ids_from_name(item.get("Name", ""))
                    | project_ids_from_name(item.get("ServerRelativeUrl", ""))
                ) & faltantes

                if ids:
                    root_file_projects[item["ServerRelativeUrl"]] = ids
                    for proj in ids:
                        project_sources.setdefault(proj, set()).add(item["ServerRelativeUrl"])
                        relatorio["localizados_teams"].setdefault(proj, []).append(item.get("Name", ""))

            # Bloqueia projeto encontrado em mais de uma origem.
            divergentes = set()
            for proj, sources in project_sources.items():
                if len(sources) > 1:
                    divergentes.add(proj)
                    relatorio["divergencias"].append({
                        "projeto": proj,
                        "origens": sorted(sources),
                    })

            # Pasta inteira.
            for folder_url, ids in folder_projects.items():
                needed = (ids & faltantes) - divergentes
                if not needed:
                    continue

                print(
                    f"\n[PREPARANDO] Pasta Teams: {folder_url.rsplit('/',1)[-1]} | "
                    f"projetos: {', '.join(sorted(needed))}"
                )

                local = baixar_pasta_teams(context, folder_url)
                upload_drive_browser(drive_page, local)

                salvar_enviado(needed)
                relatorio["uploads"].append({
                    "origem": folder_url,
                    "projetos": sorted(needed),
                    "tipo": "PASTA_INTEIRA",
                })

                shutil.rmtree(TEMP_ROOT, ignore_errors=True)

            # Arquivos soltos na raiz.
            for remote, ids in root_file_projects.items():
                needed = (ids & faltantes) - divergentes
                if not needed:
                    continue

                item = next(
                    (x for x in root_files if x.get("ServerRelativeUrl") == remote),
                    None,
                )
                if not item:
                    continue

                print(
                    f"\n[PREPARANDO] Arquivo Teams: {item.get('Name')} | "
                    f"projetos: {', '.join(sorted(needed))}"
                )

                local = baixar_arquivo_teams(context, item)
                upload_drive_browser(drive_page, local)

                salvar_enviado(needed)
                relatorio["uploads"].append({
                    "origem": remote,
                    "projetos": sorted(needed),
                    "tipo": "ARQUIVO_RAIZ",
                })

                shutil.rmtree(TEMP_ROOT, ignore_errors=True)

            localizados = set(relatorio["localizados_teams"])
            relatorio["nao_localizados_teams"] = sorted(faltantes - localizados)

        finally:
            context.close()

    RELATORIO.write_text(
        json.dumps(relatorio, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print()
    print("=" * 78)
    print("RESUMO")
    print(f"Programação: {len(projetos)}")
    print(f"Já no Drive/estado: {len(relatorio['ja_no_drive_ou_estado'])}")
    print(f"Localizados no Teams: {len(relatorio['localizados_teams'])}")
    print(f"Não localizados: {len(relatorio['nao_localizados_teams'])}")
    print(f"Divergências: {len(relatorio['divergencias'])}")
    print(f"Uploads efetuados: {len(relatorio['uploads'])}")
    print(f"Relatório: {RELATORIO}")
    print("=" * 78)

    if relatorio["nao_localizados_teams"]:
        print("\n[NÃO LOCALIZADO]")
        for proj in relatorio["nao_localizados_teams"]:
            print(" -", proj)

    if relatorio["divergencias"]:
        print("\n[DIVERGÊNCIA — NÃO ENVIADO AUTOMATICAMENTE]")
        for d in relatorio["divergencias"]:
            print(" -", d["projeto"], "=>", " | ".join(d["origens"]))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelado.")
    except Exception as exc:
        print(f"\n[ERRO] {exc}")
        sys.exit(1)
