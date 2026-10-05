from __future__ import annotations

import json
import os
import re
import shutil
import socket
import subprocess
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
PERFIL_GOOGLE_EDGE = Path(r"C:\HAGAP\GOOGLE_DRIVE_BROWSER_PROFILE")
CDP_GOOGLE_PORT = 9223
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
# GOOGLE DRIVE - EDGE NORMAL (fora do Playwright)
# ============================================================

def encontrar_edge():
    candidatos = [
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Edge/Application/msedge.exe",
    ]
    for p in candidatos:
        if p.exists():
            return p
    raise RuntimeError("Microsoft Edge não localizado no computador.")


def porta_aberta(host, port):
    try:
        with socket.create_connection((host, port), timeout=0.5):
            return True
    except OSError:
        return False


def abrir_edge_google_normal():
    """
    Abre um Edge NORMAL, sem o modo de automação do Playwright.
    Isso evita o bloqueio de login do Google em navegador automatizado.
    O perfil fica salvo em C:\HAGAP\GOOGLE_DRIVE_BROWSER_PROFILE.
    """
    if porta_aberta("127.0.0.1", CDP_GOOGLE_PORT):
        return

    PERFIL_GOOGLE_EDGE.mkdir(parents=True, exist_ok=True)
    edge = encontrar_edge()

    args = [
        str(edge),
        f"--remote-debugging-port={CDP_GOOGLE_PORT}",
        "--remote-allow-origins=*",
        f"--user-data-dir={PERFIL_GOOGLE_EDGE}",
        "--no-first-run",
        "--no-default-browser-check",
        "--new-window",
        URL_DRIVE,
    ]

    subprocess.Popen(
        args,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
    )

    inicio = time.time()
    while time.time() - inicio < 30:
        if porta_aberta("127.0.0.1", CDP_GOOGLE_PORT):
            return
        time.sleep(0.5)

    raise RuntimeError("O Edge abriu, mas não consegui conectar ao navegador do Google Drive.")


def conectar_google_drive_edge(playwright):
    abrir_edge_google_normal()

    try:
        browser = playwright.chromium.connect_over_cdp(
            f"http://127.0.0.1:{CDP_GOOGLE_PORT}",
            timeout=30000,
        )
    except Exception as exc:
        raise RuntimeError(f"Não consegui conectar ao Edge normal do Google Drive: {exc}")

    contexts = browser.contexts
    if not contexts:
        raise RuntimeError("O Edge do Google abriu sem contexto navegável.")

    context = contexts[0]

    for page in context.pages:
        if "drive.google.com" in page.url.lower():
            return browser, page

    page = context.new_page()
    page.goto(URL_DRIVE, wait_until="domcontentloaded", timeout=120000)
    return browser, page


def drive_autenticado(page):
    """
    Só considera o Drive autenticado quando a interface real do Drive aparece.
    URL sozinha NÃO é suficiente: quando o Google redireciona para login,
    a URL pode continuar parecendo a pasta do Drive.
    """
    try:
        url = (page.url or "").lower()
        titulo = (page.title() or "").lower()

        if "accounts.google.com" in url:
            return False

        if "sign in" in titulo or "fazer login" in titulo or "login" == titulo.strip():
            return False

        sinais = [
            '[aria-label*="Novo"]',
            '[aria-label*="New"]',
            '[data-tooltip*="Novo"]',
            '[data-tooltip*="New"]',
            'div[role="main"]',
            '[aria-label*="Meu Drive"]',
            '[aria-label*="My Drive"]',
        ]

        achou_interface = 0
        for seletor in sinais:
            try:
                if page.locator(seletor).count():
                    achou_interface += 1
            except Exception:
                pass

        # Exige interface real do Drive + domínio Drive.
        return "drive.google.com" in url and achou_interface >= 2

    except Exception:
        return False


def salvar_diagnostico_drive(page, motivo):
    try:
        diag = BASE / "DIAGNOSTICO_DRIVE"
        diag.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        (diag / f"{stamp}_url.txt").write_text(
            f"MOTIVO: {motivo}\nURL: {page.url}\nTITULO: {page.title()}\n",
            encoding="utf-8",
        )
        page.screenshot(path=str(diag / f"{stamp}_tela.png"), full_page=True)
        (diag / f"{stamp}_pagina.html").write_text(
            page.content(), encoding="utf-8", errors="ignore"
        )
        print(f"[DIAGNÓSTICO] Salvo em: {diag}")
    except Exception:
        pass


def aguardar_login_drive(page, limite=600):
    """
    Abre a pasta destino e aguarda o usuário concluir o login Google.
    Não avança enquanto a interface autenticada do Drive não estiver presente.
    """
    page.goto(URL_DRIVE, wait_until="domcontentloaded", timeout=120000)

    inicio = time.time()
    ultimo_aviso = 0

    while time.time() - inicio < limite:
        if drive_autenticado(page):
            # Garante que estamos na pasta destino depois do login.
            if DRIVE_FOLDER_ID not in (page.url or ""):
                try:
                    page.goto(URL_DRIVE, wait_until="domcontentloaded", timeout=120000)
                    time.sleep(2)
                except Exception:
                    pass

            if drive_autenticado(page):
                print("\n[CONFIRMADO] Google Drive autenticado e pasta destino aberta.")
                return

        agora = time.time()
        if agora - ultimo_aviso >= 5:
            print(
                "\r[PENDENTE] No Edge do Google: faça LOGIN normalmente e deixe a pasta 'projetos' aberta. "
                "O programa só continuará depois disso.       ",
                end="",
                flush=True,
            )
            ultimo_aviso = agora

        time.sleep(2)

    print()
    salvar_diagnostico_drive(page, "TIMEOUT_LOGIN")
    raise RuntimeError(
        "O Google Drive não chegou ao estado autenticado. "
        "Faça o login no Edge que o programa abriu e deixe a pasta 'projetos' visível."
    )


def _input_upload_direto(page, caminho):
    """
    Método preferencial: usa diretamente o input de upload que o Google Drive
    mantém oculto na página. Não depende do botão 'Novo'.
    """
    caminho = Path(caminho)

    if caminho.is_dir():
        seletores = [
            'input[type="file"][webkitdirectory]',
            'input[type="file"][directory]',
        ]
    else:
        seletores = [
            'input[type="file"]:not([webkitdirectory]):not([directory])',
        ]

    for seletor in seletores:
        try:
            loc = page.locator(seletor)
            qtd = loc.count()
            for i in range(qtd):
                try:
                    alvo = loc.nth(i)
                    alvo.set_input_files(str(caminho), timeout=15000)
                    return True
                except Exception:
                    continue
        except Exception:
            continue

    return False


def _clicar_novo(page):
    """
    Seletor confirmado pelo diagnóstico real do Drive:
      button[guidedhelpid="new_menu_button"][aria-disabled="false"]
    Primeiro tenta clique DOM direto para não depender de acessibilidade/role.
    """

    # Método 1 — clique DOM direto no botão ATIVO confirmado no HTML.
    try:
        clicou = page.evaluate("""
            () => {
                const btn = document.querySelector(
                    'button[guidedhelpid="new_menu_button"][aria-disabled="false"]'
                );
                if (!btn) return false;
                btn.scrollIntoView({block: 'center', inline: 'center'});
                btn.click();
                return true;
            }
        """)
        if clicou:
            page.wait_for_timeout(1200)
            return True
    except Exception:
        pass

    # Método 2 — Playwright no seletor exato.
    try:
        btn = page.locator(
            'button[guidedhelpid="new_menu_button"][aria-disabled="false"]'
        ).first
        if btn.count():
            btn.scroll_into_view_if_needed()
            btn.click(timeout=7000, force=True)
            page.wait_for_timeout(1200)
            return True
    except Exception:
        pass

    # Método 3 — fallback pelo texto visual Novo/New.
    try:
        spans = page.locator("span.jYPt8c").filter(
            has_text=re.compile(r"^\\s*(Novo|New)\\s*$", re.I)
        )
        for i in range(min(spans.count(), 10)):
            sp = spans.nth(i)
            try:
                btn = sp.locator("xpath=ancestor::button[1]")
                if (
                    btn.count()
                    and btn.get_attribute("aria-disabled") != "true"
                ):
                    btn.evaluate("(el) => el.click()")
                    page.wait_for_timeout(1200)
                    return True
            except Exception:
                continue
    except Exception:
        pass

    return False


def _menu_upload(page, pasta):
    expressoes = (
        [
            re.compile(r"Upload.*pasta", re.I),
            re.compile(r"Fazer.*upload.*pasta", re.I),
            re.compile(r"Folder.*upload", re.I),
        ]
        if pasta
        else [
            re.compile(r"Upload.*arquivo", re.I),
            re.compile(r"Fazer.*upload.*arquivo", re.I),
            re.compile(r"File.*upload", re.I),
        ]
    )

    # Aguarda o menu ser montado após clicar em Novo.
    try:
        page.wait_for_timeout(1000)
    except Exception:
        pass

    candidatos = [
        '[role="menuitem"]',
        '[role="menuitemradio"]',
        '[role="option"]',
        'div[role="menu"] *',
    ]

    for rx in expressoes:
        # Primeiro por role/acessibilidade.
        for loc in [
            page.get_by_role("menuitem", name=rx),
            page.get_by_text(rx),
        ]:
            try:
                if loc.count():
                    for i in range(min(loc.count(), 10)):
                        alvo = loc.nth(i)
                        if alvo.is_visible(timeout=1000):
                            return alvo
            except Exception:
                pass

        # Depois por texto dentro dos elementos reais do menu.
        for seletor in candidatos:
            try:
                loc = page.locator(seletor).filter(has_text=rx)
                for i in range(min(loc.count(), 20)):
                    alvo = loc.nth(i)
                    try:
                        if alvo.is_visible(timeout=800):
                            return alvo
                    except Exception:
                        continue
            except Exception:
                continue

    return None


def _selecionar_upload_via_atalho(page, caminho):
    """
    Método principal confirmado pela documentação oficial do Google Drive:
    Windows/ChromeOS
      Upload de arquivo: Alt+C, depois U
      Upload de pasta:   Alt+C, depois I
    Isso evita depender do botão Novo e do menu visual.
    """
    try:
        page.bring_to_front()
        page.locator("body").click(position={"x": 400, "y": 300}, timeout=3000, force=True)
    except Exception:
        pass

    tecla = "i" if Path(caminho).is_dir() else "u"

    try:
        with page.expect_file_chooser(timeout=15000) as fc_info:
            page.keyboard.press("Alt+C")
            page.wait_for_timeout(350)
            page.keyboard.press(tecla)

        chooser = fc_info.value
        chooser.set_files(str(caminho))
        return True

    except Exception:
        return False


def _selecionar_upload_via_menu(page, caminho):
    if not _clicar_novo(page):
        salvar_diagnostico_drive(page, "BOTAO_NOVO_ATIVO_NAO_CLICADO")
        return False

    time.sleep(1)
    menu = _menu_upload(page, caminho.is_dir())
    if menu is None:
        salvar_diagnostico_drive(page, "MENU_UPLOAD_NAO_LOCALIZADO")
        return False

    try:
        with page.expect_file_chooser(timeout=15000) as fc_info:
            menu.click()
        chooser = fc_info.value
        chooser.set_files(str(caminho))
        return True
    except Exception:
        return False


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
    time.sleep(3)

    # 1) Método principal: atalho oficial do Google Drive.
    #    Pasta: Alt+C e depois I | Arquivo: Alt+C e depois U.
    selecionado = _selecionar_upload_via_atalho(page, caminho)

    # 2) Fallback: input de upload oculto do próprio Google Drive.
    if not selecionado:
        selecionado = _input_upload_direto(page, caminho)

    # 3) Último fallback: botão/menu Novo.
    if not selecionado:
        selecionado = _selecionar_upload_via_menu(page, caminho)

    if not selecionado:
        salvar_diagnostico_drive(page, "UPLOAD_NAO_INICIADO")
        if not drive_autenticado(page):
            raise RuntimeError(
                "O Google Drive perdeu/não concluiu o login. "
                "Faça login no Edge do Google e deixe a pasta 'projetos' aberta."
            )
        raise RuntimeError(
            "O Drive está autenticado, mas o upload não iniciou. "
            "Foi salvo um diagnóstico em DIAGNOSTICO_DRIVE para eu ajustar exatamente à sua tela."
        )

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
    print("[CONFIRMADO] Google Drive será aberto em Edge normal, com perfil próprio persistente.")
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

            print("\n[Google Drive] Abrindo Edge normal para permitir o login Google...")
            drive_browser, drive_page = conectar_google_drive_edge(p)
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
