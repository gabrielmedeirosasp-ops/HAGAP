import argparse
import json
import sys
import time
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
CFG_PATH = BASE / "config.json"
STATE_PATH = BASE / "estado_sync.json"


def carregar_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


CFG = carregar_json(CFG_PATH, {})
SITE_URL = CFG["site_url"].rstrip("/")
PASTA_SERVIDOR = CFG["pasta_servidor"]
URL_INICIAL = CFG["url_inicial"]
DESTINO = Path(CFG.get("destino", r"C:\HAGAP\TEAMS_SYNC"))
PERFIL_EDGE = Path(CFG.get("perfil_edge", r"C:\HAGAP\SHAREPOINT_BROWSER_PROFILE"))


def odata_path(path):
    return str(path).replace("'", "''")


def url_api(endpoint):
    bruto = SITE_URL + endpoint
    return quote(bruto, safe=":/?&=$(),'%-")


def endpoint_arquivos(pasta):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(pasta)
        + "')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,UniqueId"
    )


def endpoint_pastas(pasta):
    return (
        "/_api/web/GetFolderByServerRelativePath(decodedUrl='"
        + odata_path(pasta)
        + "')/Folders?$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount"
    )


def endpoint_download(caminho):
    return (
        "/_api/web/GetFileByServerRelativePath(decodedUrl='"
        + odata_path(caminho)
        + "')/$value"
    )


def request_get_retry(context, url, headers=None, tentativas=6, timeout=120000):
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            resposta = context.request.get(
                url,
                headers=headers or {},
                timeout=timeout,
            )
            if resposta.ok:
                return resposta

            ultimo_erro = RuntimeError(
                f"SharePoint HTTP {resposta.status}: {resposta.text()[:300]}"
            )
        except Exception as exc:
            ultimo_erro = exc

        if tentativa < tentativas:
            espera = min(30, 2 ** (tentativa - 1))
            print(
                f"  [RETRY {tentativa}/{tentativas}] conexão interrompida; "
                f"nova tentativa em {espera}s...",
                flush=True,
            )
            time.sleep(espera)

    raise RuntimeError(
        f"Falha após {tentativas} tentativas: {ultimo_erro}"
    )


def get_json(context, endpoint):
    resposta = request_get_retry(
        context,
        url_api(endpoint),
        headers={"Accept": "application/json;odata=nometadata"},
    )
    return resposta.json()


def testar_acesso(context):
    dados = get_json(context, endpoint_arquivos(PASTA_SERVIDOR))
    return dados.get("value", [])


def aguardar_login(context, pagina, limite_segundos=300):
    pagina.goto(URL_INICIAL, wait_until="domcontentloaded", timeout=120000)

    inicio = time.time()
    ultimo_erro = ""

    while time.time() - inicio < limite_segundos:
        try:
            arquivos = testar_acesso(context)
            return arquivos
        except Exception as exc:
            ultimo_erro = str(exc)
            print(
                "\rAguardando autenticação no Edge... "
                "faça o login/MFA na janela aberta.          ",
                end="",
                flush=True,
            )
            time.sleep(3)

    print()
    raise RuntimeError(
        "Não consegui validar o SharePoint dentro do tempo limite. "
        f"Último erro: {ultimo_erro}"
    )


def listar_recursivo(context, pasta_raiz):
    arquivos = []
    pilha = [pasta_raiz]

    while pilha:
        atual = pilha.pop()

        js_arquivos = get_json(context, endpoint_arquivos(atual))
        arquivos.extend(js_arquivos.get("value", []))

        js_pastas = get_json(context, endpoint_pastas(atual))
        for pasta in js_pastas.get("value", []):
            nome = pasta.get("Name", "")
            if nome.lower() == "forms":
                continue
            caminho = pasta.get("ServerRelativeUrl")
            if caminho:
                pilha.append(caminho)

    return arquivos


def caminho_local(server_relative_url):
    marcador = "/HAGAP -/"
    if marcador in server_relative_url:
        relativo = server_relative_url.split(marcador, 1)[1]
    else:
        relativo = server_relative_url.lstrip("/")

    partes = [
        parte
        for parte in relativo.replace("\\", "/").split("/")
        if parte not in ("", ".", "..")
    ]

    destino = DESTINO.joinpath(*partes)

    raiz = DESTINO.resolve()
    pai = destino.parent.resolve()

    if raiz != pai and raiz not in pai.parents:
        raise RuntimeError("Caminho local inválido")

    return destino


def carregar_estado():
    return carregar_json(STATE_PATH, {"arquivos": {}})


def salvar_estado(estado):
    temporario = STATE_PATH.with_suffix(".tmp")
    temporario.write_text(
        json.dumps(estado, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporario.replace(STATE_PATH)


def precisa_baixar(item, estado):
    remoto = item.get("ServerRelativeUrl", "")
    if not remoto:
        return False

    destino = caminho_local(remoto)
    antigo = estado.get("arquivos", {}).get(remoto)

    tamanho = str(item.get("Length", ""))
    modificado = str(item.get("TimeLastModified", ""))

    if not destino.exists():
        return True
    if not antigo:
        return True
    if str(antigo.get("length", "")) != tamanho:
        return True
    if antigo.get("modified", "") != modificado:
        return True

    return False


def baixar_arquivo(context, item):
    remoto = item["ServerRelativeUrl"]
    destino = caminho_local(remoto)
    destino.parent.mkdir(parents=True, exist_ok=True)

    resposta = request_get_retry(
        context,
        url_api(endpoint_download(remoto)),
        timeout=180000,
    )

    dados = resposta.body()
    esperado = item.get("Length")

    if esperado not in (None, ""):
        try:
            if int(esperado) != len(dados):
                raise RuntimeError(
                    f"Tamanho divergente em {remoto}: "
                    f"recebido {len(dados)}, esperado {esperado}"
                )
        except ValueError:
            pass

    temporario = destino.with_suffix(destino.suffix + ".part")
    temporario.write_bytes(dados)
    temporario.replace(destino)

    return destino, len(dados)


def executar_teste(context):
    arquivos = listar_recursivo(context, PASTA_SERVIDOR)

    print()
    print("=" * 72)
    print("TESTE DE ACESSO SHAREPOINT — OK")
    print(f"Pasta: {PASTA_SERVIDOR}")
    print(f"Arquivos encontrados: {len(arquivos)}")
    print("=" * 72)

    for item in arquivos[:20]:
        print(
            f"- {item.get('Name')} | "
            f"{item.get('Length')} bytes | "
            f"{item.get('TimeLastModified')}"
        )

    if len(arquivos) > 20:
        print(f"... mais {len(arquivos) - 20} arquivo(s).")

    return arquivos


def executar_sync(context):
    arquivos = listar_recursivo(context, PASTA_SERVIDOR)
    estado = carregar_estado()
    estado.setdefault("arquivos", {})

    baixar = [item for item in arquivos if precisa_baixar(item, estado)]

    print()
    print("=" * 72)
    print("SINCRONIZAÇÃO AES")
    print(f"Encontrados: {len(arquivos)}")
    print(f"Novos/alterados: {len(baixar)}")
    print(f"Destino: {DESTINO}")
    print("=" * 72)

    concluidos = 0
    erros = []

    for indice, item in enumerate(baixar, start=1):
        nome = item.get("Name", item.get("ServerRelativeUrl"))
        print(f"[{indice}/{len(baixar)}] {nome}")

        try:
            destino, bytes_salvos = baixar_arquivo(context, item)
        except Exception as exc:
            remoto = item.get("ServerRelativeUrl", nome)
            erros.append({"arquivo": remoto, "erro": str(exc)})
            print(f"  [ERRO] {nome}: {exc}", flush=True)
            continue

        remoto = item["ServerRelativeUrl"]
        estado["arquivos"][remoto] = {
            "modified": str(item.get("TimeLastModified", "")),
            "length": str(item.get("Length", bytes_salvos)),
            "local": str(destino),
            "sincronizado_em": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        salvar_estado(estado)
        concluidos += 1

    print()
    print(
        f"[CONFIRMADO] Sincronizados nesta execução: {concluidos}. "
        f"Já atualizados: {len(arquivos) - len(baixar)}. "
        f"Erros pendentes: {len(erros)}."
    )

    if erros:
        print("\n[PENDENTE] Arquivos que serão tentados novamente na próxima execução:")
        for erro in erros[:30]:
            print(f"  - {erro['arquivo']} | {erro['erro']}")
        if len(erros) > 30:
            print(f"  ... mais {len(erros) - 30} arquivo(s).")


def main():
    parser = argparse.ArgumentParser(
        description="Ponte HAGAP SharePoint COPEL — somente leitura"
    )
    parser.add_argument(
        "--modo",
        choices=["teste", "sync"],
        default="teste",
        help="teste apenas lista; sync baixa novos/alterados",
    )
    args = parser.parse_args()

    DESTINO.mkdir(parents=True, exist_ok=True)
    PERFIL_EDGE.mkdir(parents=True, exist_ok=True)

    print("HAGAP — SharePoint COPEL")
    print("[CONFIRMADO] Este módulo não altera o site HAGAP.")
    print("[CONFIRMADO] Este módulo não altera o /boletim.")
    print("[CONFIRMADO] Operações no SharePoint são somente leitura (GET).")
    print()
    print(
        "Na primeira execução será aberta uma janela própria do Microsoft Edge. "
        "Faça o login da COPEL e o MFA normalmente."
    )

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=False,
        )

        try:
            pagina = context.pages[0] if context.pages else context.new_page()
            aguardar_login(context, pagina)
            print("\n[CONFIRMADO] Sessão SharePoint validada.")

            if args.modo == "teste":
                executar_teste(context)
            else:
                executar_sync(context)

            print()
            input("Pressione ENTER para fechar...")

        finally:
            context.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelado.")
    except Exception as exc:
        print(f"\n[ERRO] {exc}")
        input("Pressione ENTER para fechar...")
        sys.exit(1)
