import sys
import time
import unicodedata

import sync_sharepoint as sp

RAIZ_HAGAP = "/sites/VORUMU-OBRAS-HAGAP-/Documentos Compartilhados/HAGAP -"


def norm(txt):
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.lower().strip().split())


def subpastas(context, pasta):
    dados = sp.get_json(context, sp.endpoint_pastas(pasta))
    return [
        p for p in dados.get("value", [])
        if str(p.get("Name", "")).lower() != "forms"
    ]


def achar(context, pai, teste, descricao):
    achados = []
    for p in subpastas(context, pai):
        nome = str(p.get("Name", ""))
        if teste(norm(nome)):
            achados.append(p)

    if len(achados) == 1:
        return achados[0]["ServerRelativeUrl"]

    nomes = ", ".join(str(p.get("Name", "")) for p in subpastas(context, pai))
    if not achados:
        raise RuntimeError(
            f"{descricao} não localizada. Pastas disponíveis em {pai}: {nomes}"
        )

    raise RuntimeError(
        f"{descricao} ficou ambígua: "
        + ", ".join(str(p.get("Name", "")) for p in achados)
    )


def localizar_bmd2026(context):
    emitidos = achar(
        context,
        RAIZ_HAGAP,
        lambda n: n.startswith("8-") and "bmd" in n and "ffo" in n and "emitidos" in n,
        "Pasta 8 - BMD/FFO Emitidos",
    )

    return achar(
        context,
        emitidos,
        lambda n: n == "bmd 2026",
        "Subpasta BMD 2026",
    )


def aguardar_login(context, pagina, limite=300):
    pagina.goto(sp.URL_INICIAL, wait_until="domcontentloaded", timeout=120000)
    inicio = time.time()
    ultimo = ""

    while time.time() - inicio < limite:
        try:
            sp.get_json(context, sp.endpoint_pastas(RAIZ_HAGAP))
            return
        except Exception as exc:
            ultimo = str(exc)
            print(
                "\rAguardando autenticação no Edge... faça login/MFA.          ",
                end="",
                flush=True,
            )
            time.sleep(3)

    print()
    raise RuntimeError(f"Não consegui validar o SharePoint. Último erro: {ultimo}")


def sincronizar(context, pasta):
    arquivos = sp.listar_recursivo(context, pasta)
    estado = sp.carregar_estado()
    estado.setdefault("arquivos", {})
    baixar = [item for item in arquivos if sp.precisa_baixar(item, estado)]

    print()
    print("=" * 72)
    print("SINCRONIZAÇÃO BMD 2026")
    print(f"Pasta remota: {pasta}")
    print(f"Encontrados: {len(arquivos)}")
    print(f"Novos/alterados: {len(baixar)}")
    print(f"Destino base: {sp.DESTINO}")
    print("=" * 72)

    concluidos = 0
    for indice, item in enumerate(baixar, start=1):
        nome = item.get("Name", item.get("ServerRelativeUrl"))
        print(f"[{indice}/{len(baixar)}] {nome}")

        destino, bytes_salvos = sp.baixar_arquivo(context, item)
        remoto = item["ServerRelativeUrl"]

        estado["arquivos"][remoto] = {
            "modified": str(item.get("TimeLastModified", "")),
            "length": str(item.get("Length", bytes_salvos)),
            "local": str(destino),
            "sincronizado_em": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        sp.salvar_estado(estado)
        concluidos += 1

    print(
        f"[CONFIRMADO] BMD 2026 sincronizados nesta execução: {concluidos}. "
        f"Sem alteração: {len(arquivos) - len(baixar)}."
    )


def main():
    sp.DESTINO.mkdir(parents=True, exist_ok=True)
    sp.PERFIL_EDGE.mkdir(parents=True, exist_ok=True)

    print("HAGAP — BMD 2026")
    print("[CONFIRMADO] Somente leitura no SharePoint.")
    print("[CONFIRMADO] Site HAGAP e /boletim não são alterados.")

    with sp.sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(sp.PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=False,
        )

        try:
            pagina = context.pages[0] if context.pages else context.new_page()
            aguardar_login(context, pagina)
            pasta = localizar_bmd2026(context)
            print(f"\n[CONFIRMADO] Pasta localizada: {pasta}")
            sincronizar(context, pasta)
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
