import argparse
import sys
import time

import sync_sharepoint as sp

PASTAS = {
    "obras": "/sites/VORUMU-OBRAS-HAGAP-/Documentos Compartilhados/HAGAP -/3- Obras para Execução",
    "rmds": "/sites/VORUMU-OBRAS-HAGAP-/Documentos Compartilhados/HAGAP -/7- RMDs",
}


def aguardar_login_generico(context, pagina, pasta_teste, limite_segundos=300):
    pagina.goto(sp.URL_INICIAL, wait_until="domcontentloaded", timeout=120000)

    inicio = time.time()
    ultimo_erro = ""

    while time.time() - inicio < limite_segundos:
        try:
            sp.get_json(context, sp.endpoint_arquivos(pasta_teste))
            return
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


def listar_recursivo_com_progresso(context, pasta_raiz):
    arquivos = []
    pilha = [pasta_raiz]
    pastas_lidas = 0

    print(f"[VARREDURA] Iniciando: {pasta_raiz}", flush=True)

    while pilha:
        atual = pilha.pop()
        pastas_lidas += 1

        js_arquivos = sp.get_json(context, sp.endpoint_arquivos(atual))
        novos_arquivos = js_arquivos.get("value", [])
        arquivos.extend(novos_arquivos)

        js_pastas = sp.get_json(context, sp.endpoint_pastas(atual))
        subpastas = []
        for pasta in js_pastas.get("value", []):
            nome = pasta.get("Name", "")
            if str(nome).lower() == "forms":
                continue
            caminho = pasta.get("ServerRelativeUrl")
            if caminho:
                subpastas.append(caminho)
                pilha.append(caminho)

        print(
            f"[VARREDURA] Pastas lidas: {pastas_lidas} | "
            f"Arquivos encontrados: {len(arquivos)} | "
            f"Fila: {len(pilha)} | Atual: {atual}",
            flush=True,
        )

    print(
        f"[VARREDURA] Concluída. Pastas: {pastas_lidas} | "
        f"Arquivos: {len(arquivos)}",
        flush=True,
    )
    return arquivos


def sincronizar_pasta(context, pasta):
    arquivos = listar_recursivo_com_progresso(context, pasta)
    estado = sp.carregar_estado()
    estado.setdefault("arquivos", {})

    baixar = [item for item in arquivos if sp.precisa_baixar(item, estado)]

    print()
    print("=" * 72)
    print(f"SINCRONIZAÇÃO: {pasta.split('/')[-1]}")
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
        f"[CONFIRMADO] Sincronizados nesta execução: {concluidos}. "
        f"Sem alteração: {len(arquivos) - len(baixar)}."
    )

    return len(arquivos), concluidos


def main():
    parser = argparse.ArgumentParser(
        description="HAGAP — sincronização das pastas Obras para Execução e RMDs"
    )
    parser.add_argument(
        "--pastas",
        nargs="+",
        choices=sorted(PASTAS),
        default=["obras", "rmds"],
    )
    args = parser.parse_args()

    sp.DESTINO.mkdir(parents=True, exist_ok=True)
    sp.PERFIL_EDGE.mkdir(parents=True, exist_ok=True)

    selecionadas = [PASTAS[n] for n in args.pastas]

    print("HAGAP — SharePoint COPEL")
    print("[CONFIRMADO] Somente leitura no SharePoint.")
    print("[CONFIRMADO] O site HAGAP e /boletim não são alterados.")
    print("Pastas selecionadas:")
    for pasta in selecionadas:
        print(f"  - {pasta.split('/')[-1]}")

    with sp.sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(sp.PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=False,
        )

        try:
            pagina = context.pages[0] if context.pages else context.new_page()
            aguardar_login_generico(context, pagina, selecionadas[0])
            print("\n[CONFIRMADO] Sessão SharePoint validada.")

            total_encontrados = 0
            total_sincronizados = 0

            for pasta in selecionadas:
                encontrados, sincronizados = sincronizar_pasta(context, pasta)
                total_encontrados += encontrados
                total_sincronizados += sincronizados

            print()
            print("=" * 72)
            print("RESUMO")
            print(f"Arquivos encontrados: {total_encontrados}")
            print(f"Arquivos novos/atualizados: {total_sincronizados}")
            print(f"Destino: {sp.DESTINO}")
            print("=" * 72)
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
