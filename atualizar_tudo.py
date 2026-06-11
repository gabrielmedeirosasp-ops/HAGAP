"""
atualizar_tudo.py — HAGAP Atualização Completa
================================================
Roda os três scripts em sequência:

  PASSO 1 → leitor_bmd.py     (lê PDFs de medição → materiais_por_projeto.xlsx)
  PASSO 2 → extrair_bmd.py    (extrai rateios e datas → medicoes.xlsx + analise_financeira.xlsx)
  PASSO 3 → gerador.py        (gera db.json + resultado.xlsx + sincroniza com Render)

Como usar:
  1. Coloque este arquivo na mesma pasta que leitor_bmd.py, extrair_bmd.py e gerador.py
  2. Execute:  python atualizar_tudo.py
  3. Acompanhe o log na tela e no arquivo "atualizar_tudo.log"

Opções:
  python atualizar_tudo.py --somente-local   (não sincroniza com Render no gerador.py)
  python atualizar_tudo.py --pular 1         (pula o passo 1 e começa do 2)
  python atualizar_tudo.py --pular 1 2       (pula passos 1 e 2, só roda gerador)
"""

import sys
import os
import subprocess
import time
import datetime

# ─── Configuração ──────────────────────────────────────────────────────────────
PASTA_SCRIPTS = os.path.dirname(os.path.abspath(__file__))  # pasta deste arquivo
LOG_FILE      = os.path.join(PASTA_SCRIPTS, "atualizar_tudo.log")

PASSOS = [
    {
        "num":    1,
        "nome":   "Leitor BMD",
        "script": "leitor_bmd.py",
        "desc":   "Lê PDFs de medição e gera materiais_por_projeto.xlsx",
    },
    {
        "num":    2,
        "nome":   "Extrair BMD",
        "script": "extrair_bmd.py",
        "desc":   "Extrai rateios/datas e gera medicoes.xlsx + analise_financeira.xlsx",
    },
    {
        "num":    3,
        "nome":   "Gerador + Sync Render",
        "script": "gerador.py",
        "desc":   "Gera db.json + resultado.xlsx e sincroniza com o Render",
    },
]

# ─── Helpers ────────────────────────────────────────────────────────────────────

def _cor(texto, codigo):
    """ANSI color — desativado automaticamente no Windows sem suporte."""
    if sys.stdout.isatty():
        return f"\033[{codigo}m{texto}\033[0m"
    return texto

def verde(t):   return _cor(t, "92")
def vermelho(t):return _cor(t, "91")
def amarelo(t): return _cor(t, "93")
def azul(t):    return _cor(t, "94")
def negrito(t): return _cor(t, "1")

def log(msg, arquivo=None):
    ts  = datetime.datetime.now().strftime("%H:%M:%S")
    linha = f"[{ts}] {msg}"
    print(linha)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        # Remove códigos ANSI para o log
        import re
        limpa = re.sub(r'\033\[[0-9;]*m', '', linha)
        f.write(limpa + "\n")

def separador(char="─", largura=64):
    return char * largura

# ─── Executar passo ─────────────────────────────────────────────────────────────

def executar_passo(passo, args_extras=None):
    script = os.path.join(PASTA_SCRIPTS, passo["script"])

    if not os.path.exists(script):
        log(vermelho(f"  ❌ Script não encontrado: {script}"))
        return False

    cmd = [sys.executable, script]
    if args_extras:
        cmd += args_extras

    log(azul(f"  ▶ Executando: {passo['script']}"))
    inicio = time.time()

    try:
        resultado = subprocess.run(
            cmd,
            cwd=PASTA_SCRIPTS,
            capture_output=False,   # exibe saída em tempo real no terminal
            text=True,
        )
        duracao = time.time() - inicio

        if resultado.returncode == 0:
            log(verde(f"  ✅ Concluído em {duracao:.1f}s"))
            return True
        else:
            log(vermelho(f"  ❌ Erro (código {resultado.returncode}) em {duracao:.1f}s"))
            return False

    except Exception as e:
        log(vermelho(f"  ❌ Exceção ao executar {passo['script']}: {e}"))
        return False

# ─── Main ────────────────────────────────────────────────────────────────────────

def main():
    # Processa argumentos
    args = sys.argv[1:]
    somente_local = "--somente-local" in args
    pular = set()
    if "--pular" in args:
        idx = args.index("--pular")
        for v in args[idx+1:]:
            if v.isdigit():
                pular.add(int(v))
            else:
                break

    # Cabeçalho
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write("\n" + "=" * 64 + "\n")
        f.write(f"  INÍCIO: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 64 + "\n")

    print()
    print(negrito(separador("═")))
    print(negrito("  HAGAP — Atualização Completa"))
    print(negrito(f"  {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"))
    print(negrito(separador("═")))
    if somente_local:
        print(amarelo("  ⚠  Modo somente local (sem sync Render)"))
    if pular:
        print(amarelo(f"  ⚠  Pulando passo(s): {sorted(pular)}"))
    print()

    resultados = {}
    inicio_total = time.time()

    for passo in PASSOS:
        num = passo["num"]

        print(separador())
        log(negrito(f"PASSO {num}/3 — {passo['nome']}"))
        log(f"  {passo['desc']}")
        print()

        if num in pular:
            log(amarelo("  ⏭  Pulado (--pular)"))
            resultados[num] = "pulado"
            print()
            continue

        # Passa --somente-local para gerador.py modificando SINCRONIZAR_RENDER
        # (gerador.py não aceita argumento, então manipulamos via env var)
        extras = None
        if num == 3 and somente_local:
            os.environ["HAGAP_SOMENTE_LOCAL"] = "1"
        elif num == 3:
            os.environ.pop("HAGAP_SOMENTE_LOCAL", None)

        ok = executar_passo(passo, extras)
        resultados[num] = "ok" if ok else "erro"
        print()

        if not ok:
            resp = input(amarelo(f"  ⚠  Passo {num} falhou. Continuar mesmo assim? [s/N] ")).strip().lower()
            if resp != "s":
                log(vermelho("  ✋ Processo interrompido pelo usuário."))
                break
            print()

    # Resumo final
    duracao_total = time.time() - inicio_total
    print(separador("═"))
    log(negrito("  RESUMO FINAL"))
    print(separador())

    todos_ok = True
    for passo in PASSOS:
        num = passo["num"]
        status = resultados.get(num)
        if status == "ok":
            icone = verde("✅ OK     ")
        elif status == "pulado":
            icone = amarelo("⏭  Pulado ")
        elif status == "erro":
            icone = vermelho("❌ ERRO   ")
            todos_ok = False
        else:
            icone = amarelo("⬜ Não ran")
            todos_ok = False
        log(f"  Passo {num} — {passo['nome']:30s} {icone}")

    print(separador())
    log(f"  Tempo total: {duracao_total:.1f}s")
    if todos_ok:
        log(verde("  🎉 Tudo atualizado com sucesso!"))
    else:
        log(vermelho("  ⚠  Houve erros. Verifique acima ou o arquivo atualizar_tudo.log"))
    print(separador("═"))
    print()

    # Mantém janela aberta se rodando com duplo-clique no Windows
    if sys.platform == "win32" and not sys.stdout.isatty():
        input("Pressione ENTER para fechar...")

if __name__ == "__main__":
    main()
