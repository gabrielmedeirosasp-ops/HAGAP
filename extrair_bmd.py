
"""
extrair_bmd.py — HAGAP v2.1
Extrai valores de RATEIO HAGAP dos PDFs de medição (BMD/FFO),
cruza com resultado.xlsx (custo da AES) e gera:
  → medicoes.xlsx           (BMDs extraídos por projeto)
  → analise_financeira.xlsx (lucro / prejuízo por projeto)
"""

import os, re
import pdfplumber
import pandas as pd

# ─── CONFIGURAÇÕES ───────────────────────────────────────────────────
PASTA_PDFS      = r"C:\HAGAP\medicoes"
ARQUIVO_EXCEL   = r"C:\Users\user\Desktop\PDF AUTOMATICO\HAGAP_WEB\medicoes.xlsx"
RESULTADO_EXCEL = r"C:\Users\user\Desktop\PDF AUTOMATICO\HAGAP_WEB\resultado.xlsx"
ANALISE_EXCEL   = r"C:\Users\user\Desktop\PDF AUTOMATICO\HAGAP_WEB\analise_financeira.xlsx"
# ─────────────────────────────────────────────────────────────────────


def extrair_texto_pdf(caminho):
    texto = ""
    try:
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"
    except Exception:
        print("Erro ao ler:", caminho)
    return texto


def extrair_projeto_nome(arquivo):
    match = re.search(r"(\d{6,8}[A-Za-z]?)", arquivo)
    if match:
        return match.group(1).upper()
    return None


def extrair_tipo_medicao(texto, arquivo):
    arq_up = arquivo.upper()
    txt_up = texto.upper()
    if "FFO" in arq_up or "FINAL" in arq_up:
        return "FFO"
    if "BMD" in arq_up or "PARCIAL" in arq_up:
        return "BMD"
    palavras_final = ["MEDIÇÃO FINAL", "MEDICAO FINAL", "FFO", "FINAL"]
    palavras_parc  = ["MEDIÇÃO PARCIAL", "MEDICAO PARCIAL", "BMD", "PARCIAL"]
    pos_f = min((txt_up.find(p) for p in palavras_final if p in txt_up), default=9999)
    pos_p = min((txt_up.find(p) for p in palavras_parc  if p in txt_up), default=9999)
    return "FFO" if pos_f < pos_p else "BMD"


def extrair_data_medicao(texto):
    """
    Extrai a data de medição do BMD/FFO no formato COPEL.
    Prioridade: DT.MEDIÇÃO, DATATERMINO, competência, período, emissão.
    Retorna string normalizada DD/MM/AAAA ou vazio.
    """
    RE_DATA = re.compile(r'\b(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})\b')

    def _normalizar(d):
        """Converte DD.MM.AAAA ou DD-MM-AAAA para DD/MM/AAAA."""
        return d.replace('.', '/').replace('-', '/') if d else ''

    # 1. Campo DT.MEDIÇÃO do COPEL BMD:
    #    Header: "...DT.MEDIÇÃO MEDICAO..."  Valor na linha seguinte: "19.05.2026"
    m = re.search(
        r'DT\.MEDI[^\n]{0,40}\n[^\n]{0,120}?(\d{2}[./]\d{2}[./]\d{4})',
        texto, re.IGNORECASE
    )
    if m:
        return _normalizar(m.group(1))

    # 1b. DT.MEDIÇÃO concatenado na mesma linha (PDF sem quebra)
    m = re.search(r'DT\.MEDI[^\d]{0,30}(\d{2}[./]\d{2}[./]\d{4})', texto, re.IGNORECASE)
    if m:
        return _normalizar(m.group(1))

    # 2. DATATERMINO do FFO COPEL
    m = re.search(
        r'DATATERM[^\n]{0,40}\n[^\n]{0,120}?\d{2}[./]\d{2}[./]\d{4}\s+(\d{2}[./]\d{2}[./]\d{4})',
        texto, re.IGNORECASE
    )
    if m:
        return _normalizar(m.group(1))
    m = re.search(r'DATATERM[^\d]{0,20}(\d{2}[./]\d{2}[./]\d{4})', texto, re.IGNORECASE)
    if m:
        return _normalizar(m.group(1))

    # 3. Outras palavras-chave
    CHAVES = [
        r'compet[eê]ncia', r'per[ií]odo', r'data\s*de\s*medi[cç][aã]o',
        r'data\s*medi[cç][aã]o', r'data\s*emiss[aã]o', r'emiss[aã]o',
    ]
    for chave in CHAVES:
        m = re.search(
            rf'(?i){chave}[^\n]{{0,100}}?(\d{{1,2}}[./\-]\d{{1,2}}[./\-]\d{{2,4}})',
            texto, re.DOTALL
        )
        if m:
            return _normalizar(m.group(1))

    # 4. Mês por extenso: ex "Outubro/2023" ou "outubro de 2023"
    m2 = re.search(
        r'(?i)(janeiro|fevereiro|mar[cç]o|abril|maio|junho|julho|agosto'
        r'|setembro|outubro|novembro|dezembro)[/\s\-de]*(\d{4})',
        texto
    )
    if m2:
        meses = {
            'janeiro':'01','fevereiro':'02','marco':'03','março':'03','abril':'04',
            'maio':'05','junho':'06','julho':'07','agosto':'08','setembro':'09',
            'outubro':'10','novembro':'11','dezembro':'12'
        }
        mes_num = meses.get(m2.group(1).lower().replace('ç','c'), '01')
        return f'01/{mes_num}/{m2.group(2)}'

    return ''


def extrair_valor_rateio(texto):
    # Estratégia 1: linha com RATEIO + HAGAP
    for linha in texto.split("\n"):
        lu = linha.upper()
        if "RATEIO" in lu and "HAGAP" in lu:
            m = re.search(r"R\$\s*([\d\.]+,\d{2})", linha)
            if m:
                return _parse_brl(m.group(1))
    # Estratégia 2: linha com RATEIO + CUSTO
    for linha in texto.split("\n"):
        lu = linha.upper()
        if "RATEIO" in lu and "CUSTO" in lu:
            m = re.search(r"R\$\s*([\d\.]+,\d{2})", linha)
            if m:
                return _parse_brl(m.group(1))
    # Estratégia 3: trecho após a palavra RATEIO
    pos = texto.lower().find("rateio")
    if pos != -1:
        m = re.search(r"R\$\s*([\d\.]+,\d{2})", texto[pos:pos + 500])
        if m:
            return _parse_brl(m.group(1))
    return 0.0


def _parse_brl(s):
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fmt_brl(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_brl_r(v):
    sinal = "-" if v < 0 else ""
    return f"{sinal}R$ {_fmt_brl(abs(v))}"


# ─── EXTRAÇÃO DOS PDFs ───────────────────────────────────────────────
print("🔎 Lendo PDFs em:", PASTA_PDFS, "\n")

dados = []

for arquivo in sorted(os.listdir(PASTA_PDFS)):
    if not arquivo.lower().endswith(".pdf"):
        continue
    print("Processando:", arquivo)
    caminho = os.path.join(PASTA_PDFS, arquivo)
    projeto = extrair_projeto_nome(arquivo)
    if not projeto:
        print("  ⚠ Número de projeto não encontrado no nome do arquivo")
        continue
    texto = extrair_texto_pdf(caminho)
    valor = extrair_valor_rateio(texto)
    tipo  = extrair_tipo_medicao(texto, arquivo)
    data  = extrair_data_medicao(texto)
    print(f"  → Projeto: {projeto} | Tipo: {tipo} | Rateio: R$ {_fmt_brl(valor)} | Data: {data or '—'}")
    dados.append({"Projeto": projeto, "Tipo": tipo, "Arquivo": arquivo, "Valor": valor, "Data": data})

df_raw = pd.DataFrame(dados)

# ─── SALVAR medicoes.xlsx ────────────────────────────────────────────
if not df_raw.empty:
    df_exp = df_raw.copy()
    df_exp["Valor R$"] = df_exp["Valor"].map(_fmt_brl)
    df_exp.drop(columns=["Valor"], inplace=True)
    # Garante coluna Data na saída
    if "Data" not in df_exp.columns:
        df_exp["Data"] = ""
    df_exp = df_exp[["Projeto", "Tipo", "Arquivo", "Valor R$", "Data"]]
    df_exp.to_excel(ARQUIVO_EXCEL, index=False)
    print(f"\n✅ medicoes.xlsx salvo com {len(df_exp)} linha(s).")
else:
    print("\n⚠ Nenhum dado extraído.")
    exit()

# ─── SOMAR POR PROJETO ───────────────────────────────────────────────
df_soma = df_raw.groupby("Projeto", as_index=False)["Valor"].sum()
df_soma.rename(columns={"Valor": "Total_Medido"}, inplace=True)

# ─── LER resultado.xlsx (custo AES) ──────────────────────────────────
print("\n📂 Lendo resultado.xlsx...")

try:
    df_res = pd.read_excel(RESULTADO_EXCEL)
    df_res.columns = [c.lower().strip() for c in df_res.columns]

    def parse_custo(s):
        if pd.isna(s) or s == "":
            return 0.0
        s = re.sub(r"[^\d,\.]", "", str(s))
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    df_res["custo_float"] = df_res["custo"].apply(parse_custo) if "custo" in df_res.columns else 0.0
    df_res["projeto"] = df_res["projeto"].astype(str).str.strip().str.upper()
    df_custo = df_res[["projeto", "custo", "custo_float"]].copy()
    df_custo.rename(columns={"projeto": "Projeto", "custo": "Custo_AES_Str", "custo_float": "Custo_AES"}, inplace=True)

except FileNotFoundError:
    print("  ⚠ resultado.xlsx não encontrado.")
    df_custo = pd.DataFrame(columns=["Projeto", "Custo_AES_Str", "Custo_AES"])
except Exception as e:
    print(f"  ⚠ Erro: {e}")
    df_custo = pd.DataFrame(columns=["Projeto", "Custo_AES_Str", "Custo_AES"])

# ─── MERGE BMD x AES ─────────────────────────────────────────────────
df_soma["Projeto"] = df_soma["Projeto"].astype(str).str.strip().str.upper()
df_merge = df_soma.merge(df_custo, on="Projeto", how="left")
df_merge["Custo_AES"] = df_merge["Custo_AES"].fillna(0.0)
df_merge["Resultado"] = df_merge["Total_Medido"] - df_merge["Custo_AES"]
df_merge["Situacao"]  = df_merge["Resultado"].apply(
    lambda v: "✅ LUCRO" if v > 0 else ("❌ PREJUÍZO" if v < 0 else "—")
)
df_merge["Percentual"] = df_merge.apply(
    lambda row: f"{(row['Resultado'] / row['Custo_AES'] * 100):+.1f}%"
    if row["Custo_AES"] > 0 else "—", axis=1
)

# ─── MONTAR SAÍDA ────────────────────────────────────────────────────
df_out = pd.DataFrame({
    "Projeto":            df_merge["Projeto"],
    "Total BMD/FFO (R$)": df_merge["Total_Medido"].map(_fmt_brl),
    "Custo AES (R$)":     df_merge["Custo_AES"].map(_fmt_brl),
    "Resultado (R$)":     df_merge["Resultado"].map(_fmt_brl_r),
    "Situação":           df_merge["Situacao"],
    "Variação %":         df_merge["Percentual"],
})

# ─── SALVAR analise_financeira.xlsx ──────────────────────────────────
with pd.ExcelWriter(ANALISE_EXCEL, engine="openpyxl") as writer:
    df_out.to_excel(writer, index=False, sheet_name="Análise Financeira")
    ws = writer.sheets["Análise Financeira"]
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

print(f"\n✅ analise_financeira.xlsx salvo com {len(df_out)} projeto(s).")

# ─── RESUMO ──────────────────────────────────────────────────────────
total_medido    = df_merge["Total_Medido"].sum()
total_custo     = df_merge["Custo_AES"].sum()
resultado_total = total_medido - total_custo
situacao_geral  = "✅ LUCRO" if resultado_total > 0 else ("❌ PREJUÍZO" if resultado_total < 0 else "—")

print("\n" + "=" * 55)
print(f"  📊 RESUMO GERAL")
print(f"  Projetos analisados : {len(df_merge)}")
print(f"  Total Medido (R$)   : {_fmt_brl(total_medido)}")
print(f"  Total Custo AES (R$): {_fmt_brl(total_custo)}")
print(f"  Resultado (R$)      : {_fmt_brl_r(resultado_total)}")
print(f"  Situação            : {situacao_geral}")
print("=" * 55)
print("\n✅ FINALIZADO COM SUCESSO!")
