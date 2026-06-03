import os
import re
import pdfplumber
import pandas as pd

# =====================================================
# PASTA DOS PDFs
# =====================================================

PASTA_PDFS = r"C:\HAGAP\medicoes"

# =====================================================
# ARQUIVO DE SAÍDA
# =====================================================

ARQUIVO_EXCEL = r"C:\HAGAP\medicoes\materiais_por_projeto.xlsx"


# =====================================================
# FUNÇÕES
# =====================================================

def moeda(valor):

    valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except:
        return 0


def ler_pdf(caminho):

    texto = ""

    with pdfplumber.open(caminho) as pdf:

        for pagina in pdf.pages:

            t = pagina.extract_text(x_tolerance=2)

            if t:
                texto += "\n" + t

    return texto


def pegar_projeto(texto):

    m = re.search(
        r'FISCAL Nº PROJETO.*?\n.*?(\d{7})',
        texto,
        re.S
    )

    if m:
        return m.group(1)

    return "SEM PROJETO"


def pegar_total_material(texto):

    m = re.search(
        r'TOTAL MATERIAL\s+([\d\.,]+)',
        texto
    )

    if m:
        return moeda(m.group(1))

    return 0


def pegar_total_geral(texto):

    m = re.search(
        r'RATEIO DO CUSTO TOTAL.*?R\$\s*([\d\.,]+)',
        texto
    )

    if m:
        return moeda(m.group(1))

    return 0


def pegar_materiais(texto):

    materiais = []

    linhas = texto.split("\n")

    dentro_tabela = False

    for linha in linhas:

        if "DISCRIMINAÇÃO DOS ITENS" in linha:
            dentro_tabela = True
            continue

        if "TOTAL MATERIAL" in linha:
            dentro_tabela = False

        if not dentro_tabela:
            continue

        linha = linha.strip()

        # MATERIAL
        m = re.search(
            r'^(\d+)\s+(\d+)\s+(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$',
            linha
        )

        if m:

            item = m.group(1)
            codigo = m.group(2)
            descricao = m.group(3)
            quantidade = moeda(m.group(4))
            unitario = moeda(m.group(5))
            total = moeda(m.group(6))

            materiais.append({
                "item": item,
                "codigo": codigo,
                "descricao": descricao,
                "quantidade": quantidade,
                "valor_unitario": unitario,
                "valor_total": total
            })

    return materiais


# =====================================================
# PROCESSAMENTO
# =====================================================

resumo = []
materiais_geral = []

for arquivo in os.listdir(PASTA_PDFS):

    if not arquivo.lower().endswith(".pdf"):
        continue

    caminho = os.path.join(PASTA_PDFS, arquivo)

    print(f"\nLENDO: {arquivo}")

    try:

        texto = ler_pdf(caminho)

        projeto = pegar_projeto(texto)

        total_material = pegar_total_material(texto)

        total_geral = pegar_total_geral(texto)

        total_mao_obra = total_geral - total_material

        materiais = pegar_materiais(texto)

        # ========================================
        # RESUMO POR PROJETO
        # ========================================

        resumo.append({
            "projeto": projeto,
            "arquivo": arquivo,
            "total_geral": total_geral,
            "total_material": total_material,
            "total_mao_obra": total_mao_obra,
            "possui_material": "SIM" if total_material > 0 else "NAO"
        })

        # ========================================
        # MATERIAIS POR PROJETO
        # ========================================

        for mat in materiais:

            materiais_geral.append({
                "projeto": projeto,
                "arquivo": arquivo,
                "item": mat["item"],
                "codigo": mat["codigo"],
                "descricao": mat["descricao"],
                "quantidade": mat["quantidade"],
                "valor_unitario": mat["valor_unitario"],
                "valor_total": mat["valor_total"]
            })

        print(f"Materiais encontrados: {len(materiais)}")

    except Exception as e:

        print(f"ERRO: {e}")


# =====================================================
# DATAFRAMES
# =====================================================

df_resumo = pd.DataFrame(resumo)

df_materiais = pd.DataFrame(materiais_geral)

# =====================================================
# AGRUPA POR PROJETO
# =====================================================

df_total_projeto = (
    df_materiais
    .groupby(["projeto", "descricao"], as_index=False)
    .agg({
        "quantidade": "sum",
        "valor_total": "sum"
    })
)

# =====================================================
# EXPORTA EXCEL
# =====================================================

with pd.ExcelWriter(
    ARQUIVO_EXCEL,
    engine="openpyxl"
) as writer:

    df_resumo.to_excel(
        writer,
        sheet_name="Resumo",
        index=False
    )

    df_materiais.to_excel(
        writer,
        sheet_name="Materiais_Detalhados",
        index=False
    )

    df_total_projeto.to_excel(
        writer,
        sheet_name="Materiais_por_Projeto",
        index=False
    )

print("\n===================================")
print("FINALIZADO")
print("ARQUIVO GERADO:")
print(ARQUIVO_EXCEL)
print("===================================")