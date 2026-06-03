
"""
gerador.py — HAGAP_WEB v2.5
Roda apenas no PC local.

CORREÇÕES v2.5:
  → processar_aes: extrai MÚLTIPLOS projetos por PDF de AE
  → processar_medicoes: cria registros ORFÃOS para BMDs sem projeto
  → gerar_db: projetos com FFO → executado=True automaticamente
CORREÇÕES v2.3:
  → Deduplicação final: mesmo número de projeto de fontes diferentes → MERGE em 1 registro
  → Rejeita entradas sem número de projeto válido (SOUSERVI, FISCAL, C801057 etc.)
  → US: rejeita qualquer valor com letras ou fora do range 0–9999
  → Local: rejeita strings técnicas (/ESTADO, /UF, DadosdaObra, etc.)
  → Nunca apaga registros existentes (merge)
"""

import os, re, json, copy, traceback
from datetime import datetime, date
import pandas as pd
import pdfplumber
try:
    import requests as _requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    import fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# ══════════════════════════════════════════════════════════════════════
#  CONFIGURAÇÕES
# ══════════════════════════════════════════════════════════════════════
PASTA_AES       = r"C:\HAGAP\AES"
PASTA_PROJETOS  = r"C:\HAGAP\PROJETOS"
PASTA_MEDICOES  = r"C:\HAGAP\medicoes"
SAIDA_JSON      = "db.json"
SAIDA_EXCEL     = "resultado.xlsx"
DIAS_ALERTA     = 30
DIAS_CRITICO    = 7
# ── Sincronização automática com o Render ──────────────────────────
RENDER_URL          = "https://hagap.onrender.com"   # URL do seu site
SENHA_MIGRACAO      = "HAGAP_MIGRAR"                 # deve ser igual ao env SENHA_MIGRACAO no Render
SINCRONIZAR_RENDER  = True   # False = não envia nada ao Render (só gera db.json local)
# ID da pasta raiz HAGAP no Google Drive (subpastas = projetos)
DRIVE_FOLDER_ID = "1yWPVwMMmSfDItku95Bq6lLGlhSkLrS0u"
# Service Account JSON em Base64 (mesmo do app.py)
_GOOGLE_SA_B64  = "ewogICJ0eXBlIjogInNlcnZpY2VfYWNjb3VudCIsCiAgInByb2plY3RfaWQiOiAiaGFnYXAtZHJpdmUiLAogICJwcml2YXRlX2tleV9pZCI6ICIxN2Y5YTBlZTkzN2Y3MWU2MWY4Yjk5ZGFlNmVkNDA5YmE2NmZhNmJmIiwKICAicHJpdmF0ZV9rZXkiOiAiLS0tLS1CRUdJTiBQUklWQVRFIEtFWS0tLS0tXG5NSUlFdXdJQkFEQU5CZ2txaGtpRzl3MEJBUUVGQUFTQ0JLVXdnZ1NoQWdFQUFvSUJBUURQV3o0TjNHaTBBU2tQXG5wd09yVTJGMW00eFcyOFJCQVcxV3hxZEptckUyN2VuVlZISzlwZFNhbmlDZ29QdStyMmczN1VQSlVjTDBnUXJDXG5jeTdEYnlZUmlrdjhmZ1YxMjlQTG94SldLOWZWdnF2MWVUQno4VU1CU3FqZFBLaHhNZ0U1RUpYVCtyYjRLQnNwXG5HcUdaczBnMWo5OTNWa3VHUFhnWStUaUhGMEV0TmV1dEhKOHNuUUxXazNZMzVJSzhWQzBTc0dWTkhQRkc1RHZ5XG4ydFp2dCtRWTdHSXRiak5mNXltOGgxWXVrZ3ZmSFZiemRaTG9wSmZpUUx4ckNSU3lYQWtNZkhqM01iRmxxNDJBXG4xcUdLMzBSN2xSWndCZHJJaUVzZVB2UmhlODd1Vm4xY09VVlVNN1NVeDBKVitMQmhWWVpja1pydW1vUmFaZG1ZXG4wKy9FMEUxQkFnTUJBQUVDZ2Y5N1ZXdWg0dlhxZEwvVityeEl4bmQvOUhiR2JLZVFTU2RxOUVEMkFkNWpMQWswXG5xOTRteGEyZzNRTm43YjgxaDJLeEI1MWlJdDM5NXJ6RWFLNVA3UUZaUFhxazNBblp6cjhHU0liVVhZK1J3WHhqXG5SeU9wSDR4TTA1T0FYWkpaVHBBTzJTQXJaME1EcDMybW5UVkxGTnpHV0tVQXF4dGRzZ0loNVZxbFAzZXpTY3NyXG5kNDJ4U3EwSDBwcUQ0d09xWTBtcjZ2SFhtUU4vZEJ6dk1jSHB5b3ZpakRNY05IK3hlZFRhYXJSS0ZIWEV4NUJMXG50OS9IaVhkNWNGUDAxdDFySXBIVzNtc2FFK1FjcGFmTmZzajVWTFBRcThQU0svUWxkTHpoUFl4dXN0Zkg0SzM2XG5kMDhFNW50K3VRNXBuSWJnTjlXcVM0N2pKNjJ0RFZiUzVUVVBVMkVDZ1lFQThTTXJPcUNzYzVTRnNaOENDb1FIXG5UZGxKRTdqdjdqRitMaTZTSUpqQmFvN2lGbnlTVDN3eXB1K1lENEFDcnZ0NEFRZ2JjRWF3Qk5yVGJqUXExc2FPXG5CclNCZXZkWjFSSVhkVTNtL0RQT2dMMDF0OWl3U2JOT0lyVGtIamMzckdSMWtEZ3hOMlBlMkFTODN4aU9NQTVnXG5YQXgyaXRBWUVzWTVNeUxYUkRWQnltRUNnWUVBM0NNTjV6UnR4K09LQWJCSDFJYkV0N01JSjMvb25peTZkeE1YXG5FekxhN2FCejRkMWgvdkJ1TkVKQnA1Mm9uU3JkSVVFSDg4c0M3SVBNUCtFZDhpc1B0bi80b0NGcFdpY3hWYWkzXG5VdTFhRmRqcTk0QzJhc1hnb0luc1lrOTM0bzlrWkdkblpGT3F6VUxLRVUzWFlSVzgwTkNtT3V6MjUrMkxQemlJXG5SQWVnTHVFQ2dZQm1lYWFtWEovTTRvSktjS09DYW9kY0d3b0dPcUtrSEI1ZWErWW51aU1BTU1aOS9lU0RxT2t4XG5ya0oxMjlYSUhRL3JpRkY0em1HQlBxbDVmT3Z6cUIrMVU4SnV5bTAvc2xlTHYxWjVPVjc1c084Q09URDVzUngwXG5mV0V3UWhFRHp4RnkvNTVTeHJ0dUc5MVVCZkJ2RmJ6M2dVaGpSM25qaUovSVRuY2pBbXI2SVFLQmdRQ2tHSWxqXG5vT1I0emJmeFhFdWZORHJ0eU1vNlQ1SENYd3M4ZWdHSWNTOWJWWEVzbkE5UnNENG1QSUdlaGRyTUZjaXk0andnXG5VbVBDbE5pcmdZOEdGMjFtR0d6b1NSKzBjV1RJT3JVMVh2TDVPREttL3M1OE56Y2oxTXhkMkdsQWZLMVVYdlJtXG5pQ3ZaU0lGQ2R1a25XTUhnVXJpblVqOFhVaTZybU9PUytyQkFZUUtCZ0Y0c2dUcHRJbjdNNVc5dmY4OStCRmpxXG5hRmhjZk4veVlCSjhTaWZkTUY4QXlJSUFHV0dZQmMzY3E0ZEY5aDJaaGZScmExS0NuaXdwdytNcm5PMitkL2xmXG5XUmNXVXVydnBiK1E0TGNORVEwdVZwMmh3S3F3Z3ZuWDdyck15T0hsNkszUHNyaXNSWWl4b1l6Mkl6ejVuSEF2XG51ZUoxQ25DbnQxUXNqQ2xhR2JWNVxuLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLVxuIiwKICAiY2xpZW50X2VtYWlsIjogImhhZ2FwLXVwbG9hZGVyQGhhZ2FwLWRyaXZlLmlhbS5nc2VydmljZWFjY291bnQuY29tIiwKICAiY2xpZW50X2lkIjogIjExMzQzNzk5ODMzOTY1ODEwODIxNSIsCiAgImF1dGhfdXJpIjogImh0dHBzOi8vYWNjb3VudHMuZ29vZ2xlLmNvbS9vL29hdXRoMi9hdXRoIiwKICAidG9rZW5fdXJpIjogImh0dHBzOi8vb2F1dGgyLmdvb2dsZWFwaXMuY29tL3Rva2VuIiwKICAiYXV0aF9wcm92aWRlcl94NTA5X2NlcnRfdXJsIjogImh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL29hdXRoMi92MS9jZXJ0cyIsCiAgImNsaWVudF94NTA5X2NlcnRfdXJsIjogImh0dHBzOi8vd3d3Lmdvb2dsZWFwaXMuY29tL3JvYm90L3YxL21ldGFkYXRhL3g1MDkvaGFnYXAtdXBsb2FkZXIlNDBoYWdhcC1kcml2ZS5pYW0uZ3NlcnZpY2VhY2NvdW50LmNvbSIsCiAgInVuaXZlcnNlX2RvbWFpbiI6ICJnb29nbGVhcGlzLmNvbSIKfQo="
# ══════════════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────────────
#  LEITURA DE PDF
# ─────────────────────────────────────────────────────────────────────

def extrair_texto(caminho: str) -> str:
    texto = ""
    try:
        with pdfplumber.open(caminho) as pdf:
            for pg in pdf.pages:
                t = pg.extract_text()
                if t:
                    texto += t + "\n"
    except Exception:
        pass
    if not texto.strip() and HAS_FITZ:
        try:
            doc = fitz.open(caminho)
            for pg in doc:
                texto += pg.get_text() + "\n"
            doc.close()
        except Exception:
            pass
    return texto


def extrair_tabelas(caminho: str) -> list:
    tabelas = []
    try:
        with pdfplumber.open(caminho) as pdf:
            for pg in pdf.pages:
                tbls = pg.extract_tables()
                if tbls:
                    tabelas.extend(tbls)
    except Exception:
        pass
    return tabelas


# ─────────────────────────────────────────────────────────────────────
#  NÚMERO DO PROJETO — extração e validação
#  Formato válido COPEL: 7-8 dígitos + letra opcional  ex: 1615951 ou 1615951I
# ─────────────────────────────────────────────────────────────────────

_RE_NUM_PROJ  = re.compile(r'\b(1[5-9]\d{5}[A-Z]?)\b')
_RE_NUM_PROJ2 = re.compile(r'\b(\d{6,8}[A-Z]?)\b')

def numero_valido(s: str) -> bool:
    if not s:
        return False
    if re.fullmatch(r'1[5-9]\d{5}[A-Z]?', s):
        return True
    if re.fullmatch(r'[A-Z]{2,6}\d{6,12}', s):
        return True
    return False

def extrair_numero_do_texto(s: str) -> str:
    m = _RE_NUM_PROJ.search(s or "")
    return m.group(1) if m else ""

def numero_do_nome_pasta(nome: str) -> str:
    nome = nome.strip()
    m = re.match(r'^(1[5-9]\d{5}[A-Z]?)\b', nome)
    if m:
        return m.group(1)
    m = re.match(r'^([A-Z]{2,6}\d{5,12})\b', nome, re.I)
    if m:
        return m.group(1).upper()
    m = re.match(r'^(\d{6,8}[A-Z]?)\b', nome)
    if m:
        return m.group(1)
    return nome.split()[0]

def norm_num(s: str) -> str:
    m = _RE_NUM_PROJ.search(s or "")
    if m:
        return m.group(1).upper()
    m = re.search(r'([A-Z]{2,6}\d{6,12})', (s or "").upper())
    if m:
        return m.group(1)
    return re.sub(r'[\s\-_\.]', '', s or "").upper()

def numero_do_pdf(texto: str, tabelas: list) -> str:
    ROTULOS = ["Nº Projeto","N° Projeto","Número Projeto","Projeto","Nº Obra",
               "N° Obra","Código","OS","Contrato","Nº Contrato",
               "NºPROJETO","N°PROJETO","NOPROJETO","NUMPROJ"]
    for tbl in tabelas:
        for row in tbl:
            row_c = [str(c).strip() if c else "" for c in row]
            for i, cel in enumerate(row_c):
                for rot in ROTULOS:
                    if rot.lower() == cel.lower() and i+1 < len(row_c):
                        num = extrair_numero_do_texto(row_c[i+1])
                        if num and numero_valido(num):
                            return num
    for rot in ROTULOS:
        m = re.search(rf"(?i)\b{re.escape(rot)}\s*[:\-–]?\s*(1[5-9]\d{{5}}[A-Z]?)\b", texto)
        if m:
            return m.group(1)
    # FIX: nas AEs COPEL com múltiplos itens, o número do projeto fica na coluna PROJ.
    # Linha típica: "1 DG... I-22-2961902 D-22-1787301 InvestimentoRDU 04/08/2026 1.430,151 R$87.239,21 1520606"
    # A coluna PROJ. é o último número 7 dígitos no final da linha de item
    m_proj = re.search(
        r'Investimento[A-Z]+\s+\d{2}/\d{2}/\d{4}\s+[\d\.,]+\s+R\$[\d\.,]+\s+(1[5-9]\d{5}[A-Z]?)\b',
        texto
    )
    if m_proj:
        return m_proj.group(1)
    # Busca genérica — mas ignora números dentro de códigos PEP (ex: D-22-1787301, I-22-2961902)
    texto_sem_pep = re.sub(r'[A-Z]-\d{2}-(\d{7})', '', texto)
    m = _RE_NUM_PROJ.search(texto_sem_pep)
    return m.group(1) if m else ""


# ─────────────────────────────────────────────────────────────────────
#  FILTROS DE LIXO
# ─────────────────────────────────────────────────────────────────────

_LIXO_EXATO = {
    "dadosdaobra","dadosdosolicitante","part.financeira","dadosdoprojeto",
    "dadosgerais","informacoesgerais","dados da obra","dados do solicitante",
    "participação financeira","informações gerais","sim","não","yes","no",
    "n/a","na","s/n","—","-","/","x","participação","parte financeira",
    "us","custo proj","qtd.us custo proj","proj","fiscal","souservi",
}

_LIXO_PREFIXO = [
    "/", "dados", "part.", "infoma", "inform", "referencia:ps",
    ": cianorte", ": copel", "o;bt;", "uário:", "e806", "5 o projeto",
]

_LIXO_CONTEM = [
    "%participação", "/estado", "/uf", "cnpj", "segueparaexecução",
    "gleba são tomé", "sendo 7,50", "rol 2,000", "rol 0,800", "rol 0,400",
    "desnível", "faixas de servidão", "ntc813523",
]

def e_lixo(val: str) -> bool:
    if not val:
        return True
    v = val.strip()
    vl = v.lower().replace(" ", "")
    if not v or len(v) < 2:
        return True
    if vl in _LIXO_EXATO:
        return True
    if any(vl.startswith(p.replace(" ","")) for p in _LIXO_PREFIXO):
        return True
    if any(p in vl for p in _LIXO_CONTEM):
        return True
    if re.match(r'^[^a-z0-9]+$', vl):
        return True
    return False

# Sobrenomes comuns que indicam nomes de pessoa (não cidades)
_SOBRENOMES_PESSOA = {
    'AZEVEDO','OLIVEIRA','SILVA','SOUZA','SANTOS','FERREIRA','COSTA',
    'PEREIRA','ALVES','RODRIGUES','NASCIMENTO','LIMA','ARAÚJO','ARAUJO',
    'CARVALHO','GOMES','MARTINS','ROCHA','RIBEIRO','JESUS','SCHORK',
    'LOPES','NUNES','MOREIRA','ALMEIDA','CAMPOS','MENDES','DIAS',
}
# Palavras que indicam texto de descrição, não cidade
_DESCRICAO_PALAVRAS = {
    'PARA','ATENDER','OBRA','OBRAS','SISTEMA','VALOR','GLOSSÁRIO','GLOSSARIO',
    'AGUARDANDO','VISANDO','OPÇÃO','OPCAO','TIPO','REDE','TRIFÁSICO','TRIFASICO',
    'ACEITE','PARTICULAR','SOLICITANTE','REFERENTE','CONFORME','CONTRATO',
    'PROCESSO','SEGUINDO','FASE','ETAPA','SERVIÇO','SERVICO','INSTALAÇÃO',
    'INSTALACAO','EXECUÇÃO','EXECUCAO','PREVISTO','ESTIMADO','PARCIAL',
}

def validar_local(val: str) -> str:
    if not val:
        return ""
    v = str(val).strip()
    v = v.replace("\n", " ").replace("\r", " ")
    v = re.sub(r'\s+', ' ', v).strip()
    if not v:
        return ""
    v = v.upper()
    # Remove prefixos RDU/RDR antes de validar
    v = re.sub(r'^RD[UR]\s*[-–]?\s*', '', v).strip()
    v = re.sub(r'^RD[UR]-', '', v).strip()
    # Remove sufixo de UF (ex: -PR, /PR, - PR)
    v = re.sub(r'\s*[-/]\s*[A-Z]{2}\s*$', '', v).strip()
    if not v:
        return ""
    BLOQUEIOS = [
        "OSE","ODI","ODD","ODS","FISCAL","EMPREIT","AVISAR","PROJETISTA",
        "RESPONSÁVEL","RESPONSAVEL","DATA:","R$","PROCESSO","OBRASINCLU",
        "MMTEC","1 X","DESNÍVEL","NUMERO","NUM ","HAGAP","COPEL",
        "NO VALOR","DA OBRA","DE OBRA",
    ]
    for b in BLOQUEIOS:
        if b in v:
            return ""
    if re.search(r'\d', v):
        return ""
    if any(c in v for c in [":", ";", "/", "\\", "|", "="]):
        return ""
    if len(v) > 40:
        return ""
    letras = re.sub(r'[^A-ZÀ-Ÿ]', '', v)
    if len(letras) < 3:
        return ""
    palavras = v.split()
    if len(palavras) > 5:
        return ""
    for p in palavras:
        if len(p) <= 1:
            return ""
        # Palavra única muito longa = texto concatenado (ex: TRIFÁSICOPARAATENDER)
        if len(palavras) == 1 and len(p) > 15:
            return ""
        if len(p) > 25:
            return ""
    # Rejeita sobrenomes de pessoa
    for p in palavras:
        if p in _SOBRENOMES_PESSOA:
            return ""
    # Rejeita palavras de descrição
    for p in palavras:
        if p in _DESCRICAO_PALAVRAS:
            return ""
    # Rejeita se começa com preposição/artigo que indica descrição
    INICIO_DESCRICAO = {'NO','NA','NOS','NAS','DE','DA','DO','DAS','DOS','EM',
                        'POR','PARA','COM','SEM','SOBRE','AO','AOS',
                        'DEASCHORK','DASILVA','DAOBRA'}
    if palavras[0] in INICIO_DESCRICAO:
        return ""
    return v

def validar_us(val: str) -> str:
    if not val:
        return ""
    v = val.strip()
    v = re.sub(r'[R\$\s]', '', v)
    if re.search(r'[A-Za-z;:]', v):
        return ""
    m = re.search(r'(\d{1,3}(?:\.\d{3})*,\d+|\d+,\d+|\d+\.\d+|\d+)', v)
    if not m:
        return ""
    num_str = m.group(1)
    ns = num_str
    if ',' in ns and '.' in ns:
        ns = ns.replace('.', '').replace(',', '.')
    elif ',' in ns:
        ns = ns.replace(',', '.')
    try:
        num = float(ns)
    except ValueError:
        return ""
    if num <= 0 or num >= 1_000_000:
        return ""
    return num_str


def validar_prazo(val: str) -> str:
    if not val:
        return ""
    m = re.search(r'\b(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})\b', str(val))
    return m.group(1) if m else ""


# ─────────────────────────────────────────────────────────────────────
#  BUSCA EM TABELAS E TEXTO
# ─────────────────────────────────────────────────────────────────────

def buscar_tabela(tabelas, *chaves) -> str:
    for tbl in tabelas:
        for i, row in enumerate(tbl):
            row_c = [str(c).strip() if c else "" for c in row]
            for j, cel in enumerate(row_c):
                for chave in chaves:
                    if chave.lower() in cel.lower():
                        for delta in [1, 2]:
                            if j+delta < len(row_c):
                                v = row_c[j+delta].strip()
                                if v and not e_lixo(v):
                                    return v
                        if i+1 < len(tbl) and j < len(tbl[i+1]):
                            v = str(tbl[i+1][j] or "").strip()
                            if v and not e_lixo(v):
                                return v
    return ""

def buscar_texto(texto, *chaves) -> str:
    for chave in chaves:
        m = re.search(rf"(?i)\b{re.escape(chave)}\s*[:\-–]?\s*([^\n\r]{{1,100}})", texto)
        if m:
            v = m.group(1).strip().rstrip(".,;:")
            if v and not e_lixo(v) and len(v) < 120:
                return v
    return ""

def buscar(tabelas, texto, *chaves) -> str:
    return buscar_tabela(tabelas, *chaves) or buscar_texto(texto, *chaves)


# ─────────────────────────────────────────────────────────────────────
#  EXTRAÇÃO DE CAMPOS
# ─────────────────────────────────────────────────────────────────────

def extrair_ae(texto, tabelas):
    m = re.search(r'NÚMERO.{0,120}\n.{0,120}\n\s*(\d{5,6})\b', texto, re.DOTALL)
    if m:
        return m.group(1)
    m = re.search(r'NÚMERO.{0,200}?(\b0\d{4,5}\b)', texto, re.DOTALL)
    if m:
        return m.group(1)
    return buscar(tabelas, texto, "Nº AE","N° AE","AE nº","AE n°","Autorização de Execução","Cód. AE","Num AE")

def extrair_local(texto, tabelas):
    # PRIORIDADE 0: campo MUNICÍPIODAOBRA do BMD COPEL
    # Formato: "PR4100707-ALTOPIQUIRI" (código localidade seguido de cidade)
    m_mun = re.search(
        r'MUNIC[IÍ]PIODAOBRA[^\n]{0,40}\n[^\n]{0,30}?PR\d{7}-([A-ZÀ-ÖØ-Þ]+(?:\s+[A-ZÀ-ÖØ-Þ]+)*)',
        texto, re.IGNORECASE
    )
    if not m_mun:
        # Pode estar na mesma linha ou com formato diferente
        m_mun = re.search(r'PR\d{7}-([A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ]{2,25})(?=\s)', texto)
    if m_mun:
        cidade = m_mun.group(1).strip()
        v = validar_local(cidade)
        if v:
            return v

    # PRIORIDADE 1: campo CIDADE/ESTADO das AEs COPEL
    # Formato: linha com cabeçalho "CIDADE/ESTADO" seguida de "EMPREITEIRO UMUARAMA/PR 100-LÍDER"
    m_cidade = re.search(
        r'CIDADE/ESTADO[^\n]{0,60}\n[^\n]{0,80}?\s([A-ZÀÁÂÃÉÊÍÓÔÕÚÇ][A-ZÀÁÂÃÉÊÍÓÔÕÚÇ\s\-]{2,29})\s*/[A-Z]{2}\b',
        texto, re.IGNORECASE
    )
    if not m_cidade:
        # Formato sem quebra de linha: "HAGAPINSTALACOESELETRICASLT UMUARAMA/PR"
        m_cidade = re.search(
            r'HAGAP[^\n]{0,80}?\s([A-ZÀÁÂÃÉÊÍÓÔÕÚÇ][A-ZÀÁÂÃÉÊÍÓÔÕÚÇ\s\-]{2,29})\s*/[A-Z]{2}\b',
            texto.upper()
        )
    if m_cidade:
        cidade = m_cidade.group(1).strip()
        # Remove sufixo de UF se grudar: "UMUARAMA" de "UMUARAMA/PR"
        cidade = re.sub(r'\s*-\s*[A-Z]{2}\s*$', '', cidade).strip()
        v = validar_local(cidade)
        if v:
            return v

    for linha in texto.splitlines():
        ln = linha.strip().upper()
        if 'LOCAL:' in ln and 'OSE:' in ln:
            m = re.search(r'LOCAL:\s*([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý \-]*?)\s*OSE\s*:', ln)
            if m:
                cidade = m.group(1).strip()
                if cidade and 2 <= len(cidade) <= 40 and not re.search(r'\d', cidade):
                    return cidade
    for tbl in tabelas:
        for row in tbl:
            row_c = [str(c).strip() if c else "" for c in row]
            for i, cel in enumerate(row_c):
                if re.search(r'MUNIC[IÍ]PIO', cel.upper()):
                    for delta in [1, 2]:
                        if i + delta < len(row_c):
                            v = row_c[i + delta].strip().upper()
                            v = re.sub(r'\s*-?\s*[A-Z]{2}\s*$', '', v).strip()
                            v = re.sub(r'\s+', ' ', v)
                            if v and len(v) >= 3 and not re.search(r'\d', v) and len(v.split()) <= 4:
                                return v
    m = re.search(r'MUNIC[IÍ]PIO\s*[:\-–,]?\s*([A-ZÀÁÂÃÉÊÍÓÔÕÚÇ][A-ZÀÁÂÃÉÊÍÓÔÕÚÇ \-]{1,38})',
                  texto.upper())
    if m:
        v = m.group(1).strip()
        partes = []
        for p in v.split():
            if re.search(r'\d', p):
                break
            partes.append(p)
        v = ' '.join(partes).strip()
        v = re.sub(r'\s*-?\s*\b[A-Z]{2}\b\s*$', '', v).strip()
        if v and 3 <= len(v) <= 40 and not re.search(r'\d', v) and len(v.split()) <= 5:
            return v
    cidades_encontradas = []
    linhas = texto.splitlines()
    for linha in linhas:
        v = linha.strip()
        if not v:
            continue
        v_upper = v.upper()
        bloqueios = [
            "OSE","ODI","ODD","ODS","FISCAL","EMPREIT","AVISAR","PROJETISTA",
            "RESPONS","DATA:","R$","PROCESSO","OBRAS","MMTEC","1 X","04",":",";"," /",
        ]
        if any(b in v_upper for b in bloqueios):
            continue
        if re.search(r'\d', v_upper):
            continue
        v_upper = re.sub(r'^RDU\s+', '', v_upper)
        if not re.fullmatch(r'[A-ZÀ-Ÿ\s\-]+', v_upper):
            continue
        if len(v_upper) < 3 or len(v_upper) > 35:
            continue
        if len(v_upper.split()) > 4:
            continue
        cidades_encontradas.append(v_upper.strip())
    cidades_unicas = []
    for c in cidades_encontradas:
        if c not in cidades_unicas:
            cidades_unicas.append(c)
    if cidades_unicas:
        return cidades_unicas[-1]
    return ""

def extrair_us(texto, tabelas):
    CHAVES = ["Qtd US","Qtde US","Quantidade US","Total US","US previsto","US contratado",
              "Und. Serviço","Unidades de Serviço","US Total","QTD.US","Qtde. US"]
    v = buscar(tabelas, texto, *CHAVES)
    r = validar_us(v)
    if r:
        return r
    m = re.search(r'TOTAL:\s*(\d{1,3}(?:\.\d{3})*,\d{3})\b', texto)
    if m:
        r = validar_us(m.group(1))
        if r:
            return r
    m = re.search(r'QTD\.US[^\n]*\n[^\n]*?\b(\d{1,3}(?:\.\d{3})*,\d{3})\b', texto)
    if m:
        r = validar_us(m.group(1))
        if r:
            return r
    v2 = buscar(tabelas, texto, "US","U.S.")
    return validar_us(v2)

def extrair_prazo(texto, tabelas):
    CHAVES = ["Prazo","Vencimento","Data Limite","Data de Conclusão","Data Prevista",
              "Data Fim","Término","Termino","Data de Término","Prazo Final",
              "Data de Entrega","Prazo de Execução","Data Prazo"]
    RE_DATA = r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})'
    m = re.search(r'\bPRAZO\b[^\n]*\n[^\n]*?' + RE_DATA, texto)
    if m:
        return m.group(1)
    for tbl in tabelas:
        for i, row in enumerate(tbl):
            row_c = [str(c).strip() if c else "" for c in row]
            for j, cel in enumerate(row_c):
                for chave in CHAVES:
                    if chave.lower() in cel.lower():
                        for delta in [1, 2]:
                            if j+delta < len(row_c):
                                d = validar_prazo(row_c[j+delta])
                                if d:
                                    return d
                        if i+1 < len(tbl) and j < len(tbl[i+1]):
                            d = validar_prazo(str(tbl[i+1][j] or ""))
                            if d:
                                return d
    for chave in CHAVES:
        m = re.search(rf"(?i)\b{re.escape(chave)}\b.{{0,200}}?" + RE_DATA, texto, re.DOTALL)
        if m:
            return m.group(1)
    return ""

def extrair_custo(texto, tabelas):
    RE_VALOR = re.compile(r'R\$\s*[\d\.\s]+,\d{2}')
    # FIX: nas AEs COPEL, a linha TOTAL contém o custo total real
    # Formato: "TOTAL: 4.244,330 R$294.803,26" ou "TOTAL: 1.430,151 R$87.239,21"
    # DEVE ter precedência sobre qualquer outro valor
    m_total = re.search(
        r'\bTOTAL:\s*[\d\.,]+\s*(R\$[\d\.]+,\d{2})',
        texto, re.IGNORECASE
    )
    if m_total:
        return m_total.group(1).strip()
    # Fallback: Total Geral em tabelas
    v = buscar(tabelas, texto, "Total Geral","Valor Total","Valor Contrato","Preço Total","Custo","Valor da Obra")
    if v:
        m = RE_VALOR.search(v)
        if m:
            return m.group(0).strip()
    # Fallback final: primeiro R$ do texto (apenas se não houver linha de item)
    # Evita pegar valor de linha de item (que tem formato: data US R$valor projetonum)
    # Linha de item: "04/08/2026 1.430,151 R$87.239,21 1520606"
    # Linha que NÃO é item: não termina com número de projeto
    for linha in texto.splitlines():
        m_val = RE_VALOR.search(linha)
        if not m_val:
            continue
        # Pula linhas que são itens da tabela (terminam com número de projeto)
        if re.search(r'R\$[\d\.]+,\d{2}\s+\d{6,}', linha):
            continue
        return m_val.group(0).strip()
    return ""

def extrair_data_medicao(texto, tabelas):
    """Extrai a data de medição — suporte ao formato COPEL BMD/FFO."""

    def _norm(d):
        return d.replace('.', '/').replace('-', '/') if d else ''

    # 1. DT.MEDIÇÃO (COPEL BMD): header numa linha, valor na seguinte
    m = re.search(
        r'DT\.MEDI[^\n]{0,40}\n[^\n]{0,120}?(\d{2}[./]\d{2}[./]\d{4})',
        texto, re.IGNORECASE
    )
    if m:
        d = validar_prazo(_norm(m.group(1)))
        if d: return d

    # 1b. DT.MEDIÇÃO na mesma linha (sem quebra)
    m = re.search(r'DT\.MEDI[^\d]{0,30}(\d{2}[./]\d{2}[./]\d{4})', texto, re.IGNORECASE)
    if m:
        d = validar_prazo(_norm(m.group(1)))
        if d: return d

    # 2. DATATERMINO (FFO COPEL): segunda data na linha de valores
    m = re.search(
        r'DATATERM[^\n]{0,40}\n[^\n]{0,120}?\d{2}[./]\d{2}[./]\d{4}\s+(\d{2}[./]\d{2}[./]\d{4})',
        texto, re.IGNORECASE
    )
    if m:
        d = validar_prazo(_norm(m.group(1)))
        if d: return d
    m = re.search(r'DATATERM[^\d]{0,20}(\d{2}[./]\d{2}[./]\d{4})', texto, re.IGNORECASE)
    if m:
        d = validar_prazo(_norm(m.group(1)))
        if d: return d

    # 3. Campos genéricos via buscar()
    v = buscar(tabelas, texto,
               "Data de Medição", "Data Medição", "Competência",
               "Período", "Data Emissão", "Emissão", "Data Limite",
               "DT.MEDIÇÃO", "DT MEDIÇÃO", "DATATERMINO")
    d = validar_prazo(v)
    return d


def extrair_rateio_bmd(texto: str) -> str:
    """
    Extrai o valor do RATEIO HAGAP de um BMD/FFO.
    Suporta texto com ou sem espaços entre palavras (PDF pode juntar tokens).
    Padrão com espaços: RATEIO DO CUSTO TOTAL : HAGAP ... : R$ 2.410,16
    Padrão sem espaços: RATEIODOCUSTOTOTAL...HAGAP...R$457,30
    """
    # Padrão flexível: RATEIO + qualquer coisa + HAGAP + qualquer coisa + R$ valor
    m = re.search(
        r'RATEIO.{0,60}CUSTO.{0,60}TOTAL.{0,200}HAGAP.{0,300}R\$\s*([\d\.]+,\d{2})',
        texto, re.IGNORECASE | re.DOTALL
    )
    if m:
        return m.group(1).strip()
    # Fallback linha a linha — procura linha com RATEIO + HAGAP
    for linha in texto.splitlines():
        lu = linha.upper().replace(' ', '')
        if 'RATEIO' in lu and 'HAGAP' in lu:
            mv = re.search(r'R\$\s*([\d\.]+,\d{2})', linha)
            if mv:
                return mv.group(1).strip()
    # Fallback: linha com RATEIO + CUSTO
    for linha in texto.splitlines():
        lu = linha.upper().replace(' ', '')
        if 'RATEIO' in lu and 'CUSTO' in lu:
            mv = re.search(r'R\$\s*([\d\.]+,\d{2})', linha)
            if mv:
                return mv.group(1).strip()
    # Fallback genérico: qualquer linha com RATEIO
    for linha in texto.splitlines():
        if 'RATEIO' in linha.upper():
            mv = re.search(r'R\$\s*([\d\.]+,\d{2})', linha)
            if mv:
                return mv.group(1).strip()
    return ""


# ─────────────────────────────────────────────────────────────────────
#  TIPO DE MEDIÇÃO
# ─────────────────────────────────────────────────────────────────────

_FINAL   = ["medição final","medicao final","ffo","final de obra","encerramento","aceite final"]
_PARCIAL = ["medição parcial","medicao parcial","bmd","boletim de medição","parcial","avanço físico"]

def identificar_tipo_medicao(texto, nome):
    t = texto.lower(); n = nome.lower()
    sf = sum(1 for p in _FINAL   if p in t)
    sp = sum(1 for p in _PARCIAL if p in t)
    if sf > sp: return "ffo"
    if sp > sf: return "bmd"
    if any(p in n for p in ["ffo","final"]): return "ffo"
    return "bmd"


# ─────────────────────────────────────────────────────────────────────
#  UTILIDADES
# ─────────────────────────────────────────────────────────────────────

def parsear_data(s):
    if not s: return None
    formatos = ["%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%d/%m/%y","%d.%m.%Y","%d.%m.%y"]
    tokens = re.findall(r'\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}', str(s))
    for tok in tokens:
        for fmt in formatos:
            try: return datetime.strptime(tok, fmt).date()
            except ValueError: pass
    return None

def alerta_prazo(prazo_str):
    d = parsear_data(prazo_str)
    if not d: return ""
    diff = (d - date.today()).days
    if diff < 0:             return "vencido"
    if diff <= DIAS_CRITICO: return "critico"
    if diff <= DIAS_ALERTA:  return "alerta"
    return "ok"

def to_float(s):
    if not s: return 0.0
    s2 = re.sub(r'[^\d,\.]', '', str(s))
    if ',' in s2 and '.' in s2:
        s2 = s2.replace('.','').replace(',','.')
    elif ',' in s2:
        s2 = s2.replace(',','.')
    try: return float(s2)
    except ValueError: return 0.0


# ─────────────────────────────────────────────────────────────────────
#  MERGE E DB
# ─────────────────────────────────────────────────────────────────────

_CAMPOS_PADRAO = {
    "projeto":"","ae":"","us":"","us_num":0.0,"prazo":"","prazo_alerta":"",
    "local":"","custo":"","status":"Pendente","executado":False,
    "arquivo_ae":"","pasta_projeto":"","pasta_projeto_url":"",
    "pdfs_projeto":[],"bmds":[],"ffos":[],
    "medicao1":"","medicao2":"","final":"","origem":[],"avisos":[],
}

def carregar_db(caminho):
    if not os.path.exists(caminho): return []
    with open(caminho, encoding="utf-8") as f:
        conteudo = f.read().strip()
    if not conteudo: return []
    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError:
        try:
            dados, _ = json.JSONDecoder().raw_decode(conteudo)
        except Exception:
            return []
    if not isinstance(dados, list): return []
    for r in dados:
        for campo, padrao in _CAMPOS_PADRAO.items():
            if campo not in r:
                r[campo] = padrao
    return dados

def registro_vazio(projeto):
    r = dict(_CAMPOS_PADRAO)
    r["projeto"] = projeto
    r["pdfs_projeto"] = []
    r["bmds"] = []
    r["ffos"] = []
    r["origem"] = []
    r["avisos"] = []
    return r

def mesclar(atual, novo):
    return atual if atual else (novo or "")

def mesclar_medicoes(atual, novos):
    existentes = {i.get("arquivo","") for i in atual}
    for item in novos:
        if item.get("arquivo","") not in existentes:
            atual.append(item)
    return atual

def deduplicate_registros(registros: list) -> list:
    grupos: dict[str, list] = {}
    sem_numero = []
    for r in registros:
        proj = r.get("projeto","")
        k = norm_num(proj)
        if not numero_valido(proj) and not re.match(r'[A-Z]{2,6}\d{6,12}', proj.upper()):
            sem_numero.append(proj)
            continue
        if k not in grupos:
            grupos[k] = []
        grupos[k].append(r)
    if sem_numero:
        print(f"\n  ⚠ Rejeitados (sem nº de projeto válido): {', '.join(sem_numero[:20])}")
    resultado = []
    for k, grupo in grupos.items():
        if len(grupo) == 1:
            r = grupo[0]
        else:
            grupo_ord = sorted(grupo, key=lambda x: len(x.get("projeto","")))
            base = copy.deepcopy(grupo_ord[0])
            print(f"  🔗 Deduplicando {k}: {[g['projeto'] for g in grupo]} → '{base['projeto']}'")
            for outro in grupo_ord[1:]:
                for campo in ("ae","us","prazo","local","custo","arquivo_ae","pasta_projeto","medicao1","medicao2","final"):
                    if not base.get(campo) and outro.get(campo):
                        v = outro[campo]
                        if campo == "us":     v = validar_us(v)
                        if campo == "local":  v = validar_local(v)
                        if campo in ("prazo","medicao1","medicao2","final"): v = validar_prazo(v) or v
                        base[campo] = v
                base["bmds"]         = mesclar_medicoes(base.get("bmds",[]), outro.get("bmds",[]))
                base["ffos"]         = mesclar_medicoes(base.get("ffos",[]), outro.get("ffos",[]))
                base["pdfs_projeto"] = list(dict.fromkeys(base.get("pdfs_projeto",[]) + outro.get("pdfs_projeto",[])))
                for o in outro.get("origem",[]):
                    if o not in base.get("origem",[]): base.setdefault("origem",[]).append(o)
            r = base
        num_limpo = extrair_numero_do_texto(r["projeto"])
        if num_limpo:
            r["projeto"] = num_limpo
        resultado.append(r)
    return resultado


# ─────────────────────────────────────────────────────────────────────
#  PROCESSAMENTO DAS PASTAS
# ─────────────────────────────────────────────────────────────────────

def _extrair_todos_projetos(texto: str, tabelas: list) -> list:
    """
    Extrai TODOS os números de projeto encontrados num PDF de AE.
    Útil para AEs que cobrem múltiplos projetos numa só página.
    """
    encontrados = []
    vistos = set()

    # 1. Busca em tabelas (campo PROJ ou similar)
    for tbl in tabelas:
        for row in tbl:
            row_c = [str(c).strip() if c else "" for c in row]
            for j, cel in enumerate(row_c):
                if re.search(r'\bPROJ\.?\b', cel, re.IGNORECASE):
                    for delta in [1, 2]:
                        if j + delta < len(row_c):
                            m = _RE_NUM_PROJ.search(row_c[j + delta])
                            if m:
                                num = m.group(1)
                                k = norm_num(num)
                                if k not in vistos:
                                    vistos.add(k)
                                    encontrados.append(num)

    # 2. Linhas com padrão: data US R$valor PROJ_NUM (linha de item de AE)
    for linha in texto.splitlines():
        m = re.search(
            r'Investimento[A-Z]+\s+\d{2}/\d{2}/\d{4}\s+[\d\.,]+\s+R\$[\d\.,]+\s+(1[5-9]\d{5}[A-Z]?)\b',
            linha
        )
        if m:
            num = m.group(1)
            k = norm_num(num)
            if k not in vistos:
                vistos.add(k)
                encontrados.append(num)

    # 3. findall genérico removendo PEP codes
    texto_sem_pep = re.sub(r'[A-Z]-\d{2}-(\d{7})', '', texto)
    for m in _RE_NUM_PROJ.finditer(texto_sem_pep):
        num = m.group(1)
        k = norm_num(num)
        if k not in vistos:
            vistos.add(k)
            encontrados.append(num)

    # 4. Se nenhum achado, tenta numero_do_pdf original
    if not encontrados:
        single = numero_do_pdf(texto, tabelas)
        if single and numero_valido(single):
            encontrados.append(single)

    return encontrados


def processar_aes(pasta):
    resultado = {}
    if not os.path.isdir(pasta):
        print(f"  [AVISO] {pasta} não encontrada"); return resultado
    arquivos = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
    print(f"  {len(arquivos)} AE(s)")
    for arq in arquivos:
        caminho = os.path.join(pasta, arq)
        print(f"\n  → AE: {arq}")
        texto   = extrair_texto(caminho)
        tabelas = extrair_tabelas(caminho)
        ae      = extrair_ae(texto, tabelas)
        us      = extrair_us(texto, tabelas)
        prazo   = extrair_prazo(texto, tabelas)
        custo   = extrair_custo(texto, tabelas)

        # Extrai TODOS os projetos encontrados nesta AE
        projetos_encontrados = _extrair_todos_projetos(texto, tabelas)

        # Fallback: número no nome do arquivo
        if not projetos_encontrados:
            num_arq = extrair_numero_do_texto(arq)
            if num_arq and numero_valido(num_arq):
                projetos_encontrados = [num_arq]

        if not projetos_encontrados:
            print(f"    ⚠ Nenhum projeto encontrado — ignorado")
            continue

        print(f"    → Projetos encontrados: {projetos_encontrados}")
        print(f"    → AE:{ae or '—'} | US:{us or '—'} | Prazo:{prazo or '—'} | Custo:{custo or '—'}")

        for projeto in projetos_encontrados:
            if not numero_valido(projeto):
                continue
            k = norm_num(projeto)
            if k not in resultado:
                resultado[k] = {
                    "projeto": projeto, "ae": ae, "us": us, "prazo": prazo,
                    "local": "", "custo": custo, "arquivo_ae": arq,
                    "aes_lista": [{"arquivo": arq, "ae": ae}]
                }
            else:
                # AE adicional para mesmo projeto
                existing = resultado[k]
                existing.setdefault("aes_lista", [])
                if arq not in [a["arquivo"] for a in existing["aes_lista"]]:
                    existing["aes_lista"].append({"arquivo": arq, "ae": ae})
                print(f"    ↳ AE adicional para {projeto}: {arq}")
    return resultado


def processar_pasta_projeto(caminho_pasta):
    info = {"local":"","us":"","prazo":"","custo":"","pdfs_projeto":[]}
    if not os.path.isdir(caminho_pasta): return info
    pdfs = sorted(f for f in os.listdir(caminho_pasta) if f.lower().endswith(".pdf"))
    info["pdfs_projeto"] = pdfs
    for nome in pdfs:
        try:
            texto   = extrair_texto(os.path.join(caminho_pasta, nome))
            tabelas = extrair_tabelas(os.path.join(caminho_pasta, nome))
        except Exception as e:
            print(f"      ⚠ Erro {nome}: {e}"); continue
        if not info["local"]:  info["local"]  = extrair_local(texto, tabelas)
        if not info["us"]:     info["us"]     = extrair_us(texto, tabelas)
        if not info["prazo"]:  info["prazo"]  = extrair_prazo(texto, tabelas)
        if not info["custo"]:  info["custo"]  = extrair_custo(texto, tabelas)
        if all(info[k] for k in ("local","us","prazo","custo")): break
    return info


def buscar_local_em_outras_pastas(projeto: str) -> str:
    num = projeto[:7]
    if os.path.isdir(PASTA_PROJETOS):
        for nome_pasta in sorted(os.listdir(PASTA_PROJETOS)):
            if num not in nome_pasta:
                continue
            pasta_path = os.path.join(PASTA_PROJETOS, nome_pasta)
            if not os.path.isdir(pasta_path):
                continue
            for arq in sorted(os.listdir(pasta_path)):
                if not arq.lower().endswith(".pdf"):
                    continue
                try:
                    texto   = extrair_texto(os.path.join(pasta_path, arq))
                    tabelas = extrair_tabelas(os.path.join(pasta_path, arq))
                    local = extrair_local(texto, tabelas)
                    if local:
                        print(f"      → Local encontrado em '{nome_pasta}/{arq}': {local}")
                        return local
                except Exception:
                    continue
    for pasta in [PASTA_AES, PASTA_MEDICOES]:
        if not os.path.isdir(pasta):
            continue
        for arq in sorted(os.listdir(pasta)):
            if not arq.lower().endswith(".pdf"):
                continue
            if num not in arq and projeto not in arq:
                continue
            try:
                texto   = extrair_texto(os.path.join(pasta, arq))
                tabelas = extrair_tabelas(os.path.join(pasta, arq))
                local = extrair_local(texto, tabelas)
                if local:
                    print(f"      → Local encontrado em '{arq}': {local}")
                    return local
            except Exception:
                continue
    return ""


def processar_medicoes(pasta):
    resultado = {}
    if not os.path.isdir(pasta):
        print(f"  [AVISO] {pasta} não encontrada"); return resultado
    arquivos = sorted(f for f in os.listdir(pasta) if f.lower().endswith(".pdf"))
    print(f"  {len(arquivos)} medição(ões)")
    for arq in arquivos:
        caminho = os.path.join(pasta, arq)
        print(f"\n  → Medição: {arq}")
        texto   = extrair_texto(caminho)
        tabelas = extrair_tabelas(caminho)
        projeto = numero_do_pdf(texto, tabelas)
        if not projeto:
            projeto = extrair_numero_do_texto(arq)
        if not projeto or not numero_valido(projeto):
            # ─── ORFÃO: BMD/FFO sem projeto associado ───────────────
            nome_sem_ext = os.path.splitext(arq)[0]
            projeto_orfao = f"ORFAO_{nome_sem_ext[:30]}"
            print(f"    ⚠ Projeto não encontrado → registrando como orfão: {projeto_orfao}")
            projeto = projeto_orfao
        tipo  = identificar_tipo_medicao(texto, arq)
        us    = extrair_us(texto, tabelas)
        data  = extrair_data_medicao(texto, tabelas)
        custo = extrair_custo(texto, tabelas)
        print(f"    Projeto:{projeto} | Tipo:{tipo.upper()} | US:{us or '—'} | Data:{data or '—'}")
        rateio       = extrair_rateio_bmd(texto)
        rateio_float = to_float(rateio)
        if rateio:
            print(f"    Rateio HAGAP: R$ {rateio}")
        entrada = {"arquivo":arq,"us":us,"data":data,"valor":custo,
                   "rateio":rateio,"rateio_float":rateio_float}
        k = norm_num(projeto)
        if k not in resultado:
            resultado[k] = {"bmds":[],"ffos":[],"_projeto_nome":projeto}
        chave = tipo+"s"
        if arq not in [b["arquivo"] for b in resultado[k][chave]]:
            resultado[k][chave].append(entrada)
    return resultado


# ─────────────────────────────────────────────────────────────────────
#  BUILD PRINCIPAL
# ─────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────
#  GOOGLE DRIVE — busca URL de pasta pelo nome
# ─────────────────────────────────────────────────────────────────────

def _get_drive_service_local():
    """Autentica com a Service Account e retorna o serviço Drive."""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        import base64, json as _json
        b64 = _GOOGLE_SA_B64.strip()
        b64 += '=' * (-len(b64) % 4)
        info = _json.loads(base64.b64decode(b64).decode('utf-8'))
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        svc = build('drive', 'v3', credentials=creds)
        print(f"[Drive] ✅ Conectado como {info.get('client_email','')}")
        return svc
    except Exception as e:
        print(f"[Drive] ⚠️ Não foi possível autenticar: {e}")
        return None


_drive_cache: dict = {}   # cache nome_pasta → url

def buscar_url_pasta_drive(service, root_id: str, nome_pasta: str) -> str:
    """Busca URL webView de uma sub-pasta no Drive pelo nome."""
    if not service:
        return ""
    if nome_pasta in _drive_cache:
        return _drive_cache[nome_pasta]
    try:
        nome_esc = nome_pasta.replace("'", "\\'")
        q = (f"'{root_id}' in parents "
             f"and name='{nome_esc}' "
             f"and mimeType='application/vnd.google-apps.folder' "
             f"and trashed=false")
        res = service.files().list(
            q=q, fields="files(id,webViewLink,name)", pageSize=5
        ).execute()
        files = res.get('files', [])
        if not files:
            # tenta busca parcial pelo número do projeto (primeiros 7 dígitos)
            num = re.sub(r'\D', '', nome_pasta)[:7]
            if num:
                q2 = (f"'{root_id}' in parents "
                      f"and name contains '{num}' "
                      f"and mimeType='application/vnd.google-apps.folder' "
                      f"and trashed=false")
                res2 = service.files().list(
                    q=q2, fields="files(id,webViewLink,name)", pageSize=5
                ).execute()
                files = res2.get('files', [])
        url = files[0].get('webViewLink', '') if files else ''
        _drive_cache[nome_pasta] = url
        if url:
            print(f"    [Drive] 📂 Pasta encontrada: {nome_pasta} → {url}")
        return url
    except Exception as e:
        print(f"    [Drive] ⚠️ Erro ao buscar pasta '{nome_pasta}': {e}")
        return ''


def gerar_db():
    sep = "═"*64
    print(f"\n{sep}\n  HAGAP_WEB v2.4 — Gerador\n{sep}")

    print("\n[0/5] Carregando db.json existente...")
    db_existente = carregar_db(SAIDA_JSON)
    idx: dict[str, int] = {}
    for i, r in enumerate(db_existente):
        k = norm_num(r.get("projeto",""))
        if k: idx[k] = i
    print(f"  {len(db_existente)} registro(s) existente(s)")
    registros = [copy.deepcopy(r) for r in db_existente]

    def obter_ou_criar(proj):
        k = norm_num(proj)
        if k in idx: return registros[idx[k]]
        novo = registro_vazio(proj)
        registros.append(novo)
        idx[k] = len(registros)-1
        return novo

    print("\n[1/4] Processando AEs...")
    aes = processar_aes(PASTA_AES)
    for k_ae, ae in aes.items():
        r = obter_ou_criar(ae["projeto"])
        r["ae"]         = mesclar(r["ae"],         ae.get("ae",""))
        r["us"]         = mesclar(r["us"],         ae.get("us",""))
        r["prazo"]      = mesclar(r["prazo"],      ae.get("prazo",""))
        r["custo"]      = mesclar(r["custo"],      ae.get("custo",""))
        r["arquivo_ae"] = mesclar(r["arquivo_ae"], ae.get("arquivo_ae",""))
        # Mesclar lista de AEs (suporte a múltiplas AEs por projeto)
        r.setdefault("aes", [])
        for ae_item in ae.get("aes_lista", []):
            if ae_item["arquivo"] not in [a.get("arquivo") for a in r["aes"]]:
                r["aes"].append(ae_item)
        if "ae" not in r["origem"]: r["origem"].append("ae")

    # Tenta conectar ao Drive para buscar URLs das pastas
    print("\n[Drive] Conectando ao Google Drive...")
    drive_svc = _get_drive_service_local()

    print("\n[2/4] Escaneando pastas de projetos...")
    if not os.path.isdir(PASTA_PROJETOS):
        print(f"  [AVISO] {PASTA_PROJETOS} não encontrada"); pastas=[]
    else:
        pastas = sorted(d for d in os.listdir(PASTA_PROJETOS)
                        if os.path.isdir(os.path.join(PASTA_PROJETOS, d)))
    print(f"  {len(pastas)} pasta(s)")
    for nome_pasta in pastas:
        numero = numero_do_nome_pasta(nome_pasta)
        if not numero_valido(numero) and not re.search(r'\d{6,8}', numero):
            print(f"  ⚠ Pasta '{nome_pasta}' ignorada (número inválido)")
            continue
        print(f"\n  → '{nome_pasta}' → Projeto: '{numero}'")
        info = processar_pasta_projeto(os.path.join(PASTA_PROJETOS, nome_pasta))
        r = obter_ou_criar(numero)
        local_val = validar_local(info.get("local",""))
        if not local_val:
            local_val = buscar_local_em_outras_pastas(numero)
        r["local"]         = mesclar(r["local"],         local_val)
        r["us"]            = mesclar(r["us"],            info.get("us",""))
        r["prazo"]         = mesclar(r["prazo"],         info.get("prazo",""))
        r["custo"]         = mesclar(r["custo"],         info.get("custo",""))
        r["pasta_projeto"] = r["pasta_projeto"] or nome_pasta
        r["pdfs_projeto"]  = info.get("pdfs_projeto",[])
        # Busca URL da pasta no Drive
        if not r.get("pasta_projeto_url") and drive_svc:
            url_pasta = buscar_url_pasta_drive(drive_svc, DRIVE_FOLDER_ID, nome_pasta)
            if url_pasta:
                r["pasta_projeto_url"] = url_pasta
        if "pasta" not in r["origem"]: r["origem"].append("pasta")

    print("\n[3/4] Processando medições...")
    medicoes = processar_medicoes(PASTA_MEDICOES)
    for k_med, med in medicoes.items():
        r = obter_ou_criar(med.get("_projeto_nome", k_med))
        r["bmds"] = mesclar_medicoes(r.get("bmds",[]), med.get("bmds",[]))
        r["ffos"] = mesclar_medicoes(r.get("ffos",[]), med.get("ffos",[]))
        if "medicao" not in r["origem"]: r["origem"].append("medicao")
        bmds_ord = sorted(r["bmds"], key=lambda x: x.get("data","") or "")
        ffos_ord = sorted(r["ffos"], key=lambda x: x.get("data","") or "")
        if len(bmds_ord) >= 1 and not r.get("medicao1"):
            b = bmds_ord[0]; r["medicao1"] = b.get("data") or b.get("arquivo","")
        if len(bmds_ord) >= 2 and not r.get("medicao2"):
            b = bmds_ord[1]; r["medicao2"] = b.get("data") or b.get("arquivo","")
        if ffos_ord and not r.get("final"):
            f = ffos_ord[0]; r["final"] = f.get("data") or f.get("arquivo","")
            r["executado"] = True; r["status"] = "Final"

    print("\n[4/4] Deduplicando e validando...")
    registros = deduplicate_registros(registros)

    ae_map = {}
    for r in registros:
        r["us_num"]       = to_float(r.get("us",""))
        r["prazo_alerta"] = alerta_prazo(r.get("prazo",""))
        if r.get("ffos") or r.get("status") == "Final": r["executado"] = True
        ae = (r.get("ae") or "").strip()
        if ae: ae_map.setdefault(ae,[]).append(r["projeto"])
    for ae, projs in ae_map.items():
        if len(projs) > 1:
            msg = f"⚠ AE '{ae}' em múltiplos projetos: {', '.join(projs)}"
            for r in registros:
                if r.get("ae") == ae and msg not in r.get("avisos",[]): r.setdefault("avisos",[]).append(msg)

    print("\n"+"─"*64+"\n  RELATÓRIO\n"+"─"*64)
    def _lst(lbl, lst):
        if lst: print(f"  {lbl} ({len(lst)}): {', '.join(lst[:10])}{'...' if len(lst)>10 else ''}")
    _lst("Sem AE",     [r["projeto"] for r in registros if not r.get("ae")])
    _lst("Sem local",  [r["projeto"] for r in registros if not r.get("local")])
    _lst("Sem prazo",  [r["projeto"] for r in registros if not r.get("prazo")])
    _lst("🔴 Vencidos",[r["projeto"] for r in registros if r.get("prazo_alerta")=="vencido"])
    _lst("🟠 Crítico", [r["projeto"] for r in registros if r.get("prazo_alerta")=="critico"])
    _lst("🟡 Alerta",  [r["projeto"] for r in registros if r.get("prazo_alerta")=="alerta"])
    print(f"\n  Total final: {len(registros)} projetos")
    return registros


# ─────────────────────────────────────────────────────────────────────
#  SALVAR
# ─────────────────────────────────────────────────────────────────────

def salvar(registros):
    with open(SAIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)
    linhas = []
    for r in registros:
        row = {k:v for k,v in r.items() if k not in ("bmds","ffos","pdfs_projeto","avisos","origem")}
        row["qtd_bmds"] = len(r.get("bmds",[])); row["qtd_ffos"] = len(r.get("ffos",[])); row["qtd_pdfs"] = len(r.get("pdfs_projeto",[]))
        row["avisos"]   = " | ".join(r.get("avisos",[])); row["origem"] = ", ".join(r.get("origem",[]))
        linhas.append(row)
    df = pd.DataFrame(linhas)
    col_order = ["projeto","ae","local","us","us_num","prazo","prazo_alerta","custo","status","executado","medicao1","medicao2","final","qtd_bmds","qtd_ffos","qtd_pdfs","arquivo_ae","pasta_projeto","origem","avisos"]
    df = df.reindex(columns=[c for c in col_order if c in df.columns]+[c for c in df.columns if c not in col_order])
    df.to_excel(SAIDA_EXCEL, index=False)
    print(f"\n✅ db.json → {len(registros)} registros\n✅ resultado.xlsx salvo")
    print("\n→ git add db.json pdfs/ && git commit -m 'atualiza' && git push")


# ─────────────────────────────────────────────────────────────────────
#  SINCRONIZAÇÃO COM O RENDER (upload de PDFs + db.json)
# ─────────────────────────────────────────────────────────────────────

def _upload_arquivo(caminho, pasta_key, projeto_nome="", caminho_rel="", skip_drive=True):
    """Faz upload de um único PDF para o Render via /api/upload_pdf."""
    nome = os.path.basename(caminho)
    try:
        with open(caminho, 'rb') as f:
            data = {'pasta': pasta_key}
            if projeto_nome:   data['projeto']          = projeto_nome
            if caminho_rel:    data['caminho_relativo'] = caminho_rel
            if skip_drive:     data['skip_drive']       = '1'
            r = _requests.post(
                f"{RENDER_URL}/api/upload_pdf",
                files={'arquivo': (nome, f, 'application/pdf')},
                data=data,
                timeout=120
            )
        if r.status_code == 200:
            return True
        print(f"    ⚠️ {nome}: HTTP {r.status_code}")
        return False
    except Exception as e:
        print(f"    ❌ {nome}: {e}")
        return False


def sincronizar_render(registros):
    """Envia db.json + todos os PDFs das 3 pastas para o servidor Render."""
    if not HAS_REQUESTS:
        print("\n[Render] ⚠️  Módulo 'requests' não instalado. Execute: pip install requests")
        print("         Sem sincronização automática.")
        return

    sep = "─" * 64
    print(f"\n{sep}\n  SINCRONIZAÇÃO COM O RENDER ({RENDER_URL})\n{sep}")

    # ── 1. Sincronizar db.json ────────────────────────────────────────
    print("\n[1/4] Enviando db.json...")
    try:
        r = _requests.post(
            f"{RENDER_URL}/api/migrar_db?senha={SENHA_MIGRACAO}",
            json={"senha": SENHA_MIGRACAO, "dados": registros},
            timeout=60
        )
        if r.ok:
            print(f"  ✅ db.json OK ({len(registros)} registros)")
        else:
            print(f"  ❌ Erro ao enviar db.json: {r.status_code} — {r.text[:200]}")
            return
    except Exception as e:
        print(f"  ❌ Sem conexão com o Render: {e}")
        return

    # ── 2. AEs ───────────────────────────────────────────────────────
    print("\n[2/4] Enviando AEs...")
    if os.path.isdir(PASTA_AES):
        pdfs = sorted(f for f in os.listdir(PASTA_AES) if f.lower().endswith('.pdf'))
        print(f"  {len(pdfs)} arquivo(s)")
        ok = sum(_upload_arquivo(os.path.join(PASTA_AES, f), 'aes') for f in pdfs)
        print(f"  ✅ {ok}/{len(pdfs)} enviado(s)")
    else:
        print(f"  [AVISO] Pasta não encontrada: {PASTA_AES}")

    # ── 3. Medições ───────────────────────────────────────────────────
    print("\n[3/4] Enviando Medições (BMD/FFO)...")
    if os.path.isdir(PASTA_MEDICOES):
        pdfs = sorted(f for f in os.listdir(PASTA_MEDICOES) if f.lower().endswith('.pdf'))
        print(f"  {len(pdfs)} arquivo(s)")
        ok = sum(_upload_arquivo(os.path.join(PASTA_MEDICOES, f), 'medicoes') for f in pdfs)
        print(f"  ✅ {ok}/{len(pdfs)} enviado(s)")
    else:
        print(f"  [AVISO] Pasta não encontrada: {PASTA_MEDICOES}")

    # ── 4. Projetos (subpastas) ────────────────────────────────────────
    print("\n[4/4] Enviando Projetos (subpastas)...")
    if os.path.isdir(PASTA_PROJETOS):
        pastas = sorted(d for d in os.listdir(PASTA_PROJETOS)
                        if os.path.isdir(os.path.join(PASTA_PROJETOS, d)))
        # Conta total de PDFs antes de começar
        lista_projetos = []
        for nome_pasta in pastas:
            pasta_path = os.path.join(PASTA_PROJETOS, nome_pasta)
            for raiz, dirs, files in os.walk(pasta_path):
                dirs[:] = [d for d in dirs if not d.startswith('.')]
                for arq in sorted(files):
                    if arq.lower().endswith('.pdf'):
                        caminho_abs = os.path.join(raiz, arq)
                        caminho_rel = os.path.relpath(caminho_abs, PASTA_PROJETOS).replace('\\', '/')
                        lista_projetos.append((caminho_abs, nome_pasta, caminho_rel))
        total_arq = len(lista_projetos)
        print(f"  {total_arq} arquivo(s) em {len(pastas)} pasta(s)")
        total_ok = 0
        for i, (caminho_abs, nome_pasta, caminho_rel) in enumerate(lista_projetos, 1):
            arq = os.path.basename(caminho_abs)
            print(f"  [{i}/{total_arq}] {nome_pasta}/{arq}", end=" ... ", flush=True)
            ok = _upload_arquivo(caminho_abs, 'projetos',
                                 projeto_nome=nome_pasta,
                                 caminho_rel=caminho_rel)
            if ok:
                total_ok += 1
                print("✅")
            else:
                print("❌")
        print(f"  ✅ {total_ok}/{total_arq} enviado(s) em {len(pastas)} pasta(s)")
    else:
        print(f"  [AVISO] Pasta não encontrada: {PASTA_PROJETOS}")

    print(f"\n{'─'*64}")
    print("  ✅ Sincronização concluída! Acesse https://hagap.onrender.com")
    print(f"{'─'*64}\n")


if __name__ == "__main__":
    try:
        registros = gerar_db()
        salvar(registros)
        if SINCRONIZAR_RENDER:
            sincronizar_render(registros)
        else:
            print("\n→ Sincronização desativada (SINCRONIZAR_RENDER = False)")
    except Exception:
        print("\n❌ Erro fatal:"); traceback.print_exc()