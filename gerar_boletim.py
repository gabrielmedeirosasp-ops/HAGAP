#!/usr/bin/env python3
"""
BDO — Boletim Diário de Obras
Gera PDF idêntico à prévia do formulário HTML.
"""
import sys, json, os, time, base64, tempfile
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ── Paleta (mesma do HTML) ─────────────────────────────────────────
AZUL_ESC  = colors.HexColor("#061539")
AZUL      = colors.HexColor("#0a2557")
AZUL_MED  = colors.HexColor("#1565c0")
ROXO_ESC  = colors.HexColor("#4a0e8f")
ROXO      = colors.HexColor("#7b2ff7")
VERDE     = colors.HexColor("#1a7a3c")
VERMELHO  = colors.HexColor("#c0390b")
AMARELO_E = colors.HexColor("#c07000")
CINZA_BG  = colors.HexColor("#f0f4fb")
CINZA_BD  = colors.HexColor("#e8ecf5")
CINZA_TXT = colors.HexColor("#4a5a7a")
BRANCO    = colors.white
PRETO     = colors.HexColor("#0d1b3e")

W, H = A4   # 595 x 842 pts
LM = RM = 15*mm
CONT_W = W - LM - RM


# ── Helpers de estilo ──────────────────────────────────────────────
def st(name, **kw):
    base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=PRETO)
    return ParagraphStyle(name, **{**base, **kw})

STYLES = {
    "titulo"   : st("t1", fontName="Helvetica-Bold", fontSize=13, textColor=BRANCO),
    "subtit"   : st("t2", fontSize=8,  textColor=colors.HexColor("#b0bec5")),
    "num_bdo"  : st("t3", fontName="Helvetica-Bold", fontSize=20, textColor=BRANCO, alignment=TA_RIGHT),
    "data_r"   : st("t4", fontSize=8,  textColor=colors.HexColor("#b0bec5"), alignment=TA_RIGHT),
    "sec"      : st("s1", fontName="Helvetica-Bold", fontSize=8, textColor=BRANCO),
    "label"    : st("l1", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_TXT),
    "valor"    : st("v1", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO),
    "body"     : st("b1", fontSize=9, textColor=PRETO, leading=14),
    "li"       : st("li", fontSize=9, textColor=PRETO, leading=15, leftIndent=12),
    "obs"      : st("ob", fontSize=8, textColor=PRETO, leading=12),
    "rodape"   : st("ro", fontSize=7, textColor=CINZA_TXT, alignment=TA_CENTER),
    "sig_nome" : st("sn", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO, alignment=TA_CENTER),
    "sig_role" : st("sr", fontSize=8, textColor=CINZA_TXT, alignment=TA_CENTER),
    "status"   : st("ss", fontName="Helvetica-Bold", fontSize=10, textColor=BRANCO, alignment=TA_CENTER),
}
SP = lambda n=6: Spacer(1, n)


# ── Barra de seção colorida ────────────────────────────────────────
def barra(texto, cor=AZUL):
    t = Table([[Paragraph(texto, STYLES["sec"])]], colWidths=[CONT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), cor),
        ("ROWPADDING", (0,0), (-1,-1), 6),
    ]))
    return t


# ── Linha de campos (label / valor) ───────────────────────────────
def campos(*pares, widths=None):
    if not widths:
        u = CONT_W / len(pares) / 2
        widths = []
        for _ in pares:
            widths += [u*0.55, u*1.45]
    flat = []
    for lbl, val in pares:
        flat += [Paragraph(lbl, STYLES["label"]), Paragraph(str(val or "—"), STYLES["valor"])]
    t = Table([flat], colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), CINZA_BG),
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING", (0,0), (-1,-1), 7),
        ("VALIGN",     (0,0), (-1,-1), "TOP"),
    ]))
    return t


# ── Bloco de texto simples ─────────────────────────────────────────
def bloco_txt(texto):
    t = Table([[Paragraph(texto or "—", STYLES["body"])]], colWidths=[CONT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BRANCO),
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING", (0,0), (-1,-1), 8),
    ]))
    return t


# ── Lista de itens ─────────────────────────────────────────────────
def lista_itens(itens, cor_fundo=BRANCO):
    if not itens:
        data = [[Paragraph("—", STYLES["obs"])]]
    else:
        data = [[Paragraph(f"• {i}", STYLES["li"])] for i in itens]
    t = Table(data, colWidths=[CONT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,-1), cor_fundo),
        ("GRID",           (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING",     (0,0), (-1,-1), 5),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [cor_fundo, colors.HexColor("#f7f9ff")]),
    ]))
    return t


# ── Seção de grupo (poste ou geral) ───────────────────────────────
def grupo_poste(ps_label, srvs, mats, obss, is_geral=False):
    """Renderiza um bloco completo de poste: serviços + materiais + observações."""
    cor = AZUL if is_geral else ROXO_ESC
    elements = []
    elements.append(barra(ps_label, cor))

    if srvs:
        sub = Table([[Paragraph("🔧 SERVIÇOS", STYLES["label"])]], colWidths=[CONT_W])
        sub.setStyle(TableStyle([("ROWPADDING",(0,0),(-1,-1),4),("BACKGROUND",(0,0),(-1,-1),CINZA_BG),("GRID",(0,0),(-1,-1),.3,CINZA_BD)]))
        elements.append(sub)
        elements.append(lista_itens([s["nome"] for s in srvs]))

    if mats:
        sub = Table([[Paragraph("📦 MATERIAIS", STYLES["label"])]], colWidths=[CONT_W])
        sub.setStyle(TableStyle([("ROWPADDING",(0,0),(-1,-1),4),("BACKGROUND",(0,0),(-1,-1),CINZA_BG),("GRID",(0,0),(-1,-1),.3,CINZA_BD)]))
        elements.append(sub)
        elements.append(lista_itens([f"{m.get('descricao','')}{' — '+m.get('qtd') if m.get('qtd') else ''}" for m in mats]))

    if obss:
        sub = Table([[Paragraph("📝 OBSERVAÇÕES", STYLES["label"])]], colWidths=[CONT_W])
        sub.setStyle(TableStyle([("ROWPADDING",(0,0),(-1,-1),4),("BACKGROUND",(0,0),(-1,-1),CINZA_BG),("GRID",(0,0),(-1,-1),.3,CINZA_BD)]))
        elements.append(sub)
        elements.append(lista_itens([o.get("texto","") for o in obss]))

    return elements


# ── Cabeçalho do documento ─────────────────────────────────────────
def cabecalho(d, logo_path=None):
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    left = [
        Paragraph("BOLETIM DIÁRIO DE OBRAS", STYLES["titulo"]),
        Paragraph("Hagap Engenharia Elétrica  |  Registro de Campo", STYLES["subtit"]),
        Paragraph(f"Gerado em: {now}", STYLES["subtit"]),
    ]
    right = [
        Paragraph(f"Nº {d.get('num_projeto','---')}", STYLES["num_bdo"]),
        Paragraph(f"Data: {d.get('data_bdo','')}", STYLES["data_r"]),
    ]
    if logo_path and os.path.exists(logo_path):
        try:    logo = Image(logo_path, width=90, height=36)
        except: logo = Paragraph("HAGAP", ParagraphStyle("lx", fontName="Helvetica-Bold", fontSize=18, textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))
    else:
        logo = Paragraph("HAGAP", ParagraphStyle("lx2", fontName="Helvetica-Bold", fontSize=18, textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))
    t = Table([[logo, left, right]], colWidths=[90, 270, CONT_W-90-270])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_ESC),
        ("ROWPADDING", (0,0), (-1,-1), 10),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",      (0,0), (0,0),   "CENTER"),
    ]))
    return t


# ── Banner de status ───────────────────────────────────────────────
def status_banner(status, motivo=""):
    STATUS_MAP = {
        "conclusao":    ("✅ STATUS: CONCLUSÃO",         VERDE),
        "parcial":      ("🔄 STATUS: EXECUÇÃO PARCIAL",  AMARELO_E),
        "cancelamento": ("❌ STATUS: CANCELAMENTO",      VERMELHO),
    }
    if not status or status not in STATUS_MAP:
        return None
    txt, cor = STATUS_MAP[status]
    if motivo:
        txt += f"  —  {motivo}"
    t = Table([[Paragraph(txt, STYLES["status"])]], colWidths=[CONT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), cor),
        ("ROWPADDING", (0,0), (-1,-1), 8),
    ]))
    return t


# ── Bloco de assinaturas ───────────────────────────────────────────
def assinaturas(d, sig_img_path=None):
    def cell(nome, cargo, img=None):
        items = []
        if img and os.path.exists(img):
            try:    items.append(Image(img, width=150, height=50))
            except: items.append(Spacer(1, 50))
        else:
            items.append(Spacer(1, 50))
        items += [
            HRFlowable(width="85%", thickness=1, color=PRETO),
            Paragraph(nome or "_________________________", STYLES["sig_nome"]),
            Paragraph(cargo, STYLES["sig_role"]),
        ]
        return items
    col = CONT_W / 2
    t = Table(
        [[cell(d.get("sig_nome",""), d.get("sig_cargo","Encarregado"), sig_img_path),
          cell("", "Fiscal / Supervisor")]],
        colWidths=[col, col]
    )
    t.setStyle(TableStyle([
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING", (0,0), (-1,-1), 12),
        ("VALIGN",     (0,0), (-1,-1), "BOTTOM"),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ]))
    return t


# ── GERADOR PRINCIPAL ──────────────────────────────────────────────
def gerar_pdf(d):
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    num = d.get("num_projeto", "bdo").replace("/","_").replace(" ","_")
    filename = f"BDO_{num}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=12*mm, bottomMargin=15*mm,
        title=f"BDO {d.get('num_projeto','')}",
        author="Hagap Engenharia Elétrica"
    )

    # Logo
    logo_path = None
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for cand in [
        os.path.join(script_dir, "logo_hagap.png"),
        os.path.join(os.getcwd(), "logo_hagap.png"),
        os.path.join(os.getcwd(), "uploads", "logo_hagap.png"),
        os.path.join(script_dir, "static", "logo.png"),
        os.path.join(os.getcwd(), "static", "logo.png"),
    ]:
        if os.path.exists(cand):
            logo_path = cand
            break

    # Assinatura base64 → tmp
    sig_img_path = None
    sig_data = d.get("sig_data", "")
    if sig_data and "base64," in sig_data:
        try:
            b64 = sig_data.split("base64,")[1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(base64.b64decode(b64)); tmp.close()
            sig_img_path = tmp.name
        except: pass

    story = []

    # ── CABEÇALHO
    story.append(cabecalho(d, logo_path))
    story.append(SP(10))

    # ── IDENTIFICAÇÃO
    story.append(barra("  📋 IDENTIFICAÇÃO"))
    story.append(campos(("Projeto", d.get("num_projeto","")), ("Cidade", d.get("cidade",""))))
    story.append(campos(
        ("Data", d.get("data_bdo","")), ("Hora Início", d.get("hora_inicio","")), ("Hora Final", d.get("hora_final","")),
        widths=[50,95,55,90,55,90]
    ))
    story.append(campos(("Encarregado", d.get("encarregado","")), ("Equipe", d.get("equipe","") or "—")))

    # OMB / PLV
    if d.get("servico_livre") is False:
        story.append(campos(("Nº OMB", d.get("num_omb","")), ("Nº PLV", d.get("num_plv",""))))
    story.append(SP(8))

    # ── STATUS DE EXECUÇÃO
    st_banner = status_banner(d.get("status_execucao",""), d.get("status_motivo",""))
    if st_banner:
        story.append(st_banner)
        story.append(SP(8))

    modo_postes  = d.get("modo_postes", False)
    servicos_c   = d.get("servicos_completo", [])
    materiais    = [m for m in d.get("materiais", []) if m.get("descricao","").strip()]
    obs_lista    = [o for o in d.get("observacoes_lista", []) if o.get("texto","").strip()]
    servicos_txt = d.get("servicos", [])
    manobra      = d.get("manobra_texto","").strip()
    alteracoes   = d.get("alteracoes_texto","").strip()
    observacoes  = d.get("observacoes","").strip()

    if modo_postes:
        # ── MANOBRA / ALTERAÇÕES (sempre seção geral)
        if manobra or alteracoes:
            story.append(barra("  🔌 MANOBRA / ALTERAÇÕES"))
            txt = " — ".join(filter(None, [manobra, alteracoes]))
            story.append(bloco_txt(txt))
            story.append(SP(6))

        # ── GRUPOS POR POSTE
        todos_ps = ["geral"] + list(d.get("postes", []))
        for ps in todos_ps:
            srvs = [s for s in servicos_c if ps in s.get("postes", [])]
            mats = [m for m in materiais   if m.get("poste","geral") == ps]
            obss = [o for o in obs_lista   if o.get("poste","geral") == ps]
            if not srvs and not mats and not obss:
                continue
            ps_label = "  🌐 GERAL DO PROJETO" if ps == "geral" else f"  🏗️ {ps}"
            is_geral = (ps == "geral")
            bloco = grupo_poste(ps_label, srvs, mats, obss, is_geral)
            story.extend(bloco)
            story.append(SP(8))

    else:
        # ── LAYOUT NORMAL
        if manobra or alteracoes or servicos_txt:
            story.append(barra("  🔧 MANOBRA / ALTERAÇÕES / SERVIÇOS EXECUTADOS"))
            linhas = list(filter(None, [manobra, alteracoes] + list(servicos_txt)))
            if linhas:
                story.append(lista_itens(linhas))
            story.append(SP(6))

        if materiais:
            story.append(barra("  📦 MATERIAIS UTILIZADOS"))
            story.append(lista_itens([
                f"{m.get('descricao','')}{' — '+m.get('qtd') if m.get('qtd') else ''}"
                for m in materiais
            ]))
            story.append(SP(6))

        obs_texto = observacoes or " ".join(o.get("texto","") for o in obs_lista)
        if obs_texto.strip():
            story.append(barra("  📝 OBSERVAÇÕES"))
            story.append(bloco_txt(obs_texto))
            story.append(SP(6))

    # ── ASSINATURAS
    story.append(barra("  ✍️ ASSINATURAS"))
    story.append(assinaturas(d, sig_img_path))
    story.append(SP(8))

    # ── RODAPÉ
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BD))
    story.append(SP(3))
    story.append(Paragraph(
        f"Documento gerado automaticamente  |  BDO: {d.get('num_projeto','')}  |  "
        f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Hagap Engenharia Elétrica",
        STYLES["rodape"]
    ))

    doc.build(story)

    if sig_img_path and os.path.exists(sig_img_path):
        try: os.unlink(sig_img_path)
        except: pass

    print(json.dumps({"type":"generated_file","path":filepath,"name":filename}))
    return filename


# ── MAIN ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    dados = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {
        "num_projeto": "PRJ-2026-001", "cidade": "Londrina / PR",
        "data_bdo": "29/05/2026", "hora_inicio": "07:30", "hora_final": "17:00",
        "encarregado": "Carlos Eduardo Silva",
        "servico_livre": False, "num_omb": "OMB-45678", "num_plv": "PLV-12345",
        "manobra_texto": "Feito abertura de chaves e GLV pela equipe executora.",
        "alteracoes_texto": "Executado conforme projeto.",
        "status_execucao": "conclusao", "status_motivo": "",
        "modo_postes": True,
        "postes": ["PS1","PS2"],
        "servicos_completo": [
            {"nome":"🏗️ Reforma de calçada (2 pontos)", "postes":["PS1"]},
            {"nome":"🌳 Podas",                          "postes":["PS1","PS2"]},
            {"nome":"⚡ Malha de aterramento — 35m",     "postes":["geral"]},
        ],
        "materiais": [
            {"descricao":"Conector tipo A","qtd":"4 un","poste":"PS1"},
            {"descricao":"Cabo 10mm","qtd":"50m","poste":"geral"},
        ],
        "observacoes_lista": [
            {"texto":"Rede normalizada às 17h00.","poste":"geral"},
            {"texto":"Poste com inclinação de 5°.","poste":"PS2"},
        ],
        "sig_nome":"Carlos Eduardo Silva","sig_cargo":"Encarregado Elétrico","sig_data":"",
    }
    gerar_pdf(dados)
