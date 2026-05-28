
#!/usr/bin/env python3
"""
BDO — Boletim Diário de Obras
Gerador de PDF com ReportLab + Logo da empresa
"""

import sys, json, os, time, base64, tempfile
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ── Paleta ────────────────────────────────────────────────────────
AZUL_ESCURO   = colors.HexColor("#0f3460")
AZUL_CLARO    = colors.HexColor("#eef4ff")
LARANJA       = colors.HexColor("#e67e22")
CINZA_CLARO   = colors.HexColor("#f0f4fb")
CINZA_BORDA   = colors.HexColor("#d0daf0")
CINZA_TEXTO   = colors.HexColor("#6c757d")
AMARELO_CLARO = colors.HexColor("#fffdf0")
AMARELO_BORDA = colors.HexColor("#f0e68c")
BRANCO        = colors.white
PRETO         = colors.HexColor("#1a1a2e")

W, H = A4  # 595 x 842 pts

# ── Estilos ────────────────────────────────────────────────────────
def estilos():
    base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=PRETO)
    def st(name, **kw):
        return ParagraphStyle(name, **{**base, **kw})
    return {
        "titulo"  : st("titulo",   fontName="Helvetica-Bold", fontSize=13, textColor=BRANCO),
        "subtit"  : st("subtit",   fontSize=8,  textColor=colors.HexColor("#b0bec5")),
        "num_bdo" : st("num",      fontName="Helvetica-Bold", fontSize=20, textColor=BRANCO, alignment=TA_RIGHT),
        "sec_lbl" : st("sec_lbl",  fontName="Helvetica-Bold", fontSize=8,  textColor=BRANCO),
        "label"   : st("label",    fontName="Helvetica-Bold", fontSize=7,  textColor=CINZA_TEXTO),
        "valor"   : st("valor",    fontName="Helvetica-Bold", fontSize=9,  textColor=PRETO),
        "body"    : st("body",     fontSize=9,  textColor=PRETO, leading=14),
        "list_item": st("li",      fontSize=9,  textColor=PRETO, leading=15, leftIndent=10),
        "obs"     : st("obs",      fontSize=8,  textColor=PRETO, leading=12),
        "rodape"  : st("rod",      fontSize=7,  textColor=CINZA_TEXTO, alignment=TA_CENTER),
        "sig_nome": st("sn",       fontName="Helvetica-Bold", fontSize=9,  textColor=PRETO, alignment=TA_CENTER),
        "sig_role": st("sr",       fontSize=8,  textColor=CINZA_TEXTO, alignment=TA_CENTER),
        "auto_txt": st("at",       fontSize=9,  textColor=AZUL_ESCURO, leading=14, fontName="Helvetica-Oblique"),
    }

# ── Barra de seção ─────────────────────────────────────────────────
def sec_bar(titulo, st):
    t = Table([[Paragraph(titulo, st["sec_lbl"])]], colWidths=[W - 30*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_ESCURO),
        ("ROWPADDING", (0,0), (-1,-1), 6),
    ]))
    return t

# ── Linha de campos ────────────────────────────────────────────────
def campos_row(pares, st, col_w=None):
    if not col_w:
        unit = (W - 30*mm) / len(pares) / 2
        col_w = []
        for _ in pares:
            col_w += [unit * 0.60, unit * 1.40]
    flat = []
    for lbl, val in pares:
        flat += [Paragraph(lbl, st["label"]), Paragraph(val or "—", st["valor"])]
    t = Table([flat], colWidths=col_w)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), CINZA_CLARO),
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BORDA),
        ("ROWPADDING", (0,0), (-1,-1), 7),
        ("VALIGN",     (0,0), (-1,-1), "TOP"),
    ]))
    return t

# ── Bloco de texto automático ──────────────────────────────────────
def texto_bloco(texto, st):
    t = Table([[Paragraph(texto, st["auto_txt"])]], colWidths=[W - 30*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_CLARO),
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BORDA),
        ("ROWPADDING", (0,0), (-1,-1), 9),
    ]))
    return t

# ── Tabela de serviços ─────────────────────────────────────────────
def servicos_table(servicos, st):
    if not servicos:
        data = [[Paragraph("Nenhum serviço marcado.", st["obs"])]]
    else:
        data = [[Paragraph(f"• {s}", st["list_item"])] for s in servicos]
    t = Table(data, colWidths=[W - 30*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,-1), BRANCO),
        ("GRID",           (0,0), (-1,-1), 0.3, CINZA_BORDA),
        ("ROWPADDING",     (0,0), (-1,-1), 5),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [BRANCO, colors.HexColor("#f5f7fc")]),
    ]))
    return t

# ── Assinaturas ────────────────────────────────────────────────────
def assinatura_bloco(d, st, sig_img_path=None):
    def sig_cell(nome, cargo, img_path=None):
        items = []
        if img_path and os.path.exists(img_path):
            try:
                items.append(Image(img_path, width=150, height=50))
            except:
                items.append(Spacer(1, 50))
        else:
            items.append(Spacer(1, 50))
        items += [
            HRFlowable(width="85%", thickness=1, color=PRETO),
            Paragraph(nome or "_________________________", st["sig_nome"]),
            Paragraph(cargo, st["sig_role"]),
        ]
        return items

    col = (W - 30*mm) / 2
    t = Table(
        [[sig_cell(d.get("sig_nome",""), d.get("sig_cargo","Encarregado"), sig_img_path),
          sig_cell("", "Fiscal / Supervisor")]],
        colWidths=[col, col]
    )
    t.setStyle(TableStyle([
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BORDA),
        ("ROWPADDING", (0,0), (-1,-1), 12),
        ("VALIGN",     (0,0), (-1,-1), "BOTTOM"),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ]))
    return t

# ── Cabeçalho com logo ─────────────────────────────────────────────
def header_bdo(d, st, logo_path=None):
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    left_items = [
        Paragraph("BOLETIM DIÁRIO DE OBRAS", st["titulo"]),
        Paragraph("Hagap Engenharia Elétrica  |  Registro de Campo", st["subtit"]),
        Paragraph(f"Gerado em: {now}", st["subtit"]),
    ]
    right_items = [
        Paragraph(f"Nº {d.get('num_projeto','---')}", st["num_bdo"]),
        Paragraph(f"Data: {d.get('data_bdo','')}", ParagraphStyle(
            "dr", fontName="Helvetica", fontSize=8,
            textColor=colors.HexColor("#b0bec5"), alignment=TA_RIGHT)),
    ]

    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path, width=90, height=36)
        except:
            logo = Paragraph("HAGAP", ParagraphStyle("lp", fontName="Helvetica-Bold", fontSize=18,
                             textColor=LARANJA, alignment=TA_CENTER))
    else:
        logo = Paragraph("HAGAP", ParagraphStyle("lp2", fontName="Helvetica-Bold", fontSize=18,
                         textColor=LARANJA, alignment=TA_CENTER))

    row = [[logo, left_items, right_items]]
    t = Table(row, colWidths=[100, 270, 165])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_ESCURO),
        ("ROWPADDING", (0,0), (-1,-1), 10),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",      (0,0), (0,0),   "CENTER"),
    ]))
    return t

# ── GERADOR PRINCIPAL ──────────────────────────────────────────────
def gerar_pdf(d):
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    num_proj = d.get("num_projeto", "bdo").replace("/", "_").replace(" ", "_")
    filename = f"BDO_{num_proj}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=12*mm, bottomMargin=15*mm,
        title=f"BDO {d.get('num_projeto','')}",
        author="Hagap Engenharia Elétrica"
    )

    # ── Logo
    logo_path = None
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for candidate in [
        os.path.join(script_dir, "logo_hagap.png"),
        os.path.join(os.getcwd(), "logo_hagap.png"),
        os.path.join(os.getcwd(), "uploads", "logo_hagap.png"),
    ]:
        if os.path.exists(candidate):
            logo_path = candidate
            break

    # ── Assinatura (base64 → tmp)
    sig_img_path = None
    sig_data = d.get("sig_data", "")
    if sig_data and "base64," in sig_data:
        try:
            b64 = sig_data.split("base64,")[1]
            sig_bytes = base64.b64decode(b64)
            tmp_sig = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp_sig.write(sig_bytes)
            tmp_sig.close()
            sig_img_path = tmp_sig.name
        except:
            sig_img_path = None

    st = estilos()
    story = []
    SP = lambda n=6: Spacer(1, n)

    # Cabeçalho
    story.append(header_bdo(d, st, logo_path))
    story.append(SP(10))

    # Identificação
    story.append(sec_bar("  IDENTIFICAÇÃO DO BOLETIM", st))
    story.append(campos_row([("Nº Projeto", d.get("num_projeto","")), ("Cidade", d.get("cidade",""))], st))
    story.append(campos_row(
        [("Data", d.get("data_bdo","")), ("Hora Início", d.get("hora_inicio","")), ("Hora Final", d.get("hora_final",""))],
        st, col_w=[55, 100, 55, 95, 55, 95]
    ))
    story.append(campos_row([("Equipe", d.get("equipe","")), ("Encarregado", d.get("encarregado",""))], st))

    # OMB / PLV
    sl = d.get("servico_livre")
    if sl is False or sl == "false" or sl == False:
        story.append(campos_row([("Nº OMB", d.get("num_omb","")), ("Nº PLV", d.get("num_plv",""))], st))
    story.append(SP(8))

    # Manobra
    manobra_txt = d.get("manobra_texto", "").strip()
    if manobra_txt:
        story.append(sec_bar("  MANOBRA", st))
        story.append(texto_bloco(manobra_txt, st))
        story.append(SP(8))

    # Alterações
    alt_txt = d.get("alteracoes_texto", "").strip()
    if alt_txt:
        story.append(sec_bar("  ALTERAÇÕES / MATERIAIS", st))
        story.append(texto_bloco(alt_txt, st))
        story.append(SP(8))

    # Serviços
    story.append(sec_bar("  SERVIÇOS EXECUTADOS", st))
    story.append(servicos_table(d.get("servicos", []), st))
    story.append(SP(8))

    # Observações
    obs = d.get("observacoes", "").strip()
    if obs:
        story.append(sec_bar("  OBSERVAÇÕES ADICIONAIS", st))
        obs_t = Table([[Paragraph(obs, st["obs"])]], colWidths=[W - 30*mm])
        obs_t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), AMARELO_CLARO),
            ("GRID",       (0,0), (-1,-1), 0.3, AMARELO_BORDA),
            ("ROWPADDING", (0,0), (-1,-1), 9),
        ]))
        story.append(obs_t)
        story.append(SP(8))

    # Assinaturas
    story.append(sec_bar("  ASSINATURAS", st))
    story.append(assinatura_bloco(d, st, sig_img_path))
    story.append(SP(8))

    # Rodapé
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BORDA))
    story.append(SP(3))
    rodape = (f"Documento gerado automaticamente  |  BDO: {d.get('num_projeto','')}  |  "
              f"{datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Hagap Engenharia Elétrica")
    story.append(Paragraph(rodape, st["rodape"]))

    doc.build(story)

    # Limpar tmp
    if sig_img_path and os.path.exists(sig_img_path):
        try: os.unlink(sig_img_path)
        except: pass

    print(json.dumps({"type": "generated_file", "path": filepath, "name": filename}))
    print(f"\n✅ PDF gerado: {filepath}")
    return filename


# ── MAIN ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        dados = json.loads(sys.argv[1])
    else:
        dados = {
            "num_projeto":       "PRJ-2026-001",
            "cidade":            "Londrina / PR",
            "data_bdo":          "27/05/2026",
            "hora_inicio":       "07:30",
            "hora_final":        "17:00",
            "equipe":            "Equipe Alpha — T01",
            "encarregado":       "Carlos Eduardo Silva",
            "servico_livre":     False,
            "num_omb":           "OMB-45678",
            "num_plv":           "PLV-12345",
            "manobra_texto":     "Feito abertura de chaves e GLV pela equipe executora.",
            "alteracoes_texto":  "Executado conforme projeto.",
            "servicos": [
                "Reforma de calçada (2 pontos)",
                "Reinstalação de luminárias — 4 un.",
                "Malha de aterramento — 35 metros",
            ],
            "observacoes": "Serviço executado sem intercorrências. Rede normalizada às 17h00.",
            "sig_nome":  "Carlos Eduardo Silva",
            "sig_cargo": "Encarregado Elétrico",
            "sig_data":  "",
        }
    gerar_pdf(dados)
