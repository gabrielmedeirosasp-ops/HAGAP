
#!/usr/bin/env python3
"""
BDO — Boletim Diário de Obras
Gera PDF usando ReportLab (puro Python, sem dependências de sistema).
Compatível com Render, Railway, Heroku, etc.
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

# ── Paleta de cores ────────────────────────────────────────────────
AZUL_ESC  = colors.HexColor("#061539")
AZUL      = colors.HexColor("#0a2557")
ROXO_ESC  = colors.HexColor("#4a0e8f")
VERDE     = colors.HexColor("#1a7a3c")
VERMELHO  = colors.HexColor("#c0390b")
AMARELO   = colors.HexColor("#c07000")
CINZA_BG  = colors.HexColor("#f0f4fb")
CINZA_BD  = colors.HexColor("#e8ecf5")
CINZA_T   = colors.HexColor("#4a5a7a")
BRANCO    = colors.white
PRETO     = colors.HexColor("#0d1b3e")

W, H = A4
LM = RM = 15 * mm
CW = W - LM - RM


# ── Helpers ────────────────────────────────────────────────────────

def _st(name, **kw):
    base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=PRETO)
    return ParagraphStyle(name, **{**base, **kw})

def _sp(n=6):
    return Spacer(1, n)

def _barra(txt, cor=AZUL):
    t = Table([[Paragraph(txt, _st("sec", fontName="Helvetica-Bold", fontSize=8, textColor=BRANCO))]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), cor),
        ("ROWPADDING", (0,0), (-1,-1), 7),
    ]))
    return t

def _campo_linha(*pares, widths=None):
    if not widths:
        u = CW / len(pares) / 2
        widths = [u * 0.45, u * 1.55] * len(pares)
    flat = []
    for lbl, val in pares:
        flat += [
            Paragraph(str(lbl), _st("lbl", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_T)),
            Paragraph(str(val or "—"), _st("val", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO)),
        ]
    t = Table([flat], colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,-1), CINZA_BG),
        ("GRID",        (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING",  (0,0), (-1,-1), 7),
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
    ]))
    return t

def _lista_itens(itens, fundo=BRANCO):
    if not itens:
        data = [[Paragraph("—", _st("li_vz", fontSize=9, textColor=CINZA_T))]]
    else:
        data = [[Paragraph(f"• {i}", _st("li", fontSize=9, textColor=PRETO, leading=14, leftIndent=10))]
                for i in itens]
    t = Table(data, colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,-1), fundo),
        ("GRID",           (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING",     (0,0), (-1,-1), 6),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [fundo, colors.HexColor("#f7f9ff")]),
    ]))
    return t

def _bloco_txt(txt):
    t = Table([[Paragraph(str(txt or "—"), _st("b", fontSize=9, textColor=PRETO, leading=14))]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BRANCO),
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING", (0,0), (-1,-1), 8),
    ]))
    return t


# ── Cabecalho ──────────────────────────────────────────────────────

def _cabecalho(d, logo_path=None):
    now = datetime.now().strftime('%d/%m/%Y %H:%M')

    if logo_path and os.path.exists(logo_path):
        try:
            logo_cell = Image(logo_path, width=85, height=34)
        except:
            logo_cell = Paragraph("HAGAP", _st("lx", fontName="Helvetica-Bold", fontSize=18,
                                                textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))
    else:
        logo_cell = Paragraph("HAGAP", _st("lx", fontName="Helvetica-Bold", fontSize=18,
                                           textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))

    col_meio = [
        Paragraph("BOLETIM DIARIO DE OBRAS",
                  _st("h1", fontName="Helvetica-Bold", fontSize=13, textColor=BRANCO)),
        _sp(3),
        Paragraph("Hagap Engenharia Eletrica  |  Registro de Campo",
                  _st("h2", fontSize=8, textColor=colors.HexColor("#b0bec5"))),
        Paragraph(f"Gerado em: {now}",
                  _st("h3", fontSize=7, textColor=colors.HexColor("#8090a0"))),
    ]

    col_dir = [
        Paragraph(str(d.get("num_projeto", "---")),
                  _st("np", fontName="Helvetica-Bold", fontSize=22, textColor=BRANCO, alignment=TA_RIGHT)),
        Paragraph(f"Data: {d.get('data_bdo', '')}",
                  _st("dt", fontSize=8, textColor=colors.HexColor("#b0bec5"), alignment=TA_RIGHT)),
    ]

    t = Table([[logo_cell, col_meio, col_dir]], colWidths=[90, 270, CW - 360])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), AZUL_ESC),
        ("ROWPADDING", (0,0), (-1,-1), 12),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",      (0,0), (0,0),   "CENTER"),
    ]))
    return t


# ── Banner de status ────────────────────────────────────────────────

def _status_banner(status, motivo=""):
    STATUS_MAP = {
        "conclusao":    ("STATUS: CONCLUSAO",        VERDE),
        "parcial":      ("STATUS: EXECUCAO PARCIAL", AMARELO),
        "cancelamento": ("STATUS: CANCELAMENTO",     VERMELHO),
    }
    if not status or status not in STATUS_MAP:
        return None
    txt, cor = STATUS_MAP[status]
    if motivo:
        txt += f"  -  {motivo}"
    t = Table([[Paragraph(txt, _st("st", fontName="Helvetica-Bold", fontSize=10,
                                   textColor=BRANCO, alignment=TA_CENTER))]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), cor),
        ("ROWPADDING", (0,0), (-1,-1), 9),
    ]))
    return t


# ── Grupo por poste ─────────────────────────────────────────────────

def _grupo_poste(label, srvs, mats, obss, cor=ROXO_ESC):
    els = [_barra(label, cor)]
    sub_st = _st("sub", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_T)

    if srvs:
        sub = Table([[Paragraph("SERVICOS", sub_st)]], colWidths=[CW])
        sub.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CINZA_BG),
                                  ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
                                  ("ROWPADDING",(0,0),(-1,-1),5)]))
        els.append(sub)
        els.append(_lista_itens([s["nome"] for s in srvs]))

    if mats:
        sub = Table([[Paragraph("MATERIAIS", sub_st)]], colWidths=[CW])
        sub.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CINZA_BG),
                                  ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
                                  ("ROWPADDING",(0,0),(-1,-1),5)]))
        els.append(sub)
        els.append(_lista_itens([
            f"{m.get('descricao','')}{' - '+m.get('qtd') if m.get('qtd') else ''}"
            for m in mats
        ]))

    if obss:
        sub = Table([[Paragraph("OBSERVACOES", sub_st)]], colWidths=[CW])
        sub.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CINZA_BG),
                                  ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
                                  ("ROWPADDING",(0,0),(-1,-1),5)]))
        els.append(sub)
        els.append(_lista_itens([o.get("texto","") for o in obss]))

    return els


# ── Assinaturas ─────────────────────────────────────────────────────

def _assinaturas(d, sig_img_path=None):
    col = CW / 2

    def _cel(nome, cargo, img_path=None):
        items = []
        if img_path and os.path.exists(img_path):
            try:
                items.append(Image(img_path, width=150, height=50))
            except:
                items.append(_sp(50))
        else:
            items.append(_sp(50))
        items.append(HRFlowable(width="85%", thickness=1, color=PRETO))
        items.append(Paragraph(nome or "_________________________",
                                _st("sn", fontName="Helvetica-Bold", fontSize=9,
                                    textColor=PRETO, alignment=TA_CENTER)))
        items.append(Paragraph(cargo,
                                _st("sr", fontSize=8, textColor=CINZA_T, alignment=TA_CENTER)))
        return items

    t = Table(
        [[_cel(d.get("sig_nome",""), d.get("sig_cargo","Encarregado"), sig_img_path),
          _cel("", "Fiscal / Supervisor")]],
        colWidths=[col, col]
    )
    t.setStyle(TableStyle([
        ("GRID",       (0,0), (-1,-1), 0.3, CINZA_BD),
        ("ROWPADDING", (0,0), (-1,-1), 12),
        ("VALIGN",     (0,0), (-1,-1), "BOTTOM"),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ]))
    return t


# ── GERADOR PRINCIPAL ───────────────────────────────────────────────

def gerar_pdf(d):
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    num = (d.get("num_projeto","bdo") or "bdo").replace("/","_").replace(" ","_")
    filename = f"BDO_{num}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    # Logo
    logo_path = None
    sd = os.path.dirname(os.path.abspath(__file__))
    for cand in [
        os.path.join(sd, "static", "logo.png"),
        os.path.join(os.getcwd(), "static", "logo.png"),
        os.path.join(sd, "logo_hagap.png"),
        os.path.join(os.getcwd(), "logo_hagap.png"),
    ]:
        if os.path.exists(cand):
            logo_path = cand
            break

    # Assinatura base64
    sig_img_path = None
    sig_data = d.get("sig_data", "")
    if sig_data and "base64," in sig_data:
        try:
            b64 = sig_data.split("base64,")[1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(base64.b64decode(b64))
            tmp.close()
            sig_img_path = tmp.name
        except:
            pass

    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=LM, rightMargin=RM,
        topMargin=12*mm, bottomMargin=15*mm,
        title=f"BDO {d.get('num_projeto','')}",
        author="Hagap Engenharia Eletrica"
    )

    story = []

    # Cabecalho
    story.append(_cabecalho(d, logo_path))
    story.append(_sp(10))

    # Status
    st_banner = _status_banner(d.get("status_execucao",""), d.get("status_motivo",""))
    if st_banner:
        story.append(st_banner)
        story.append(_sp(8))

    # Identificacao
    story.append(_barra("  IDENTIFICACAO"))
    story.append(_campo_linha(
        ("Projeto", d.get("num_projeto","")),
        ("Cidade",  d.get("cidade","")),
    ))
    story.append(_campo_linha(
        ("Data",        d.get("data_bdo","")),
        ("Hora Inicio", d.get("hora_inicio","")),
        ("Hora Final",  d.get("hora_final","")),
        widths=[45, 105, 55, 100, 52, 78]
    ))
    story.append(_campo_linha(
        ("Encarregado", d.get("encarregado","")),
        ("Equipe",      d.get("equipe","") or "—"),
    ))
    if d.get("servico_livre") is False:
        story.append(_campo_linha(
            ("N OMB", d.get("num_omb","")),
            ("N PLV", d.get("num_plv","")),
        ))
    story.append(_sp(8))

    # Corpo
    modo_postes  = d.get("modo_postes", False)
    servicos_c   = d.get("servicos_completo", [])
    materiais    = [m for m in d.get("materiais", []) if (m.get("descricao","") or "").strip()]
    obs_lista    = [o for o in d.get("observacoes_lista", []) if (o.get("texto","") or "").strip()]
    servicos_txt = d.get("servicos", [])
    manobra      = (d.get("manobra_texto","") or "").strip()
    alteracoes   = (d.get("alteracoes_texto","") or "").strip()
    observacoes  = (d.get("observacoes","") or "").strip()

    if modo_postes:
        if manobra or alteracoes:
            story.append(_barra("  MANOBRA / ALTERACOES"))
            story.append(_bloco_txt(" — ".join(filter(None, [manobra, alteracoes]))))
            story.append(_sp(6))

        todos_ps = ["geral"] + list(d.get("postes", []))
        for ps in todos_ps:
            srvs = [s for s in servicos_c if ps in s.get("postes", [])]
            mats = [m for m in materiais   if m.get("poste","geral") == ps]
            obss = [o for o in obs_lista   if o.get("poste","geral") == ps]
            if not srvs and not mats and not obss:
                continue
            ps_label = "  GERAL DO PROJETO" if ps == "geral" else f"  POSTE: {ps}"
            cor = AZUL if ps == "geral" else ROXO_ESC
            story.extend(_grupo_poste(ps_label, srvs, mats, obss, cor))
            story.append(_sp(6))

    else:
        linhas = list(filter(None, [manobra, alteracoes] + list(servicos_txt)))
        if linhas:
            story.append(_barra("  SERVICOS / MANOBRA / ALTERACOES"))
            story.append(_lista_itens(linhas))
            story.append(_sp(6))

        if materiais:
            story.append(_barra("  MATERIAIS UTILIZADOS"))
            story.append(_lista_itens([
                f"{m.get('descricao','')}{' - '+m.get('qtd') if m.get('qtd') else ''}"
                for m in materiais
            ]))
            story.append(_sp(6))

        obs_texto = observacoes or " ".join(o.get("texto","") for o in obs_lista)
        if obs_texto.strip():
            story.append(_barra("  OBSERVACOES"))
            story.append(_bloco_txt(obs_texto))
            story.append(_sp(6))

    # Assinaturas
    story.append(_barra("  ASSINATURAS"))
    story.append(_assinaturas(d, sig_img_path))
    story.append(_sp(8))

    # Rodape
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BD))
    story.append(_sp(3))
    story.append(Paragraph(
        f"Documento gerado automaticamente  |  BDO: {d.get('num_projeto','')}  |  {now}  |  Hagap Engenharia Eletrica",
        _st("rod", fontSize=7, textColor=CINZA_T, alignment=TA_CENTER)
    ))

    doc.build(story)

    if sig_img_path and os.path.exists(sig_img_path):
        try:
            os.unlink(sig_img_path)
        except:
            pass

    print(json.dumps({"type": "generated_file", "path": filepath, "name": filename}))
    return filename


# ── MAIN ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    dados = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {
        "num_projeto": "PRJ-2026-001", "cidade": "Londrina / PR",
        "data_bdo": "29/05/2026", "hora_inicio": "07:30", "hora_final": "17:00",
        "encarregado": "Carlos Eduardo Silva", "equipe": "Equipe Alpha",
        "servico_livre": False, "num_omb": "OMB-45678", "num_plv": "PLV-12345",
        "manobra_texto": "Abertura de chaves e GLV pela equipe executora.",
        "alteracoes_texto": "Executado conforme projeto.",
        "status_execucao": "conclusao", "status_motivo": "",
        "modo_postes": True, "postes": ["PS1", "PS2"],
        "servicos_completo": [
            {"nome": "Reforma de calcada 2 pontos", "postes": ["PS1"]},
            {"nome": "Podas",                        "postes": ["PS1","PS2"]},
            {"nome": "Malha de aterramento 35m",     "postes": ["geral"]},
        ],
        "materiais": [
            {"descricao": "Conector tipo A", "qtd": "4 un",  "poste": "PS1"},
            {"descricao": "Cabo 10mm",       "qtd": "50m",   "poste": "geral"},
        ],
        "observacoes_lista": [
            {"texto": "Rede normalizada as 17h00.", "poste": "geral"},
            {"texto": "Poste com inclinacao 5 graus.", "poste": "PS2"},
        ],
        "sig_nome": "Carlos Eduardo Silva", "sig_cargo": "Encarregado Eletrico", "sig_data": "",
    }
    gerar_pdf(dados)
