
#!/usr/bin/env python3
"""
BDO — Boletim Diario de Obras
Gera PDF usando APENAS ReportLab (puro Python, sem WeasyPrint, sem libs de sistema).
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
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

import re as _re

def _strip_emoji(text):
    """Remove emojis e caracteres Unicode especiais que o ReportLab nao sabe renderizar."""
    if not text:
        return text
    # Emojis acima do BMP (U+10000+)
    text = _re.sub(r'[\U00010000-\U0010FFFF]', '', text)
    # Blocos de emojis no BMP
    text = _re.sub(r'[\U0001F000-\U0001FFFF]', '', text)
    text = _re.sub(r'[\u2600-\u27BF\u2300-\u23FF\u2B50-\u2B55\u231A-\u231B]', '', text)
    # Modificadores de variacao de emoji (U+FE00-U+FE0F)
    text = _re.sub(r'[\uFE00-\uFE0F]', '', text)
    # Zero Width Joiner e combinadores
    text = _re.sub(r'[\u200D\u200B\u200C\u200E\u200F\uFEFF]', '', text)
    return text.strip()

def _clean(v):
    """Limpa emojis de qualquer valor antes de enviar ao ReportLab."""
    if isinstance(v, str):
        return _strip_emoji(v)
    return v

AZUL_ESC = colors.HexColor("#061539")
AZUL     = colors.HexColor("#0a2557")
ROXO_ESC = colors.HexColor("#4a0e8f")
VERDE    = colors.HexColor("#1a7a3c")
VERMELHO = colors.HexColor("#c0390b")
AMARELO  = colors.HexColor("#c07000")
CINZA_BG = colors.HexColor("#f0f4fb")
CINZA_BD = colors.HexColor("#e8ecf5")
CINZA_T  = colors.HexColor("#4a5a7a")
BRANCO   = colors.white
PRETO    = colors.HexColor("#0d1b3e")

W, H = A4
LM = RM = 15 * mm
CW = W - LM - RM


def _st(name, **kw):
    base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=PRETO)
    return ParagraphStyle(name, **{**base, **kw})

def _sp(n=6):
    return Spacer(1, n)

def _barra(txt, cor=AZUL):
    t = Table([[Paragraph(_clean(txt), _st("s", fontName="Helvetica-Bold", fontSize=8, textColor=BRANCO))]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),cor),("ROWPADDING",(0,0),(-1,-1),7)]))
    return t

def _campos(*pares, widths=None):
    if not widths:
        u = CW / len(pares) / 2
        widths = [u*0.45, u*1.55] * len(pares)
    flat = []
    for lbl, val in pares:
        flat += [
            Paragraph(_clean(str(lbl)), _st("l", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_T)),
            Paragraph(_clean(str(val or "—")), _st("v", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO)),
        ]
    t = Table([flat], colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),CINZA_BG),
        ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
        ("ROWPADDING",(0,0),(-1,-1),7),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    return t

def _lista(itens):
    if not itens:
        data = [[Paragraph("—", _st("li0", fontSize=9, textColor=CINZA_T))]]
    else:
        data = [[Paragraph(f"• {_clean(str(i))}", _st("li", fontSize=9, textColor=PRETO, leading=14, leftIndent=10))] for i in itens]
    t = Table(data, colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),BRANCO),
        ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
        ("ROWPADDING",(0,0),(-1,-1),6),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[BRANCO, colors.HexColor("#f7f9ff")]),
    ]))
    return t

def _bloco(txt):
    t = Table([[Paragraph(_clean(str(txt or "—")), _st("b", fontSize=9, textColor=PRETO, leading=14))]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),BRANCO),("GRID",(0,0),(-1,-1),0.3,CINZA_BD),("ROWPADDING",(0,0),(-1,-1),8)]))
    return t

def _cabecalho(d, logo_path=None):
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    if logo_path and os.path.exists(logo_path):
        try:    logo_cel = Image(logo_path, width=85, height=34)
        except: logo_cel = Paragraph("HAGAP", _st("lx", fontName="Helvetica-Bold", fontSize=18, textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))
    else:
        logo_cel = Paragraph("HAGAP", _st("lx", fontName="Helvetica-Bold", fontSize=18, textColor=colors.HexColor("#e67e22"), alignment=TA_CENTER))

    meio = [
        Paragraph("BOLETIM DIARIO DE OBRAS", _st("h1", fontName="Helvetica-Bold", fontSize=13, textColor=BRANCO)),
        _sp(3),
        Paragraph("Hagap Engenharia Eletrica  |  Registro de Campo", _st("h2", fontSize=8, textColor=colors.HexColor("#b0bec5"))),
        Paragraph(f"Gerado em: {now}", _st("h3", fontSize=7, textColor=colors.HexColor("#8090a0"))),
    ]
    dire = [
        Paragraph(str(d.get("num_projeto","---")), _st("np", fontName="Helvetica-Bold", fontSize=22, textColor=BRANCO, alignment=TA_RIGHT)),
        Paragraph(f"Data: {d.get('data_bdo','')}", _st("dt", fontSize=8, textColor=colors.HexColor("#b0bec5"), alignment=TA_RIGHT)),
    ]
    t = Table([[logo_cel, meio, dire]], colWidths=[90, 270, CW-360])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),AZUL_ESC),("ROWPADDING",(0,0),(-1,-1),12),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(0,0),"CENTER")]))
    return t

def _status_banner(status, motivo=""):
    MAP = {"conclusao":("STATUS: CONCLUSAO",VERDE),"parcial":("STATUS: EXECUCAO PARCIAL",AMARELO),"cancelamento":("STATUS: CANCELAMENTO",VERMELHO)}
    if not status or status not in MAP: return None
    txt, cor = MAP[status]
    if motivo: txt += f"  -  {motivo}"
    t = Table([[Paragraph(txt, _st("st", fontName="Helvetica-Bold", fontSize=10, textColor=BRANCO, alignment=TA_CENTER))]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),cor),("ROWPADDING",(0,0),(-1,-1),9)]))
    return t

def _grupo_poste(label, srvs, mats, obss, cor=ROXO_ESC):
    els = [_barra(label, cor)]
    sub_s = _st("sub", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_T)
    def _sub(txt):
        t = Table([[Paragraph(txt, sub_s)]], colWidths=[CW])
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CINZA_BG),("GRID",(0,0),(-1,-1),0.3,CINZA_BD),("ROWPADDING",(0,0),(-1,-1),5)]))
        return t
    if srvs:
        els.append(_sub("SERVICOS"))
        els.append(_lista([s["nome"] for s in srvs]))
    if mats:
        els.append(_sub("MATERIAIS"))
        els.append(_lista([f"{m.get('descricao','')}{' - '+m.get('qtd') if m.get('qtd') else ''}" for m in mats]))
    if obss:
        els.append(_sub("OBSERVACOES"))
        els.append(_lista([o.get("texto","") for o in obss]))
    return els

def _assinaturas(d, sig_img_path=None):
    col = CW / 2
    def _cel(nome, cargo, img=None):
        items = []
        if img and os.path.exists(img):
            try: items.append(Image(img, width=150, height=50))
            except: items.append(_sp(50))
        else:
            items.append(_sp(50))
        items.append(HRFlowable(width="85%", thickness=1, color=PRETO))
        items.append(Paragraph(nome or "_________________________", _st("sn", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO, alignment=TA_CENTER)))
        items.append(Paragraph(cargo, _st("sr", fontSize=8, textColor=CINZA_T, alignment=TA_CENTER)))
        return items
    t = Table([[_cel(d.get("sig_nome",""), d.get("sig_cargo","Encarregado"), sig_img_path), _cel("","Fiscal / Supervisor")]], colWidths=[col,col])
    t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,CINZA_BD),("ROWPADDING",(0,0),(-1,-1),12),("VALIGN",(0,0),(-1,-1),"BOTTOM"),("ALIGN",(0,0),(-1,-1),"CENTER")]))
    return t


def gerar_pdf(d):
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    num = (d.get("num_projeto","bdo") or "bdo").replace("/","_").replace(" ","_")
    filename = f"BDO_{num}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    # Logo
    logo_path = None
    sd = os.path.dirname(os.path.abspath(__file__))
    for cand in [os.path.join(sd,"static","logo.png"), os.path.join(os.getcwd(),"static","logo.png"),
                 os.path.join(sd,"logo_hagap.png"), os.path.join(os.getcwd(),"logo_hagap.png")]:
        if os.path.exists(cand): logo_path = cand; break

    # Assinatura
    sig_img_path = None
    sig_data = d.get("sig_data","")
    if sig_data and "base64," in sig_data:
        try:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(base64.b64decode(sig_data.split("base64,")[1]))
            tmp.close()
            sig_img_path = tmp.name
        except: pass

    doc = SimpleDocTemplate(filepath, pagesize=A4, leftMargin=LM, rightMargin=RM,
                            topMargin=12*mm, bottomMargin=15*mm,
                            title=f"BDO {d.get('num_projeto','')}", author="Hagap Engenharia Eletrica")
    story = []

    story.append(_cabecalho(d, logo_path))
    story.append(_sp(10))

    sb = _status_banner(d.get("status_execucao",""), d.get("status_motivo",""))
    if sb: story += [sb, _sp(8)]

    story.append(_barra("  IDENTIFICACAO"))
    story.append(_campos(("Projeto", d.get("num_projeto","")), ("Cidade", d.get("cidade",""))))
    story.append(_campos(("Data", d.get("data_bdo","")), ("Hora Inicio", d.get("hora_inicio","")), ("Hora Final", d.get("hora_final","")), widths=[45,105,55,100,52,78]))
    story.append(_campos(("Encarregado", d.get("encarregado","")), ("Equipe", d.get("equipe","") or "—")))
    if d.get("servico_livre") is False:
        story.append(_campos(("N OMB", d.get("num_omb","")), ("N PLV", d.get("num_plv",""))))
    story.append(_sp(8))

    modo_postes  = d.get("modo_postes", False)
    servicos_c   = d.get("servicos_completo", [])
    materiais    = [m for m in d.get("materiais",[]) if (m.get("descricao","") or "").strip()]
    obs_lista    = [o for o in d.get("observacoes_lista",[]) if (o.get("texto","") or "").strip()]
    servicos_txt = d.get("servicos", [])
    manobra      = (d.get("manobra_texto","") or "").strip()
    alteracoes   = (d.get("alteracoes_texto","") or "").strip()
    observacoes  = (d.get("observacoes","") or "").strip()

    if modo_postes:
        if manobra or alteracoes:
            story.append(_barra("  MANOBRA / ALTERACOES"))
            story.append(_bloco(" — ".join(filter(None,[manobra,alteracoes]))))
            story.append(_sp(6))
        for ps in ["geral"] + list(d.get("postes",[])):
            srvs = [s for s in servicos_c if ps in s.get("postes",[])]
            mats = [m for m in materiais   if m.get("poste","geral") == ps]
            obss = [o for o in obs_lista   if o.get("poste","geral") == ps]
            if not srvs and not mats and not obss: continue
            label = "  GERAL DO PROJETO" if ps == "geral" else f"  POSTE: {ps}"
            cor   = AZUL if ps == "geral" else ROXO_ESC
            story.extend(_grupo_poste(label, srvs, mats, obss, cor))
            story.append(_sp(6))
    else:
        linhas = list(filter(None, [manobra, alteracoes] + list(servicos_txt)))
        if linhas:
            story.append(_barra("  SERVICOS / MANOBRA / ALTERACOES"))
            story.append(_lista(linhas))
            story.append(_sp(6))
        if materiais:
            story.append(_barra("  MATERIAIS UTILIZADOS"))
            story.append(_lista([f"{m.get('descricao','')}{' - '+m.get('qtd') if m.get('qtd') else ''}" for m in materiais]))
            story.append(_sp(6))
        obs = observacoes or " ".join(o.get("texto","") for o in obs_lista)
        if obs.strip():
            story.append(_barra("  OBSERVACOES"))
            story.append(_bloco(obs))
            story.append(_sp(6))

    story.append(_barra("  ASSINATURAS"))
    story.append(_assinaturas(d, sig_img_path))
    story.append(_sp(8))

    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BD))
    story.append(_sp(3))
    story.append(Paragraph(
        f"Documento gerado automaticamente  |  BDO: {d.get('num_projeto','')}  |  {now}  |  Hagap Engenharia Eletrica",
        _st("rod", fontSize=7, textColor=CINZA_T, alignment=TA_CENTER)
    ))

    doc.build(story)

    if sig_img_path:
        try: os.unlink(sig_img_path)
        except: pass

    print(json.dumps({"type":"generated_file","path":filepath,"name":filename}))
    return filename


if __name__ == "__main__":
    dados = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {
        "num_projeto":"PRJ-2026-001","cidade":"Londrina / PR",
        "data_bdo":"29/05/2026","hora_inicio":"07:30","hora_final":"17:00",
        "encarregado":"Carlos Eduardo Silva","equipe":"Equipe Alpha",
        "servico_livre":False,"num_omb":"OMB-45678","num_plv":"PLV-12345",
        "manobra_texto":"Abertura de chaves e GLV.","alteracoes_texto":"Conforme projeto.",
        "status_execucao":"conclusao","status_motivo":"",
        "modo_postes":True,"postes":["PS1","PS2"],
        "servicos_completo":[
            {"nome":"Reforma de calcada","postes":["PS1"]},
            {"nome":"Podas","postes":["PS1","PS2"]},
            {"nome":"Malha de aterramento 35m","postes":["geral"]},
        ],
        "materiais":[{"descricao":"Conector tipo A","qtd":"4 un","poste":"PS1"},{"descricao":"Cabo 10mm","qtd":"50m","poste":"geral"}],
        "observacoes_lista":[{"texto":"Rede normalizada as 17h00.","poste":"geral"}],
        "sig_nome":"Carlos Eduardo Silva","sig_cargo":"Encarregado Eletrico","sig_data":"",
    }
    gerar_pdf(dados)


def gerar_pdf_from_html(html_content, css_extra="", data=None):
    """
    Gera PDF usando WeasyPrint a partir do HTML da prévia (idêntico ao que aparece na tela).
    """
    from weasyprint import HTML as WP_HTML
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    d = data or {}
    num = (d.get("num_projeto", "bdo") or "bdo").replace("/", "_").replace(" ", "_")
    filename = f"BDO_{num}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    full_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  :root{{--primary:#0a2557;--accent:#1565c0;--bg:#e8eef8;--border:#b8c8e8;
        --text:#0d1b3e;--muted:#4a5a7a;--success:#1a7a3c;--danger:#c0392b;
        --warning:#d68910;--ps:#7b2ff7;}}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{font-family:Arial,'Segoe UI',sans-serif;color:#0d1b3e;background:#fff;padding:0;margin:0;}}
  @page{{size:A4;margin:12mm 14mm 14mm 14mm;}}
  .bdo-preview{{background:white;font-size:12px;line-height:1.6;width:100%;}}
  .bdo-hbar{{background:linear-gradient(135deg,#061539,#0a2f7a);color:white;padding:14px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px;}}
  .bdo-hbar img{{height:50px;background:white;border-radius:6px;padding:3px 8px;}}
  .bdo-hinfo{{flex:1;}}
  .bdo-hinfo h2{{font-size:14px;font-weight:800;margin-bottom:2px;}}
  .bdo-hinfo p{{font-size:10px;opacity:.75;}}
  .bdo-num{{font-size:22px;font-weight:900;opacity:.9;text-align:right;white-space:nowrap;}}
  .bdo-sec{{padding:5px 14px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:white;}}
  .bdo-sec.azul{{background:#0a2557;}}
  .bdo-sec.roxo{{background:linear-gradient(90deg,#4a0e8f,#7b2ff7);}}
  .bdo-sec.verde{{background:#1a7a3c;}}
  .bdo-sec.laranja{{background:#c0390b;}}
  .bdo-sec.amarelo{{background:#c07000;}}
  .bdo-row{{display:grid;grid-template-columns:1fr 1fr;}}
  .bdo-field{{padding:7px 14px;border-bottom:1px solid #e8ecf5;}}
  .bdo-field:nth-child(odd){{border-right:1px solid #e8ecf5;}}
  .bdo-field .fl{{font-size:9px;font-weight:700;color:#4a5a7a;text-transform:uppercase;margin-bottom:2px;}}
  .bdo-field .fv{{font-size:12px;font-weight:700;}}
  .bdo-block{{padding:10px 14px;border-bottom:1px solid #e8ecf5;}}
  .bdo-block .bl{{font-size:9px;font-weight:700;color:#4a5a7a;text-transform:uppercase;margin-bottom:4px;}}
  .bdo-sig{{display:grid;grid-template-columns:1fr 1fr;}}
  .bdo-sig-cell{{padding:12px;border-right:1px solid #e8ecf5;text-align:center;}}
  .bdo-sig-cell:last-child{{border-right:none;}}
  .bdo-sig-line{{border-top:1.5px solid #333;margin:30px 20px 6px;}}
  .bdo-sig-name{{font-size:11px;font-weight:700;}}
  .bdo-sig-role{{font-size:10px;color:#4a5a7a;}}
  .bdo-footer{{background:#f0f4fb;padding:6px 14px;font-size:9px;color:#4a5a7a;text-align:center;border-top:1px solid #b8c8e8;}}
  ul{{margin-left:16px;margin-top:4px;}}
  li{{margin-bottom:3px;}}
  {css_extra}
</style>
</head>
<body>
<div class="bdo-preview">
{html_content}
</div>
</body>
</html>"""

    WP_HTML(string=full_html).write_pdf(filepath)
    print(json.dumps({"type": "generated_file", "path": filepath, "name": filename}))
    return filename
