#!/usr/bin/env python3
"""
BDO — Boletim Diário de Obras
Gera PDF idêntico à prévia HTML.
Tenta WeasyPrint primeiro; se não disponível, usa ReportLab como fallback.
"""
import sys, json, os, time, base64, tempfile
from datetime import datetime


# ═══════════════════════════════════════════════════════════════════
#  GERADOR HTML  (usado pelo WeasyPrint)
# ═══════════════════════════════════════════════════════════════════

def _esc(s):
    return str(s or "—").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _row(label, val, bg="#f0f4fb"):
    return f'''
    <tr style="background:{bg};">
      <td style="padding:9px 14px;font-weight:700;color:#4a5a7a;font-size:12px;width:35%;
                 border-bottom:1px solid #e0e8f5;">{_esc(label)}</td>
      <td style="padding:9px 14px;font-weight:600;color:#0d1b3e;font-size:13px;
                 border-bottom:1px solid #e0e8f5;">{_esc(val)}</td>
    </tr>'''

def _section(titulo, cor="#0a2557"):
    return (f'<div style="background:{cor};color:white;padding:9px 14px;font-weight:700;'
            f'font-size:12px;margin-top:14px;border-radius:6px 6px 0 0;">{titulo}</div>')

def _lista(itens):
    if not itens:
        return ('<div style="background:#fff;padding:10px 14px;color:#888;font-size:12px;'
                'border:1px solid #e0e8f5;border-top:none;border-radius:0 0 6px 6px;">—</div>')
    lis = "".join(f'<li style="padding:4px 0;font-size:12px;">{_esc(i)}</li>' for i in itens)
    return (f'<ul style="background:#fff;padding:10px 14px 10px 28px;border:1px solid #e0e8f5;'
            f'border-top:none;border-radius:0 0 6px 6px;margin:0;">{lis}</ul>')

def _bloco_txt(txt):
    return (f'<div style="background:#fff;padding:10px 14px;border:1px solid #e0e8f5;'
            f'border-top:none;border-radius:0 0 6px 6px;font-size:12px;color:#0d1b3e;">{_esc(txt)}</div>')

def _grupo_poste(label, srvs, mats, obss, cor="#4a0e8f"):
    html = _section(label, cor)
    html += ('<div style="border:1px solid #e0e8f5;border-top:none;'
             'border-radius:0 0 6px 6px;margin-bottom:4px;">')
    sub = ('<div style="background:#f0f4fb;padding:6px 14px;font-size:11px;'
           'font-weight:700;color:#4a5a7a;border-bottom:1px solid #e0e8f5;">')
    if srvs:
        html += sub + "🔧 SERVIÇOS</div>"
        html += "<ul style='padding:8px 14px 8px 28px;margin:0;'>"
        for s in srvs:
            html += f'<li style="font-size:12px;padding:2px 0;">{_esc(s["nome"])}</li>'
        html += "</ul>"
    if mats:
        html += sub + "📦 MATERIAIS</div>"
        html += "<ul style='padding:8px 14px 8px 28px;margin:0;'>"
        for m in mats:
            desc = m.get("descricao","")
            qtd  = m.get("qtd","")
            html += f'<li style="font-size:12px;padding:2px 0;">{_esc(desc)}{(" — "+_esc(qtd)) if qtd else ""}</li>'
        html += "</ul>"
    if obss:
        html += sub + "📝 OBSERVAÇÕES</div>"
        html += "<ul style='padding:8px 14px 8px 28px;margin:0;'>"
        for o in obss:
            html += f'<li style="font-size:12px;padding:2px 0;">{_esc(o.get("texto",""))}</li>'
        html += "</ul>"
    html += "</div>"
    return html

def gerar_html(d, logo_b64=None):
    now = datetime.now().strftime('%d/%m/%Y %H:%M')
    logo_tag = (f'<img src="data:image/png;base64,{logo_b64}" '
                f'style="height:40px;object-fit:contain;background:white;border-radius:6px;padding:2px 5px;">'
                if logo_b64 else
                '<span style="font-size:20px;font-weight:900;color:#e67e22;">HAGAP</span>')

    STATUS_MAP = {
        "conclusao":    ("✅ STATUS: CONCLUSÃO",         "#1a7a3c"),
        "parcial":      ("🔄 STATUS: EXECUÇÃO PARCIAL",  "#c07000"),
        "cancelamento": ("❌ STATUS: CANCELAMENTO",      "#c0390b"),
    }
    st = d.get("status_execucao","")
    motivo = d.get("status_motivo","")
    status_html = ""
    if st in STATUS_MAP:
        stxt, scor = STATUS_MAP[st]
        if motivo: stxt += f" — {motivo}"
        status_html = (f'<div style="background:{scor};color:white;padding:10px 14px;'
                       f'font-weight:700;font-size:13px;text-align:center;border-radius:6px;'
                       f'margin-top:14px;">{stxt}</div>')

    omb_plv_html = ""
    if d.get("servico_livre") is False:
        omb_plv_html = _row("📄 Nº OMB", d.get("num_omb",""), "#fff") + _row("📄 Nº PLV", d.get("num_plv",""))

    modo_postes  = d.get("modo_postes", False)
    servicos_c   = d.get("servicos_completo", [])
    materiais    = [m for m in d.get("materiais", []) if m.get("descricao","").strip()]
    obs_lista    = [o for o in d.get("observacoes_lista", []) if o.get("texto","").strip()]
    servicos_txt = d.get("servicos", [])
    manobra      = (d.get("manobra_texto","") or "").strip()
    alteracoes   = (d.get("alteracoes_texto","") or "").strip()
    observacoes  = (d.get("observacoes","") or "").strip()

    corpo = ""
    if modo_postes:
        if manobra or alteracoes:
            corpo += _section("🔌 MANOBRA / ALTERAÇÕES")
            corpo += _bloco_txt(" — ".join(filter(None, [manobra, alteracoes])))
        todos_ps = ["geral"] + list(d.get("postes", []))
        for ps in todos_ps:
            srvs = [s for s in servicos_c if ps in s.get("postes", [])]
            mats = [m for m in materiais   if m.get("poste","geral") == ps]
            obss = [o for o in obs_lista   if o.get("poste","geral") == ps]
            if not srvs and not mats and not obss:
                continue
            ps_label = "🌐 GERAL DO PROJETO" if ps == "geral" else f"🏗️ {ps}"
            cor = "#0a2557" if ps == "geral" else "#4a0e8f"
            corpo += _grupo_poste(ps_label, srvs, mats, obss, cor)
    else:
        linhas = list(filter(None, [manobra, alteracoes] + list(servicos_txt)))
        if linhas:
            corpo += _section("🔧 MANOBRA / ALTERAÇÕES / SERVIÇOS EXECUTADOS")
            corpo += _lista(linhas)
        if materiais:
            corpo += _section("📦 MATERIAIS UTILIZADOS")
            linha_mats = [f"{m.get('descricao','')}{' — '+m.get('qtd') if m.get('qtd') else ''}" for m in materiais]
            corpo += _lista(linha_mats)
        obs_texto = observacoes or " ".join(o.get("texto","") for o in obs_lista)
        if obs_texto.strip():
            corpo += _section("📝 OBSERVAÇÕES")
            corpo += _bloco_txt(obs_texto)

    sig_html = ""
    sig_data = d.get("sig_data","")
    if sig_data and "base64," in sig_data:
        sig_html = f'<img src="{sig_data}" style="max-height:55px;max-width:180px;display:block;margin:0 auto 4px;">'

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  @page {{ size: A4; margin: 14mm 15mm 16mm 15mm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: Arial, 'Helvetica Neue', sans-serif; color: #0d1b3e; background: white; }}
  table {{ border-collapse: collapse; width: 100%; }}
  td {{ vertical-align: top; }}
</style>
</head>
<body>
<table style="background:linear-gradient(135deg,#061539,#0a2557);border-radius:10px 10px 0 0;width:100%;">
  <tr>
    <td style="padding:14px 18px;width:90px;">{logo_tag}</td>
    <td style="padding:14px 18px;">
      <div style="color:white;font-weight:800;font-size:16px;">BOLETIM DIÁRIO DE OBRAS</div>
      <div style="color:#b0bec5;font-size:11px;margin-top:2px;">Hagap Engenharia Elétrica · Registro de Campo</div>
      <div style="color:#b0bec5;font-size:10px;margin-top:2px;">Gerado em: {now}</div>
    </td>
    <td style="padding:14px 18px;text-align:right;">
      <div style="color:white;font-weight:900;font-size:24px;">{_esc(d.get("num_projeto","---"))}</div>
      <div style="color:#b0bec5;font-size:11px;">Data: {_esc(d.get("data_bdo",""))}</div>
    </td>
  </tr>
</table>

{status_html}

{_section("📋 IDENTIFICAÇÃO")}
<table style="border:1px solid #e0e8f5;border-top:none;border-radius:0 0 6px 6px;">
  {_row("📁 Projeto",     d.get("num_projeto",""))}
  {_row("📍 Cidade",      d.get("cidade",""), "#fff")}
  {_row("📅 Data",        d.get("data_bdo",""))}
  {_row("⏰ Horário",      f"{d.get('hora_inicio','—')} → {d.get('hora_final','—')}", "#fff")}
  {_row("👷 Encarregado", d.get("encarregado",""))}
  {_row("👥 Equipe",      d.get("equipe","") or "—", "#fff")}
  {omb_plv_html}
</table>

{corpo}

<div style="margin-top:18px;">
{_section("✍️ ASSINATURAS")}
<table style="border:1px solid #e0e8f5;border-top:none;border-radius:0 0 6px 6px;">
  <tr>
    <td style="padding:16px;width:50%;text-align:center;border-right:1px solid #e0e8f5;">
      {sig_html}
      <div style="border-top:1px solid #0d1b3e;padding-top:6px;margin-top:50px;">
        <div style="font-weight:700;font-size:12px;">{_esc(d.get("sig_nome","") or "_________________________")}</div>
        <div style="color:#4a5a7a;font-size:11px;">{_esc(d.get("sig_cargo","") or "Encarregado")}</div>
      </div>
    </td>
    <td style="padding:16px;text-align:center;">
      <div style="border-top:1px solid #0d1b3e;padding-top:6px;margin-top:50px;">
        <div style="font-weight:700;font-size:12px;">_________________________</div>
        <div style="color:#4a5a7a;font-size:11px;">Fiscal / Supervisor</div>
      </div>
    </td>
  </tr>
</table>
</div>

<div style="margin-top:14px;border-top:1px solid #e0e8f5;padding-top:6px;text-align:center;
            font-size:10px;color:#4a5a7a;">
  Documento gerado automaticamente · BDO: {_esc(d.get("num_projeto",""))} · {now} · Hagap Engenharia Elétrica
</div>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════════
#  FALLBACK REPORTLAB (quando WeasyPrint não disponível)
# ═══════════════════════════════════════════════════════════════════

def _gerar_pdf_reportlab(d, filepath):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT

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
    LM = RM = 15*mm
    CW = W - LM - RM
    now = datetime.now().strftime('%d/%m/%Y %H:%M')

    def st(name, **kw):
        base = dict(fontName="Helvetica", fontSize=9, leading=13, textColor=PRETO)
        return ParagraphStyle(name, **{**base, **kw})

    S = {
        "titulo": st("t1", fontName="Helvetica-Bold", fontSize=13, textColor=BRANCO),
        "subtit":  st("t2", fontSize=8, textColor=colors.HexColor("#b0bec5")),
        "num":     st("t3", fontName="Helvetica-Bold", fontSize=20, textColor=BRANCO, alignment=TA_RIGHT),
        "data_r":  st("t4", fontSize=8, textColor=colors.HexColor("#b0bec5"), alignment=TA_RIGHT),
        "sec":     st("s1", fontName="Helvetica-Bold", fontSize=8, textColor=BRANCO),
        "label":   st("l1", fontName="Helvetica-Bold", fontSize=7, textColor=CINZA_T),
        "valor":   st("v1", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO),
        "body":    st("b1", fontSize=9, textColor=PRETO, leading=14),
        "li":      st("li", fontSize=9, textColor=PRETO, leading=15, leftIndent=12),
        "rodape":  st("ro", fontSize=7, textColor=CINZA_T, alignment=TA_CENTER),
        "sig_n":   st("sn", fontName="Helvetica-Bold", fontSize=9, textColor=PRETO, alignment=TA_CENTER),
        "sig_r":   st("sr", fontSize=8, textColor=CINZA_T, alignment=TA_CENTER),
        "status":  st("ss", fontName="Helvetica-Bold", fontSize=10, textColor=BRANCO, alignment=TA_CENTER),
    }
    SP = lambda n=6: Spacer(1, n)

    def barra(txt, cor=AZUL):
        t = Table([[Paragraph(txt, S["sec"])]], colWidths=[CW])
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),cor),("ROWPADDING",(0,0),(-1,-1),6)]))
        return t

    def campo_linha(*pares, widths=None):
        if not widths:
            u = CW / len(pares) / 2
            widths = [u*0.55, u*1.45] * len(pares)
        flat = []
        for lbl, val in pares:
            flat += [Paragraph(lbl, S["label"]), Paragraph(str(val or "—"), S["valor"])]
        t = Table([flat], colWidths=widths)
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),CINZA_BG),
            ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
            ("ROWPADDING",(0,0),(-1,-1),7),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        return t

    def lista_itens(itens):
        data = [[Paragraph(f"• {i}", S["li"])] for i in itens] if itens else [[Paragraph("—", S["body"])]]
        t = Table(data, colWidths=[CW])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BRANCO),
            ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
            ("ROWPADDING",(0,0),(-1,-1),5),
        ]))
        return t

    def bloco_txt(txt):
        t = Table([[Paragraph(txt or "—", S["body"])]], colWidths=[CW])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),BRANCO),
            ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
            ("ROWPADDING",(0,0),(-1,-1),8),
        ]))
        return t

    doc = SimpleDocTemplate(filepath, pagesize=A4,
                            leftMargin=LM, rightMargin=RM,
                            topMargin=12*mm, bottomMargin=15*mm)

    # Logo
    logo_path = None
    sd = os.path.dirname(os.path.abspath(__file__))
    for c in [os.path.join(sd,"static","logo.png"), os.path.join(sd,"logo_hagap.png"),
              os.path.join(os.getcwd(),"static","logo.png")]:
        if os.path.exists(c):
            logo_path = c; break

    # Cabeçalho
    from reportlab.platypus import Image
    if logo_path:
        try: logo_cell = Image(logo_path, width=80, height=32)
        except: logo_cell = Paragraph("HAGAP", ParagraphStyle("lx",fontName="Helvetica-Bold",fontSize=16,textColor=colors.HexColor("#e67e22")))
    else:
        logo_cell = Paragraph("HAGAP", ParagraphStyle("lx",fontName="Helvetica-Bold",fontSize=16,textColor=colors.HexColor("#e67e22")))

    cab = Table([
        [logo_cell,
         [Paragraph("BOLETIM DIÁRIO DE OBRAS", S["titulo"]),
          Paragraph("Hagap Engenharia Elétrica · Registro de Campo", S["subtit"]),
          Paragraph(f"Gerado em: {now}", S["subtit"])],
         [Paragraph(f"Nº {d.get('num_projeto','---')}", S["num"]),
          Paragraph(f"Data: {d.get('data_bdo','')}", S["data_r"])]]
    ], colWidths=[80, 260, CW-340])
    cab.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),AZUL_ESC),
        ("ROWPADDING",(0,0),(-1,-1),10),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))

    story = [cab, SP(10)]

    # Status
    STATUS_MAP = {
        "conclusao":    ("✅ STATUS: CONCLUSÃO", VERDE),
        "parcial":      ("🔄 STATUS: EXECUÇÃO PARCIAL", AMARELO),
        "cancelamento": ("❌ STATUS: CANCELAMENTO", VERMELHO),
    }
    st_key = d.get("status_execucao","")
    if st_key in STATUS_MAP:
        stxt, scor = STATUS_MAP[st_key]
        motivo = d.get("status_motivo","")
        if motivo: stxt += f" — {motivo}"
        st_t = Table([[Paragraph(stxt, S["status"])]], colWidths=[CW])
        st_t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),scor),("ROWPADDING",(0,0),(-1,-1),8)]))
        story += [st_t, SP(8)]

    # Identificação
    story.append(barra("  📋 IDENTIFICAÇÃO"))
    story.append(campo_linha(("Projeto", d.get("num_projeto","")), ("Cidade", d.get("cidade",""))))
    story.append(campo_linha(
        ("Data", d.get("data_bdo","")),
        ("Hora Início", d.get("hora_inicio","")),
        ("Hora Final", d.get("hora_final","")),
        widths=[45,100,55,95,50,90]
    ))
    story.append(campo_linha(("Encarregado", d.get("encarregado","")), ("Equipe", d.get("equipe","") or "—")))
    if d.get("servico_livre") is False:
        story.append(campo_linha(("Nº OMB", d.get("num_omb","")), ("Nº PLV", d.get("num_plv",""))))
    story.append(SP(8))

    # Corpo
    modo_postes  = d.get("modo_postes", False)
    servicos_c   = d.get("servicos_completo", [])
    materiais    = [m for m in d.get("materiais", []) if m.get("descricao","").strip()]
    obs_lista    = [o for o in d.get("observacoes_lista", []) if o.get("texto","").strip()]
    servicos_txt = d.get("servicos", [])
    manobra      = (d.get("manobra_texto","") or "").strip()
    alteracoes   = (d.get("alteracoes_texto","") or "").strip()
    observacoes  = (d.get("observacoes","") or "").strip()

    if modo_postes:
        if manobra or alteracoes:
            story.append(barra("  🔌 MANOBRA / ALTERAÇÕES"))
            story.append(bloco_txt(" — ".join(filter(None,[manobra,alteracoes]))))
            story.append(SP(6))
        todos_ps = ["geral"] + list(d.get("postes",[]))
        for ps in todos_ps:
            srvs = [s for s in servicos_c if ps in s.get("postes",[])]
            mats = [m for m in materiais   if m.get("poste","geral") == ps]
            obss = [o for o in obs_lista   if o.get("poste","geral") == ps]
            if not srvs and not mats and not obss: continue
            ps_label = "  🌐 GERAL" if ps == "geral" else f"  🏗️ {ps}"
            cor = AZUL if ps == "geral" else ROXO_ESC
            story.append(barra(ps_label, cor))
            if srvs: story.append(lista_itens([s["nome"] for s in srvs]))
            if mats: story.append(lista_itens([f"{m.get('descricao','')}{' — '+m.get('qtd') if m.get('qtd') else ''}" for m in mats]))
            if obss: story.append(lista_itens([o.get("texto","") for o in obss]))
            story.append(SP(6))
    else:
        linhas = list(filter(None, [manobra, alteracoes] + list(servicos_txt)))
        if linhas:
            story.append(barra("  🔧 SERVIÇOS / MANOBRA / ALTERAÇÕES"))
            story.append(lista_itens(linhas))
            story.append(SP(6))
        if materiais:
            story.append(barra("  📦 MATERIAIS UTILIZADOS"))
            story.append(lista_itens([f"{m.get('descricao','')}{' — '+m.get('qtd') if m.get('qtd') else ''}" for m in materiais]))
            story.append(SP(6))
        obs_texto = observacoes or " ".join(o.get("texto","") for o in obs_lista)
        if obs_texto.strip():
            story.append(barra("  📝 OBSERVAÇÕES"))
            story.append(bloco_txt(obs_texto))
            story.append(SP(6))

    # Assinaturas
    story.append(barra("  ✍️ ASSINATURAS"))
    col = CW / 2
    sig_cells = []
    for nome_sig, cargo_sig in [
        (d.get("sig_nome",""), d.get("sig_cargo","Encarregado")),
        ("", "Fiscal / Supervisor")
    ]:
        cel = [SP(40), HRFlowable(width="85%", thickness=1, color=PRETO),
               Paragraph(nome_sig or "_________________________", S["sig_n"]),
               Paragraph(cargo_sig, S["sig_r"])]
        sig_cells.append(cel)
    st_t = Table([sig_cells], colWidths=[col, col])
    st_t.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.3,CINZA_BD),
        ("ROWPADDING",(0,0),(-1,-1),12),
        ("VALIGN",(0,0),(-1,-1),"BOTTOM"),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
    ]))
    story += [st_t, SP(8)]

    # Rodapé
    story.append(HRFlowable(width="100%", thickness=0.5, color=CINZA_BD))
    story.append(SP(3))
    story.append(Paragraph(
        f"Documento gerado automaticamente · BDO: {d.get('num_projeto','')} · {now} · Hagap Engenharia Elétrica",
        S["rodape"]
    ))

    doc.build(story)


# ═══════════════════════════════════════════════════════════════════
#  FUNÇÃO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════

def gerar_pdf(d):
    output_dir = os.path.join(os.getcwd(), "files")
    os.makedirs(output_dir, exist_ok=True)

    num = d.get("num_projeto", "bdo").replace("/","_").replace(" ","_")
    filename = f"BDO_{num}_{int(time.time())}.pdf"
    filepath = os.path.join(output_dir, filename)

    # ── Tenta WeasyPrint (idêntico à prévia) ──────────────────────
    weasyprint_ok = False
    try:
        from weasyprint import HTML as WHTML

        # Logo como base64
        logo_b64 = None
        sd = os.path.dirname(os.path.abspath(__file__))
        for cand in [os.path.join(sd,"static","logo.png"),
                     os.path.join(os.getcwd(),"static","logo.png"),
                     os.path.join(sd,"logo_hagap.png")]:
            if os.path.exists(cand):
                try:
                    with open(cand,"rb") as f:
                        logo_b64 = base64.b64encode(f.read()).decode()
                    break
                except: pass

        html_str = gerar_html(d, logo_b64)
        WHTML(string=html_str).write_pdf(filepath)
        weasyprint_ok = True
        print(f"[PDF] ✅ Gerado com WeasyPrint: {filename}", file=sys.stderr)

    except Exception as e:
        print(f"[PDF] ⚠️ WeasyPrint falhou ({e}), usando ReportLab...", file=sys.stderr)

    # ── Fallback: ReportLab ───────────────────────────────────────
    if not weasyprint_ok:
        _gerar_pdf_reportlab(d, filepath)
        print(f"[PDF] ✅ Gerado com ReportLab: {filename}", file=sys.stderr)

    print(json.dumps({"type":"generated_file","path":filepath,"name":filename}))
    return filename


# ── MAIN ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    dados = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {
        "num_projeto": "PRJ-2026-001", "cidade": "Londrina / PR",
        "data_bdo": "29/05/2026", "hora_inicio": "07:30", "hora_final": "17:00",
        "encarregado": "Carlos Eduardo Silva",
        "servico_livre": False, "num_omb": "OMB-45678", "num_plv": "PLV-12345",
        "manobra_texto": "Feito abertura de chaves e GLV pela equipe executora.",
        "alteracoes_texto": "Executado conforme projeto.",
        "status_execucao": "conclusao", "status_motivo": "",
        "modo_postes": True, "postes": ["PS1","PS2"],
        "servicos_completo": [
            {"nome":"🏗️ Reforma de calçada (2 pontos)","postes":["PS1"]},
            {"nome":"🌳 Podas","postes":["PS1","PS2"]},
            {"nome":"⚡ Malha de aterramento — 35m","postes":["geral"]},
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
