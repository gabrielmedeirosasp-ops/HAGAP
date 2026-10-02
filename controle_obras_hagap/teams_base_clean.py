from __future__ import annotations

import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


ROOT = Path(r"C:\HAGAP\TEAMS_SYNC")
OUT = Path(__file__).resolve().parent / "base_teams.json"
REPORT = Path(__file__).resolve().parent / "base_teams_relatorio.txt"

RE_PROJECT = re.compile(r"(?<!\d)(\d{7})(?!\d)")
RE_DATE = re.compile(r"\b(\d{2}/\d{2}/\d{4})\b")
RE_AES_FILE = re.compile(r"\bAES\s*0*(\d{4,6})\b", re.I)


def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()


def all_pdfs_under(root: Path, folder_match) -> list[Path]:
    if not root.exists():
        return []
    out = []
    for p in root.rglob("*.pdf"):
        rel = norm(str(p.relative_to(root)))
        if folder_match(rel):
            out.append(p)
    return sorted(set(out))


def text_pdf(path: Path) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF não instalado.")
    try:
        doc = fitz.open(path)
        return "\n".join(page.get_text("text") or "" for page in doc)
    except Exception:
        return ""


def tables_pdf(path: Path):
    if pdfplumber is None:
        return []
    out = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    for table in page.extract_tables() or []:
                        if table:
                            out.append(table)
                except Exception:
                    continue
    except Exception:
        return []
    return out


def valid_date_br(s: str) -> bool:
    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", s or "")
    if not m:
        return False
    d, mo, y = map(int, m.groups())
    if not (1 <= mo <= 12 and 1 <= d <= 31 and 2000 <= y <= 2100):
        return False
    import datetime as dt
    try:
        dt.date(y, mo, d)
        return True
    except ValueError:
        return False


def aes_number(path: Path, text: str) -> str:
    m = RE_AES_FILE.search(path.stem)
    if m:
        return m.group(1)
    # Só aceita número quando houver rótulo explícito.
    for pat in [
        r"\bNUMERO\s*[:\-]?\s*0*(\d{4,6})\b",
        r"\bAES\s*(?:N[ºO°.]*)?\s*[:\-]?\s*0*(\d{4,6})\b",
    ]:
        m = re.search(pat, norm(text))
        if m:
            return m.group(1)
    return ""


def parse_aes_rows(path: Path):
    text = text_pdf(path)
    aes = aes_number(path, text)
    found = []

    # Prioridade: tabela, usando a coluna PRAZO da mesma linha do PROJ.
    for table in tables_pdf(path):
        header_idx = None
        proj_col = None
        prazo_col = None

        for ri, row in enumerate(table):
            cells = [norm(c) for c in (row or [])]
            for ci, cell in enumerate(cells):
                if "PROJ" in cell and proj_col is None:
                    proj_col = ci
                if cell == "PRAZO" or " PRAZO" in (" " + cell):
                    prazo_col = ci
            if proj_col is not None and prazo_col is not None:
                header_idx = ri
                break

        if header_idx is not None:
            for row in table[header_idx + 1:]:
                cells = ["" if c is None else str(c) for c in (row or [])]
                if proj_col >= len(cells) or prazo_col >= len(cells):
                    continue
                projs = RE_PROJECT.findall(cells[proj_col] or "")
                dates = RE_DATE.findall(cells[prazo_col] or "")
                if len(projs) == 1 and len(dates) == 1 and valid_date_br(dates[0]):
                    found.append((projs[0], dates[0], aes, "TABELA:PRAZO/PROJ"))
            continue

        # Fallback controlado: mesma linha precisa ter exatamente 1 projeto e 1 data.
        for row in table:
            linha = " | ".join("" if c is None else str(c) for c in (row or []))
            projs = RE_PROJECT.findall(linha)
            dates = [d for d in RE_DATE.findall(linha) if valid_date_br(d)]
            if len(projs) == 1 and len(dates) == 1:
                found.append((projs[0], dates[0], aes, "TABELA:FALLBACK_MESMA_LINHA"))

    # Deduplica o mesmo achado.
    out = []
    seen = set()
    for item in found:
        k = item[:3]
        if k not in seen:
            seen.add(k)
            out.append(item)
    return out


def municipio_from_text(text: str) -> str:
    if not text:
        return ""

    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines()]
    for i, raw in enumerate(lines):
        if "MUNICIPIO DA OBRA" not in norm(raw):
            continue

        candidatos = []
        resto = re.sub(r"(?i).*MUNIC[IÍ]PIO\s+DA\s+OBRA\s*", "", raw).strip(" :-")
        if resto:
            candidatos.append(resto)
        candidatos += [x for x in lines[i + 1:i + 4] if x]

        for c in candidatos:
            # Ex.: PR 4106605 - CRUZEIRO DO OESTE
            m = re.search(r"\b(?:PR\s*)?\d{6,8}\s*-\s*(.+)$", c, re.I)
            if m:
                city = re.sub(r"\s+", " ", m.group(1)).strip(" -")
                if city:
                    return city.upper()

            # Aceita texto puro somente imediatamente após o rótulo e sem números.
            cc = re.sub(r"\s+", " ", c).strip(" -")
            if cc and not re.search(r"\d", cc) and 2 <= len(cc) <= 80:
                return cc.upper()

    return ""


def project_from_text(text: str) -> str:
    t = norm(text)
    # Só aceita projeto quando houver rótulo explícito.
    patterns = [
        r"\bPROJETO\s*[:\-]?\s*(\d{7})\b",
        r"\bN[ºO°.]?\s*PROJETO\s*[:\-]?\s*(\d{7})\b",
        r"\bPROJ[.]?\s*[:\-]?\s*(\d{7})\b",
    ]
    achados = []
    for pat in patterns:
        achados += re.findall(pat, t)
    unicos = []
    for x in achados:
        if x not in unicos:
            unicos.append(x)
    return unicos[0] if len(unicos) == 1 else ""


def project_from_path(path: Path) -> str:
    # Fallback apenas quando o caminho inteiro contém UM único número de projeto.
    candidates = RE_PROJECT.findall(str(path))
    uniq = []
    for x in candidates:
        if x not in uniq:
            uniq.append(x)
    return uniq[0] if len(uniq) == 1 else ""


def main():
    if not ROOT.exists():
        print(f"[PENDENTE] Pasta Teams não localizada: {ROOT}")
        print("Execute primeiro a sincronização do Teams/SharePoint.")
        sys.exit(2)

    projects = defaultdict(lambda: {
        "aes_rows": [],
        "municipios": [],
        "arquivosProjeto": [],
        "bmds": [],
        "ffos": [],
        "fontes": [],
        "divergencias": [],
    })

    # 1) AES EMITIDA — prazo deve vir da MESMA LINHA do projeto.
    aes_pdfs = all_pdfs_under(ROOT, lambda rel: "AES EMITIDA" in rel)
    for pdf in aes_pdfs:
        rows = parse_aes_rows(pdf)
        for projeto, prazo, aes, metodo in rows:
            p = projects[projeto]
            p["aes_rows"].append({
                "aes": aes,
                "prazo": prazo,
                "arquivo": pdf.name,
                "origem": str(pdf.relative_to(ROOT)),
                "metodo": metodo,
            })
            p["fontes"].append(str(pdf.relative_to(ROOT)))

    # 2) OBRAS PARA EXECUCAO — município SOMENTE do campo MUNICIPIO DA OBRA.
    proj_pdfs = all_pdfs_under(ROOT, lambda rel: "OBRAS PARA EXECUCAO" in rel)
    for pdf in proj_pdfs:
        texto = text_pdf(pdf)
        projeto = project_from_text(texto) or project_from_path(pdf)
        if not projeto:
            continue

        p = projects[projeto]
        p["arquivosProjeto"].append({"nome": pdf.name, "url": ""})
        p["fontes"].append(str(pdf.relative_to(ROOT)))

        city = municipio_from_text(texto)
        if city:
            p["municipios"].append({"municipio": city, "arquivo": pdf.name})
        else:
            p["divergencias"].append(
                "[NÃO LOCALIZADO] MUNICÍPIO DA OBRA não encontrado em " + pdf.name
            )

    # 3) BMD/FFO — associação somente quando o nome/caminho contém UM projeto.
    med_pdfs = all_pdfs_under(
        ROOT,
        lambda rel: ("BMD" in rel or "FFO" in rel or "FF0" in rel)
    )
    for pdf in med_pdfs:
        projeto = project_from_path(pdf)
        if not projeto:
            continue
        nome_n = norm(pdf.name)
        item = {"arquivo": pdf.name, "url": "", "origem": str(pdf.relative_to(ROOT))}
        if re.search(r"\bBMD\b", nome_n):
            if not any(x["origem"] == item["origem"] for x in projects[projeto]["bmds"]):
                projects[projeto]["bmds"].append(item)
                projects[projeto]["fontes"].append(item["origem"])
        if re.search(r"\bFFO\b|\bFF0\b", nome_n):
            if not any(x["origem"] == item["origem"] for x in projects[projeto]["ffos"]):
                projects[projeto]["ffos"].append(item)
                projects[projeto]["fontes"].append(item["origem"])

    output = []
    report = [
        "HAGAP — BASE TEAMS LIMPA",
        "Fonte: somente C:\\HAGAP\\TEAMS_SYNC (SharePoint/Teams).",
        "db.json/site HAGAP: NÃO UTILIZADO. Fonte técnica exclusiva: Teams/SharePoint.",
        "",
    ]

    for projeto in sorted(projects):
        p = projects[projeto]

        aes_values = sorted({x["aes"] for x in p["aes_rows"] if x["aes"]})
        prazo_values = sorted({x["prazo"] for x in p["aes_rows"] if x["prazo"]})
        mun_values = sorted({x["municipio"] for x in p["municipios"] if x["municipio"]})

        aes = aes_values[0] if len(aes_values) == 1 else (" / ".join(aes_values) if aes_values else "")
        prazo = prazo_values[0] if len(prazo_values) == 1 else ""
        municipio = mun_values[0] if len(mun_values) == 1 else ""

        div = list(p["divergencias"])
        if not p["aes_rows"]:
            div.append("[NÃO LOCALIZADO] Projeto não localizado em linha válida de AES.")
        if len(aes_values) > 1:
            div.append("[DIVERGÊNCIA] Mais de uma AES localizada: " + " / ".join(aes_values))
        if len(prazo_values) > 1:
            div.append("[DIVERGÊNCIA] Prazos diferentes na(s) AES: " + " / ".join(prazo_values))
        if not mun_values:
            div.append("[NÃO LOCALIZADO] Município não confirmado no campo MUNICÍPIO DA OBRA.")
        if len(mun_values) > 1:
            div.append("[DIVERGÊNCIA] Municípios diferentes em projetos: " + " / ".join(mun_values))

        output.append({
            "projeto": projeto,
            "aes": aes,
            "prazoAes": prazo,
            "municipio": municipio,
            "arquivoAes": p["aes_rows"][0]["arquivo"] if p["aes_rows"] else "",
            "arquivosProjeto": sorted(p["arquivosProjeto"], key=lambda x: x["nome"]),
            "bmds": p["bmds"],
            "ffos": p["ffos"],
            "fontes": sorted(set(p["fontes"])),
            "divergencias": div,
        })

        if div:
            report.append(projeto + ": " + " | ".join(div))

    OUT.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    REPORT.write_text("\n".join(report), encoding="utf-8")

    print("=" * 72)
    print("BASE TEAMS GERADA")
    print(f"Projetos: {len(output)}")
    print(f"AES PDFs lidos: {len(aes_pdfs)}")
    print(f"Projetos PDFs lidos: {len(proj_pdfs)}")
    print(f"BMD/FFO PDFs lidos: {len(med_pdfs)}")
    print(f"Saída: {OUT}")
    print(f"Relatório: {REPORT}")
    print("=" * 72)


if __name__ == "__main__":
    main()
