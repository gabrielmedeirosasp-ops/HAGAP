import re
import unicodedata
from datetime import datetime


PROJECT_RE = re.compile(r"(?<!\d)(\d{7})([A-Za-z])?(?!\d)")
REF_RE = re.compile(r"(?<!\d)(\d{1,5})\s*[/\-]\s*(20\d{2})(?!\d)")


def clean_text(value):
    return " ".join(str(value or "").replace("\xa0", " ").split())


def norm(value):
    txt = unicodedata.normalize("NFKD", str(value or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return clean_text(txt).upper()


def extract_projects(*values):
    found = []
    seen = set()
    for value in values:
        text = str(value or "")
        for m in PROJECT_RE.finditer(text):
            project = m.group(1) + (m.group(2).upper() if m.group(2) else "")
            if project not in seen:
                seen.add(project)
                found.append(project)
    return found


def extract_projects_from_pdf(pdf_text):
    """
    Extrai projeto de PDF com prioridade para campos rotulados.
    Se o documento não trouxer um campo claro e houver mais de um número
    de 7 dígitos, NÃO escolhe silenciosamente.
    """
    text = str(pdf_text or "")

    patterns = [
        r"N[º°]\s*Projeto\s*[:.]?\s*(\d{7}[A-Za-z]?)",
        r"N[º°]\s*PROJETO\s*[:.]?\s*(\d{7}[A-Za-z]?)",
        r"FISCAL\s+N[º°]\s*PROJETO[^\n]*\n[^\n]*?\b(\d{7}[A-Za-z]?)\b",
        r"N[º°]\s*Projeto[^\n]*\n(?:[^\n]*\n){0,4}?[^\n]*?\b(\d{7}[A-Za-z]?)\b",
    ]

    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            return [m.group(1).upper()]

    candidates = extract_projects(text)
    if len(candidates) == 1:
        return candidates

    return []


def classify_subject(subject):
    s = norm(subject)

    if "DOCUMENTOS PARCIAL" in s or "DOCUMENTO PARCIAL" in s:
        return "DOC_PARCIAL"
    if "DOCUMENTOS" in s or s.startswith("DOCUMENTO "):
        return "DOC_FINAL"
    if "PDE NUMERO:" in s or s.startswith("PDE "):
        return "PDE"
    if "PLV NUMERO:" in s or s.startswith("PLV "):
        return "PLV"
    if "OMB ENVIADO INTEGRACAO" in s:
        return "OMB"

    has_bmd = bool(re.search(r"\bBMD\b", s))
    has_ffo = bool(re.search(r"\bFFO\b|\bFF0\b", s))
    if has_bmd and has_ffo:
        return "BMD_FFO"
    if has_ffo:
        return "FFO"
    if has_bmd:
        return "BMD"

    return None


def parse_status_from_subject(subject, prefix):
    s = clean_text(subject)
    pat = re.compile(
        rf"{re.escape(prefix)}\s*(?:n[uú]mero\s*:)?\s*\d+\s*[/\-]\s*\d{{4}}\s*-\s*([^-]+)",
        re.I,
    )
    m = pat.search(s)
    return clean_text(m.group(1)) if m else ""


def parse_ref(text, label):
    patterns = [
        rf"\b{re.escape(label)}\s*(?:n[º°o.]*)?\s*[:\-]?\s*0*(\d{{1,5}})\s*[/\-]\s*(20\d{{2}})",
        rf"\b{re.escape(label)}\s+0*(\d{{1,5}})\s*[/\-]\s*(20\d{{2}})",
    ]
    for p in patterns:
        m = re.search(p, str(text or ""), re.I)
        if m:
            return f"{int(m.group(1))}/{m.group(2)}"
    return ""


def parse_municipality_from_subject(subject):
    s = clean_text(subject)
    parts = [clean_text(x) for x in s.split(" - ")]

    ignorar = {
        "ABERTO", "APROVADO", "CANCELADO", "REPROVADO",
        "ENCERRADO", "FECHADO",
    }

    for candidate in parts[1:]:
        n = norm(candidate)
        if not candidate:
            continue
        if n in ignorar:
            continue
        if "RESP" in n or "ENVIADO POR" in n:
            continue
        return candidate

    return ""


def parse_date_value(text, label):
    patterns = [
        rf"{re.escape(label)}[^\d]*(\d{{2}}/\d{{2}}/\d{{4}})",
        rf"{re.escape(label)}[^\d]*(\d{{2}}\.\d{{2}}\.\d{{4}})",
    ]
    for p in patterns:
        m = re.search(p, str(text or ""), re.I)
        if m:
            return m.group(1).replace(".", "/")
    return ""


def parse_period(text):
    m = re.search(
        r"(\d{2}/\d{2}/\d{4}).{0,60}?(\d{1,2}:\d{2}).{0,30}?(?:ÀS|AS|às|as)\s*(\d{1,2}:\d{2})",
        str(text or ""),
        re.I | re.S,
    )
    if not m:
        return {}
    return {
        "data": m.group(1),
        "inicio": m.group(2),
        "fim": m.group(3),
    }


def parse_pde(subject, pdf_text):
    ref = parse_ref(subject, "PDE") or parse_ref(pdf_text, "PDE")
    projects = extract_projects_from_pdf(pdf_text)

    municipio = parse_municipality_from_subject(subject)
    if not municipio:
        m = re.search(
            r"C[oó]d\.\s*/\s*Localidade\s*/\s*Munic[ií]pio\s*:\s*[^\n]*/\s*([^\n]+)",
            pdf_text or "",
            re.I,
        )
        if m:
            municipio = clean_text(m.group(1))

    return {
        "kind": "PDE",
        "ref_number": ref,
        "projects": projects,
        "status": parse_status_from_subject(subject, "PDE"),
        "municipality": municipio,
        "details": {
            "data_solicitada": parse_date_value(pdf_text, "Data Solicitada"),
            "data_confirmada": parse_date_value(pdf_text, "Data Confirmada"),
        },
    }


def parse_plv(subject, pdf_text):
    ref = parse_ref(subject, "PLV") or parse_ref(pdf_text, "PLV")
    projects = extract_projects_from_pdf(pdf_text)
    return {
        "kind": "PLV",
        "ref_number": ref,
        "projects": projects,
        "status": parse_status_from_subject(subject, "PLV"),
        "municipality": parse_municipality_from_subject(subject),
        "details": {
            "data_solicitada": parse_date_value(pdf_text, "Data Solicitada"),
            "data_confirmada": parse_date_value(pdf_text, "Data Confirmada"),
        },
    }


def parse_omb(subject, pdf_text):
    ref = parse_ref(subject, "OMB") or parse_ref(pdf_text, "OMB")
    pde_ref = parse_ref(pdf_text, "PDE")
    return {
        "kind": "OMB",
        "ref_number": ref,
        "related_ref": pde_ref,
        "projects": [],
        "status": "Emitida",
        "municipality": parse_municipality_from_subject(subject),
        "details": parse_period(pdf_text),
    }


def parse_measurement(kind, subject, pdf_text, attachment_name=""):
    # Em BMD/FFO, assunto/nome do anexo têm prioridade porque o PDF também
    # contém outros números de 7 dígitos (ex.: número de fornecedor).
    projects = extract_projects(attachment_name)
    if not projects:
        projects = extract_projects(subject)
    if not projects:
        projects = extract_projects_from_pdf(pdf_text)

    upper = norm(subject + " " + attachment_name + " " + (pdf_text or "")[:2000])
    details = {
        "parcial": "PARCIAL" in upper,
        "final": "FINAL" in upper,
    }
    return {
        "kind": kind,
        "projects": projects,
        "status": "Parcial" if details["parcial"] else ("Final" if details["final"] else ""),
        "municipality": "",
        "details": details,
    }


def parse_documents(kind, subject, body, attachment_names):
    projects = extract_projects(subject, body, " ".join(attachment_names or []))
    return {
        "kind": kind,
        "projects": projects,
        "status": "Parcial" if kind == "DOC_PARCIAL" else "Enviado",
        "municipality": "",
        "details": {},
    }
