from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
import cgi
import json
import os
import mimetypes
import csv
import subprocess
import sys
import traceback
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
import unicodedata

import win32com.client


HOST = "127.0.0.1"
PORT = 8765
DATA_DIR = Path(__file__).with_name("data")
REGISTRY_PATH = DATA_DIR / "projects_registry.json"
REGISTRY_CSV_PATH = DATA_DIR / "projects_registry.csv"
OUTBOX_DIR = DATA_DIR / "outbox"
PROJECTS_DIR = DATA_DIR / "projects"
WORKER_PATH = Path(__file__).with_name("outlook_send_worker.py")


def send_via_outlook(to_address: str, subject: str, body: str, attachments: list[Path]) -> None:
    """Run Outlook automation in a dedicated worker process (fresh interpreter).
    This avoids freezing the HTTP server when Outlook/COM hangs.
    """
    if not WORKER_PATH.exists():
        raise RuntimeError(f"Worker manquant: {WORKER_PATH.name}")

    payload = {
        "to": to_address,
        "subject": subject,
        "body": body,
        "attachments": [str(a) for a in attachments],
    }

    try:
        completed = subprocess.run(
            [sys.executable, str(WORKER_PATH)],
            input=json.dumps(payload).encode("utf-8"),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=25,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "Outlook ne repond pas (timeout). Fermez Outlook completement puis relancez-le (pas en admin), puis reessayez."
        ) from exc

    if completed.returncode == 0:
        return

    err = (completed.stdout.decode("utf-8", errors="ignore") + "\n" + completed.stderr.decode("utf-8", errors="ignore")).strip()
    err = err or f"Erreur Outlook (code {completed.returncode})"

    if (
        "-2146959355" in err
        or "Échec de l’exécution du serveur" in err
        or "Echec de l'execution du serveur" in err
        or "Echec de l’execution du serveur" in err
    ):
        raise RuntimeError(
            "Outlook ne peut pas etre automatise sur ce PC (echec COM).\n"
            "Actions recommandees :\n"
            "1) Fermez Outlook completement (toutes les fenetres), puis relancez Outlook normalement (pas en administrateur).\n"
            "2) Reessayez.\n"
            "3) Si ca bloque encore : redemarrez Windows.\n"
        )

    raise RuntimeError(err)


def _ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _load_registry() -> list[dict]:
    _ensure_data_dir()
    if not REGISTRY_PATH.exists():
        return []
    try:
        raw = REGISTRY_PATH.read_text(encoding="utf-8")
        parsed = json.loads(raw) if raw.strip() else []
        return parsed if isinstance(parsed, list) else []
    except Exception:
        return []


def _save_registry(items: list[dict]) -> None:
    _ensure_data_dir()
    REGISTRY_PATH.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")

    # Also keep a CSV beside the JSON so Excel users can open it directly.
    headers = ["code","created_at","updated_at","source","titre","porteur","email","unite","mr004","cmt","zone"]
    with REGISTRY_CSV_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for r in items:
            w.writerow({h: r.get(h, "") for h in headers})


def _upsert_registry_row(row: dict) -> None:
    code = str(row.get("code") or "").strip()
    if not code:
        return
    items = _load_registry()
    idx = next((i for i, r in enumerate(items) if str(r.get("code") or "") == code), -1)
    if idx >= 0:
        items[idx] = row
    else:
        items.insert(0, row)
    _save_registry(items)


W_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}


def _iter_sdts(root: ET.Element):
    for sdt in root.iterfind(".//w:sdt", W_NS):
        yield sdt


def _sdt_text(sdt: ET.Element) -> str:
    texts = []
    for t in sdt.iterfind(".//w:sdtContent//w:t", W_NS):
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def _set_sdt_text(sdt: ET.Element, value: str) -> None:
    content = sdt.find("w:sdtContent", W_NS)
    if content is None:
        return
    ts = list(content.iterfind(".//w:t", W_NS))
    if not ts:
        return
    ts[0].text = value
    for t in ts[1:]:
        t.text = ""


def _set_checkbox(sdt: ET.Element, checked: bool) -> None:
    pr = sdt.find("w:sdtPr", W_NS)
    if pr is None:
        return
    cb = pr.find("w14:checkbox", W_NS)
    if cb is None:
        return
    checked_el = cb.find("w14:checked", W_NS)
    if checked_el is None:
        checked_el = ET.SubElement(cb, f"{{{W_NS['w14']}}}checked")
    checked_el.set(f"{{{W_NS['w14']}}}val", "1" if checked else "0")


def _fill_by_placeholder_sequence(doc_root: ET.Element, values: list[str]) -> None:
    # Replace SDT placeholders in order (first N occurrences of the standard Word placeholder text)
    placeholder = "Cliquez ou appuyez ici pour entrer du texte."
    sdts = [s for s in _iter_sdts(doc_root) if placeholder in _sdt_text(s)]
    for i, v in enumerate(values):
        if i >= len(sdts):
            break
        if not str(v or "").strip():
            continue
        _set_sdt_text(sdts[i], v)


def _checkbox_sdts(doc_root: ET.Element) -> list[ET.Element]:
    return [sdt for sdt in _iter_sdts(doc_root) if sdt.find("w:sdtPr/w14:checkbox", W_NS) is not None]


def _set_checkbox_by_position(doc_root: ET.Element, position: int, checked: bool) -> None:
    checkboxes = _checkbox_sdts(doc_root)
    if 0 <= position < len(checkboxes):
        _set_checkbox(checkboxes[position], checked)


def _text(value: object) -> str:
    return str(value or "").strip()


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", _text(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().split())


def _yes_no(flag: object) -> str:
    return "Oui" if bool(flag) else "Non"


def _join_non_empty(*values: object, sep: str = " - ") -> str:
    return sep.join(_text(value) for value in values if _text(value))


def _data_categories(data: dict) -> str:
    labels = [
        ("mr_data_pathology", "Description de la pathologie"),
        ("mr_data_treatments", "Description des traitements"),
        ("mr_data_genetics_somatic", "Donnees genetiques somatiques"),
        ("mr_data_genetics_germline", "Donnees genetiques constitutionnelles"),
        ("mr_data_imaging", "Donnees d'imagerie"),
        ("mr_data_slides", "Lames virtuelles d'anapath"),
        ("mr_data_samples", "Echantillons biologiques humains"),
        ("mr_data_pgeb_contacted", "PGEB contactee"),
        ("mr_data_social", "Donnees sociales ou mode de vie"),
    ]
    parts = [label for key, label in labels if data.get(key)]
    if data.get("mr_data_health_other"):
        parts.append(_text(data.get("mr_data_health_other_txt")) or "Autres donnees de sante")
    if data.get("mr_data_nonhealth_other"):
        parts.append(_text(data.get("mr_data_nonhealth_other_txt")) or "Autres donnees hors sante")
    if _text(data.get("mr_data_schema")):
        parts.append(f"Schema general: {_text(data.get('mr_data_schema'))}")
    return "; ".join(parts)


def _population_categories(data: dict) -> str:
    parts = []
    if data.get("mr_population_patients"):
        parts.append("Patients")
    if data.get("mr_population_aidants"):
        parts.append("Aidants")
    if data.get("mr_population_pros"):
        parts.append("Professionnels de sante")
    if data.get("mr_population_other"):
        parts.append(_text(data.get("mr_population_other_txt")) or "Autres")
    return "; ".join(parts)


def _tumor_types(data: dict) -> str:
    labels = [
        ("mr_tumor_all", "Tout cancer"),
        ("mr_tumor_brain", "Tumeur du cerveau"),
        ("mr_tumor_colorectal", "Tumeur colorectale"),
        ("mr_tumor_stomach", "Tumeur de l'estomac"),
        ("mr_tumor_liver", "Tumeur du foie"),
        ("mr_tumor_small_intestine", "Tumeur de l'intestin grele"),
        ("mr_tumor_eye", "Tumeur de l'oeil"),
        ("mr_tumor_orl", "Tumeur ORL / VADS"),
        ("mr_tumor_bone", "Tumeur de l'os / cartilage"),
        ("mr_tumor_ovary", "Tumeur de l'ovaire"),
        ("mr_tumor_pancreas", "Tumeur du pancreas"),
        ("mr_tumor_skin", "Tumeur de la peau"),
        ("mr_tumor_pleura", "Tumeur de la plevre"),
        ("mr_tumor_lung", "Tumeur du poumon"),
        ("mr_tumor_prostate", "Tumeur de la prostate"),
        ("mr_tumor_kidney", "Tumeur du rein"),
        ("mr_tumor_hematology", "Tumeur hematologique"),
        ("mr_tumor_breast", "Tumeur du sein"),
        ("mr_tumor_testicle", "Tumeur du testicule"),
        ("mr_tumor_thyroid", "Tumeur thyroide / endocrine"),
        ("mr_tumor_uterus", "Tumeur col / endometre"),
        ("mr_tumor_bladder", "Tumeur de la vessie"),
        ("mr_tumor_soft_tissue", "Tumeur des tissus mous"),
        ("mr_tumor_other_solid", "Autres tumeurs solides"),
        ("mr_tumor_unknown_primary", "Site primitif inconnu"),
    ]
    parts = [label for key, label in labels if data.get(key)]
    if data.get("mr_tumor_other"):
        parts.append(_text(data.get("mr_tumor_other_txt")) or "Autre tumeur")
    if not parts and _text(data.get("mr_tumor_type")):
        parts.append(_text(data.get("mr_tumor_other_txt")) if data.get("mr_tumor_type") == "other" else _text(data.get("mr_tumor_type")))
    return "; ".join(parts)


def _mr_case_details(data: dict) -> list[str]:
    case = _text(data.get("mr_case"))
    details: list[str] = []
    if case == "other":
        details.append(_text(data.get("mr_case_other_detail")))
    if case == "1":
        details.extend([
            _text(data.get("mr_case1_internal_teams")),
            _text(data.get("mr_case1_subcontractors")),
            _text(data.get("mr_case1_stats_team")),
        ])
    if case == "2":
        details.extend([
            _text(data.get("mr_case2_centers")),
            _text(data.get("mr_case2_internal_teams")),
            _text(data.get("mr_case2_subcontractors")),
            _text(data.get("mr_case2_stats_team")),
        ])
    if case in {"3", "4", "5", "6"}:
        details.extend([
            _text(data.get("mr_case36_internal_teams")),
            _text(data.get("mr_case_host_country")),
        ])
    details.extend([
        _text(data.get("mr_flow_other")),
        _text(data.get("mr_case_collection_responsible")),
        _text(data.get("mr_case_collection_tool")),
        _text(data.get("mr_contract_drafted")),
        _text(data.get("mr_contract_validated_by")),
    ])
    return [value for value in details if value]


def _information_mode(data: dict) -> str:
    parts = []
    if data.get("mr_info_patient_level5"):
        parts.append("Patients CLB: note d'information aux patients ayant un niveau d'information < 5")
    if data.get("mr_info_patient_notice"):
        parts.append("Remise de la note d'information a chaque patient")
    if data.get("mr_info_patient_other"):
        parts.append(_text(data.get("mr_info_patient_other_txt")) or "Autres methodes patients")
    if data.get("mr_info_nonpatient_oral"):
        parts.append("Information orale")
    if data.get("mr_info_nonpatient_written"):
        parts.append("Notice ecrite individuelle")
    if data.get("mr_info_nonpatient_other"):
        parts.append(_text(data.get("mr_info_nonpatient_other_txt")) or "Autres methodes hors patients")
    return "; ".join(parts)


def _ethics_need(data: dict) -> str:
    parts = []
    if data.get("mr_ethics_certificate"):
        parts.append("Demande de certification d'instruction RGPD FR/EN")
    if data.get("mr_ethics_opinion"):
        topic = _text(data.get("mr_ethics_topic"))
        parts.append(f"Avis ethique CMT: {topic}" if topic else "Avis ethique CMT demande")
    return "; ".join(parts)


def _cmt_platforms(data: dict) -> str:
    labels = [
        ("cmt_pf_biopath", "BIOPATH"),
        ("cmt_pf_par", "PAR"),
        ("cmt_pf_pgeb", "PGEB"),
        ("cmt_pf_onco3d", "Onco-3D"),
        ("cmt_pf_pgc", "PGC"),
        ("cmt_pf_pgt", "PGT"),
        ("cmt_pf_licl", "LICL"),
        ("cmt_pf_pathec", "PATHEC"),
    ]
    parts = [label for key, label in labels if data.get(key)]
    if data.get("cmt_pf_other"):
        parts.append(_text(data.get("cmt_pf_other_txt")) or "Autre")
    if _text(data.get("cmt_platform_details")):
        parts.append(_text(data.get("cmt_platform_details")))
    return "; ".join(parts)


def _first_tc_text(tc: ET.Element) -> str:
    return " ".join(
        "".join(t.text or "" for t in p.iterfind(".//w:t", W_NS)).strip()
        for p in tc.iterfind("w:p", W_NS)
    ).strip()


def _clear_and_set_cell_text(tc: ET.Element, value: str) -> None:
    value = _text(value)
    paragraphs = list(tc.iterfind("w:p", W_NS))
    if not paragraphs:
        p = ET.SubElement(tc, f"{{{W_NS['w']}}}p")
    else:
        p = paragraphs[0]
    for child in list(p):
        p.remove(child)
    run = ET.SubElement(p, f"{{{W_NS['w']}}}r")
    text_el = ET.SubElement(run, f"{{{W_NS['w']}}}t")
    text_el.text = value
    for extra in paragraphs[1:]:
        tc.remove(extra)


def _set_table_value_by_label(doc_root: ET.Element, label_substr: str, value: str) -> bool:
    if not _text(value):
        return False
    needle = _norm(label_substr)
    for tr in doc_root.iterfind(".//w:tr", W_NS):
        cells = list(tr.iterfind("w:tc", W_NS))
        for idx, tc in enumerate(cells[:-1]):
            if needle in _norm(_first_tc_text(tc)):
                _clear_and_set_cell_text(cells[idx + 1], value)
                return True
    return False


def _set_first_empty_row_cells(doc_root: ET.Element, header_substr: str, values: list[str]) -> bool:
    needle = _norm(header_substr)
    for tbl in doc_root.iterfind(".//w:tbl", W_NS):
        rows = list(tbl.iterfind("w:tr", W_NS))
        if not rows:
            continue
        header = _norm(" ".join(_first_tc_text(tc) for tc in rows[0].iterfind("w:tc", W_NS)))
        if needle not in header:
            continue
        if len(rows) < 2:
            return False
        cells = list(rows[1].iterfind("w:tc", W_NS))
        for idx, value in enumerate(values[: len(cells)]):
            if _text(value):
                _clear_and_set_cell_text(cells[idx], value)
        return True
    return False


def _set_single_cell_after_label(doc_root: ET.Element, label_substr: str, value: str) -> bool:
    if not _text(value):
        return False
    needle = _norm(label_substr)
    for tbl in doc_root.iterfind(".//w:tbl", W_NS):
        cells = list(tbl.iterfind(".//w:tc", W_NS))
        for idx, tc in enumerate(cells):
            if needle in _norm(_first_tc_text(tc)):
                target = cells[idx + 1] if idx + 1 < len(cells) else tc
                if target is tc:
                    _clear_and_set_cell_text(target, f"{_first_tc_text(tc)}\n{value}")
                else:
                    _clear_and_set_cell_text(target, value)
                return True
    return False


def _paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.iterfind(".//w:t", W_NS)).strip()


def _new_text_paragraph(value: str) -> ET.Element:
    p = ET.Element(f"{{{W_NS['w']}}}p")
    run = ET.SubElement(p, f"{{{W_NS['w']}}}r")
    text_el = ET.SubElement(run, f"{{{W_NS['w']}}}t")
    text_el.text = _text(value)
    return p


def _insert_paragraph_after_label(doc_root: ET.Element, label_substr: str, value: str) -> bool:
    if not _text(value):
        return False
    body = doc_root.find("w:body", W_NS)
    if body is None:
        return False
    needle = _norm(label_substr)
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag != f"{{{W_NS['w']}}}p":
            continue
        if needle in _norm(_paragraph_text(child)):
            body.insert(idx + 1, _new_text_paragraph(value))
            return True
    return False


def _set_platform_rows(doc_root: ET.Element, data: dict) -> None:
    platform_map = {
        "BIOPATH": data.get("cmt_pf_biopath"),
        "PAR": data.get("cmt_pf_par"),
        "PGEB": data.get("cmt_pf_pgeb"),
        "Onco-3D": data.get("cmt_pf_onco3d"),
        "PGC": data.get("cmt_pf_pgc"),
        "PGT": data.get("cmt_pf_pgt"),
        "LICL": data.get("cmt_pf_licl"),
        "PATHEC": data.get("cmt_pf_pathec"),
    }
    details = _text(data.get("cmt_platform_details"))
    other = _text(data.get("cmt_pf_other_txt"))
    for tbl in doc_root.iterfind(".//w:tbl", W_NS):
        rows = list(tbl.iterfind("w:tr", W_NS))
        if not rows:
            continue
        header = _norm(" ".join(_first_tc_text(tc) for tc in rows[0].iterfind("w:tc", W_NS)))
        if "sollicitation" not in header or "contact" not in header:
            continue
        for tr in rows:
            cells = list(tr.iterfind("w:tc", W_NS))
            if len(cells) < 4:
                continue
            label = _first_tc_text(cells[0])
            for platform, checked in platform_map.items():
                if _norm(platform) == _norm(label):
                    _clear_and_set_cell_text(cells[1], _yes_no(checked))
                    if checked and details:
                        _clear_and_set_cell_text(cells[2], details)
            if "autre" in _norm(label) and (data.get("cmt_pf_other") or other):
                _clear_and_set_cell_text(cells[0], f"autre: {other or 'Autre'}")
                _clear_and_set_cell_text(cells[1], "Oui")
                if details:
                    _clear_and_set_cell_text(cells[2], details)
        return


def _container_text(el: ET.Element) -> str:
    parts: list[str] = []
    for t in el.iterfind(".//w:t", W_NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _fill_first_placeholder_in_container(doc_root: ET.Element, label_substr: str, value: str) -> bool:
    """
    The provided DOCX templates don't expose stable SDT tags/aliases.
    Best effort: find a paragraph or table cell whose text contains `label_substr`
    and fill the first SDT placeholder in that same container.
    """
    if not value:
        return False
    placeholder = "Cliquez ou appuyez ici pour entrer du texte."
    needle = (label_substr or "").strip().lower()
    if not needle:
        return False

    # Prefer table cells first (CMT is mostly table-based)
    containers = list(doc_root.iterfind(".//w:tc", W_NS)) + list(doc_root.iterfind(".//w:p", W_NS))
    for c in containers:
        text = _container_text(c).lower()
        if needle not in text:
            continue
        for sdt in c.iterfind(".//w:sdt", W_NS):
            if placeholder in _sdt_text(sdt):
                _set_sdt_text(sdt, value)
                return True
    return False


def generate_mr004_docx(data: dict) -> Path:
    template_path = Path(__file__).with_name("MR004.docx")
    with zipfile.ZipFile(template_path, "r") as zin:
        xml_bytes = zin.read("word/document.xml")
        root = ET.fromstring(xml_bytes)

    # Minimal, deterministic filling: only shared "project identity" fields.
    porteur = _text(data.get("porteur"))
    email = _text(data.get("email"))
    unite = _text(data.get("unite"))
    titre = _text(data.get("titre"))
    resume = _text(data.get("resume"))

    porteur_line = porteur
    if unite:
        porteur_line = f"{porteur} ({unite})".strip()

    period_start = _text(data.get("date_debut"))
    period_end = _text(data.get("date_fin"))
    case_details = _mr_case_details(data)
    sensitive = _text(data.get("mr_sensitive"))
    sensitive_yes_no = "Oui" if sensitive == "yes" else ("Non" if sensitive == "no" else sensitive)

    # Fill text controls in the Word form order. Empty values are skipped so the original placeholder remains visible.
    _fill_by_placeholder_sequence(
        root,
        [
            date.today().strftime("%d/%m/%Y"),
            _text(data.get("q6_site_other_txt")),
            porteur_line,  # Entite(s) responsable(s) / porteur(s)
            _text(data.get("q6_resp_scientifique")) or porteur_line,
            _text(data.get("q6_contact")) or _join_non_empty(porteur, email),
            titre,  # Nom du projet
            _text(data.get("mr_eds_number")),
            resume[:400],  # Description grand public (approx)
            _text(data.get("q6_objective")) or resume,
            _text(data.get("mr_population_other_txt")),
            _text(data.get("mr_population_desc")),
            _text(data.get("mr_population_counts")),
            _tumor_types(data),
            period_start,
            period_end,
            _text(data.get("mr_case_other_detail")),
            _text(data.get("mr_case1_internal_teams")),
            _text(data.get("mr_case1_subcontractors")),
            _text(data.get("mr_case1_stats_team")),
            _text(data.get("mr_case2_centers")),
            _text(data.get("mr_case2_internal_teams")),
            _text(data.get("mr_case2_subcontractors")),
            _text(data.get("mr_case2_stats_team")),
            _text(data.get("mr_flow_other")),
            _text(data.get("mr_case36_internal_teams")),
            _text(data.get("mr_case_host_country")),
            _text(data.get("mr_case_collection_responsible")),
            _text(data.get("mr_case_collection_tool")),
            _text(data.get("mr_contract_drafted")),
            _text(data.get("mr_contract_validated_by")),
            _text(data.get("mr_legal_basis")),
            _text(data.get("mr_legal_basis_other")),
            _text(data.get("mr_data_health_other_txt")),
            _text(data.get("mr_data_nonhealth_other_txt")),
            _text(data.get("mr_data_schema")),
            sensitive_yes_no,
            _text(data.get("mr_sensitive_types")),
            _text(data.get("mr_sensitive_justification")),
            _text(data.get("mr_questionnaire_tool")),
            _text(data.get("mr_pseudonymisation")),
            _text(data.get("mr_identity_removal")),
            _text(data.get("mr_retention_detail")),
            _information_mode(data),
            _text(data.get("mr_info_patient_other_txt")),
            _text(data.get("mr_info_nonpatient_other_txt")),
            _ethics_need(data),
            _text(data.get("mr_ethics_topic")),
            "; ".join(case_details),
            _data_categories(data),
        ],
    )

    # First checkbox groups in the MR004 template: sites, population, case, flows and key declarations.
    for pos, flag in enumerate([data.get("q6_site_clb"), data.get("q6_site_ihope"), data.get("q6_site_other")]):
        _set_checkbox_by_position(root, pos, bool(flag))
    for offset, key in enumerate(["mr_population_patients", "mr_population_aidants", "mr_population_pros", "mr_population_other"], start=3):
        _set_checkbox_by_position(root, offset, bool(data.get(key)))
    case = _text(data.get("mr_case"))
    for idx, value in enumerate(["1", "2", "3", "4", "5", "6"], start=18):
        _set_checkbox_by_position(root, idx, case == value)
    for idx, key in enumerate(["mr_flow_ecrf", "mr_flow_owncloud", "mr_flow_mss"], start=63):
        _set_checkbox_by_position(root, idx, bool(data.get(key)))
    for idx, key in enumerate(["mr_info_patient_level5", "mr_info_patient_notice", "mr_info_patient_other"], start=100):
        _set_checkbox_by_position(root, idx, bool(data.get(key)))
    for idx, key in enumerate(["mr_info_nonpatient_oral", "mr_info_nonpatient_written", "mr_info_nonpatient_other"], start=103):
        _set_checkbox_by_position(root, idx, bool(data.get(key)))
    for idx, key in enumerate(["mr_ethics_certificate", "mr_ethics_opinion"], start=106):
        _set_checkbox_by_position(root, idx, bool(data.get(key)))

    # Write new docx to temp
    with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        out_path = Path(tmp.name)

    with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, ET.tostring(root, encoding="utf-8", xml_declaration=True))
            else:
                zout.writestr(item, zin.read(item.filename))

    return out_path


def generate_cmt_docx(data: dict) -> Path:
    template_path = Path(__file__).with_name("Fiche_CMT.docx")
    with zipfile.ZipFile(template_path, "r") as zin:
        xml_bytes = zin.read("word/document.xml")
        root = ET.fromstring(xml_bytes)

    porteur = _text(data.get("porteur"))
    email = _text(data.get("email"))
    unite = _text(data.get("unite"))
    titre = _text(data.get("titre"))
    partnerships = _text(data.get("cmt_partnerships")) or _text(data.get("trf_dest"))
    funding = _text(data.get("cmt_funding")) or _join_non_empty(data.get("aap_org"), data.get("fin_montant"), sep=" / ")
    cmt_summary = _text(data.get("cmt_summary")) or _text(data.get("resume"))
    cmt_objective = _text(data.get("cmt_data_objective")) or _text(data.get("q6_objective"))

    _set_table_value_by_label(root, "TITRE du projet", titre)
    _set_table_value_by_label(root, "Porteur", porteur)
    _set_table_value_by_label(root, "Laboratoire / Etablissement", unite)
    _set_table_value_by_label(root, "Coordonn", _join_non_empty(email, data.get("q6_contact")))
    _set_table_value_by_label(root, "Collaboration(s) / Partenariat(s)", partnerships)
    _set_table_value_by_label(root, "clinicien", _text(data.get("cmt_clinician")))
    _set_table_value_by_label(root, "Financement du projet", funding)

    _set_first_empty_row_cells(
        root,
        "Type d'echantillon",
        [
            _text(data.get("cmt_sample_type")),
            _text(data.get("cmt_sample_site")),
            _text(data.get("cmt_sample_pathology")),
            _text(data.get("cmt_sample_count")),
            _text(data.get("cmt_criteria_clinical")),
            _text(data.get("cmt_criteria_quantitative")),
            _text(data.get("cmt_criteria_quality_detail")) or _text(data.get("cmt_criteria_quality")),
            _join_non_empty(_yes_no(data.get("cmt_matching")), data.get("cmt_matching_detail")),
        ],
    )
    _insert_paragraph_after_label(root, "Autre precision utile", _text(data.get("cmt_selection_note")))
    _insert_paragraph_after_label(root, "Resume du projet", cmt_summary)
    _set_single_cell_after_label(root, "Lister des donnees associees", _text(data.get("cmt_data_list")))
    _set_single_cell_after_label(root, "objectif", cmt_objective)
    _set_single_cell_after_label(root, "Ethique", _text(data.get("cmt_ethics_impact")))
    _insert_paragraph_after_label(root, "Plateformes technologiques", _cmt_platforms(data))
    _set_platform_rows(root, data)

    with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        out_path = Path(tmp.name)

    with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, ET.tostring(root, encoding="utf-8", xml_declaration=True))
            else:
                zout.writestr(item, zin.read(item.filename))

    return out_path


class OutlookBridgeHandler(BaseHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        if self.path in ("/health", "/health/"):
            return self._send_json(200, {"ok": True})

        if self.path.startswith("/registry/"):
            return self._handle_registry_get()

        # Serve the static site in "local mode" so the browser is on http://127.0.0.1:8765
        return self._serve_static()

    def do_POST(self) -> None:
        if self.path == "/send-mail":
            return self._handle_send_mail()
        if self.path == "/send-docx":
            return self._handle_send_docx()
        if self.path == "/generate-docx":
            return self._handle_generate_docx()
        if self.path == "/registry/upsert":
            return self._handle_registry_upsert()
        self._send_json(404, {"error": "Route inconnue"})

    def _handle_registry_get(self) -> None:
        try:
            if self.path.startswith("/registry/list"):
                items = _load_registry()
                # return lightweight rows
                lite = []
                for r in items:
                    lite.append({
                        "code": r.get("code", ""),
                        "titre": r.get("titre", ""),
                        "porteur": r.get("porteur", ""),
                        "updated_at": r.get("updated_at", ""),
                        "mr004": bool(r.get("mr004")),
                        "cmt": bool(r.get("cmt")),
                    })
                return self._send_json(200, lite)

            if self.path.startswith("/registry/get"):
                # /registry/get?code=...
                from urllib.parse import urlparse, parse_qs
                q = parse_qs(urlparse(self.path).query)
                code = (q.get("code") or [""])[0].strip()
                if not code:
                    return self._send_json(400, {"error": "code manquant"})
                items = _load_registry()
                row = next((r for r in items if str(r.get("code") or "") == code), None)
                if not row:
                    return self._send_json(404, {"error": "introuvable"})
                return self._send_json(200, row)

            if self.path.startswith("/registry/export.csv"):
                return self._send_registry_csv()

            return self._send_json(404, {"error": "Route registre inconnue"})
        except Exception as exc:
            return self._send_json(500, {"error": str(exc), "trace": traceback.format_exc()})

    def _handle_registry_upsert(self) -> None:
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0") or "0"))
            payload = json.loads(raw.decode("utf-8"))
            if not isinstance(payload, dict):
                return self._send_json(400, {"error": "payload invalide"})
            _upsert_registry_row(payload)
            return self._send_json(200, {"ok": True})
        except Exception as exc:
            return self._send_json(500, {"error": str(exc), "trace": traceback.format_exc()})

    def _send_registry_csv(self) -> None:
        items = _load_registry()
        headers = ["code","created_at","updated_at","source","titre","porteur","email","unite","mr004","cmt","zone"]
        lines = [",".join(headers)]
        for r in items:
            row = []
            for h in headers:
                v = r.get(h, "")
                s = "" if v is None else str(v)
                row.append('"' + s.replace('"', '""') + '"')
            lines.append(",".join(row))
        data = ("\n".join(lines)).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", 'attachment; filename="registre_projets.csv"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_static(self) -> None:
        try:
            root = Path(__file__).resolve().parent
            path = self.path.split("?", 1)[0]
            if path in ("/", ""):
                path = "/index.html"

            safe = path.lstrip("/").replace("/", os.sep)
            file_path = (root / safe).resolve()

            # prevent directory traversal
            if root not in file_path.parents and file_path != root:
                self.send_response(403)
                self.end_headers()
                return

            if not file_path.exists() or not file_path.is_file():
                self.send_response(404)
                self.end_headers()
                return

            ctype, _ = mimetypes.guess_type(str(file_path))
            ctype = ctype or "application/octet-stream"
            data = file_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception:
            self.send_response(500)
            self.end_headers()

    def _handle_send_docx(self) -> None:
        temp_files: list[Path] = []
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0") or "0"))
            payload = json.loads(raw.decode("utf-8"))
            to_address = str(payload.get("to", "")).strip()
            subject = str(payload.get("subject", "")).strip()
            body = str(payload.get("body", "")).strip()
            doc_type = str(payload.get("docType", "")).strip().lower()
            data = payload.get("data") or {}

            if not to_address:
                self._send_json(400, {"error": "Destinataire manquant"})
                return
            if doc_type not in ("mr004", "cmt"):
                self._send_json(400, {"error": "docType invalide (mr004|cmt)"})
                return
            if not isinstance(data, dict):
                self._send_json(400, {"error": "data invalide"})
                return

            if doc_type == "mr004":
                temp_files.append(generate_mr004_docx(data))
            else:
                temp_files.append(generate_cmt_docx(data))

            send_via_outlook(to_address, subject or "Fiche", body, temp_files)
            self._send_json(200, {"ok": True, "attachments": len(temp_files)})
        except Exception as exc:
            self._send_json(500, {"error": str(exc), "trace": traceback.format_exc()})
        finally:
            for path in temp_files:
                try:
                    os.unlink(path)
                except OSError:
                    pass

    def _handle_generate_docx(self) -> None:
        temp_files: list[Path] = []
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0") or "0"))
            payload = json.loads(raw.decode("utf-8"))
            doc_type = str(payload.get("docType", "")).strip().lower()
            data = payload.get("data") or {}

            if doc_type not in ("mr004", "cmt"):
                self._send_json(400, {"error": "docType invalide (mr004|cmt)"})
                return
            if not isinstance(data, dict):
                self._send_json(400, {"error": "data invalide"})
                return

            if doc_type == "mr004":
                temp_files.append(generate_mr004_docx(data))
                filename = "MR004.docx"
            else:
                temp_files.append(generate_cmt_docx(data))
                filename = "CMT.docx"

            b = temp_files[0].read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)
        except Exception as exc:
            self._send_json(500, {"error": str(exc), "trace": traceback.format_exc()})
        finally:
            for path in temp_files:
                try:
                    os.unlink(path)
                except OSError:
                    pass

    def _handle_send_mail(self) -> None:
        if self.path != "/send-mail":
            self._send_json(404, {"error": "Route inconnue"})
            return

        temp_files: list[Path] = []
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )

            to_address = form.getfirst("to", "").strip()
            subject = form.getfirst("subject", "").strip()
            body = form.getfirst("body", "").strip()

            if not to_address:
                self._send_json(400, {"error": "Destinataire manquant"})
                return

            for key in form.keys():
                if not key.startswith("attachment_"):
                    continue
                field = form[key]
                if isinstance(field, list):
                    items = field
                else:
                    items = [field]
                for item in items:
                    if not getattr(item, "filename", ""):
                        continue
                    suffix = Path(item.filename).suffix or ".pdf"
                    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(item.file.read())
                        temp_files.append(Path(tmp.name))

            send_via_outlook(to_address, subject, body, temp_files)
            self._send_json(200, {"ok": True, "attachments": len(temp_files)})
        except Exception as exc:
            self._send_json(500, {"error": str(exc), "trace": traceback.format_exc()})
        finally:
            for path in temp_files:
                try:
                    os.unlink(path)
                except OSError:
                    pass

    def log_message(self, format: str, *args) -> None:
        return

    def _send_json(self, status: int, payload: dict) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), OutlookBridgeHandler)
    print(f"Outlook mail bridge running on http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
