from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
import cgi
import json
import os
import traceback
import zipfile
import xml.etree.ElementTree as ET

import win32com.client


HOST = "127.0.0.1"
PORT = 8765


def send_via_outlook(to_address: str, subject: str, body: str, attachments: list[Path]) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = subject
    mail.Body = body
    for attachment in attachments:
      mail.Attachments.Add(str(attachment))
    mail.Send()


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
        _set_sdt_text(sdts[i], v)


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
    porteur = str(data.get("porteur", "")).strip()
    email = str(data.get("email", "")).strip()
    unite = str(data.get("unite", "")).strip()
    titre = str(data.get("titre", "")).strip()
    resume = str(data.get("resume", "")).strip()

    porteur_line = porteur
    if unite:
        porteur_line = f"{porteur} ({unite})".strip()

    # Fill the first placeholders in a stable sequence (matches the Word form order)
    # Only fill shared identity + project description fields; leave the rest untouched (template placeholder).
    _fill_by_placeholder_sequence(
        root,
        [
            "",  # Date d'envoi au DPD
            "",  # Entite patient (CLB/IHOPe/Autre) - laisse vide
            porteur_line,  # Entite(s) responsable(s) / porteur(s)
            porteur_line,  # Responsable scientifique CLB
            f"{porteur} - {email}".strip(" -"),  # Acteur operationnel
            titre,  # Nom du projet
            "",  # Numero EDS
            resume[:400],  # Description grand public (approx)
            resume,  # Description detaillee
        ],
    )

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

    porteur = str(data.get("porteur", "")).strip()
    email = str(data.get("email", "")).strip()
    unite = str(data.get("unite", "")).strip()
    titre = str(data.get("titre", "")).strip()
    partnerships = str(data.get("cmt_partnerships", "")).strip() or str(data.get("trf_dest", "")).strip()
    funding = str(data.get("cmt_funding", "")).strip()

    # Only fill shared identity fields. Everything else stays as the template placeholder.
    _fill_first_placeholder_in_container(root, "TITRE du projet", titre)
    _fill_first_placeholder_in_container(root, "Porteur", porteur)
    _fill_first_placeholder_in_container(root, "Laboratoire / Etablissement", unite)
    _fill_first_placeholder_in_container(root, "Coordonn", email)
    _fill_first_placeholder_in_container(root, "Collaboration(s) / Partenariat(s)", partnerships)
    _fill_first_placeholder_in_container(root, "Financement du projet", funding)

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

    def do_POST(self) -> None:
        if self.path == "/send-mail":
            return self._handle_send_mail()
        if self.path == "/send-docx":
            return self._handle_send_docx()
        self._send_json(404, {"error": "Route inconnue"})

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
    server = HTTPServer((HOST, PORT), OutlookBridgeHandler)
    print(f"Outlook mail bridge running on http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
