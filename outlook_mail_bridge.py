from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
import cgi
import json
import os
import mimetypes
import traceback
import zipfile
import xml.etree.ElementTree as ET

import win32com.client


HOST = "127.0.0.1"
PORT = 8765
DATA_DIR = Path(__file__).with_name("data")
REGISTRY_PATH = DATA_DIR / "projects_registry.json"


def send_via_outlook(to_address: str, subject: str, body: str, attachments: list[Path]) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_address
    mail.Subject = subject
    mail.Body = body
    for attachment in attachments:
      mail.Attachments.Add(str(attachment))
    mail.Send()


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
