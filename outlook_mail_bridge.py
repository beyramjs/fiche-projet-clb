from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
import cgi
import json
import os
import traceback

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
