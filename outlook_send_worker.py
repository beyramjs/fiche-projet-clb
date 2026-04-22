from __future__ import annotations

import json
import sys

import win32com.client


def main() -> int:
    try:
        raw = sys.stdin.buffer.read()
        payload = json.loads(raw.decode("utf-8"))
        to_address = str(payload.get("to", "")).strip()
        subject = str(payload.get("subject", "")).strip()
        body = str(payload.get("body", "")).strip()
        attachments = payload.get("attachments") or []

        if not to_address:
            print("Destinataire manquant", file=sys.stderr)
            return 2

        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_address
        mail.Subject = subject
        mail.Body = body
        for p in attachments:
            if not p:
                continue
            mail.Attachments.Add(str(p))
        mail.Send()
        return 0
    except Exception as exc:
        print(repr(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

