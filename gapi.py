"""Små hjælpere til Google (Drive, Gmail, Calendar) uden OAuth-login, der udløber.

Hemmeligheder (GitHub Secrets -> miljøvariabler):
  GOOGLE_SA_JSON      hele JSON-nøglen for servicekontoen (Drive + kalender).
                      KJ27-mappen og Automatik-mappen deles med servicekontoens mail som Redaktør,
                      og Michaels kalender deles med den (Se alle oplysninger om begivenheder).
  GMAIL_APP_PASSWORD  app-adgangskode til michaelmarkdnk@gmail.com (sender via SMTP, søger via IMAP).
"""
import email, email.header, imaplib, json, os, re, smtplib
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
import requests

_orig_rfs = requests.Response.raise_for_status


def _rfs(self):
    """Vis Googles fejltekst i loggen, så fejl kan forstås."""
    if self.status_code >= 400:
        print(f"Google-fejl {self.status_code} {self.url.split('?')[0]}: {self.text[:600]}")
    return _orig_rfs(self)


requests.Response.raise_for_status = _rfs

ME = "michaelmarkdnk@gmail.com"
DRV = "https://www.googleapis.com/drive/v3"
UPL = "https://www.googleapis.com/upload/drive/v3"
CAL = "https://www.googleapis.com/calendar/v3"
_tok = {}


def _oauth(kind):
    """Midlertidig login med refresh token (udløber efter 7 dage, mens Google-appen er i testtilstand)."""
    if kind not in _tok:
        r = requests.post("https://oauth2.googleapis.com/token", timeout=30, data=dict(
            client_id=os.environ["GOOGLE_CLIENT_ID"], client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
            refresh_token=os.environ[f"REFRESH_TOKEN_{kind}"], grant_type="refresh_token"))
        r.raise_for_status()
        _tok[kind] = r.json()["access_token"]
    return {"Authorization": "Bearer " + _tok[kind]}


USE_OAUTH = bool(os.environ.get("REFRESH_TOKEN_DRIVE")) and not os.environ.get("GOOGLE_SA_JSON")
if not USE_OAUTH and not os.environ.get("GOOGLE_SA_JSON"):
    raise SystemExit("Mangler GitHub-hemmeligheder: enten REFRESH_TOKEN_MAIL + REFRESH_TOKEN_DRIVE + GOOGLE_CLIENT_ID + "
                     "GOOGLE_CLIENT_SECRET, eller GOOGLE_SA_JSON + GMAIL_APP_PASSWORD.")
CAL_ID = "primary" if USE_OAUTH else ME


def _h(kind="DRIVE"):
    if USE_OAUTH:
        return _oauth(kind)
    if "sa" not in _tok:
        from google.oauth2 import service_account
        from google.auth.transport.requests import Request
        c = service_account.Credentials.from_service_account_info(
            json.loads(os.environ["GOOGLE_SA_JSON"]),
            scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/calendar.readonly"])
        c.refresh(Request())
        _tok["sa"] = c.token
    return {"Authorization": "Bearer " + _tok["sa"]}


# ------------------------------------------------------------------ Drive
def find(folder, name):
    """Nyeste fil med præcis dette navn i mappen."""
    q = f"'{folder}' in parents and name = '{name}' and trashed = false"
    r = requests.get(DRV + "/files", headers=_h(), timeout=30, params=dict(
        q=q, orderBy="createdTime desc", fields="files(id,name,createdTime,modifiedTime,md5Checksum)",
        supportsAllDrives="true", includeItemsFromAllDrives="true"))
    r.raise_for_status()
    f = r.json()["files"]
    if not f:
        raise FileNotFoundError(f"{name} findes ikke i Drive-mappen (er mappen delt med servicekontoen?)")
    return f[0]


def meta(fid):
    r = requests.get(f"{DRV}/files/{fid}", headers=_h(), timeout=30, params=dict(fields="id,name,modifiedTime,md5Checksum"))
    r.raise_for_status()
    return r.json()


def download(fid, path):
    r = requests.get(f"{DRV}/files/{fid}", headers=_h(), params=dict(alt="media"), timeout=120)
    r.raise_for_status()
    open(path, "wb").write(r.content)
    return path


def update(fid, path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """Overskriver indholdet i den eksisterende fil (samme fil-id, ingen kopi)."""
    r = requests.patch(f"{UPL}/files/{fid}", headers={**_h(), "Content-Type": mime},
                       params=dict(uploadType="media"), data=open(path, "rb").read(), timeout=120)
    r.raise_for_status()
    return r.json()


# ------------------------------------------------------------------ Gmail (SMTP/IMAP med app-adgangskode)
def send(to, subject, body, html=None, cc=None, attachments=()):
    m = EmailMessage()
    m["From"] = f"Michael Mark <{ME}>"
    m["To"] = ", ".join(to)
    if cc:
        m["Cc"] = ", ".join(cc)
    m["Subject"] = subject
    m.set_content(body)
    if html:
        m.add_alternative(html, subtype="html")
    for fname, data, mime in attachments:
        mt, st = mime.split("/", 1)
        m.add_attachment(data, maintype=mt, subtype=st, filename=fname)
    if USE_OAUTH:
        import base64
        raw = base64.urlsafe_b64encode(m.as_bytes()).decode()
        r = requests.post("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", headers=_oauth("MAIL"),
                          json={"raw": raw}, timeout=60)
        r.raise_for_status()
        return r.json()["id"]
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(ME, os.environ["GMAIL_APP_PASSWORD"].replace(" ", ""))
        s.send_message(m)
    return m["Subject"]


def _imap():
    im = imaplib.IMAP4_SSL("imap.gmail.com")
    im.login(ME, os.environ["GMAIL_APP_PASSWORD"].replace(" ", ""))
    _, boxes = im.list()
    allbox = next(re.search(rb'"([^"]+)"$', b).group(1).decode() for b in boxes if rb"\All" in b)
    im.select(f'"{allbox}"', readonly=True)
    return im


def search(q, n=10):
    """Gmail-søgning (samme syntaks som i Gmail). Returnerer de nyeste først."""
    if USE_OAUTH:
        r = requests.get("https://gmail.googleapis.com/gmail/v1/users/me/messages", headers=_oauth("MAIL"),
                         params=dict(q=q, maxResults=n), timeout=30)
        r.raise_for_status()
        return [m["id"] for m in r.json().get("messages", [])]
    im = _imap()
    try:
        im.literal = q.encode("utf-8")          # sendes som IMAP-literal, så æøå virker
        _, d = im.uid("SEARCH", "CHARSET", "UTF-8", "X-GM-RAW")
        ids = d[0].split()[::-1][:n]
        return [i.decode() for i in ids]
    finally:
        im.logout()


def plain_text(uid):
    if USE_OAUTH:
        import base64
        r = requests.get(f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{uid}", headers=_oauth("MAIL"),
                         params=dict(format="full"), timeout=30)
        r.raise_for_status()
        msg = r.json()

        def walk(p):
            if p.get("mimeType") == "text/plain" and p.get("body", {}).get("data"):
                return base64.urlsafe_b64decode(p["body"]["data"]).decode("utf-8", "replace")
            for c in p.get("parts", []) or []:
                t = walk(c)
                if t:
                    return t
            return ""
        return walk(msg["payload"]), int(msg.get("internalDate", 0))
    im = _imap()
    try:
        _, d = im.uid("FETCH", uid, "(RFC822)")
        msg = email.message_from_bytes(d[0][1])
        part = msg.get_body(preferencelist=("plain",)) if hasattr(msg, "get_body") else None
        if part is None:
            for p in msg.walk():
                if p.get_content_type() == "text/plain":
                    part = p
                    break
        txt = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace") if part else ""
        ts = int(parsedate_to_datetime(msg["Date"]).timestamp() * 1000) if msg["Date"] else 0
        return txt, ts
    finally:
        im.logout()


# ------------------------------------------------------------------ Kalender (Michaels kalender delt med servicekontoen)
def events(time_min, time_max, q=None):
    r = requests.get(f"{CAL}/calendars/{CAL_ID}/events", headers=_h("MAIL"), timeout=30, params=dict(
        timeMin=time_min, timeMax=time_max, q=q or "", singleEvents="true", orderBy="startTime", maxResults=50))
    r.raise_for_status()
    return r.json().get("items", [])


# ------------------------------------------------------------------ fælles
# Kontaktes aldrig om KJ27 (Michaels faste regel)
EXCLUDE = re.compile(r"hertz|simonsigurd|simon\.sigurd|vibeke|rottgarn", re.I)
KJ27_FOLDER = "1PAUBz_TCoWy-crokJYMgFPPdoavSvmya"
AUTO_FOLDER = "1un6f-XzIeIs4FO2_eIWyDRUoF2pcbxZC"


def active_recipients(leaders):
    out = []
    for l in leaders.values():
        if l["aktiv"]:
            for e in re.split(r"[;,]", l["email"] or ""):
                e = e.strip()
                if e and not EXCLUDE.search(e) and e.lower() not in [x.lower() for x in out]:
                    out.append(e)
    return out


# ------------------------------------------------------------------ hele mails (til indskrivning af svar)
def messages(q, n=100, skip=()):
    """Mails der matcher Gmail-søgningen, ældste først:
    dict(id (Gmail-id i hex), thread, sender_email, sender_name, subject, ts (ms), text)."""
    out = []
    if USE_OAUTH:
        import base64
        import time
        ids = [i for i in search(q, n) if i not in skip]       # kun mails der ikke er behandlet
        for mid in ids:
            for forsoeg in range(5):
                r = requests.get(f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{mid}", headers=_oauth("MAIL"),
                                 params=dict(format="full"), timeout=30)
                if r.status_code in (403, 429) and "uota" in r.text:
                    time.sleep(15 * (forsoeg + 1))             # Gmails minutkvote: vent og prøv igen
                    continue
                break
            r.raise_for_status()
            time.sleep(0.3)
            m = r.json()
            hd = {h["name"].lower(): h["value"] for h in m["payload"].get("headers", [])}

            def walk(p):
                if p.get("mimeType") == "text/plain" and p.get("body", {}).get("data"):
                    return base64.urlsafe_b64decode(p["body"]["data"]).decode("utf-8", "replace")
                for c in p.get("parts", []) or []:
                    t = walk(c)
                    if t:
                        return t
                return ""
            out.append(_msg(m["id"], m["threadId"], hd.get("from", ""), hd.get("subject", ""),
                            int(m.get("internalDate", 0)), walk(m["payload"])))
    else:
        im = _imap()
        try:
            im.literal = q.encode("utf-8")
            _, d = im.uid("SEARCH", "CHARSET", "UTF-8", "X-GM-RAW")
            for uid in d[0].split()[-n:]:
                _, f0 = im.uid("FETCH", uid, "(X-GM-MSGID)")
                if format(int(re.search(rb"X-GM-MSGID (\d+)", f0[0]).group(1)), "x") in skip:
                    continue
                _, f = im.uid("FETCH", uid, "(X-GM-MSGID X-GM-THRID RFC822)")
                meta = f[0][0].decode()
                gid = format(int(re.search(r"X-GM-MSGID (\d+)", meta).group(1)), "x")
                tid = format(int(re.search(r"X-GM-THRID (\d+)", meta).group(1)), "x")
                msg = email.message_from_bytes(f[0][1])
                part = next((p for p in msg.walk() if p.get_content_type() == "text/plain"), None)
                txt = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace") if part else ""
                ts = int(parsedate_to_datetime(msg["Date"]).timestamp() * 1000) if msg["Date"] else 0
                out.append(_msg(gid, tid, msg.get("From", ""), str(email.header.make_header(email.header.decode_header(msg.get("Subject", "")))), ts, txt))
        finally:
            im.logout()
    return sorted(out, key=lambda m: m["ts"])


def _msg(mid, tid, frm, subject, ts, text):
    from email.utils import parseaddr
    name, addr = parseaddr(frm)
    return dict(id=mid, thread=tid, sender_email=addr.lower(), sender_name=name or addr, subject=subject, ts=ts, text=text)
