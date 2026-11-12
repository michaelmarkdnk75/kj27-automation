"""Små hjælpere til Google (Drive, Gmail, Calendar) uden OAuth-login, der udløber.

Hemmeligheder (GitHub Secrets -> miljøvariabler):
  GOOGLE_SA_JSON      hele JSON-nøglen for servicekontoen (Drive + kalender).
                      KJ27-mappen og Automatik-mappen deles med servicekontoens mail som Redaktør,
                      og Michaels kalender deles med den (Se alle oplysninger om begivenheder).
  GMAIL_APP_PASSWORD  app-adgangskode til michaelmarkdnk@gmail.com (sender via SMTP, søger via IMAP).
"""
import email, imaplib, json, os, re, smtplib
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
import requests

ME = "michaelmarkdnk@gmail.com"
DRV = "https://www.googleapis.com/drive/v3"
UPL = "https://www.googleapis.com/upload/drive/v3"
CAL = "https://www.googleapis.com/calendar/v3"
_tok = {}


def _h():
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
    """Gmail-søgning (samme syntaks som i Gmail). Returnerer de nyeste uid'er først."""
    im = _imap()
    try:
        im.literal = q.encode("utf-8")          # sendes som IMAP-literal, så æøå virker
        _, d = im.uid("SEARCH", "CHARSET", "UTF-8", "X-GM-RAW")
        ids = d[0].split()[::-1][:n]
        return [i.decode() for i in ids]
    finally:
        im.logout()


def plain_text(uid):
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
    r = requests.get(f"{CAL}/calendars/{ME}/events", headers=_h(), timeout=30, params=dict(
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
