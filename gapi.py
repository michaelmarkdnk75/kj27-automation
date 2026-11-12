"""Små hjælpere til Google API (Drive, Gmail, Calendar) via REST og refresh tokens.

Hemmeligheder (GitHub Secrets -> miljøvariabler):
  GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET
  REFRESH_TOKEN_MAIL   (michaelmarkdnk@gmail.com: gmail.send, gmail.readonly, calendar.readonly)
  REFRESH_TOKEN_DRIVE  (koegejernmand@gmail.com: drive) – ejer af KJ27-filerne
"""
import base64, os, re
from email.message import EmailMessage
import requests

_tok = {}
DRV = "https://www.googleapis.com/drive/v3"
UPL = "https://www.googleapis.com/upload/drive/v3"
GM = "https://gmail.googleapis.com/gmail/v1/users/me"
CAL = "https://www.googleapis.com/calendar/v3"


def _token(kind):
    if kind not in _tok:
        r = requests.post("https://oauth2.googleapis.com/token", timeout=30, data=dict(
            client_id=os.environ["GOOGLE_CLIENT_ID"], client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
            refresh_token=os.environ[f"REFRESH_TOKEN_{kind}"], grant_type="refresh_token"))
        r.raise_for_status()
        _tok[kind] = r.json()["access_token"]
    return _tok[kind]


def _h(kind):
    return {"Authorization": "Bearer " + _token(kind)}


# ------------------------------------------------------------------ Drive (som koegejernmand)
def find(folder, name):
    """Nyeste fil med præcis dette navn i mappen."""
    q = f"'{folder}' in parents and name = '{name}' and trashed = false"
    r = requests.get(DRV + "/files", headers=_h("DRIVE"), timeout=30, params=dict(
        q=q, orderBy="createdTime desc", fields="files(id,name,createdTime,modifiedTime,md5Checksum)"))
    r.raise_for_status()
    f = r.json()["files"]
    if not f:
        raise FileNotFoundError(f"{name} findes ikke i Drive-mappen")
    return f[0]


def meta(fid):
    r = requests.get(f"{DRV}/files/{fid}", headers=_h("DRIVE"), timeout=30,
                     params=dict(fields="id,name,modifiedTime,md5Checksum"))
    r.raise_for_status()
    return r.json()


def download(fid, path):
    r = requests.get(f"{DRV}/files/{fid}", headers=_h("DRIVE"), params=dict(alt="media"), timeout=120)
    r.raise_for_status()
    open(path, "wb").write(r.content)
    return path


def update(fid, path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
    """Overskriver indholdet i den eksisterende fil (samme fil-id, ingen kopi)."""
    r = requests.patch(f"{UPL}/files/{fid}", headers={**_h("DRIVE"), "Content-Type": mime},
                       params=dict(uploadType="media"), data=open(path, "rb").read(), timeout=120)
    r.raise_for_status()
    return r.json()


# ------------------------------------------------------------------ Gmail (som michaelmarkdnk)
def send(to, subject, body, html=None, cc=None, attachments=()):
    m = EmailMessage()
    m["From"] = "Michael Mark <michaelmarkdnk@gmail.com>"
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
    raw = base64.urlsafe_b64encode(m.as_bytes()).decode()
    r = requests.post(GM + "/messages/send", headers=_h("MAIL"), json={"raw": raw}, timeout=60)
    r.raise_for_status()
    return r.json()["id"]


def search(q, n=10):
    r = requests.get(GM + "/messages", headers=_h("MAIL"), params=dict(q=q, maxResults=n), timeout=30)
    r.raise_for_status()
    return [m["id"] for m in r.json().get("messages", [])]


def plain_text(mid):
    r = requests.get(f"{GM}/messages/{mid}", headers=_h("MAIL"), params=dict(format="full"), timeout=30)
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


# ------------------------------------------------------------------ Calendar (som michaelmarkdnk)
def events(time_min, time_max, q=None):
    r = requests.get(CAL + "/calendars/primary/events", headers=_h("MAIL"), timeout=30, params=dict(
        timeMin=time_min, timeMax=time_max, q=q or "", singleEvents="true", orderBy="startTime", maxResults=50))
    r.raise_for_status()
    return r.json().get("items", [])


# ------------------------------------------------------------------ fælles
ME = "michaelmarkdnk@gmail.com"
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
