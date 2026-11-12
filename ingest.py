"""Skriver områdeledernes svar ind i planen (Gantt) – uden tokens.
Læser alle mails med emnet "KJ27 opdatering – <område>" og "[KJ27 dato]", som ikke står i fanen
"Behandlede mails", og bruger kj27.py (parse + apply) til at skrive entydige rettelser ind.
Kun afsendere fra fanerne Områdeledere/Personer. Michaels egne mails og uklare svar skrives ikke ind,
men står i opsummeringen. Planen og konfigurationen skrives tilbage i de samme filer (ingen kopier).
"""
import datetime as dt, json, os, re
import gapi
from gapi import ME

DELETE_WORDS = re.compile(r"\b(fjern|slet|slettes|fjernes|ikke relevant|står 2 gange|dublet)", re.I)


def strip_signature(t):
    t = re.split(r"\n(Mange hilsner|Med venlig hilsen|Venlig hilsen|Best regards|Mvh)\b", t)[0]
    return t


def run(c, kj, plan, cfg, cfg_meta, plan_meta, today):
    """Returnerer (antal indskrevne mails, liste med linjer til opsummering)."""
    import openpyxl
    wb, leaders = kj.load_config(cfg)
    people = kj.load_people(wb, leaders)
    known_mail = {p["email"].lower(): p["navn"] for p in people}
    for l in leaders.values():
        for e in re.split(r"[;,]", l["email"] or ""):
            if e.strip():
                known_mail.setdefault(e.strip().lower(), l["navn"])
    done = set(kj.processed(cfg))
    areas = {a.lower().replace("_", " "): a for a in kj.SHEETS}
    msgs = gapi.messages('newer_than:60d (subject:"KJ27 opdatering" OR subject:"KJ27 dato")', 300)
    lines, applied_ids, all_updates, marks = [], [], [], []
    for m in msgs:
        if m["id"] in done:
            continue
        if m["sender_email"] == ME:
            marks.append((m["id"], "Michael Mark", "", 0))      # egne/automatiske mails tæller ikke som svar
            continue
        subj = m["subject"]
        ma = re.search(r"KJ27 opdatering\s*[–-]\s*(.+)$", subj, re.I)
        area = areas.get(ma.group(1).strip().lower().replace("_", " "), "") if ma else ""
        who = known_mail.get(m["sender_email"])
        if not who:
            lines.append(f"Ikke indskrevet (afsender ikke på listen): {m['sender_name']} <{m['sender_email']}> – {subj}")
            marks.append((m["id"], m["sender_name"], area, 0))
            continue
        body = strip_signature(m["text"])
        p = os.path.join(c.dir, f"m_{m['id']}.txt")
        open(p, "w", encoding="utf8").write(body)
        date = dt.datetime.fromtimestamp(m["ts"] / 1000).date().isoformat()
        res = kj.parse_reply(plan, open(p, encoding="utf8").read(), who, date)
        ups = res.get("updates", [])
        # [KJ27 dato] <ID>: kommentar uden id i teksten -> knyt til id i emnet
        md = re.search(r"\[KJ27 dato\]\s*([A-ZÆØÅ]{3}-\d{3})", subj)
        if md and not ups:
            kom = re.search(r"Kommentar:\s*(.+)", body, re.S)
            nd = re.search(r"(\d{1,2}-\d{1,2}-\d{4})", body)
            u = dict(id=md.group(1), fra=who, dato=date)
            if kom and kom.group(1).strip():
                u["kommentar"] = " ".join(kom.group(1).split())
            if nd:
                d, mo, y = nd.group(1).split("-")
                u["slut"] = f"{y}-{int(mo):02d}-{int(d):02d}"
            if len(u) > 3:
                ups = [u]
        for u in ups:
            all_updates.append(u)
            what = ", ".join(f"{k} {v if k != 'pct' else str(round(v * 100)) + '%'}" for k, v in u.items()
                             if k not in ("id", "fra", "dato"))
            lines.append(f"Indskrevet: {u['id']} ({who}): {what[:160]}")
            if DELETE_WORDS.search(str(u.get("kommentar", ""))):
                lines.append(f"  → {u['id']}: {who} ønsker opgaven slettet. Slet rækken i Excel, hvis du er enig.")
        for n in res.get("nye_opgaver", []):
            lines.append(f"Ny opgave foreslået ({who}, {area}): {json.dumps(n, ensure_ascii=False)[:200]} – ikke indskrevet, tjek selv")
        for pr in res.get("problemer", []):
            lines.append(f"Uklart ({who}): {pr} – ikke indskrevet")
        if not ups and not res.get("nye_opgaver") and not res.get("problemer"):
            lines.append(f"Ingen planændring fundet ({who}, {subj}): {' '.join(body.split())[:140]}")
        marks.append((m["id"], who, area, len(ups)))
    if not marks:
        return 0, []
    if c.test:
        return len(all_updates), ["[TEST – intet skrevet til planen]"] + lines
    out = os.path.join(c.dir, "plan_ny.xlsx")
    if all_updates:
        upd = os.path.join(c.dir, "upd.json")
        json.dump(all_updates, open(upd, "w", encoding="utf8"), ensure_ascii=False)
        rep = kj.apply_updates(plan, all_updates, out, today, cfg)
        if rep.get("ukendte"):
            lines.append(f"Ukendte opgave-id'er (ikke indskrevet): {', '.join(rep['ukendte'])}")
        if gapi.meta(plan_meta["id"])["modifiedTime"] != plan_meta["modifiedTime"]:
            lines.append("Planen blev ændret imens – intet skrevet. Prøver igen ved næste kørsel.")
            return 0, lines
        gapi.update(plan_meta["id"], out)
    if gapi.meta(cfg_meta["id"])["modifiedTime"] != cfg_meta["modifiedTime"]:
        gapi.download(cfg_meta["id"], cfg)          # konfigurationen ændret imens: brug den nye
    for mid, who, area, n in marks:
        kj.mark_processed(cfg, mid, who, area or "-", n, today)
    gapi.update(cfg_meta["id"], cfg)
    c.new_plan = out if all_updates else None
    return len(all_updates), lines
