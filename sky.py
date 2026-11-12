"""KJ27 sky: påmindelser (A), 14-dages status (B) og budget (C). Erstatter Claude-opgaven
"KJ27 sky – påmindelser, status, budget". Del D (mails om budgettet) ligger stadig i en lille Claude-opgave.

morning(): A + B (mandag) + C + opsummeringsmail til Michael (kun hvis noget skete eller fejlede)
evening(): kun C
TEST=ja: alle mails går kun til Michael med [TEST] i emnet, og intet skrives til Drive.
"""
import datetime as dt, html, os, shutil, subprocess, tempfile, traceback
from urllib.parse import quote
import openpyxl
import gapi
from gapi import ME

NY_ARBEJDSGANG = ("Får I en ekstra indtægt eller udgift til Køge Jernmand 2027, som ikke står i budgettet, eller ændrer "
                  "et beløb sig, så send mig en mail om det. Skriv hvad det drejer sig om, beløbet, om det er inkl. eller "
                  "ekskl. moms, og hvilket område det hører under. Vedhæft gerne tilbud eller faktura. Jeg skriver det ind "
                  "i budgettet, og I får den opdaterede version på mail.")


def sbtn(href, label, color="#1F4E78"):
    """Tydelig, omkranset knap (virker i Gmail, også uden baggrundsfarver)."""
    return (f'<a href="{html.escape(href)}" style="display:inline-block;margin:6px 6px 0 0;padding:9px 14px;'
            f'border:2px solid {color};border-radius:8px;color:{color};font-weight:bold;text-decoration:none;'
            f'font-family:Arial,sans-serif;font-size:14px">{html.escape(label)}</a>')


def mailto(subject, body=""):
    return f"mailto:{ME}?subject={quote(subject, safe='')}&body={quote(body, safe='')}"


BUDGET_FELTER = "Post: \nOmråde: \nBeløb (kr.): \nInkl./ekskl. moms: \nKommentar: "


def budget_buttons():
    return ('<p style="margin:14px 0 4px"><b>Ret i budgettet</b> – tryk på en knap, udfyld og send:</p>'
            + sbtn(mailto("KJ27 budget – ret beløb", BUDGET_FELTER), "Ret et beløb")
            + sbtn(mailto("KJ27 budget – ny indtægt/udgift", "Type (indtægt/udgift): \n" + BUDGET_FELTER), "Tilføj indtægt/udgift", "#516B46")
            + sbtn(mailto("KJ27 budget – slet post", "Post: \nBegrundelse: "), "Slet en post", "#C62828"))


def plan_buttons(kj, tasks):
    """Knapper pr. område til status-/ugemails: opdatér områdets opgaver, tilføj, slet."""
    rows = []
    for a in kj.SHEETS:
        ts = [t for t in tasks if t["area"] == a and t["pct"] < 1]
        b = sbtn(kj.mailto(a, kj.NEW_TASK_BLOCK), "+ Ny opgave", "#516B46")
        href = kj._all_href(a, ts) if ts else None   # None når listen er for lang til ét mail-link
        b = sbtn(href or kj.mailto(a, "<ID> <hvad skal ændres>\n"), "Opdatér opgaver") + b
        b += sbtn(kj.mailto(a, "Slet opgave: <ID>\nBegrundelse: "), "Slet opgave", "#C62828")
        rows.append(f'<div style="margin:10px 0 0"><b>{html.escape(a.replace("_", " og "))}</b><br>{b}</div>')
    return ('<div style="margin:16px 0 0;font-family:Arial,sans-serif;font-size:14px"><b>Ret i planen</b> – '
            'tryk på en knap, ret teksten i mailen og send:' + "".join(rows) + "</div>")


def to_pdf(xlsx, outdir):
    """Laver en PDF af budgettet med LibreOffice (installeres kun når der er brug for det)."""
    if not shutil.which("soffice"):
        subprocess.run("sudo apt-get update -qq && sudo apt-get install -y -qq libreoffice-calc-nogui",
                       shell=True, check=True, timeout=600)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, xlsx], check=True, timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(xlsx))[0] + ".pdf")


class Ctx:
    def __init__(self, test):
        self.test = test
        self.dir = tempfile.mkdtemp()
        self.log, self.errors = [], []

    def send(self, to, subject, body, html_=None, cc=None, attachments=()):
        if self.test:
            body = f"[TEST – ville være sendt til: {', '.join(to)}{' cc ' + ', '.join(cc) if cc else ''}]\n\n" + body
            return gapi.send([ME], "[TEST] " + subject, body, html_, None, attachments)
        return gapi.send(to, subject, body, html_, cc, attachments)

    def already_sent(self, subject):
        return bool(gapi.search(f'in:sent subject:"{subject}"', 1))


def load_kj27(c):
    """Henter kj27.py fra Drive (Automatik), så scriptet altid er den aktuelle version."""
    f = gapi.find(gapi.AUTO_FOLDER, "kj27.py")
    p = gapi.download(f["id"], os.path.join(c.dir, "kj27.py"))
    import importlib.util
    spec = importlib.util.spec_from_file_location("kj27", p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def fetch(c, folder, name):
    f = gapi.find(folder, name)
    return f, gapi.download(f["id"], os.path.join(c.dir, name))


# ------------------------------------------------------------------ A: påmindelser
def part_a(c, kj, today, plan, cfg_meta, cfg):
    due = kj.compute_due(plan, cfg, today)
    sent = []
    for d in due:
        if d.get("skipped"):
            c.log.append(f"Ingen påmindelse til {d['area']}: {d['skipped']} ({', '.join(i['id'] for i in d['items'])})")
            continue
        try:
            c.send(d["to"], d["subject"], d["body"], d["html"], d["cc"])
            sent.append(d)
            c.log.append(f"Påmindelse sendt: {d['area']} til {', '.join(d['to'])}"
                         + (f" (cc {', '.join(d['cc'])})" if d["cc"] else "")
                         + f" – opgaver {', '.join(i['id'] for i in d['items'])}")
        except Exception as e:
            c.errors.append(f"Påmindelse til {d['area']} fejlede: {e}")
    if sent:
        miss = {m["navn"]: m for m in kj.coverage(plan, cfg)}
        ids = {i["id"] for d in sent for i in d["items"]}
        for n, m in miss.items():
            if ids & set(m["opgaver"]):
                c.log.append(f"Ansvarlig uden mailadresse: {n} ({', '.join(sorted(ids & set(m['opgaver'])))})")
    if sent and not c.test:
        # Gem loggen i den samme fil på Drive (ingen kopi). Er filen ændret imens, hentes den igen.
        for _ in range(2):
            now = gapi.meta(cfg_meta["id"])
            if now["modifiedTime"] != cfg_meta["modifiedTime"]:
                gapi.download(cfg_meta["id"], cfg)
                cfg_meta = now
            kj.log_sent(cfg, sent, today)
            if gapi.meta(cfg_meta["id"])["modifiedTime"] == cfg_meta["modifiedTime"]:
                gapi.update(cfg_meta["id"], cfg)
                return len(sent)
        c.errors.append("Påmindelsesloggen kunne ikke gemmes (konfigurationen blev ændret imens). "
                        "Sendte påmindelser står ovenfor.")
    return len(sent)


# ------------------------------------------------------------------ B: status (mandage, hver 14. dag)
def part_b(c, kj, today, plan, cfg):
    st = kj.status_mail(plan, cfg, today)
    if not st["send"]:
        c.log.append(f"Statusmail: ikke sendt ({st['grund']})")
        return False
    if c.already_sent(st["subject"]):
        c.log.append("Statusmail: allerede sendt i dag")
        return False
    to = [e for e in st["to"] if not gapi.EXCLUDE.search(e)]
    h = st["html"].replace("<p>Mange hilsner", plan_buttons(kj, kj.load_tasks(plan)) + "<p>Mange hilsner", 1)
    c.send(to, st["subject"], st["body"], h, [ME])
    c.log.append(f"Statusmail sendt til {len(to)} modtagere")
    return True


# ------------------------------------------------------------------ C: budget
def _cells(path, data_only):
    wb = openpyxl.load_workbook(path, data_only=data_only)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value not in (None, ""):
                    out[(ws.title, cell.coordinate)] = cell.value
    return wb, out


def _kr(v):
    return "–" if v is None else f"{v:,.0f} kr.".replace(",", ".")


def budget_totals(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Budget 2027"]
    head = [c.value for c in ws[4]]
    col = head.index("Budget 2027 (kr.)")
    res = {}
    for r in ws.iter_rows(min_row=5, values_only=True):
        if r[0] in ("INDTÆGTER I ALT", "UDGIFTER I ALT", "RESULTAT"):
            res[r[0]] = r[col]
    return res


def part_c(c, today):
    bf = gapi.find(gapi.KJ27_FOLDER, "Køge-Jernmand-2027-budget.xlsx")
    sf = gapi.find(gapi.AUTO_FOLDER, "Budget-2027-sidst-sendt.xlsx")
    if bf.get("md5Checksum") and bf.get("md5Checksum") == sf.get("md5Checksum"):
        return 0
    new = gapi.download(bf["id"], os.path.join(c.dir, "budget.xlsx"))
    old = gapi.download(sf["id"], os.path.join(c.dir, "sidst.xlsx"))
    _, fn = _cells(new, False)
    _, fo = _cells(old, False)
    wbn, vn = _cells(new, True)
    _, vo = _cells(old, True)
    diffs = []
    for k in sorted(set(fn) | set(fo), key=lambda k: (k[0], k[1])):
        if fn.get(k) != fo.get(k):
            ws = wbn[k[0]] if k[0] in wbn.sheetnames else None
            label = ""
            if ws is not None:
                a = ws[f"A{''.join(ch for ch in k[1] if ch.isdigit())}"].value
                label = f" ({str(a)[:60]})" if a and k[1][0] != "A" else ""
            show = lambda v: _kr(v) if isinstance(v, (int, float)) else str(v if v is not None else "tom")[:80]
            diffs.append(f"{k[0]} {k[1]}{label}: {show(vo.get(k, fo.get(k)))} → {show(vn.get(k, fn.get(k)))}")
    if not diffs:
        return 0
    tot = budget_totals(new)
    _, leaders = c.kj.load_config(c.cfg)
    to = gapi.active_recipients(leaders)
    vis = diffs[:25] + ([f"og {len(diffs) - 25} ændringer mere"] if len(diffs) > 25 else [])
    k = lambda key: _kr(tot[key]) if tot.get(key) is not None else "kunne ikke læses (åbn og gem budgettet i Excel)"
    tal = [f"Indtægter i alt: {k('INDTÆGTER I ALT')}", f"Udgifter i alt: {k('UDGIFTER I ALT')}", f"Resultat: {k('RESULTAT')}"]
    body = ("Hej alle,\n\nBudgettet for Køge Jernmand 2027 er opdateret, og den nye version er vedhæftet.\n\n"
            + "\n".join(tal) + "\n\nÆndret siden sidste udsendelse:\n" + "\n".join("- " + d for d in vis)
            + "\n\nSig til, hvis noget skal rettes – svar på denne mail, eller brug knapperne i mailen.\n\nNy arbejdsgang for ekstra indtægter og udgifter\n"
            + NY_ARBEJDSGANG + "\n\nMange hilsner\nMichael Mark")
    e = html.escape
    h = ("<div style=\"font-family:Arial,sans-serif;font-size:14px\"><p>Hej alle,</p>"
         "<p>Budgettet for Køge Jernmand 2027 er opdateret, og den nye version er vedhæftet.</p><p>"
         + "<br>".join(e(t) for t in tal) + "</p><p>Ændret siden sidste udsendelse:</p><ul>"
         + "".join(f"<li>{e(d)}</li>" for d in vis) + "</ul>" + budget_buttons()
         + "<p><b>Ny arbejdsgang for ekstra indtægter og udgifter</b><br>" + e(NY_ARBEJDSGANG) + "</p>"
         "<p>Mange hilsner<br>Michael Mark</p></div>")
    att = []
    try:
        pdf = to_pdf(new, c.dir)
        att = [(f"Køge-Jernmand-2027-budget {today:%d-%m-%Y}.pdf", open(pdf, "rb").read(), "application/pdf")]
    except Exception as ex:
        c.errors.append(f"Budget-PDF kunne ikke laves ({ex}); mailen er sendt uden vedhæftning")
        body = body.replace("og den nye version er vedhæftet", "men PDF'en kunne ikke vedhæftes denne gang")
        h = h.replace("og den nye version er vedhæftet", "men PDF'en kunne ikke vedhæftes denne gang")
    c.send(to, "Opdateret budget Køge Jernmand 2027", body, h, [ME], att)
    if not c.test:
        gapi.update(sf["id"], new)   # sammenligningskopien opdateres i samme fil
    c.log.append(f"Opdateret budget sendt til {len(to)} modtagere ({len(diffs)} ændrede celler)")
    return len(to)


# ------------------------------------------------------------------ kørsler
def _safe(c, name, fn):
    try:
        return fn()
    except Exception as e:
        c.errors.append(f"{name}: {e}")
        traceback.print_exc()


def do_ingest(c, today, plan, plan_meta, cfg_meta):
    """Skriv områdeledernes svar ind i planen før alt andet. Returnerer (plan-sti, ny cfg-meta)."""
    import ingest
    c.ingest_n = 0
    try:
        n, lines = ingest.run(c, c.kj, plan, c.cfg, cfg_meta, plan_meta, today)
        c.ingest_n = n
        if lines:
            c.log.append("Svar fra områdelederne:")
            c.log += ["  " + x for x in lines]
        if getattr(c, "new_plan", None):
            plan = c.new_plan
        cfg_meta = gapi.meta(cfg_meta["id"])
    except Exception as e:
        c.errors.append(f"Indskrivning af svar: {e}")
        traceback.print_exc()
    return plan, cfg_meta


def ingest_only(today, test):
    """Kun indskrivning af svar + mail til Michael (sender aldrig til udvalget)."""
    c = Ctx(test)
    c.kj = load_kj27(c)
    plan_meta, plan = fetch(c, gapi.KJ27_FOLDER, "Køge-Jernmand-2027-planlægnig.xlsx")
    cfg_meta, c.cfg = fetch(c, gapi.AUTO_FOLDER, "KJ27-områdeledere.xlsx")
    do_ingest(c, today, plan, plan_meta, cfg_meta)
    body = "\n".join(c.log or ["Ingen nye svar fra områdelederne."]) + \
        ("\n\nFejl:\n" + "\n".join(c.errors) if c.errors else "") + "\n\nMange hilsner"
    gapi.send([ME], f"{'[TEST] ' if test else ''}KJ27 svar indskrevet i planen {today:%d-%m-%Y}", body)
    print(f"KJ27 indskrivning: {c.ingest_n} rettelser; fejl: {len(c.errors)}")


def morning(today, test):
    c = Ctx(test)
    a = b = 0
    try:
        c.kj = load_kj27(c)
        plan_meta, plan = fetch(c, gapi.KJ27_FOLDER, "Køge-Jernmand-2027-planlægnig.xlsx")
        cfg_meta, c.cfg = fetch(c, gapi.AUTO_FOLDER, "KJ27-områdeledere.xlsx")
        plan, cfg_meta = do_ingest(c, today, plan, plan_meta, cfg_meta)
        a = _safe(c, "Påmindelser", lambda: part_a(c, c.kj, today, plan, cfg_meta, c.cfg)) or 0
        if today.weekday() == 0:
            b = _safe(c, "Statusmail", lambda: part_b(c, c.kj, today, plan, c.cfg))
    except Exception as e:
        c.errors.append(f"Filer kunne ikke hentes: {e}")
        traceback.print_exc()
    n = _safe(c, "Budget", lambda: part_c(c, today)) if hasattr(c, "kj") else None
    if a or b or n or c.errors or getattr(c, "ingest_n", 0) or any("Svar fra" in x for x in c.log):
        lines = c.log + ([""] + ["Fejl: " + x for x in c.errors] if c.errors else [])
        body = "\n".join(lines) + "\n\nMange hilsner"
        c.send([ME], f"KJ27 påmindelser – status {today:%d-%m-%Y}", body)
    print(f"KJ27 sky morgen: {a} påmindelser, status {'sendt' if b else 'ikke sendt'}, "
          f"budget {'sendt til ' + str(n) if n else 'ingen ændringer'}; fejl: {len(c.errors)}")


def evening(today, test):
    c = Ctx(test)
    n = None
    try:
        c.kj = load_kj27(c)
        plan_meta, plan = fetch(c, gapi.KJ27_FOLDER, "Køge-Jernmand-2027-planlægnig.xlsx")
        cfg_meta, c.cfg = fetch(c, gapi.AUTO_FOLDER, "KJ27-områdeledere.xlsx")
        do_ingest(c, today, plan, plan_meta, cfg_meta)
        n = part_c(c, today)
    except Exception as e:
        c.errors.append(f"Budget: {e}")
        traceback.print_exc()
    if c.errors or c.log:
        gapi.send([ME], f"KJ27 aften {today:%d-%m-%Y}: svar indskrevet / budget",
                  "\n".join(c.log + ["Fejl: " + x for x in c.errors]) + "\n\nMange hilsner")
    print(f"KJ27 sky aften: budget {'sendt til ' + str(n) if n else 'ingen ændringer'}; fejl: {len(c.errors)}")
