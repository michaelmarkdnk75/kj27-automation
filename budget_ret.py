"""Engangsrettelse af budget 2027 (12-11-2026): telt til målområdet og Stines sponsorerede timer.

Rækker indsættes med LibreOffice, så alle formler og henvisninger (også fanen Budget vs. faktisk)
flytter korrekt med. Filen opdateres på stedet i Drive – der laves ingen kopi.
Kør:  python budget_ret.py <sti-til-xlsx>   (kun lokal rettelse)
      JOB=budget-ret python run.py          (henter, retter og gemmer i Drive)
"""
import os, subprocess, sys, time, datetime

NOTE = "12-11-2026"
ROWS = [
    # (overskrift for afsnittet, indsæt før denne post, post, leverandør, manuel kr., bemærkning)
    ("Målområdet", "Uforudsete udgifter", "Telt til målområdet (min. 20 x 12 m)", None, 135000,
     "Leje 120.000–135.000 kr. jf. sponsormøde 11-11-2026. Sponsor søges (Stark-telt via Nyholm). "
     "Alternativ: sejlklubbens lokaler"),
    ("Annoncer og tryk", "SoMe-annoncering", "Rutetegninger (Stine Seiden)", "Seidenfaden GD", 0,
     "10 timer sponsoreret – timer derudover faktureres"),
]
LOG = [
    ("Telt til målområdet (min. 20 x 12 m)", "Ny post", 0, 135000, "Sponsormøde 11-11-2026, beslutning fra Michael Mark"),
    ("Rutetegninger (Stine Seiden)", "Ny post – 10 timer sponsoreret, timer derudover faktureres", 0, 0,
     "Aftale med Stine Seiden 29-10-2026"),
]


def edit(path):
    import uno
    from com.sun.star.beans import PropertyValue
    port = 2002
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--nologo", "--norestore",
                             f"--accept=socket,host=localhost,port={port};urp;"])
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def pv(n, v):
            p = PropertyValue(); p.Name = n; p.Value = v; return p

        url = uno.systemPathToFileUrl(os.path.abspath(path))
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (pv("Hidden", True),))
        ws = doc.Sheets.getByName("Budget 2027")

        def col_a():
            return [ws.getCellByPosition(0, r).getString() for r in range(0, 200)]

        for section, before, post, supplier, amount, note in ROWS:
            a = col_a()
            if post in a:
                print("Findes allerede:", post); continue
            s = a.index(section)
            r = next(i for i in range(s + 1, len(a)) if a[i] == before)  # 0-baseret række
            ws.Rows.insertByIndex(r, 1)
            src = r + 1  # posten der blev skubbet ned – formler kopieres derfra
            for c in range(9):
                f = ws.getCellByPosition(c, src).getFormula()
                if c in (6, 7) and f.startswith("="):
                    ws.getCellByPosition(c, r).setFormula(f.replace(str(src + 1), str(r + 1)))
            ws.getCellByPosition(0, r).setString(post)
            if supplier: ws.getCellByPosition(1, r).setString(supplier)
            ws.getCellByPosition(2, r).setValue(0)
            ws.getCellByPosition(3, r).setValue(0)
            ws.getCellByPosition(4, r).setString("Manuel")
            ws.getCellByPosition(5, r).setValue(amount)
            ws.getCellByPosition(8, r).setString(note)
            print("Indsat:", post, "række", r + 1)

        a = col_a()
        # Sørg for at subtotalen dækker alle poster i afsnittet (også en ny øverste post)
        for section, *_ in ROWS:
            s0 = a.index(section); t = a.index(section + " i alt")
            for c, L in ((2, "C"), (3, "D"), (6, "G"), (7, "H")):
                ws.getCellByPosition(c, t).setFormula(f"=SUM({L}{s0 + 2}:{L}{t})")
        r3 = 2
        ws.getCellByPosition(0, r3).setString(
            f"Sidst opdateret: {NOTE} af Michael Mark (telt til målområdet og rutetegninger tilføjet). "
            "Alle ændringer står på fanen Log.")

        lg = doc.Sheets.getByName("Log")
        la = [lg.getCellByPosition(2, r).getString() for r in range(200)]
        # Log-linjen fra 07-10 skal vise totalen dengang, ikke den aktuelle
        for i in range(200):
            if lg.getCellByPosition(5, i).getFormula().endswith("'Budget 2027'.G110") or "G1" in lg.getCellByPosition(5, i).getFormula() and lg.getCellByPosition(2, i).getString().startswith("Alle udgiftsposter"):
                lg.getCellByPosition(5, i).setValue(1289071.91)
        nxt = max(i for i in range(200) if lg.getCellByPosition(0, i).getString() or lg.getCellByPosition(2, i).getString()) + 1
        tot = a.index("UDGIFTER I ALT") + 1
        for post, change, before_kr, after_kr, src in LOG:
            if post in la:
                continue
            lg.getCellByPosition(0, nxt).setFormula("=DATE(2026;11;12)")
            lg.getCellByPosition(0, nxt).NumberFormat = lg.getCellByPosition(0, nxt - 1).NumberFormat
            lg.getCellByPosition(1, nxt).setString("Michael Mark")
            lg.getCellByPosition(2, nxt).setString(post)
            lg.getCellByPosition(3, nxt).setString(change)
            lg.getCellByPosition(4, nxt).setValue(before_kr)
            lg.getCellByPosition(5, nxt).setValue(after_kr)
            lg.getCellByPosition(6, nxt).setString(src)
            nxt += 1

        doc.calculateAll()
        g = lambda name: ws.getCellByPosition(6, a.index(name)).getValue()
        print("Målområdet i alt:", g("Målområdet i alt"), "| Annoncer og tryk i alt:", g("Annoncer og tryk i alt"),
              "| UDGIFTER I ALT:", g("UDGIFTER I ALT"), "| RESULTAT:", g("RESULTAT"))
        res = g("RESULTAT")
        doc.storeToURL(url, (pv("FilterName", "Calc MS Excel 2007 XML"), pv("Overwrite", True)))
        doc.close(True)
        return res
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        time.sleep(2)
        proc.kill()


def run(test):
    import shutil, tempfile, gapi
    ME = gapi.ME
    subprocess.run("sudo apt-get update -qq && sudo apt-get install -y -qq libreoffice-calc-nogui python3-uno",
                   shell=True, check=True, timeout=900)
    fid = gapi.find(gapi.KJ27_FOLDER, "Køge-Jernmand-2027-budget.xlsx")["id"]
    meta = gapi.meta(fid)
    d = tempfile.mkdtemp(); p = os.path.join(d, "budget.xlsx")
    gapi.download(fid, p)
    # UNO findes kun til systemets Python
    out = subprocess.run(["/usr/bin/python3", os.path.abspath(__file__), p], check=True, timeout=600,
                         capture_output=True, text=True).stdout
    print(out)
    res = out.strip().splitlines()[-1]
    if test:
        gapi.send([ME], "[TEST] KJ27 budget rettet (ikke gemt i Drive)", res,
                  attachments=[p])
        return
    if gapi.meta(fid).get("modifiedTime") != meta.get("modifiedTime"):
        sys.exit("Budgettet blev ændret under kørslen – intet gemt.")
    gapi.update(fid, p)
    print("Budget gemt i Drive. Resultat:", res)


if __name__ == "__main__":
    edit(sys.argv[1])
