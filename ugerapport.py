"""KJ27 ugerapport – eneste fælles projektmail til stævneudvalget (mandag 07:23).
Erstatter Claude-opgaven "KJ27 ugerapport". Bygger alt ud fra plan, budget og kalender; gætter intet.
TEST=ja: mailen går kun til Michael med [TEST] i emnet.
"""
import datetime as dt, html, os, re, tempfile
import openpyxl
import gapi
from gapi import ME

RACE = dt.date(2027, 6, 13)
AREAS = ["Målområdet", "Sponsor", "Løb_Cykling", "Service Team", "Svøm", "Handyman team", "Skiftezone", "Rutetest", "Økonomi"]
SHOW = {"Løb_Cykling": "Løb og cykling"}
BASELINE = ["SVO-006", "SVO-013", "SVO-016", "LOB-015", "LOB-028", "HAN-002", "HAN-016"]
BLACK, TURK, ORANGE, GREEN, BG, LINE = "#1A1819", "#74C5D6", "#E6943E", "#516B46", "#F5F5F2", "#E6E6E1"
RAG = {"Grøn": "#2E7D32", "Gul": "#F9A825", "Rød": "#C62828"}
MAXL = 8
DAGE = ["mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag", "søndag"]
e = html.escape


def f(d):
    return d.strftime("%d-%m-%Y") if d else "–"


def kr(v):
    return "–" if v is None else f"{v:,.0f} kr.".replace(",", ".")


def area_name(a):
    return SHOW.get(a, a)


# ------------------------------------------------------------------ data
def classify(tasks, today):
    for t in tasks:
        p = t["pct"]
        if p >= 1:
            t["st"] = "Færdig"
        elif t["slut"] and t["slut"] < today:
            t["st"] = "Forsinket"
        elif p == 0 and t["start"] and t["start"] > today:
            t["st"] = "Ikke startet"
        else:
            t["st"] = "I gang"
        t["datofejl"] = bool(t["slut"] and t["slut"] > RACE)
        a = (t["ansvarlig"] or "").strip()
        t["uden_ejer"] = not a or a in AREAS or a in SHOW.values()
        t["dage_til_slut"] = (t["slut"] - today).days if t["slut"] else None


def area_stats(ts, today):
    n = len(ts)
    s = {k: sum(t["st"] == k for t in ts) for k in ("Færdig", "I gang", "Forsinket", "Ikke startet")}
    prog = sum(min(t["pct"], 1) for t in ts) / n if n else 0
    nxt = min([t["start"] for t in ts if t["st"] != "Færdig" and t["start"] and t["start"] >= today], default=None)
    df = sum(t["datofejl"] for t in ts)
    ue = sum(t["uden_ejer"] for t in ts)
    if s["Forsinket"] or any(t["st"] != "Færdig" and t["dage_til_slut"] is not None and 0 <= t["dage_til_slut"] <= 14
                             and t["pct"] < 0.5 for t in ts):
        rag = "Rød"
    elif any(t["st"] != "Færdig" and t["dage_til_slut"] is not None and 0 <= t["dage_til_slut"] <= 30 for t in ts) or df or ue:
        rag = "Gul"
    else:
        rag = "Grøn"
    return dict(n=n, prog=prog, nxt=nxt, df=df, ue=ue, rag=rag, **s)


def budget(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Budget 2027"]
    head = [c.value for c in ws[4]]
    cb, cr = head.index("Budget 2027 (kr.)"), head.index("Regnskab 2026 (kr.)")
    tot, areas, in_udg = {}, [], False
    for r in ws.iter_rows(min_row=5, values_only=True):
        lab = str(r[0] or "").strip()
        if lab == "UDGIFTER":
            in_udg = True
        if lab in ("INDTÆGTER I ALT", "UDGIFTER I ALT", "RESULTAT"):
            tot[lab] = (r[cb], r[cr])
            if lab == "UDGIFTER I ALT":
                in_udg = False
        elif in_udg and lab.endswith(" i alt"):
            areas.append((lab[:-6], r[cb] or 0))
    areas.sort(key=lambda x: -x[1])
    afk, on = [], False
    for r in wb["Antagelser"].iter_rows(values_only=True):
        v = str(r[0] or "").strip()
        if v == "Afklaringspunkter":
            on = True
        elif on and re.match(r"^\d+\.", v):
            afk.append(v)
        elif on and v:
            break
    return tot, areas, afk


def prev_done(today):
    q = f'from:me subject:"KJ27 ugerapport" -subject:FORHÅNDSVISNING before:{today:%Y/%m/%d}'
    for mid in gapi.search(q, 5):
        txt, _ = gapi.plain_text(mid)
        m = re.search(r"Færdige-id:\s*(.*)", txt)
        if m:
            return [x.strip() for x in m.group(1).split(",") if x.strip()], True
    return BASELINE, False


def meetings(today):
    from zoneinfo import ZoneInfo
    dk = ZoneInfo("Europe/Copenhagen")
    out = []
    for ev in gapi.events(f"{today}T00:00:00Z", f"{today + dt.timedelta(days=14)}T23:59:59Z", "KJ27"):
        s = ev.get("start", {})
        st = s.get("dateTime") or s.get("date")
        d = dt.datetime.fromisoformat(st.replace("Z", "+00:00")).astimezone(dk).replace(tzinfo=None) if "T" in st \
            else dt.datetime.fromisoformat(st)
        out.append((d, ev.get("summary", "")))
    return out


# ------------------------------------------------------------------ html-byggesten
def sec(title, inner):
    return (f'<tr><td style="padding:18px 20px 6px;font:bold 13px Arial;letter-spacing:1px;color:{GREEN}">{e(title.upper())}</td></tr>'
            f'<tr><td style="padding:0 20px 8px;font:14px/1.5 Arial;color:{BLACK}">{inner}</td></tr>')


def dot(rag):
    return f'<span style="display:inline-block;width:10px;height:10px;border-radius:5px;background-color:{RAG[rag]}"></span>'


def ul(items):
    return "<br>".join(items) if items else ""


def cap(lst, n=MAXL):
    return lst[:n] + ([f"og {len(lst) - n} mere"] if len(lst) > n else [])


# ------------------------------------------------------------------ rapport
def build(kj, plan, bpath, today):
    tasks = [t for t in kj.load_tasks(plan) if not re.fullmatch(r"[A-ZÆØÅ]{3}-\d{3}(\s+[A-ZÆØÅ]{3}-\d{3})*", t["aktivitet"])]
    classify(tasks, today)
    stats = {a: area_stats([t for t in tasks if t["area"] == a], today) for a in AREAS}
    tot = area_stats(tasks, today)
    worst = "Rød" if any(s["rag"] == "Rød" for s in stats.values()) else "Gul" if any(s["rag"] == "Gul" for s in stats.values()) else "Grøn"
    dage = (RACE - today).days
    lc = lambda t: (kj.latest_comment(t) or "")[:120]

    action = sorted([t for t in tasks if t["st"] == "Forsinket" or (t["st"] != "Færdig" and t["dage_til_slut"] is not None
                     and 0 <= t["dage_til_slut"] <= 30)], key=lambda t: t["slut"])
    act_l = [f"<b>{e(t['id'])} {e(t['aktivitet'])}</b> · {e(t['ansvarlig'] or 'ejer mangler')} · slut {f(t['slut'])} · "
             f"{round(t['pct'] * 100)}%{' · ' + e(lc(t)) if lc(t) else ''}" for t in action]
    upc = sorted([t for t in tasks if t["st"] != "Færdig" and t["start"] and t["start"] >= today], key=lambda t: (t["start"], t["id"]))
    nxt5 = upc[:5]
    extra = len([t for t in upc[5:] if nxt5 and t["start"] == nxt5[-1]["start"]])
    fix = [t for t in tasks if t["datofejl"] or t["uden_ejer"]]
    fix_l = []
    for a in AREAS:
        for t in [t for t in fix if t["area"] == a]:
            why = "slutdato efter stævnet" if t["datofejl"] else "ejer mangler"
            fix_l.append(f"{e(area_name(a))}: {e(t['id'])} {e(t['aktivitet'])} · {e(t['ansvarlig'] or '–')} · "
                         f"{f(t['start'])}–{f(t['slut'])} ({why})")
    done = sorted([t for t in tasks if t["st"] == "Færdig"], key=lambda t: t["id"])
    prev, found = prev_done(today)
    new_closed = [t for t in done if t["id"] not in prev]
    btot, bareas, afk = budget(bpath)
    mtg = meetings(today)

    # risici (kun ud fra data)
    risks = []
    late = [t for t in tasks if t["st"] == "Forsinket"]
    if late:
        risks.append(f"{len(late)} forsinkede opgaver ({', '.join(t['id'] for t in late[:5])}). Konsekvens: efterfølgende opgaver "
                     f"skubbes. Ejer: {e(', '.join(sorted({t['ansvarlig'] or 'ejer mangler' for t in late})[:3]))}. Næste skridt: ny slutdato eller hjælp.")
    soon = [t for t in tasks if t["st"] != "Færdig" and t["dage_til_slut"] is not None and 0 <= t["dage_til_slut"] <= 14 and t["pct"] < 0.5]
    for t in soon[:2]:
        risks.append(f"{e(t['id'])} {e(t['aktivitet'])} skal være færdig {f(t['slut'])} og står på {round(t['pct'] * 100)}%. "
                     f"Ejer: {e(t['ansvarlig'] or 'ejer mangler')}. Næste skridt: bekræft at fristen holder.")
    if tot["df"]:
        risks.append(f"{tot['df']} opgaver har slutdato efter stævnet. Konsekvens: Gantt-diagrammet viser forkert. "
                     "Ejer: områdelederne. Næste skridt: ret datoerne (se \"Ret i planen\").")
    for p in afk:
        if len(risks) >= 5:
            break
        if re.search(r"sponsor", p, re.I):
            risks.append(f"Ubekræftede sponsorindtægter: {e(p)} Ejer: Økonomi. Næste skridt: bekræft aftalerne.")
    risks = risks[:5]

    nxt_meet = next((f"{DAGE[d.weekday()]} {d:%d-%m}" for d, s in mtg if re.search(r"stævneudvalg", s, re.I)), "næste møde")
    dec = [f"{e(p)} – beslutter: Økonomi ({nxt_meet})" for p in afk[:3]]
    if tot["ue"]:
        dec.append(f"{tot['ue']} opgaver uden ejer skal have en ansvarlig – beslutter: områdelederne ({nxt_meet})")
    if tot["df"]:
        dec.append(f"{tot['df']} slutdatoer efter stævnet skal rettes – beslutter: områdelederne ({nxt_meet})")
    dec = dec[:5]

    stu = next(((d, s) for d, s in mtg if re.search(r"stævneudvalg", s, re.I) and (d.date() - today).days <= 7), None)
    stu_line = f"Status til stævneudvalgsmødet {DAGE[stu[0].weekday()]} {stu[0]:%d-%m-%Y}" + (
        f" kl. {stu[0]:%H:%M}" if stu[0].hour or stu[0].minute else "") if stu else ""

    # ---------------------------------------------------- HTML
    kpi = lambda big, small, col=BLACK: (f'<td style="background-color:#ffffff;border:1px solid {LINE};padding:10px;text-align:center;width:33%">'
                                         f'<div style="font:bold 22px Arial;color:{col}">{big}</div><div style="font:12px Arial;color:#555">{small}</div></td>')
    rows_area = []
    for a in AREAS:
        s = stats[a]
        w = round(s["prog"] * 100)
        rows_area.append(
            f'{dot(s["rag"])} <b>{e(area_name(a))}</b> <span style="color:{GREEN};font-weight:bold">{w}%</span>'
            f'<div style="background-color:{LINE};height:6px;margin:3px 0"><div style="background-color:{ORANGE};height:6px;width:{w}%"></div></div>'
            f'<span style="font-size:12px;color:#555">{s["n"]} opgaver · {s["Færdig"]} færdige · {s["I gang"]} i gang · {s["Forsinket"]} forsinkede · '
            f'{s["df"]} datofejl · Næste start {f(s["nxt"])}</span>')
    nxt_html = [f'<b>{e(t["aktivitet"])}</b> <span style="float:right">{f(t["start"])}</span><br>'
                f'<span style="font-size:12px;color:#555">{e(area_name(t["area"]))} · {e(t["ansvarlig"] or "–")} · {e(t["id"])}</span>' for t in nxt5]
    if extra:
        nxt_html.append(f"Yderligere {extra} opgaver starter samme dag ({f(nxt5[-1]['start'])})")
    check = f'<span style="color:{RAG["Grøn"]}">&#10003;</span>'
    exp = [x for x in bareas[:5]]
    rest = bareas[5:]
    bmax = max([v for _, v in bareas] or [1]) or 1
    bars = [f'{e(n)} {kr(v)}<div style="background-color:{LINE};height:6px;margin:2px 0 6px"><div style="background-color:{GREEN};height:6px;width:{round(v / bmax * 100)}%"></div></div>' for n, v in exp]
    if rest:
        bars.append(f"Øvrige {len(rest)} udgiftsområder: {kr(sum(v for _, v in rest))}")
    bt = lambda k: btot.get(k, (None, None))
    budget_html = (f"Indtægter: {kr(bt('INDTÆGTER I ALT')[0])} (regnskab 2026: {kr(bt('INDTÆGTER I ALT')[1])})<br>"
                   f"Udgifter: {kr(bt('UDGIFTER I ALT')[0])} (regnskab 2026: {kr(bt('UDGIFTER I ALT')[1])})<br>"
                   f"Resultat: {kr(bt('RESULTAT')[0])} (regnskab 2026: {kr(bt('RESULTAT')[1])})<br><br>"
                   + "".join(bars)
                   + ("<br><b>Åbne afklaringspunkter</b><br>" + "<br>".join(e(p) for p in cap(afk)) if afk else "")
                   + '<br><br><span style="font-size:12px;color:#555">Kun budgetbeløb; der er endnu ingen regnskabstal for 2027.</span>')
    rag_txt = {"Rød": "mindst ét område har forsinkede opgaver eller en deadline inden for 14 dage med lav fremdrift.",
               "Gul": "der er deadlines inden for 30 dage, datofejl eller opgaver uden ejer.",
               "Grøn": "ingen forsinkelser og ingen kritiske deadlines."}[worst]
    sections = [
        sec("Samlet status", f'<span style="background-color:{RAG[worst]};color:#ffffff;padding:2px 8px;border-radius:3px">{worst}</span> '
            f'Samlet status er {worst.lower()}: {rag_txt}'
            f'<table width="100%" cellpadding="0" cellspacing="6" style="margin-top:8px"><tr>'
            + kpi(f"{tot['Færdig']}/{tot['n']}", "Opgaver færdige") + kpi(f"{round(tot['prog'] * 100)}%", "Samlet fremdrift")
            + kpi(str(tot["Forsinket"]), "Forsinkede", RAG["Rød"] if tot["Forsinket"] else BLACK) + "</tr></table>"
            f'<div style="font-size:12px;color:#555">{tot["I gang"]} i gang · {tot["Ikke startet"]} ikke startet · {tot["df"]} datoer skal rettes</div>'),
        sec("Kræver handling", ul(cap(act_l)) or "Ingen forsinkede og ingen deadlines de næste 30 dage."),
        sec("Top-risici", ul([f"{i + 1}. {r}" for i, r in enumerate(risks)]) or "Ingen risici fundet i data."),
        sec("Beslutninger der kræves", ul(dec) or "Ingen."),
        sec("Status pr. område", "<br><br>".join(rows_area)),
        sec("Næste opgaver der starter", "<br>".join(nxt_html) or "Ingen planlagte starter."),
        sec("Ret i planen", ul(cap(fix_l, 12)) or "Intet at rette."),
        sec("Lukket siden sidst", (ul([f"{check} {e(t['aktivitet'])} · {e(area_name(t['area']))} · {e(t['id'])}" for t in new_closed])
                                   or "Ingen nye lukkede siden forrige rapport")
            + ("" if found else "<br><span style=\"font-size:12px;color:#555\">Sammenlignet med 05-11-2026.</span>")),
        sec(f"Færdige opgaver ({len(done)})", ul(cap([f"{check} {e(t['aktivitet'])} · {e(area_name(t['area']))} · "
                                                   f"{e(t['ansvarlig'] or '–')} · {e(t['id'])}" for t in done])) or "Ingen endnu."),
        sec("Budget 2027", budget_html),
        sec("Kommende møder", ul([f"{e(s)} – {DAGE[d.weekday()]} {d:%d-%m-%Y}" + (f" kl. {d:%H:%M}" if d.hour or d.minute else "")
                                  for d, s in mtg]) or "Ingen KJ27-møder de næste 14 dage."),
    ]
    head = (f'<tr><td style="background-color:{BLACK};padding:20px">'
            f'<div style="font:bold 11px Arial;letter-spacing:1px;color:{TURK}">KØGE JERNMAND 2027 · UGERAPPORT</div>'
            f'<div style="font:bold 30px Arial;color:#ffffff;margin-top:6px">{dage} DAGE TIL START</div>'
            f'<div style="font:13px Arial;color:#dddddd">Søndag 13. juni 2027 · Køge · Status {f(today)}</div></td></tr>'
            f'<tr><td><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="background-color:{TURK};height:5px"></td>'
            f'<td style="background-color:{ORANGE};height:5px"></td><td style="background-color:{GREEN};height:5px"></td></tr></table></td></tr>')
    intro = ((f'<tr><td style="padding:14px 20px 0;font:bold 14px Arial;color:{GREEN}">{e(stu_line)}</td></tr>' if stu_line else "")
             + '<tr><td style="padding:14px 20px 0;font:14px/1.5 Arial;color:#1A1819">Hej alle,<br>Her er ugens samlede status for '
               'Køge Jernmand 2027. Ret gerne ved at svare på denne mail.</td></tr>')
    end = '<tr><td style="padding:16px 20px 22px;font:14px Arial">Mange hilsner<br>Michael Mark</td></tr>'
    htm = (f'<div style="background-color:{BG};padding:12px 0"><table align="center" width="100%" cellpadding="0" cellspacing="0" '
           f'style="max-width:620px;background-color:#ffffff;border:1px solid {LINE}">' + head + intro + "".join(sections) + end + "</table></div>")

    # ---------------------------------------------------- tekst
    strip = lambda s: re.sub(r"<[^>]+>", "", html.unescape(s))
    txt = ["KØGE JERNMAND 2027 · UGERAPPORT", f"{dage} DAGE TIL START", f"Søndag 13. juni 2027 · Køge · Status {f(today)}", ""]
    if stu_line:
        txt += [stu_line, ""]
    txt += ["Hej alle,", "Her er ugens samlede status for Køge Jernmand 2027. Ret gerne ved at svare på denne mail.", "",
            f"SAMLET STATUS: {worst} – {rag_txt}",
            f"{tot['Færdig']}/{tot['n']} færdige · {round(tot['prog'] * 100)}% fremdrift · {tot['Forsinket']} forsinkede · "
            f"{tot['I gang']} i gang · {tot['Ikke startet']} ikke startet · {tot['df']} datoer skal rettes", ""]
    for title, lst in (("KRÆVER HANDLING", cap(act_l) or ["Ingen forsinkede og ingen deadlines de næste 30 dage."]),
                       ("TOP-RISICI", risks or ["Ingen risici fundet i data."]), ("BESLUTNINGER DER KRÆVES", dec or ["Ingen."]),
                       ("STATUS PR. OMRÅDE", [f"{area_name(a)} ({stats[a]['rag']}): {round(stats[a]['prog'] * 100)}% · {stats[a]['n']} opgaver · "
                                              f"{stats[a]['Færdig']} færdige · {stats[a]['Forsinket']} forsinkede · {stats[a]['df']} datofejl" for a in AREAS]),
                       ("NÆSTE OPGAVER DER STARTER", [f"{f(t['start'])}: {t['aktivitet']} · {area_name(t['area'])} · {t['ansvarlig'] or '–'} · {t['id']}"
                                                      for t in nxt5] + ([strip(nxt_html[-1])] if extra else []) or ["Ingen."]),
                       ("RET I PLANEN", cap(fix_l, 12) or ["Intet at rette."]),
                       ("LUKKET SIDEN SIDST", [f"{t['id']} {t['aktivitet']}" for t in new_closed] or ["Ingen nye lukkede siden forrige rapport"]),
                       ("BUDGET 2027", [strip(x) for x in budget_html.split("<br>") if strip(x).strip()]),
                       ("KOMMENDE MØDER", [f"{s} – {d:%d-%m-%Y %H:%M}" for d, s in mtg] or ["Ingen."])):
        txt += [title] + [strip(x) for x in lst] + [""]
    txt += [f"Færdige-id: {', '.join(t['id'] for t in done)}", "", "Mange hilsner", "Michael Mark"]
    return htm, "\n".join(txt)


def run(today, test):
    d = tempfile.mkdtemp()
    subject = f"KJ27 ugerapport {today:%d-%m-%Y}"
    try:
        if not test and gapi.search(f'in:sent subject:"{subject}"', 1):
            print(f"{subject}: allerede sendt")
            return
        import sky
        c = sky.Ctx(test)
        c.dir = d
        kj = sky.load_kj27(c)
        _, plan = sky.fetch(c, gapi.KJ27_FOLDER, "Køge-Jernmand-2027-planlægnig.xlsx")
        _, bud = sky.fetch(c, gapi.KJ27_FOLDER, "Køge-Jernmand-2027-budget.xlsx")
        _, cfg = sky.fetch(c, gapi.AUTO_FOLDER, "KJ27-områdeledere.xlsx")
        _, leaders = kj.load_config(cfg)
        to = gapi.active_recipients(leaders)
        htm, txt = build(kj, plan, bud, today)
        c.send(to, subject, txt, htm, [ME])
        print(f"{subject}: sendt til {len(to)} modtagere")
    except Exception as ex:
        gapi.send([ME], f"KJ27 ugerapport kunne ikke sendes {today:%d-%m-%Y}",
                  f"Rapporten blev ikke sendt til udvalget. Fejl: {ex}\n\nMange hilsner")
        raise
