"""Startpunkt for GitHub Actions.
Planlagte kørsler: hver tid findes to gange i workflowet (sommer- og vintertid, fordi GitHub kun kender UTC).
Kun den, der passer med dansk tid i dag, kører videre.
Manuel kørsel: JOB = sky-morgen | sky-aften | ugerapport, TEST = ja/nej.
"""
import datetime as dt, os, sys
from zoneinfo import ZoneInfo

DK = ZoneInfo("Europe/Copenhagen")
# cron (UTC) -> (job, krævet UTC-forskydning i timer)
SCHEDULES = {
    "47 4 * * *": ("sky-morgen", 2), "47 5 * * *": ("sky-morgen", 1),
    "47 18 * * *": ("sky-aften", 2), "47 19 * * *": ("sky-aften", 1),
    "23 5 * * 1": ("ugerapport", 2), "23 6 * * 1": ("ugerapport", 1),
}


def main():
    now = dt.datetime.now(DK)
    today = now.date()
    sched = os.environ.get("SCHEDULE", "").strip()
    test = os.environ.get("TEST", "nej").strip().lower() in ("ja", "yes", "true", "1")
    if sched:
        job, off = SCHEDULES[sched]
        if now.utcoffset() != dt.timedelta(hours=off):
            print(f"{sched}: ikke dansk tid i dag (sommer/vintertid) – springes over")
            return
    else:
        job = os.environ.get("JOB", "").strip()
    if os.path.exists(os.path.join(os.path.dirname(__file__), "PAUSE")) and not test:
        print("PAUSE-fil findes: kører i testtilstand (kun mail til Michael, intet skrives til Drive)")
        test = True
    if job == "sky-morgen":
        import sky; sky.morning(today, test)
    elif job == "sky-aften":
        import sky; sky.evening(today, test)
    elif job == "ugerapport":
        import ugerapport; ugerapport.run(today, test)
    else:
        sys.exit(f"Ukendt job: {job!r}")


if __name__ == "__main__":
    main()
