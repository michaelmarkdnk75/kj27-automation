# kj27-automation

Kører KJ27-opgaverne i GitHub Actions uden pc og uden tokens:

| Job | Tid (dansk) | Hvad |
|---|---|---|
| sky-morgen | dagligt 06:47 | Påmindelser til områdeledere, statusmail hver 14. mandag, budget-udsendelse ved ændring, opsummering til Michael |
| sky-aften | dagligt 20:47 | Budget-udsendelse ved ændring |
| ugerapport | mandag 07:23 | Ugentlig projektstatus til stævneudvalget |

`kj27.py` hentes fra Drive (KJ27/Automatik) ved hver kørsel, så rettelser der slår igennem med det samme.
Filer skrives altid tilbage i den samme fil på Drive (ingen kopier).

Manuel test: Actions → KJ27 → Run workflow → vælg job og `test = ja` (alt går kun til Michael, intet skrives til Drive).
