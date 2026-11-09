# Timeafstemning

Start den lokale app med `python -m streamlit run app.py` fra projektmappen.
Vælg **1:1** (22 timer/dag) eller **2:1** (44 timer/dag), og vælg HTML eller
én Excel-fil med personfaner. **Normal vagtkalender** bruger den oprindelige
eksport uden ekstra faner.

HTML behandles først af den eksisterende eksport. Excel-input læses direkte.
Afstemningen bruger kun personfanernes reducerede timer og omfatter alle
institutioner under én dagsnorm. Resultatet er et øjebliksbillede i fanerne
**Samlet** og **Samlet Oversigts**. Ændringer i personfanerne efter eksport
opdaterer ikke automatisk afstemningen; upload filen igen for at beregne på ny.

Formeltimer skal have et gemt resultat. Ved fejl med cellehenvisning: genberegn
og gem arbejdsbogen i Excel, og upload den igen. Eksisterende personformler og
deres gemte resultater bevares. Uploadfilen ændres ikke.

Download er tilgængelig efter Streamlit-genkørsler i samme session. Ændrede
input eller indstillinger, et nyt eksportforsøg og **Ryd felter** fjerner det
tidligere downloadresultat.

## Verifikation

Kør fra projektmappen:

```text
python -m pytest -q -p no:cacheprovider --basetemp=.test-tmp-local
```

Testene bruger syntetiske data. De oprindelige 64 tests bestod før ændringerne.
Eksport- og beregningskoden i `workhours_tool/scraper.py` er ikke ændret.

Den lokale accepttest afventer eksempelfilen. Opbevar den eventuelt i
`local_acceptance/`, som er udelukket fra versionsstyring. Kontrollér:

| Kontrol | Forventning |
|---|---|
| Vagtrækker | 79, eksklusive overskrift, eventuelle tomme dage og total |
| Periode | 16.09.2026–15.10.2026, 30 kalenderdage |
| Aktuelle timer | 727,87 |
| 1:1 | 660,00 optimale timer; +67,87 difference |
| 2:1 | 1.320,00 optimale timer; −592,13 difference |

Der er ikke oprettet eller ændret et GitHub-repository eller foretaget deployment.
