# TournamentCalculator
Beregner poeng for VM og EM, basert på et mal-excelark -> `EM2021_Fornavn_Etternavn.xlsx` som sendes ut til brukerne.

* På rota legges en katalog med navn `Leagues`.
* Inni Leagues legges kataloger for alle ligaer man ønsker, f.eks FooLiga1, FooLiga2 + Fasit.xlsx (som gjelder for alle ligaene).
* Hver ligakatalog skal inneholde en katalog `Tippeforslag` og `Resultater`.
* Alle innsendte Excelark legges i `Tippeforslag`. `Resultater` inneholder json av output fra kjøringen, og man behøver vanligvis ikke denne.

### EM - Poengscore gis etter følgende kriterier:
**Innledende kamper:** Hvert riktige tippetegn (HUB) gir 2 poeng. Riktig resultat/målscore gir 2 ekstra poeng.

**Etter gruppespillet:** Hvert lag med riktig plassering i gruppen gir 2 poeng. I sluttspillet gjelder det kun å ha riktige lag videre til hver runde, resultatene har dermed ikke noe å si for poengscoren.

**1/8-dels finaler:** Hvert riktige lag i 1/8-delsfinalene gir 4 poeng.<br/>
**Kvartfinaler:** Hvert riktige lag i kvartfinalene gir 6 poeng.<br/>
**Semifinaler:** Hvert riktige lag i semifinalene gir 8 poeng.<br/>
**Finalen:** Hvert riktige lag i finalen gir 10 poeng.<br/>
**Vinner:** Riktig vinner gir 12 poeng.<br/>

### VM - Poengscore gis etter følgende kriterier:

**Innledende kamper:** Hvert riktige tippetegn (HUB) gir 2 poeng. Riktig resultat/målscore gir 2 ekstra poeng, totalt 192 oppnåelige poeng

**Etter gruppespillet:** Hvert lag med riktig plassering i gruppen gir 2 poeng, totalt 64 oppnåelige poeng. I sluttspillet gjelder det kun å ha riktige lag videre til hver runde, resultatene har dermed ikke noe å si for poengscoren.

**1/8-dels finaler:** Hvert riktige lag i 1/8-delsfinalene gir 4 poeng.<br/>
**Kvartfinaler:** Hvert riktige lag i kvartfinalene gir 6 poeng.<br/>
**Semifinaler:** Hvert riktige lag i semifinalene gir 8 poeng.<br/>
**Bronsefinale:** Hvert riktige lag i bronsefinalen gir 10 poeng.<br/>
**Bronsevinner:** Riktig bronsevinner gir 14 poeng.<br/>
**Finalen:** Hvert riktige lag i finalen gir 12 poeng.<br/>
**Vinner:** Riktig vinner gir 16 poeng.
