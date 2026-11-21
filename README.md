# Ladetid data

Jon Pedersen, november 2026

Data til appen Ladetid. To gange om dagen gemmes de prognoser for elprisen, som appen bruger, sammen med vejrudsigten for vind, sol og temperatur og om dagen er en helligdag. Prognoserne sammenlignes med de faktiske priser, når de kommer. Ud fra fejlene lærer programmet en rettelse, som appen bruger, hvis den har ramt bedre end den rå prognose de sidste 4 uger. Rettelsen flytter aldrig prognosen mere end 10 % og højst 15 øre væk fra elpriser.org og Strømligning.

## Filer

* `data/prognoser/` gemte prognoser, én fil pr. dag og prisområde
* `data/faktiske_priser.json` faktiske spotpriser fra Energi Data Service
* `data/rapport.json` hvor godt prognoserne har ramt de sidste 4 uger
* `data/rettelse.json` rettelsen som appen henter
* `excel/` en Excel fil pr. uge med alle fejl. `seneste.xlsx` er den nyeste hele uge, `denne-uge.xlsx` er ugen indtil nu

## Kortsigtet og langsigtet

* Kortsigtet: prognoser lavet 1 til 4 dage før
* Langsigtet: prognoser lavet 5 til 9 dage før

Alle priser er spotpris i kr pr. kWh uden moms, tariffer og afgifter.

Data: [Energi Data Service](https://www.energidataservice.dk) (Energinet, CC BY 4.0), [elpriser.org](https://elpriser.org) og [Strømligning](https://stromligning.dk).
