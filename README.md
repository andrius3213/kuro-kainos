# Kuro kainų robotas („Mano garažas“)

Kasdien (09:23 ir 17:23 Lietuvos laiku) GitHub Actions:

1. atsisiunčia oficialų **Lietuvos energetikos agentūros (LEA)** degalų kainų Excel failą
   (nuoroda randama puslapyje https://www.ena.lt/dk-pr-pr-duomenys/);
2. kiekvienai degalinei paima naujausias 95 benzino, dyzelino ir dujų (SND/LPG) kainas
   (degalinės, kurios kainų neteikė ilgiau nei 45 d., praleidžiamos);
3. suranda koordinates per OpenStreetMap (Nominatim adresas → artimiausia OSM degalinė
   to paties tinklo). Rasti adresai įsimenami `geocache.json`, todėl kasdien ieškomi tik nauji;
4. išsaugo `fuel.json` šakoje **`data`**. Programėlė jį skaito iš
   `https://raw.githubusercontent.com/<savininkas>/<repo>/data/fuel.json`.

Saugikliai: jei degalinių < 300, kainos senesnės nei 10 d. arba jų staiga sumažėjo > 40 %,
naujas failas **neišsaugomas** – lieka vakarykštis.

## Šaltinių nurodymas (privaloma)

- Kainos: *Lietuvos energetikos agentūra (LEA)* ir degalines valdančios įmonės
  (teisinis pavadinimas saugomas masyve `L`, programėlė jį rodo prie kiekvienos degalinės).
- Koordinatės: *© OpenStreetMap bendraautoriai*, ODbL.

## `fuel.json` formatas

```
{ "v":1, "source":"…LEA…", "geo":"© OpenStreetMap…", "prices_date":"2026-11-13",
  "generated":"…", "N":[tinklai], "L":[įmonės], "C":[savivaldybės], "D":[datos],
  "s":[[N, adresas, C, lat, lng, kaina95, dyzelinas, dujos, D, L, apytikslė], …] }
```

`lat`/`lng` gali būti `null` (vieta nerasta). `apytikslė = 1` – rasta tik gyvenvietė.

## Vietoje

```
pip install -r requirements.txt pytest
python -m pytest -q tests
python robotas.py --xlsx DK-2026.xlsx --no-geo --out out/fuel.json   # be interneto paieškų
```
