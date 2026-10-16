# Hiányzó csoport és hibás CSV próbák

Külön üres riportmappát használj. Minden fájl UTF–16, tabulátorral tagolt; a helyes tételek HUF-ban vannak.

## Hat választás
A hianyzo_csoport_hat_valasztas.csv hat számlatételt tartalmaz, hat külön bizonylattal. A felugró ablakok CSV-sorrendben jönnek:

| Cikkszám | Választás | Nettó érték |
|---|---|---:|
| VALASZD-UJ | Új alkatrész | 1000 Ft |
| VALASZD-BONTOTT | Bontott alkatrész | 2000 Ft |
| VALASZD-MUNKADIJ | Munkadíj | 3000 Ft |
| VALASZD-SZALLITAS | Szállítási költség | 4000 Ft |
| VALASZD-TRAILER | Trailer | 5000 Ft |
| VALASZD-EGYIK-SEM | Egyik sem | 6000 Ft |

Az előnézet 2026. szeptember 18-ra új alkatrész 1000, bontott 2000, munkadíj 3000, szállítás 4000 és Trailer 5000 Ft növekményt mutasson. A többi napi összeg és a havi SANY/DIAG/KJ nulla. Műhely-alkatrészösszeg nincs, mert a munkadíj más bizonylaton szerepel. Új bizonylatok száma: 6; az Egyik sem választás tétele is nyilvántartásba kerül, de nincs összegzési kategóriája. Mentés után a riportot újranyitva ellenőrizd a napi értékeket.

## X-es bezárás
A hianyzo_csoport_megszakitas.csv importjánál zárd be X-szel a csoportválasztó ablakot. Megszakításról szóló üzenetet várunk, nem importhibát. Nem lehet előnézet vagy mentés; az importgomb újra használható. Újbóli importkor ugyanaz a tétel újra választást kérjen.

## Nem számla ügyletek
A hianyzo_csoport_nem_szamla.csv nem kérhet csoportot: a beszerzés 700 Ft-tal növeli a kiadásokat, a leltár kimarad. Új bizonylatok száma: 1.

## Hibás CSV
A hibas_csv_netto.csv, hibas_csv_datum.csv, hibas_csv_penznem.csv és hibas_csv_bizonylat.csv fájlokban az első adatsor helyes, a második szándékosan hibás. A hiba a fájl 3. sorára mutasson, és nevezze meg a hibás értéket vagy a hiányzó bizonylatszámot. Előnézet és mentés nem történhet, az első helyes sort sem szabad részlegesen importálni.

A hibas_csv_hianyzo_oszlop.csv fájlból a Pénznem fejléc hiányzik. Hiányzó CSV-oszlopra vonatkozó üzenetet várunk; ehhez nem szükséges sorszám.

A hibaablak bezárása után a program új CSV-választást indíthat. Válaszd a csv_jo_ujraprobalashoz.csv fájlt: 1000 Ft újalkatrész-növekménnyel és 1 új bizonylattal importálható legyen. Ez a jó fájl ugyanazt a bizonylatot tartalmazza, mint a hibás fájlok első helyes sora, így bizonyítja, hogy a hibás beolvasás nem mentett részlegesen. Minden hibásfájl-próbát külön üres mappában végezz, vagy a jó próbafájlt csak az összes hibás próba után importáld egyszer.
