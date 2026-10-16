# CSV–Excel riportautomatizálás – használati útmutató

Az alkalmazás árumozgási CSV-exportok alapján tölti ki az éves Excel értékesítési riportokat. A program a napi és havi nettó összegeket forintban írja be, az EUR-tételeket a tárolt napi árfolyammal váltja át.

A fejlesztési munkaterv, az ellenőrzések állapota és a részletes feldolgozási szabályok a [Statement_of_Work.md](Statement_of_Work.md) fájlban találhatók.

## A bemutatott változat

Ez a magyar nyelvű, anonimizált portfólióváltozat dokumentációja. A projekt egy valós riportkészítési folyamat automatizálását mutatja be, személynevek és ügyfélhez köthető feliratok nélkül. A bemutatóadatok fiktív ügyleteket tartalmaznak.

A magyar és az angol változat külön Git-ágon szerepel. A GitHub ágkiválasztójával lehet köztük váltani.

A magyar változatban a CSV-fejlécek, kategóriák és a felület magyar elnevezései megmaradnak. Az alábbi futtatható fájlnevek és beállítási útvonalak a jelenlegi technikai azonosítókat követik.

## Indítás és első beállítás

1. Indítsd el az átadott `AxelProRiport.exe` fájlt. A csomagolt alkalmazáshoz nem kell Pythont telepíteni, és külső sablonfájl sem szükséges.
2. Első indításkor tallózással válaszd ki a már létező, számodra írható mappát, ahová az éves riportokat szeretnéd menteni. Például hozz létre egy **Értékesítési riportok** mappát a **Dokumentumok** alatt. A `C:\` gyökerét ne válaszd célmappának; a normál használathoz nem kell rendszergazdaként indítani a programot.
3. Add meg a további újalkatrész-csoportok neveit pontosvesszővel elválasztva, például `Szűrők;Kenőanyagok`, majd kattints a **Kész** gombra. Ha nincs további csoport, ezt a mezőt üresen hagyhatod.
4. Megjelenik a főmenü. A beállításokat a program megőrzi a következő indításhoz.

Az `Új alkatrész` csoportot a program eleve új alkatrészként kezeli, ezt nem kell külön felvenni. A csoportnevek felismerése kis- és nagybetűtől független. A nevek az exportban szereplő csoportnevekkel egyezzenek meg.

Az `AxelProRiport-dev.exe` azonos grafikus felületet indít konzolablakkal; a normál használathoz az `AxelProRiport.exe` szolgál.

## CSV előkészítése és szerkesztése

Az eredeti exportot közvetlenül importálhatod. Ha előtte meg szeretnéd nézni vagy módosítani a tartalmát, **Rons Data Edit** használatát javasoljuk, amelyet a Windowsos próbák során is használtunk.

A Rons Data Edit táblázatként mutatja a fájlt, az értékeket szövegként kezeli, és támogatja az export UTF–16 kódolását. Ez azért fontos, mert az Excel CSV-megnyitás és újramentés közben automatikusan átalakíthatja a dátumokat, számokat vagy azonosítókat; az ilyen módosítás miatt a riportprogram elutasíthatja a fájlt.

- [Hivatalos termékoldal](https://www.ronsplace.ca/products/ronsdataedit)
- [Hivatalos letöltés](https://www.ronsplace.ca/products/download) – a **Rons Data Edit** résznél válaszd a **Download Install** lehetőséget. A **Download Portable** változat telepítés nélküli használatra szolgál.
- Az ingyenes **Lite** változat legfeljebb **2500 sort tud menteni**. Nagyobb fájl módosításához a korlátozás nélküli változat szükséges. Ez a CSV-szerkesztő korlátja, nem a riportprogram importkorlátja.

Szerkesztéskor őrizd meg az **UTF–16 kódolást**, a **tabulátoros elválasztást**, az eredeti oszlopfejléceket és a változatlan mezők szövegét. Ha a szerkesztő beállítást kér, az elválasztó legyen **Tab**, a kódolás **UTF–16**. Ne futtass automatikus dátum- vagy számátalakítást. CSV-t Excelből ne ments újra ehhez az importhoz; az elkészült `.xlsx` riportokat viszont Excelben használhatod.

A szükséges oszlopok:

| Oszlop | Elvárt tartalom |
| --- | --- |
| Árumozgás időpontja | Az export eredeti dátumformátuma, például `2026. 09. 18. 10:00:00`. |
| Ügylet típus | Az export szerinti ügylettípus. A leltártételeket az import kihagyja. |
| Bizonylatszám | A feldolgozott tétel nem üres bizonylatszáma. |
| Partner | Kötelező oszlop; a két kizárt partner tételeit a program automatikusan kihagyja. |
| Csoport | Csoportnév; üres számlacsoportnál választást kér, az üres csoportú beszerzés összegét kihagyja. |
| Cikkszám | A tétel cikkszáma; a `SANY`, `DIAG`, `KJ` kódokhoz külön havi összeg tartozik. |
| Ne. érték | Nettó érték, például `1250,50`. |
| Pénznem | `HUF` vagy `EUR`. |

### Automatikusan kihagyott tételek

A program a leltártételek mellett **ügylettípustól függetlenül kihagyja** azokat a sorokat, amelyek Partner mezőjében az alábbi nevek valamelyike szerepel:

- Cég1
- Cég2

A `Cég1` és `Cég2` fiktív megnevezések: a publikus CSV-beolvasó kizárási értékeinek és a próba-CSV-k partnerneveinek is ezekkel kell egyezniük. A táblázat feliratainak átnevezése önmagában nem módosítja a beolvasó szűrését.

Az összehasonlítás teljes névegyezést használ, a szélső szóközöket és a kis-/nagybetűk eltérését figyelmen kívül hagyja. E tételek nem kerülnek az összesítésekbe vagy a bizonylatnyilvántartásba. Kizárt partner üres csoportú számlájánál sem kér a program besorolást.

A **beszerzésekből csak a csoporttal rendelkező tételek összege számít bele** a napi beszerzési összegbe. Az üres vagy csak szóközös csoportú beszerzések kimaradnak ebből az összegből, de bizonylataik feldolgozottként nyilvántartásba kerülnek. Beszerzésnél nincs csoportválasztó ablak. A szabály HUF- és EUR-tételekre egyaránt érvényes.

## Importálás

1. Zárd be Excelben az érintett éves riportokat, hogy a program menteni tudjon beléjük.
2. Kattints az **Importálás** gombra, és tallózd be a CSV-fájlt.
3. Várd meg a feldolgozást. Az importgomb addig nem használható. Több év esetén a program külön éves riportokkal dolgozik.
4. Ha a program csoportválasztást vagy árfolyamot kér, töltsd ki a megjelenő ablakot az alábbiak szerint.
5. Az évenkénti előnézetben ellenőrizd a célfájlt és a hozzáadandó összegeket.
6. **Elfogadás** esetén a program elmenti az adott évet. **Elutasítás** esetén annak az évnek a változásait nem menti, és folytatja a következő évvel.
7. A folyamat végén az importgomb ismét használható. A mentett riportot megnyithatod Excelben, ahol a képletek újraszámolódnak.

A fájl neve `Napi_értékesítési_riport_ÉV.xlsx`, például `Napi_értékesítési_riport_2026.xlsx`. Új fájlban mind a tizenkét hónap, az **Összesítő**, az **Euro**, valamint a rejtett árfolyam- és bizonylatnyilvántartás szerepel.

### A havi riport használata

A havi munkalapokon az **első három sor és az A–C oszlop rögzített**. Görgetéskor ezek láthatók maradnak; az első görgethető cella **D4**.

A bemutató sablon egyes mezői semleges nevet vagy `-` feliratot kaptak. A kötőjel csak a feliratot helyettesíti: a mögöttes cellák és a megmaradt képletek továbbra is működnek.

### Hiányzó csoport

Üres csoportú, partner alapján nem kizárt számlatételnél az ablak kiírja a bizonylatszámot és a cikkszámot. Válassz a hat gomb közül:

- Új alkatrész
- Bontott alkatrész
- Munkadíj
- Szállítási költség
- Trailer
- Egyik sem

Az **Egyik sem** üresen hagyja a csoportot. Ettől a tétel csoportalapú összesítésbe nem kerül, de a külön cikkszám szerinti havi összesítés továbbra is érvényesülhet. A választás csak az adott import tételét módosítja, az eredeti CSV-t nem írja át. A választóablak **X** gombja megszakítja az importot.

### Árfolyam megadása

A program lekéri az aktuális MNB EUR/HUF árfolyamot, és felhasználja a riportokban már tárolt napi árfolyamokat. Hiányzó naphoz a legutóbbi, a célnapnál nem későbbi ismert forrásdátumú árfolyamot használja. Ha ilyen nincs, vagy az aktuális lekérés nem sikerül, kézi bevitelt kér.

Adj meg pozitív árfolyamot, például `380,25`, és valódi forrásdátumot `ÉÉÉÉ-HH-NN` formátumban, például `2026-09-18`. Kézi bevitelkor a forrásdátum eltérhet a tétel napjától, akár későbbi is lehet. Hibás szám vagy dátum esetén a program új bevitelt kér.

Az **Euro** munkalap havi árfolyamait a kézzel kitöltött EUR-mezők képletei használják. Új munkafüzetben ezek nulla értékkel indulnak. Az import csak az aktuális év aktuális hónapjának havi árfolyamát frissíti; egy korábbi hónap kézi EUR-összegeinek használata előtt ellenőrizd és szükség szerint töltsd ki az adott hónap árfolyamát. A CSV EUR-tételeinek napi átváltása ettől külön működik.

### Ismételt import és megszakítás

A program **év és bizonylatszám alapján** tartja nyilván a már importált bizonylatokat. Ugyanazon teljes fájl ismételt importja nem növeli újra az összegeket. Az ismert és új bizonylatokat együtt tartalmazó fájlból csak az új bizonylatok tételei kerülnek be.

Már importált bizonylathoz utólag hozzáadott vagy módosított CSV-sor nem frissíti a riportot: az adott bizonylat összes tétele kimarad. Az ilyen helyesbítést külön kell egyeztetni; a CSV újraimportálása nem korrekciós művelet.

A főablak bezárása megszakítást kér. A program megvárja a háttérfeldolgozás leállását. A még el nem mentett évek változásai nem maradnak meg, a korábban már elfogadott és sikeresen mentett évek viszont igen.

## Főmenü

| Gomb | Működés |
| --- | --- |
| Importálás | CSV kiválasztása, feldolgozása és évenkénti előnézet. |
| Újalkatrész-csoport hozzáadása | További újalkatrész-csoportok megadása pontosvesszővel elválasztva. |
| Újalkatrész-csoport eltávolítása | Korábban felvett további csoportok eltávolítása. |
| Riportmappa módosítása | Másik létező célmappa kiválasztása. A korábbi riportokat nem helyezi át. |
| Jelenlegi beállítások megtekintése | A riportmappa és a további újalkatrész-csoportok megjelenítése. |
| Kilépés | Az alkalmazás bezárása. |
| Alaphelyzetbe állítás és bezárás | Megerősítés után a beállítások törlése és az alkalmazás bezárása. |

A hozzáadás és eltávolítás ablakában legalább egy csoportot adj meg. Hibás bevitel esetén javítsd az adatokat, vagy válaszd a **Megszakítás** lehetőséget.

**Az alaphelyzetbe állítás nem törli a riportmappát vagy az Excel-riportokat.** A mentett célmappa-beállítást és a csoportlistát törli. Következő indításkor a program újra bekéri őket. Ha ugyanazt a riportmappát választod, az ottani éves fájlok árfolyam- és bizonylatnyilvántartása megmarad.

## Hibák és figyelmeztetések

Hiba esetén olvasd el az üzenetet, majd zárd be az **OK** gombbal. Hibás CSV-nél az üzenet a problémás mezőt, és sorhiba esetén a sorszámot is megjelöli. Javítsd a fájlt Rons Data Editben, vagy készíts új exportot, majd indítsd újra az importot.

Ha a program nem tud menteni, ellenőrizd, hogy a riport nincs-e megnyitva Excelben, és a célmappa elérhető-e, illetve írható-e. Többéves importban egy későbbi mentési hiba előtt sikeresen mentett évek megmaradnak.

A munkafüzet szerkezeti vagy képlethibája megszakítja az importot. Az üzenet alapján állítsd helyre a megjelölt munkalapot vagy cellát. A havi lapok közül az importtal érintett hónapokat ellenőrzi a program.

Formázási vagy méreteltérés esetén az import folytatódhat. A program csak az első ilyen figyelmeztetést mutatja meg; a megjelölt helyen ellenőrizd a táblázatot. Az Excel egyszerű újramentéséből eredő, jelentésükben azonos képlet- és formázásváltozatokat elfogadja.

## A riportok megőrzése

A riportokban a kézi adatbevitelre szánt mezők továbbra is kitölthetők Excelben. A munkalapneveket, dátumokat, fejléceket, összevonásokat, képleteket és rejtett nyilvántartásokat őrizd meg.

A program ideiglenes fájlon keresztül ment, **automatikus biztonsági másolatot nem készít**. A riportok és a bennük tárolt bizonylatnyilvántartás megőrzéséről a felhasználó gondoskodik. A program nettó összegeket összesít és egész forintra kerekít; részletekben importálva a kerekítés miatt kis eltérés lehet az egyszeri teljes importhoz képest.

A beállításfájl Windows alatt: `%APPDATA%\axelpro_riport\settings.json`.
