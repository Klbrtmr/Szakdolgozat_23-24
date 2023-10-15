# Szakdolgozat tervezet

## PLC áramkörök jeleinek feldolgozása és monitorozása
## *Processing and monitoring PLC circuit signals*
### *Szakdolgozat tématerv*
A rendszer feladata, hogy képesek legyünk PLC áramkörök jeleinek feldolgozására Excel fájlokból beimportálni, a program fel tudja dolgozni a benne szereplő adatokat, majd pedig kívánt diagramon megjelenítse őket. A diagramokat tudjuk szerkeszteni, például színét, stílusát. Illetve ezeket a diagramokat tudjuk majd exportálni. A PLC-k bizonyos lépésközönként tudnak mérni, és diagram kirajzolás esetén fogunk tudni váltani lépésköz (azaz „sample”) vagy idő („time”) mód között, ez a tengelyeken fog változtatni elsősorban.
A feladat alapötlete az Evosoft Hungary Kft. cégtől ered, ahol dolgozom gyakornokként, és napi szinten foglalkozom PLC eszközök által adott jelek diagramon való kirajzolásával.
### Importálás
Vezetett lépéseken keresztül az input adatokat tartalmazó Excel fájl beolvasása (melyik cellától melyik celláig, milyen adatokat olvasson be, illetve milyen stílusú diagramm készüljön belőle). Az adatok beolvasását és feldolgozását egy C# osztálykönyvtárral tervezem megvalósítani.
### Diagramm szerkesztése
Színek, stílusok szerkesztésére, megváltoztatására való lehetőség. 
### Exportálás
Az elkészített diagramm képként való exportálása. Tervként szerepel egy plusz fájl exportálása is, ami tartalmazza a diagram adatait, hogy a diagram is importálható legyen, és folytatható legyen a szerkesztése. Exportálás bővítése szerepel a tervek között, miszerint az elkészített diagramokat nyomtatni és akár PDF fájlokba is exportálhatóak legyenek.

### Technikai követelmények
C# .NET Framework 4.8

### Fejlesztési sprintek megtervezése
Egyszemélyes agilis módszerben dolgozva minden sprint végén működő szoftver az elvárt. Folyamatos tesztelési lehetőségek. Scrum board használata.

| Hónap |	Feladat |	Risk |
|:-------:|:---------:|:------:|
| Szeptember | Feladatok, tervek előzetes meghatározása | Alacsony |
| Október |	Minimális GUI létrehozása |	Alacsony |
| November |	Importálási lépések előre meghatározott sémán létrehozott excelen |	Közepes |
| December |	Importálási lépések előre meghatározott bármilyen sémán létrehozott excelen |	Nagyon magas |
| Január |	Diagram exportálás és szerkeszthetőségének létrehozása |	Magas |
| Február |	Plusz diagram (tervezetten xml) fájl létrehozása importálás miatt |	Közepes |
| Március |	Plusz fájl importálásának megvalósítása, szakdolgozatírás |	Közepes |
| Április |	Tesztelés, bug fixing, szakdolgozatírás |	Alacsony |
| Május |	Kész szoftver, szakdolgozatírás |	Közepes |


Szeged, 2023. október 3.
