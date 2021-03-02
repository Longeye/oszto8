Nyolcadikos szóbeli felvételi vizsgabeosztó program

A program beolvas egy megfelelően előkészített Excel fájlt, majd az abban található adatok alapján létrehoz egy másikat, amiben a felvételizők vizsgabeosztása látható, különböző munkalapokon. Mivel az eredeti fájlhoz nem nyúl, az szabadon módosítható és a program újra lefuttatható rajta. Így ki lehet kísérletezni a lehető legoptimálisabb vizsgabeosztást.
Az Excel fájlt a központilag kapott táblázatból kell átalakítani, a következő képpen:

1. A fölösleges oszlopokat töröljük, csak az OM azonosítók, a nevek és a tantárgyak vizsgáit jelölő, "x"-eket tartalmazó oszlopok maradjanak.
2. Le kell szűrni azokat, akiknek szóbeli vizsgát kell tenniük. A "minta.xlsx" fájlban található erre vonatkozóan "példamódszer" az "Export" munkalapon.
3. A leszűrt adatokat másoljuk át egy új munkalapra!
4. Az új munkalapot át kell nevezni "Adatok"-ra. Figyelem! A kisbetű/nagybetű számít!
5. Be kell szúrni egy új oszlopot a neveket tartalmazó oszlop mögé. Tehát az A oszlopban lesznek az OM azonosítók, a B oszlopban a nevek, a C pedig egy új, üres oszlop lesz. A többi tartalmazza az "x"-eket.
6. Opcionális: a C oszlopba be lehet írni, hogy egy-egy tanuló legkorábban hanyadikként kerüljön sorra. Ez akkor lehet hasznos, ha tudjuk valakiről, hogy nem érhet be a vizsgára korábban, mert pl. külföldről érkezik, vagy előre jelezte a késést.
7. Át kell nevezni a vizsgakódokat tartalmazó fejléceket a vizsgatárgyaknak megfelelően. Figyelem! Az ide beírt feliratokat máshol is használjuk, ezért nem lehet benne elgépelés! Az üres oszlopokat, vagyis amiből nem lesz szóbeli vizsga, törölni kell.
8. Létre kell hozni a vizsganapokhoz tartozó munkalapokat. Ezeket úgy kell elnevezni, hogy azok utaljanak a napra és ugyanezt a nevet, a programba is be kell írni az elején található "napok" nevű tömbbe (listába), az ott látható módon.
9. A vizsganap munkalapját a "minta.xlsx" nevű Excel tábla alapján kell kitölteni. Az egyes napokhoz tartozó munkalapoknak nem kell egyformának lenniük, bár kezdetben érdemes úgy indulni, hogy lemásoljuk az elsőt és a keletkezett beosztások alapján módosítjuk a többit.
Az első sor tartalmazza a bizottságok neveit, a második sor a vizsgaterem megnevezését, a harmadik azt az épületrészt, ahol a vizsga zajlik.
Ez nem jelenik meg sehol, de a vizsgák beosztásánál a program figyelembe veszi, ezért ugyanazt az épületrészt ugyanúgy kell elnevezni. Nem lehet egy karakternyi eltérés sem, mert az már külön épületrészként értelmeződik.
A negyedik sorban a tantárgy neve látható, aminek meg kell egyeznie az "Adatok" munkalapon szereplőkkel.
Az egyes bizottságok rugalmasan változtathatóak. Bármikor lehet törölni egyet, vagy újat felvenni. Nem kell egymás mellett lennie az azonos tantárgyú bizottságoknak. Az oszlopok sorrendje tetszés szerint változtatható.
A táblázat első oszlopában fel kell venni az időpontokat. Itt addig lehet lehúzni a sort, ameddig vizsgáztatni akarunk. A program ezt figyelembe veszi. Az kezdőidőpont is lehet bármi, nem kell reggel 8:00-tól kezdődnie. A későbbiekben az itt található időpontoknak megfelelően készülnek el a beosztások.
A program a táblázatban található üres mezőket fogja feltölteni, tehát, ha bármit írunk bele, azt a mezőt figyelmen kívül hagyja. Ha za első három mezőbe beírjuk, hogy felkészülési idő, akkor oda nem oszt be feleletre senkit és úgy veszi, hogy ezek a tantárgyak felkészülési időt igényelnek. Erre mindíg figyelni fog.
Ebédidőt is beiktathatunk, csak a megfelelő sávban minden cellába be kell írni, hogy ebédidő, vagy kaja, vagy bármi mást.

A diákok közötti névazonosságot úgy kezeli, hogy az azonos nevűek neve mögé berakja az OM azonosító utolsó négy karakterét, zárójelben.

A létrejövő munkalapok:
- Felvételizők vizsgái napokra bontva
- Vizsgabizottságok időbeosztása

A vizsgabeosztás szempontjai:
- egy diáknak csak max. 3 vizsgája lehet egy nap. Ha több van, akkor a "többletet" átdobja a következő napra
- lehető leggyorsabban végezzen a felvételiző
- amennyiben megoldható, ugyanabban az épületrészben (folyosó, szárny, stb.) maradjon
- amelyik tantárgynál szükséges, mindenhol biztosítja a 30 perces felkészülési időt
- egy feleletnyi idő alatt kell átérnie a következő helyszínre a felvételizőnek

Jelenlegi korlátok:
Nagyon oda kell figyelni azoknál a tantárgyaknál, ahol felkészülési idő van, hogy a kiadott sorrendnek megfelelően kezdődjön a szóbeliztetés. Még akkor is, ha valaki jelzi, hogy hamarabb készen áll a megmérettetésre. Ha nem így történik, borulhat az egész beosztás, mert a diák nem ér oda a következő vizsgájának a helyszínére. Ez mondjuk kézi beosztás készítésénél is fennáll...
Régi megszokás, hogy azoknál a tantárgyaknál, ahol felkészülési idő van, behívják az első három vizsgázót és azok egyszerre kezdik a tételek kidolgozását. Emiatt a harmadik felelőnek majdnem egy órája is van a felkészülésre. Eredetileg a program úgy működött, hogy ezeket a vizsgázókat arra a 20 perc többletidőre beosztotta előtte egy felkészülési időt nem igénylő tantárgyhoz, így nekik is csak 30 perc marad a felkészülésre. További előny még, hogy a diák is 20 perccel hamarabb végez és a bizottság is.
Most, kérésre a hagyományos rendet követi a program, de a következő verziókban tervezem opcióként a beállíthatóságát.

További tervek, ötletek:
- az objektumhierarchia és az osztályok szerkezetének átgondolása, letisztítása (v2, pipa)
- a kód karbantartása, optimalizálása, amit az első verzióban már nem tudtam, mertem megtenni (v2, pipa)
- önellenőrző programrész fejlesztése, ami visszaellenőrzi a vizsgabeosztást. (Jelenleg csak excelben, megfelelő függvényekkel lehet. Ez macerás és hosszadalmas.)
- grafikus felület
- a forrásfájl automatikus feldolgozása, azaz nem lesz szükség előkészítő munkára, illetve csak minimálisra (v2, pipa)
- e-mailes értesítés kiküldése a jelentkezőknek a vizsgabeosztásukról
- lehetőség legyen időpont foglalásra.
- jelenleg a várakozási idő mindenhol 3 "vizsgaegységnyi". Ez is rugalmasan változtatható legyen (v2, pipa)
- felkészülési időoptimalizáció opcionálisan bekapcsolható legyen
- lehetőség szerint, tömörítse a vizsgákat, azaz, semmiképpen se maradjanak üres vizsgahelyek, még akkor sem, ha nincs elég egyvizsgás diák, akikkel fel lehetne tölteni
- az első vizsganap dátuma és a vizsganapok száma alapján a forrásadatok beolvasásakor kialakítsa a naponkénti vizsgabeosztások elkészítéséhez szükséges üres táblázatokat
- a kimenet formázottan jelenjen meg választható módon HTML és/vagy PDF formátumban. A formázáshoz lehessen megadni Excel táblát, amit sablonként használ.

