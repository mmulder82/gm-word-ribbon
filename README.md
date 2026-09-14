# Gemeente Gooise Meren - Word Add-in

Office Add-in (task pane + ribbon-knoppen, gebouwd met Office.js) die een
klein setje tekststijlen toevoegt aan Word - ook in **Word op het web
(Word Online)**, niet alleen desktop-Word.

## Wat zit erin (v4 - bewust klein gehouden)

| Groep | Knoppen |
|---|---|
| Tekststijlen | Titel, Standaardtekst |
| Kopteksten | Koptekst 1 t/m Koptekst 6 (zonder nummering) |
| Kopteksten genummerd | Koptekst 1 t/m Koptekst 4 (mét automatische nummering) |
| Instellen | Stijlen installeren, Vernummeren |

Er is ook een taakvenster (klik "Gemeente-opties" op het Start-tabblad, of
"Help" op het eigen tabblad) met dezelfde knoppen als terugval-optie.

**Dit is een bewuste inperking** ten opzichte van eerdere versies: Klembord,
Tabelstijlen, Taal, Bijhouden en Bijlage zijn eruit gehaald om eerst een
kleine, betrouwbare basis neer te zetten. Zie "Waarom deze knoppen en niet
meer" hieronder voor de reden.

### Lettertype en formaat per stijl

Alles is **Corbel 11**, behalve de kopteksten en de titel. Titel en alle
kopteksten krijgen de gemeentelijke huisstijlkleur (`#441D42`);
Standaardtekst blijft zwart:

| Stijl | Lettertype/formaat | Kleur | Nummering |
|---|---|---|---|
| Titel | Corbel 18, vet | #441D42 | - |
| Standaardtekst | Corbel 11 | zwart | - |
| Koptekst 1 | Corbel 16, vet | #441D42 | - |
| Koptekst 2 | Corbel 14, vet | #441D42 | - |
| Koptekst 3 | Corbel 12, vet | #441D42 | - |
| Koptekst 4 | Corbel 11, vet | #441D42 | - |
| Koptekst 5 | Corbel 11, cursief | #441D42 | - |
| Koptekst 6 | Corbel 11 (geen extra opmaak) | #441D42 | - |
| Koptekst 1 genummerd | Corbel 16, vet | #441D42 | `1.` |
| Koptekst 2 genummerd | Corbel 14, vet | #441D42 | `1.1` |
| Koptekst 3 genummerd | Corbel 12, vet | #441D42 | `1.1.1` |
| Koptekst 4 genummerd | Corbel 11, vet | #441D42 | `1.1.1.1` |

## Hoe de stijlen en de nummering werken

**Architectuurgeschiedenis (kort):** een eerdere, grotere versie van dit
project probeerde stijlen als kant-en-klaar Word-bestand te "importeren",
en later om Word's eigen ingebouwde Kop 1-6/Standaard-stijlen te
herprogrammeren. Beide liepen in de praktijk tegen problemen aan die vanaf
hier niet betrouwbaar te reproduceren/debuggen waren: Word Online weigerde
het geïmporteerde bestand (`ooxmlIsMalformated`), en het herprogrammeren
van ingebouwde stijlen liep vast doordat Word Online geen bruikbare
`paragraphFormat`-eigenschap op een stijl-object blijkt te ondersteunen
("Cannot read properties of undefined (reading 'set')"), en doordat
zo'n stijl opzoeken een *gelokaliseerde* naam vereist (in het Nederlands
bijvoorbeeld "Kop 1" i.p.v. "Heading 1").

Deze v4-versie vermijdt beide problemen bewust:

- **Alle stijlen hierboven zijn eigen, custom stijlen** (aangemaakt via
  Word's eigen `context.document.addStyle(...)`), nooit Word's ingebouwde
  Kop-/Standaard-stijlen. Daardoor is er geen gelokaliseerde naam nodig om
  een stijl terug te vinden - de naam is precies wat wij hebben gekozen
  ("Titel", "Koptekst 3", "Koptekst 2 genummerd", ...).
- **Er wordt alleen het lettertype (`font`) op een stijl gezet**, nooit
  `paragraphFormat` - dat laatste bleek niet betrouwbaar te zetten op een
  stijl-object. Dat is voor deze opgave ook niet nodig: er was geen
  regelafstand/witruimte gevraagd, alleen lettertype/formaat/vet/cursief
  en nummering.
- **Elke stijlknop past de opmaak bovendien rechtstreeks toe op de
  geselecteerde tekst zelf**, náást het bijwerken van de gedeelde
  stijl-definitie. Zo werkt de knop altijd zichtbaar correct, zelfs als er
  ooit weer iets misgaat met de gedeelde stijl.
- **Elke stijlknop herstelt de opmaak bij elk gebruik**, niet alleen de
  eerste keer - een stijl die al bestond maar verkeerd was opgemaakt wordt
  dus gewoon gecorrigeerd in plaats van overgeslagen.
- De nummering (`1.` / `1.1` / `1.1.1` / `1.1.1.1`) is **geen "levende"
  Word-multilevel-lijst** - de add-in berekent het nummer zelf en zet het
  als platte tekst + tab vooraan de alinea, opnieuw voor het hele document
  elke keer dat je een genummerde knop gebruikt. Kopteksten verplaatsen,
  verwijderen of kopiëren/plakken betekent dus: klik daarna op
  **"Vernummeren"** (in het tabblad of het taakvenster) om de nummers weer
  kloppend te maken - dat gebeurt niet vanzelf.
  - **Dit was eerst anders (v4.4) en is bewust teruggedraaid (v4.5).**
    Koptekst 1-4 genummerd gebruikten kort Word's eigen, "levende"
    nummering (`Word.List`) in plaats van deze platte tekst. Dat werkte
    wel automatisch, maar de Word JavaScript API bleek geen ondersteunde
    manier te hebben om een niveau de nummers van de niveaus erboven te
    laten meetonen - `Word.List.setLevelNumbering()` stelt alleen het
    format van een niveau's éígen, onafhankelijke teller in. Het resultaat
    was dus "1., 2., 3., ..." per niveau, niet de samengestelde
    "1.1"/"1.1.1"-notatie.
  - Toen is uitgezocht of Word Online zelf (buiten de add-in-API om) een
    manier biedt om dit wél te krijgen (Word's "Koppelen niveau aan
    stijl"-functie bij een meerniveaulijst, die dit normaal automatisch
    zou doen). Uitkomst: nee - Word Online's eigen meerniveaulijst-
    functionaliteit is zelf beperkt en volgens Microsoft's eigen
    supportforums en een MVP onbetrouwbaar (lijsten die niet correct
    herstarten, een MVP die zegt dat Word Online meerniveaulijsten kan
    "vernietigen"); desktop Word wordt aangeraden voor dit soort werk.
    Er was dus geen Word Online-native mechanisme om op terug te vallen.
  - Conclusie: de platte-tekst-aanpak (met "Vernummeren") is de enige
    manier gebleken om betrouwbaar de exacte gevraagde notatie te krijgen
    in Word Online. Niet "levend", maar wel voorspelbaar en correct.

## Word's eigen stijlen (Normaal, Kop 1-9, ...) neutraliseren

Om te voorkomen dat je per ongeluk een van Word's eigen ingebouwde stijlen
te pakken krijgt en zo buiten Corbel om typt, doet de knop **"Stijlen
installeren"** meer dan alleen onze eigen stijlen aanmaken: hij overschrijft
óók Word's eigen ingebouwde stijlen, zoals die in Word Online's
stijlengalerij staan:

- **"Normaal"** (soms "Standaard" in oudere Word-versies - beide worden
  geprobeerd) krijgt dezelfde opmaak als Standaardtekst; **"Kop 1"** t/m
  **"Kop 9"** krijgen dezelfde opmaak als de overeenkomstige Koptekst
  (Kop 7-9 hebben geen eigen knop, die krijgen dezelfde opmaak als
  Koptekst 6).
- Alle overige stijlen uit Word Online's standaardgalerij - **Geen
  afstand, Ondertitel, Nadruk, Sterk, Citaat, Subtiele verwijzing,
  Intensieve verwijzing, Titel van boek, Lijstalinea** - krijgen alléén
  hun lettertype overschreven naar Corbel; hun eigen grootte, kleur,
  cursief/vet etc. blijft ongemoeid (dat was hier ook alles wat gevraagd
  was: "als lettertype Corbel te hebben").

Pak je per ongeluk toch een van deze ingebouwde stijlen, dan staat de tekst
dus nog steeds (in elk geval) in Corbel.

**Belangrijke beperking, eerlijk gezegd:** er bestaat geen ondersteunde
manier in de Word JavaScript API om een ingebouwde stijl echt te
verwijderen, of om Word's eigen documentbeveiliging "Opmaak beperken tot
geselecteerde stijlen" aan te zetten (dát zou pas echt voorkomen dat de
andere stijlen te kiezen zijn) - dat zit domweg niet in de API die een
add-in tot zijn beschikking heeft. Wat de add-in wél probeert, als extra,
best-effort stap: de ingebouwde stijlen uit Word's "Snelstijlen"-galerie in
het lint halen. Dat lukt mogelijk niet in elke Word Online-versie (dit
project is al eerder tegen een niet-ondersteunde stijl-eigenschap
aangelopen, zie de architectuurgeschiedenis hieronder) - en zelfs als het
lukt, blijven alle ingebouwde stijlen gewoon nog vindbaar en te kiezen via
het volledige "Stijlen"-deelvenster (Ctrl+Alt+Shift+S). Kortom: dit is een
praktische vangnet-oplossing (altijd Corbel, ook bij een verkeerde keuze),
geen echte, harde blokkade van de andere stijlen.

Klik "Stijlen installeren" dus één keer per document (het maakt niet uit of
dat vóór of na het typen is - bestaande tekst met een ingebouwde stijl
verandert gewoon mee).

## Waarom deze knoppen en niet meer

De vorige, bredere versie (Klembord, Tabelstijlen, Taal, Bijhouden,
Bijlage, plus dit alles) leverde bij live testen in Word Online herhaald
moeilijk te reproduceren fouten op, deels doordat de add-in daar
Word-eigen ingebouwde stijlen probeerde te herprogrammeren. Om eerst een
kleine, betrouwbare basis te hebben die echt goed werkt, is de scope
teruggebracht tot alleen Titel, Standaardtekst en Kopteksten
(genummerd/ongenummerd) - de rest kan er later weer bij, nu de aanpak
(alleen custom stijlen, alleen `font`, nooit ingebouwde Word-stijlen)
zich bewezen heeft.

## Projectstructuur

```
manifest.xml                     - add-in manifest (ribbon-definitie)
src/
  shared/ggm-actions.js          - alle knop-logica (1x, gedeeld)
  commands/commands.html/.js     - ribbon-knoppen (draaien zonder taakvenster)
  taskpane/taskpane.html/.css/.js- taakvenster (backup-UI)
assets/
  icon-16/32/64/80.png           - ribbon-iconen
scripts/
  make_icons.py                  - genereert de PNG-iconen
```

Geen build-stap nodig (geen webpack/React) - gewoon statische bestanden,
makkelijk te hosten en te doorgronden.

## Hosten op GitHub Pages (voor delen/testen)

Repo: <https://github.com/mmulder82/gm-word-ribbon/>. `manifest.xml` wijst
al naar **`https://mmulder82.github.io/gm-word-ribbon`** - dat is de
standaard GitHub Pages-URL voor deze repo, ervan uitgaande dat Pages
uitgeeft vanaf de `main`-branch, map `/ (root)`.

1. Push deze map naar die repo (branch `main`, bestanden in de root - dus
   `manifest.xml`, `assets/`, `src/`, ... direct in de repo-root, niet in
   een submap).
2. Zet in de repo-instellingen **Settings > Pages** de bron op de
   `main`-branch, map `/ (root)`.
3. Wacht tot Pages live is (meestal <1 min), open
   <https://mmulder82.github.io/gm-word-ribbon/manifest.xml> in de browser
   om te controleren dat het manifest gewoon laadt (en check ook
   `.../assets/icon-32.png`).
4. **Belangrijk na een update:** GitHub Pages/CDN en Word Online's eigen
   add-in-cache kunnen een net gepushte wijziging een paar minuten laten
   hangen. Een InPrivate/incognitovenster omzeilt alleen je LOKALE
   browsercache, niet die twee. Geef het na een push een paar minuten, en
   overweeg de add-in in Word echt te verwijderen en opnieuw te
   sideloaden als een fix niet lijkt aan te komen.

Gebruik je toch een andere branch, submap of custom domain, dan moet je
elke `mmulder82.github.io/gm-word-ribbon` in `manifest.xml` handmatig
aanpassen naar de URL die je Pages-instellingen je echt geven.

## Sideloaden om te testen

**Word desktop (Windows/Mac):** Invoegen > Invoegtoepassingen > Mijn
invoegtoepassingen > Uploaden vanuit bestand... > kies `manifest.xml`.

**Word Online:** open Word in de browser > Invoegen > Invoegtoepassingen >
Mijn invoegtoepassingen > Invoegtoepassing uploaden > kies `manifest.xml`.
(Vereist dat de URL's in het manifest al naar de live GitHub Pages-URL
wijzen - Word Online kan geen `localhost`-bestanden laden.)

## Later: uitrollen via de Microsoft 365 app catalog

Zodra dit getest en goedgekeurd is, kan een beheerder `manifest.xml` (met de
definitieve hosting-URL) uploaden in **Microsoft 365 beheercentrum >
Instellingen > Geïntegreerde apps** (of de klassieke "Centraal beheerde
invoegtoepassingen"), zodat alle medewerkers het tabblad automatisch in Word
zien staan, in zowel desktop als Word Online - zonder zelf iets te hoeven
sideloaden.

## Stijlen aanpassen

Lettertype/formaat/vet/cursief/kleur per stijl staan bovenaan in
`src/shared/ggm-actions.js` (`TITEL_SPEC`, `STANDAARD_SPEC`,
`KOPTEKST_SPECS`, en de losse `BRAND_COLOR`/`BLACK`-constanten).

Stijlen die al in een document zijn geïnstalleerd passen zich niet met
terugwerkende kracht aan - dat geldt voor bestaande documenten, niet voor
nieuwe.
