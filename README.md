# PPTimagePoint

Panelet er på engelsk og bruker de vanlige engelske fagordene (hue, saturation, clarity osv.). Denne beskrivelsen er på norsk.

Et PowerPoint-tillegg som gjør bildebehandling enkelt, rett i PowerPoint:

* **Fjern bakgrunn med fargenøkkel.** Marker bildet, klikk på bakgrunnsfargen, ferdig. Shift+klikk legger til flere farger (skygger, fargeoverganger). «Auto detect background» leser fargen langs kantene.
* **Color and light:** exposure, brightness, contrast, highlights, shadows, temperature, tint, hue, saturation, vibrance, clarity og sharpness.
* **Color mixer:** hue, saturation og luminance for åtte fargeområder hver for seg (for eksempel rødt til oransje).
* **Advanced:** Levels med histogram, og Curves for RGB og hver fargekanal.
* Resultatet settes inn **på nøyaktig samme sted og i samme størrelse**, og originalen erstattes (kan slås av).

Alt skjer lokalt i PowerPoint. Ingen bilder sendes til noen server.

## Slik bruker du det

1. Klikk **Hjem > Picture Tools** for å åpne panelet.
2. Marker et bilde på lysbildet. Det hentes automatisk inn i panelet, og når du klikker på et annet bilde, følger panelet etter.
3. Klikk på bakgrunnen i forhåndsvisningen. Bakgrunnen forsvinner med en gang.
4. Finjuster:
   * **Tolerance**: hvor like bakgrunnsfargen pikslene må være for å fjernes.
   * **Softness**: myk kant i stedet for hard kant.
   * **Keying mode**: «Contiguous background» fjerner bare bakgrunn som henger sammen med det du klikket på eller med bildekantene, slik at like farger inne i motivet beholdes. «Color everywhere in the image» fjerner fargen i hele bildet (nyttig for hull, f.eks. mellom armen og kroppen).
   * **Choke / expand edge** og **Feather** rydder i kanten.
   * **Despill** fjerner f.eks. grønt skjær fra greenscreen.
   * **Erase / Restore** er pensler for å rette opp det fargenøkkelen ikke klarer alene. **Brush size** og **Hardness** vises når en pensel er valgt. Lav hardhet gir myk pensel.
   * Knappen med pil ut av en boks åpner forhåndsvisningen i et **stort eget vindu**, der det er lettere å tegne. Alt du gjør der, skjer i panelet og på lysbildet.
   * Hold inne øyeknappen for å sammenligne med originalen.
   * **Reset** øverst til høyre legger originalbildet tilbake på lysbildet, på opprinnelig plass, og nullstiller alt.
   * Alle knapper har forklaring når du holder musepekeren over dem.
   * **Ctrl+Z** angrer og **Ctrl+Y** gjør om (eller knappene **Undo** og **Redo**). Snarveiene virker når panelet har fokus. Står markøren i lysbildet, går Ctrl+Z til PowerPoint i stedet.
5. Med **Live update on slide** (på som standard) oppdateres bildet på lysbildet automatisk et øyeblikk etter at du slipper en glidebryter eller penselen. Slår du det av, bruker du **Insert into slide** når du er ferdig.

### Beskjære, flytte og skalere

Bruk PowerPoints egne verktøy (Bildeformat > Beskjær, eller dra i hjørnene). Panelet leser av endringene neste gang det oppdaterer bildet, så beskjæring, plassering og størrelse beholdes. Det som er beskåret bort vises mørkt i forhåndsvisningen.

### Skjule forhåndsvisningen

Med øyeknappen til høyre under forhåndsvisningen jobber du bare mot lysbildet (med Live påslått). Pipetten og penselen trenger forhåndsvisningen, fordi PowerPoint ikke lar tillegg ta imot klikk på selve lysbildet. Valget huskes.

### Designforslag (Designer) dukker opp

PowerPoint Designer foreslår oppsett hver gang et bilde settes inn, og live-oppdatering setter inn bildet på nytt. Slå det av under **Fil > Alternativer > Generelt > PowerPoint Designer**, og fjern haken for å vise designideer automatisk. På Mac: **PowerPoint > Innstillinger > Generelt**.

### Om live-oppdatering

PowerPoint lar ikke tillegg tegne direkte i et bilde, så live-oppdatering betyr at bildet på lysbildet byttes ut med en ny versjon i full oppløsning. Det har tre følger:

* Panelet jobber alltid fra originalpikslene, så justeringene kan endres og nullstilles uten tap.
* Bildet får et usynlig merke i presentasjonen, og originalen, innstillingene og angrehistorikken lagres lokalt på maskinen. Klikker du på bildet igjen senere, også etter at panelet har vært lukket (for eksempel av Utforming), fortsetter du der du slapp. De 30 sist redigerte bildene huskes.
* Hver oppdatering blir et eget angresteg i PowerPoint. Angrer du med Ctrl+Z i lysbildet (eller gjør om med Ctrl+Y), retter panelet seg etter versjonen som ligger på lysbildet.
* Rotasjon, rammer og effekter på bildet følger ikke med når det byttes ut.

Panelet henter originalfilen som ligger inne i presentasjonen, ikke en skjermkopi, så du jobber alltid i full oppløsning (opptil 6000 px på lengste side). Beskjæring du har gjort i PowerPoint blir tatt med.

## Installere tillegget

Tillegget er en nettside (mappen `src/`) og et manifest (`manifest.xml`) som forteller PowerPoint hvor nettsiden ligger.

### 1. Publiser nettsiden (en gang)

Manifestet peker på `https://sirbull.github.io/PPTimagePoint/`.

1. Slå sammen koden til `main`.
2. Gå til **Settings > Pages** i GitHub-repoet og velg **Source: GitHub Actions**.
3. Arbeidsflyten `.github/workflows/pages.yml` publiserer `src/` automatisk ved hver push til `main`.

### 2. Legg til tillegget i PowerPoint

**PowerPoint på nett**
Hjem > Tillegg > Flere tillegg > Mine tillegg > **Last opp mitt tillegg**, og velg `manifest.xml`.

**Mac**
Kopier `manifest.xml` til
`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
(lag mappen `wef` hvis den ikke finnes). Start PowerPoint på nytt, og finn tillegget under Hjem > Tillegg > Mine tillegg.

**Windows**
1. Legg `manifest.xml` i en delt mappe (høyreklikk på mappen > Egenskaper > Deling).
2. I PowerPoint: Fil > Alternativer > Klareringssenter > Innstillinger for klareringssenter > **Klarerte tilleggskataloger**. Lim inn nettverksbanen (f.eks. `\\MASKIN\manifester`), klikk Legg til katalog og kryss av for **Vis i meny**.
3. Start PowerPoint på nytt. Hjem > Tillegg > Mine tillegg > **Delt mappe**.

**Hele organisasjonen**
En Microsoft 365-administrator kan rulle det ut via Admin center > Innstillinger > Integrerte apper > Last opp egendefinerte apper.

## Utvikling

```bash
npm install
npm run certs      # lokalt HTTPS-sertifikat for localhost (en gang)
npm run serve      # serverer src/ på https://localhost:3000
```

Last så inn `manifest.local.xml` i stedet for `manifest.xml`. Endres adressen, rediger `scripts/build-manifests.py` og kjør `npm run manifests`.

Åpner du `taskpane.html` i en vanlig nettleser, kjører panelet i frittstående modus der du kan åpne en bildefil og laste ned resultatet. Det er praktisk for å prøve ut justeringer.

`npm test` lager testbilder og en testpresentasjon og kjører hele panelet i Chromium mot en etterligning av PowerPoint-API-et.

### Oppbygning

| Fil | Innhold |
| --- | --- |
| `src/js/imageops.js` | Bildemotoren: fargenøkkel (Lab-avstand, flomfyll, kantbehandling), fargejusteringer, Lanczos-skalering |
| `src/js/pptx-extract.js` | Henter originalbildet ut av lysbildet (eksportert som .pptx) |
| `src/js/office-bridge.js` | All kontakt med PowerPoint |
| `src/js/pipeline.js` | Felles behandlingskjede for forhåndsvisning og full oppløsning |
| `src/js/preview.js`, `src/preview.html` | Det store forhåndsvisningsvinduet |
| `src/js/ui.js` | Ikoner og verktøytips |
| `src/js/app.js` | Panelet og arbeidsflyten |
| `scripts/make-icons.py` | Tegner ikonene |
| `manifest.template.xml` | Mal for manifestene |

### Krav og begrensninger

* Henting av det markerte bildet krever PowerPoint API 1.8 (Microsoft 365, oppdatert versjon på Windows, Mac og nett). På eldre versjoner kan du bruke «Open file» eller lime inn bildet med Ctrl+V i panelet.
* Rotasjon, rammer og effekter på originalbildet følger ikke med når bildet erstattes. Legg dem på igjen etterpå.
* Fargenøkkel fungerer best når bakgrunnen er noenlunde ensfarget (studio, produktbilder, logoer, greenscreen). For travle fotobakgrunner trengs penselen, eller en AI-basert modell (se under).

### Mulige utvidelser

* AI-basert bakgrunnsfjerning for vanlige fotografier (f.eks. en ONNX-modell som kjører lokalt i panelet).
* Lagrede forhåndsinnstillinger.
* Kjøre behandlingen i en Web Worker så panelet aldri fryser på store bilder.
