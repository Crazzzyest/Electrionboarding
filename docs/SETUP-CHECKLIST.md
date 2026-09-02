# Oppsett-sjekkliste (for deg, utvikler)

Alt herfra er ting som må gjøres/innhentes før hvert steg kan kjøre mot ekte tjenester i
stedet for `DEMO_MODE`. Appen kjører og er fullt ut demoerbar uten noe av dette — se
README-notatet i bunnen av denne filen for hvordan du kjører den nå.

For Microsoft Graph: se den separate [MICROSOFT-ADMIN-SETUP.md](MICROSOFT-ADMIN-SETUP.md),
som er skrevet til Electis egen IT/tenant-admin, ikke til deg.

## DocuSign

**Status 2026-08-28: autentisering fungerer end-to-end og er verifisert live** (ekte JWT-token
hentet, ekte kladde-konvolutt opprettet fra malen med riktige felt, deretter makulert). To
konkrete ting gjenstår, begge beskrevet under punkt 6.

1. ✅ Developer-konto finnes (`account-d.docusign.com`, konto "Reistad tjenester").
2. ✅ Malen "Arbeidsavtale" er bygget. `templateId` og rollenavn er bekreftet og lagt inn i
   [src/config.js](../src/config.js):
   - `templateId`: `29ce5c69-2d6c-4032-a83f-900179d302c5`
   - Rollenavn: `Ansatt` (matcher `config.docusign.signerRoleName`)

   **Viktig oppdagelse:** malen ble bygget med DocuSigns **automatiske feltgjenkjenning**, ikke
   ved å plassere felt manuelt med selvvalgte navn. Det betyr at feltene i malen IKKE heter det
   klammeteksten i Word-dokumentet sa (`[[STILLINGSDEL]]` osv.) — DocuSign genererte sine egne
   interne navn. Full oversikt over de 7 opprinnelige feltene, verifisert direkte mot malen:

   | Felt | Hvordan det fylles | Status |
   |---|---|---|
   | Navn | Automatisk, DocuSigns eget "Full Name"-felt (fra `name` i koden) | ✅ Virker, ingen kode nødvendig |
   | Stilling | DocuSigns eget "Title"-felt, men trenger likevel sitt genererte tab-navn (`config.docusign.tabLabels.stilling`) | ✅ Verifisert |
   | Telefon | Egendefinert tekstfelt, genrert tab-navn (`tabLabels.telefon`) | ✅ Verifisert |
   | Stillingsdel | Egendefinert tekstfelt, generert tab-navn (`tabLabels.stillingsprosent`) | ✅ Verifisert |
   | Nærmeste leder | Egendefinert tekstfelt, generert tab-navn (`tabLabels.naermesteLeder`) | ✅ Verifisert |
   | E-post | **Se punkt 6a — feil feltype i malen** | ⚠️ Må rettes i DocuSign |
   | Tiltredelse (dato) | **Ble ikke gjenkjent i det hele tatt** | ❌ Mangler i malen, se punkt 6b |

   **Skjør detalj:** de genererte tab-navnene (typen `atb.docusignFields.label-text
   7yu4n1po5qrmtcmdtis`) hører til **akkurat denne prosesserte versjonen** av malen. Lastes
   dokumentet opp på nytt i DocuSign (ny fil, regenerert), vil disse mest sannsynlig endre seg,
   og `config.docusign.tabLabels` må oppdateres på nytt (kjør en `listTabs`-sjekk mot malen for
   å finne de nye verdiene — spør meg, det tar to minutter).
3. ✅ App og Integration Key opprettet ("Arbeidsavtale", Integration Type = Private custom
   integration). `DOCUSIGN_INTEGRATION_KEY` lagt i `.env`.
4. ✅ RSA-nøkkelpar generert i DocuSign og **privat nøkkel korrekt lagret** som
   `docusign-private-key.pem` (verifisert med `openssl rsa -check`, og at den offentlige nøkkelen
   den utleder er identisk med den DocuSign viser).

   **Advarsel som kostet oss en runde:** DocuSigns "Generate RSA"-dialog viser BÅDE offentlig og
   privat nøkkel i samme vindu, offentlig nøkkel øverst. Det er lett å kopiere feil boks. En
   offentlig nøkkel limt inn som "privat nøkkel" gir ingen tydelig feilmelding ved bruk, bare en
   generisk autentiseringsfeil lenger ned i kjeden — så dobbeltsjekk alltid med
   `openssl rsa -in docusign-private-key.pem -check -noout` (skal si `RSA key ok`) etter at en ny
   nøkkel er lagret.
5. ✅ Redirect URI lagt til (`https://www.docusign.com` — brukes aldri til noe reelt, DocuSign
   krever bare at én er registrert for at samtykke-skjermen skal vises i det hele tatt).
6. ✅ Engangssamtykket er gitt og bekreftet fungerende (verifisert med et ekte token-kall).

   **Gjenstår, begge krever at du gjør noe inne i DocuSigns malredigering (jeg kan ikke gjøre
   dette via API):**

   **a) E-post-feltet i malen har feil type.** Det ble automatisk gjenkjent som DocuSigns
   innebygde "Email"-felt, som alltid speiler adressen kontrakten faktisk sendes til
   (`privatEpost` — må være slik, siden @electi.no-postkassen ikke finnes når kontrakten sendes).
   Men de ekte kontraktene viser **@electi.no-adressen** i det feltet, ikke privatadressen. Disse
   to kan ikke være det samme feltet. Løsning: åpne malen i DocuSign, klikk på E-post-feltet i
   dokumentet, og bytt felttype fra "Email" til vanlig "Text". Gi meg beskjed når det er gjort,
   så henter jeg det nye genererte tab-navnet og kobler det til `microsoftUpn` i koden (samme
   mønster som de andre tekstfeltene over).

   **b) Tiltredelsesdato-feltet finnes ikke i malen i det hele tatt.** Klammeteksten
   `[[TILTREDELSESDATO]]` ble ikke gjenkjent av DocuSigns automatikk, sannsynligvis fordi den
   ikke lignet noe DocuSign kjenner igjen som et datofelt. Løsning: åpne malen, finn stedet i
   dokumentet (rett etter "...med tiltredelse fra"), og legg til et vanlig tekstfelt (eller
   datofelt) manuelt der. Gi meg beskjed når det er gjort også, så kobler jeg den til
   `startdato`.
5. Noter `accountId` (finnes under Apps and Keys, eller i respons fra OAuth userinfo).
6. Fyll `.env`: `DOCUSIGN_INTEGRATION_KEY`, `DOCUSIGN_USER_ID` (bruker-GUID, ikke e-post),
   `DOCUSIGN_ACCOUNT_ID`, og enten `DOCUSIGN_PRIVATE_KEY_BASE64` eller la
   `DOCUSIGN_PRIVATE_KEY_PATH` peke på filen.
7. Admin → Connect → Add Configuration:
   - Events: **Envelope Completed**, **Envelope Declined**, **Envelope Voided**.
   - URL: `https://<din-deployede-url>/webhooks/docusign`.
   - Sett en HMAC-nøkkel og legg samme verdi i `.env` som `DOCUSIGN_CONNECT_HMAC_KEY`.
   - **Verifiser JSON-formen** på minst én ekte test-konvolutt mot
     `parseWebhookEvent()` i [src/docusign.js](../src/docusign.js) — den er skrevet
     defensivt mot standard per-event-format, men Connect-konfigurasjonen avgjør nøyaktig
     nøsting, og det er ikke bekreftet mot en ekte hendelse ennå.
8. Fyll `config.docusign.templateId` i [src/config.js](../src/config.js) (ikke en env-var —
   den er ikke hemmelig, bare en stabil ID).
9. Gjenta alt over mot **produksjons**-DocuSign-kontoen før go-live, og sett
   `DOCUSIGN_ENV=production` i `.env`.

## SalesScreen

1. Magnus har API-nøkkelen klar (bekreftet 2026-08-21) — få den fra ham over en sikker
   kanal (ikke e-post/Slack i klartekst) og legg den i `.env` som `SALESSCREEN_API_KEY`.
   Auth-header heter bokstavelig talt `apiKey`.
2. Team bekreftet: nye brukere skal inn i teamet **"Telenor"** i SalesScreen
   (`config.salesscreen.teamName`, allerede satt i [src/config.js](../src/config.js)).
   Fortsatt uavklart: om SalesScreens API vil ha en **team-ID** i stedet for navnet — sjekk
   dette samtidig som du finner endepunktet under, og fyll evt. inn `teamId`.
3. **Viktig, fortsatt uavklart:** det konkrete endepunktet for å opprette/slå opp brukere er
   *ikke* bekreftet. Deres offentlige dokumentasjon (docs.salesscreen.com) er en JS-app som
   ikke lot seg lese automatisk, og søk fant ikke et konkret skjema. Før du går videre:
   - Be om et Postman-eksempel eller konkret API-referanse fra SalesScreen support, eller
     logg inn med den ekte nøkkelen og finn API-dokumentasjonslenken i egen konto.
   - Fyll `config.salesscreen.baseUrl`, `createUserEndpoint`, `getUserEndpoint` i
     [src/config.js](../src/config.js) når du har den. Så lenge disse er tomme, nekter
     `src/salesscreen.js` bevisst å gjette en URL og returnerer en tydelig feil i stedet.
4. **Testmiljø, avklares med Magnus:** uklart ennå om SalesScreen har et eget
   demo-/sandkasse-miljø adskilt fra skarpe data, eller om testing må gjøres forsiktig i det
   vanlige miljøet med en tydelig merket testbruker som slettes etterpå. Spør SalesScreen
   support direkte hvis Magnus ikke vet.
5. E-post brukes som stabil bruker-ID (SalesScreens egen anbefaling) — se
   `salesscreenUserId` i datamodellen.

## Telenor

**Mottakere og format er nå hentet fra en ekte bestilling** Magnus sendte 21.08.2026, og lagt
inn i `config.telenor` + [src/telenor.js](../src/telenor.js):

- Til: `c-view.admin@telenor.no`, `kasper-maehlum.kvesetberget@telenor.no`
- Kopi: `audun.paulsen@electi.no`, `daniel.kolstad@electi.no`
- Emne: `Ny selger`
- Innhold: kun navn, mobil og den nyansattes **@electi.no-adresse**

Gjenstående avklaringer:

1. **"Salesforce" og "Ransel" nevnes ikke i den ekte bestillingen.** Det opprinnelige forslaget
   sa at bestillingen dekker begge, men mailen Magnus faktisk sender ber bare om "brukere" hos
   C-View. Spør om disse bestilles et annet sted, eller om C-View-brukeren dekker alt.
2. **Avsenderadresse.** Magnus sender fra `magnus.buran@electi.no`. Appen sender som
   `config.email.sendAsMailbox` (se Microsoft/Graph-seksjonen under). Sjekk at Telenor godtar
   bestillinger fra den avsenderen, ellers bør `sendAsMailbox` settes til en adresse de
   kjenner igjen.
3. Bekreft at mottakerlisten er stabil (er Kasper en fast mottaker, eller var han med kun på
   denne ene bestillingen?).

## Lagring (Excel i SharePoint/OneDrive) og e-post

**Avgjort 2026-08-29:** lagring og e-postutsending flyttet fra Google (Sheets/Gmail) til
Microsoft, siden Electi uansett er en Microsoft-shop og appen allerede har testet, fungerende
Graph-tilgang for brukeropprettelse. Det betyr **én mindre kontotype å holde styr på**, ikke
to (Google + Microsoft) — alt kjører nå på samme `MICROSOFT_TENANT_ID` /
`MICROSOFT_CLIENT_ID` / `MICROSOFT_CLIENT_SECRET` som allerede står i `.env`. `google.js` og
`auth-setup.js` er fjernet fra prosjektet, de trengs ikke lenger.

**Nytt, ikke bekreftet ennå — dette blokkerer ekte lagring og e-post:**

1. **Hvor skal Excel-filen ligge?** Trenger fra Electi: hvilket SharePoint-nettsted (eller om
   det heller skal være en bestemt persons OneDrive), og hvilken mappe/filnavn. Fyll deretter
   inn i [src/config.js](../src/config.js) under `excel`:
   - `siteId` (SharePoint-nettsted, `GET /sites/{hostname}:/{sti}` i Graph gir ID-en), **eller**
     `driveId` hvis det heller skal være en OneDrive.
   - `itemPath`, f.eks. `/Onboarding/Electi-Onboarding.xlsx`.
2. **Opprett selve Excel-filen** på den plasseringen, med to faner: `Ansatte` og `Logg`.
   I hver fane, sett inn en **Excel-tabell** (merk celleområdet → Sett inn → Tabell — dette må
   være en ekte tabell, ikke bare et celleområde, siden Graphs rad-API krever det) med
   kolonneoverskrifter som i `HEADERS`/`LOGG_HEADERS` i
   [src/columns.js](../src/columns.js) (rad 1). Gi tabellene navnene `Ansatte` og `Logg` (eller
   juster `config.excel.ansatteTable`/`loggTable` til det du faktisk kalte dem).
3. **Hvilken postboks skal automatiske e-poster sendes fra?** (Telenor-bestilling,
   velkomstmail, bursdagsvarsel.) App-only `Mail.Send` krever en ekte, lisensiert postboks —
   det finnes ikke noe "tjenestekonto uten innboks"-alternativ. Fyll inn i
   `config.email.sendAsMailbox` i [src/config.js](../src/config.js), f.eks.
   `onboarding@electi.no` eller `magnus.buran@electi.no`.
4. **Electis admin må gi to nye Graph-tillatelser** utover det som allerede er gitt — se
   oppdatert [MICROSOFT-ADMIN-SETUP.md](MICROSOFT-ADMIN-SETUP.md).
5. **Ikke testet mot en ekte fil ennå.** `src/excel-storage.js` er skrevet mot Graphs
   dokumenterte Excel-API (samme mønster som `src/microsoft.js`, som *er* verifisert live),
   men selve rad-lese/skrive-logikken bør kjøres én gang mot en ekte testfil før den brukes på
   ekte kandidatdata — spesielt `updateCells()`, som leser hele raden og skriver den tilbake
   (Excel sin rad-API har ingen "oppdater kun én celle"-operasjon slik Sheets har).

## Generelt

- Sett `TEST_MODE=true` og `DEMO_MODE=false` og kjør én full runde (registrer → signer →
  alle steg) mot sandkasse-/testkontoer for hver tjeneste før noe kobles til skarpe kontoer.
- Fyll `config.microsoft.groupIdsByAvdeling` og `config.microsoft.licenseSkuId` når
  Electis admin har levert dem (se MICROSOFT-ADMIN-SETUP.md).
- Fyll `config.email.managementEmail` (hvem som skal ha bursdagsvarsler).

## Kjøre appen nå (uten noe av det over)

```bash
npm install
DEMO_MODE=true npm start
```

Åpne `http://localhost:3000`. Hele kandidat-løpet er interaktivt demoerbart: registrer en
kandidat, se kontrakten "sendes", trykk "Simuler signering (demo)", se
Microsoft365/Telenor/SalesScreen fullføre uavhengig av hverandre og Velkommen fullføre
etter Microsoft365, trykk "Kjør på nytt" på et allerede grønt steg og se at det er en trygg
no-op, og kall `POST /api/check-birthdays` for å se et bursdagsvarsel (én demo-kandidat har
fødselsdato satt til i dag). Ingenting av dette lagres noe sted — data nullstilles hver gang
serveren starter på nytt.
