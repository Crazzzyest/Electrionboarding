# Til Electis IT-administrator: tilgang for onboarding-appen

Dette dokumentet er skrevet til den som administrerer Electis Microsoft 365-tenant (Entra ID
/ Azure AD-administrator) — ikke til utvikleren. Onboarding-appen bruker Microsoft til mer enn
å opprette brukerkontoer for nyansatte: den lagrer også kandidatdata i en Excel-fil i
SharePoint/OneDrive og sender e-post (Telenor-bestilling, velkomstmail, bursdagsvarsel) via en
ekte postboks, alt gjennom det samme settet med tilganger. Det trengs noen tilganger som
**kun en administrator kan gi** — utvikleren har ikke og kan ikke skaffe seg disse selv.

Appen kjører uten en innlogget person til stede (den kjører automatisk når en kontrakt
signeres), så den bruker det som kalles "app-only"/"client credentials"-autentisering — én
egen, navngitt applikasjon i Entra ID med sine egne rettigheter, ikke en vanlig
brukerkonto.

## Steg 1 — Registrer applikasjonen

1. Gå til [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** →
   **App registrations** → **New registration**.
2. Navn: f.eks. «Electi Onboarding».
3. Supported account types: **Single tenant** (kun kontoer i Electis egen organisasjon).
4. Redirect URI: la stå tom/ubrukt — appen logger ingen mennesker inn, den trenger ingen
   redirect.
5. Etter opprettelse, noter fra oversiktssiden:
   - **Application (client) ID**
   - **Directory (tenant) ID**

## Steg 2 — Lag en klienthemmelighet (client secret)

1. I app-registreringen: **Certificates & secrets** → **New client secret**.
2. Kopier **verdien** med det samme — den vises kun én gang.
3. Gi denne til utvikleren over en sikker kanal (f.eks. et delt passord-verktøy), ikke i
   ren tekst på e-post/Slack/Teams.
4. Noter utløpsdatoen et sted dere husker den — hemmeligheten må fornyes før den utløper,
   ellers stopper appen å virke uten forvarsel.

## Steg 3 — Gi applikasjonstillatelser (ikke delegerte) og godkjenn dem

1. I app-registreringen: **API permissions** → **Add a permission** → **Microsoft Graph**
   → **Application permissions** (ikke "Delegated permissions" — appen kjører uten en
   innlogget bruker).
2. Legg til:
   - `User.ReadWrite.All` — oppretter brukerkontoen og tildeler lisens.
   - `Group.ReadWrite.All` — legger brukeren til riktig(e) gruppe(r). (Kan snevres til
     `GroupMember.ReadWrite.All` hvis dere ønsker en litt mer begrenset tillatelse — samme
     funksjon for det appen faktisk gjør.)
   - `Organization.Read.All` — trengs for å slå opp riktig lisens-SKU.
   - `Files.ReadWrite.All` — **nytt 2026-08-29.** Appen lagrer kandidatdata og status i en
     Excel-fil i SharePoint/OneDrive i stedet for en ekstern tjeneste, og trenger dette for å
     lese/skrive den filen. (Alternativ med snevrere omfang: `Sites.ReadWrite.All`, hvis dere
     heller vil begrense til SharePoint-nettsteder spesifikt — funksjonelt likt for det appen
     gjør.)
   - `Mail.Send` — **nytt 2026-08-29.** Sender Telenor-bestillingen, velkomstmailen til
     nyansatte og bursdagsvarselet til ledelsen. Sendes alltid som én bestemt, ekte postboks
     (se `config.email.sendAsMailbox` — utvikleren avklarer hvilken med dere).
   - `Mail.Read` — **nytt 2026-08-29.** Brukes kun til å sjekke om en Telenor-bestilling
     allerede er sendt (unngår duplikater ved automatiske gjenforsøk), ikke til å lese
     generell e-post.
3. Trykk **Grant admin consent for Electi**. Dette steget er **obligatorisk og kun mulig
   for en administrator** — appen fungerer ikke før dette er gjort, uansett hvilke
   tillatelser som er lagt til i listen. **Er de fire første tillatelsene fra tidligere
   allerede godkjent, må du likevel trykke denne knappen på nytt** etter å ha lagt til de tre
   nye — godkjenning dekker kun de tillatelsene som var lagt til i det den ble gitt.

## Steg 4 — Bekreft lisens

Hvilken Microsoft 365/Exchange Online-lisens skal nyansatte ha (så de får e-postkonto)?
Bekreft at det finnes ledige lisenser, og gi utvikleren lisensens **SKU-ID** (finnes under
**Billing → Licenses**, eller via `GET /subscribedSkus` i Graph).

> **Status 25.08.2026 — må avklares.** SKU-ID-en som først ble oppgitt
> (`1f2f344a-700d-42c9-9427-5cea1d5d7ba6`) er **Microsoft Stream**, som *ikke* gir postkasse.
> Den eneste Exchange-tjenesten i den lisensen er `EXCHANGE_S_FOUNDATION`, som bare er en
> teknisk grunnkomponent uten e-postkasse. Brukes den, får den nyansatte konto men ingen
> e-post.
>
> Verifisert mot tenanten finnes disse lisensene med ekte postkasse:
>
> | Lisens | SKU-ID | Ledige (25.08.2026) |
> |---|---|---|
> | `EXCHANGEDESKLESS` (Exchange Online Kiosk) | `80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82` | **4** av 45 |
> | `O365_BUSINESS_PREMIUM` | `f245ecc8-75af-4f8e-b61f-27d8114de5f3` | **1** av 7 |
> | `Microsoft_Teams_Exploratory_Dept` | `e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8` | 0 av 4 |
>
> Appen er foreløpig satt opp med **EXCHANGEDESKLESS**, fordi 41 av 45 plasser er i bruk,
> noe som passer med antall selgere. **Bekreft at dette er riktig lisens for nyansatte
> selgere.** Merk også at det bare var **4 ledige plasser** — flere ansettelser enn det
> krever innkjøp av nye lisenser før onboardingen kan fullføres.

## Steg 5 — Grupper

Gi utvikleren **Object ID**-en til gruppen(e) nyansatte skal inn i (f.eks. en
«Alle ansatte»-gruppe, og eventuelt egne grupper per avdeling/team).

> **Status 25.08.2026 — må avklares.** Object ID-en som først ble oppgitt
> (`457752a7-b65a-48b0-bcff-1506c75a70fd`) finnes ikke i tenanten; oppslag mot Graph gir
> «Resource does not exist», både som gruppe og som bruker. Den er antagelig kopiert fra feil
> sted, eller gjelder et slettet objekt.
>
> Disse gruppene finnes faktisk:
>
> | Gruppe | Object ID |
> |---|---|
> | Salg | `b5adad24-9fdc-47e2-95a4-a85084c0c08d` |
> | All Company | `e0fe9397-35a2-457a-8e9c-2b2273acf22c` |
> | Electi Business Partner AS | `bab2c260-6036-4564-b48a-2a0a32350c17` |
> | Diverse | `e5af881c-c0a2-41c7-90dd-c389560fc10b` |
>
> Appen er foreløpig satt opp med **Salg**. Bekreft om nyansatte selgere også skal inn i
> «All Company» eller andre grupper.

## Steg 6 — Send tilbake til utvikleren

- Tenant ID
- Client ID
- Client secret (over sikker kanal)
- Lisens-SKU-ID
- Gruppe-Object-ID-er
- Bekreftelse på at admin consent er gitt (steg 3), inkludert de tre nye tillatelsene
- Hvilket SharePoint-nettsted (eller OneDrive) Excel-filen med kandidatdata skal ligge i
- Hvilken postboks automatisk e-post skal sendes fra (må være en ekte, lisensiert postboks)

## Notater

- Siden dette er single-tenant, app-only og admin-godkjent, vil aldri sluttbrukere se noen
  "ubekreftet utgiver"-advarsel — det er ikke en vanlig innloggingsapp.
- Det kan ta **opptil ca. 10 minutter** før en ny postboks faktisk er tilgjengelig etter at
  en lisens er tildelt — dette er normalt, ikke en feil i appen.
- Hvis onboarding plutselig begynner å feile for alle nye kandidater med
  tillatelsesfeil, er første sjekk om admin consent (steg 3) på noe vis er trukket
  tilbake — det kan skje ved enkelte endringer i tenantens sikkerhetspolicyer, uten at
  noen aktivt "slettet" tilgangen.
- Client secret utløper (dere satte selv utløpsdatoen i steg 2) — sett gjerne en påminnelse
  et par uker før, siden appen ellers stopper å virke uten varsel den dagen den utløper.
