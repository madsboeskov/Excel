# Spørgeskema med automatisk kvitteringsmail

En lille hjemmeside med et spørgeskema. Så snart nogen sender skemaet, bliver
der automatisk sendt en standardmail til den e-mailadresse, de har skrevet i
skemaet. Alle besvarelser gemmes samtidig, så du kan hente dem som en
CSV-fil og åbne dem i Excel.

* Spørgsmålene står i én fil — `config/survey.json` — og siden bygger sig selv ud fra den.
* Mailteksten står i `templates/` og kan skrives om uden at røre koden.
* Virker på mobil, i mørk tilstand og uden at brugeren skal oprette noget.
* Kører i "console"-tilstand indtil du udfylder SMTP-oplysningerne, så du kan
  prøve det hele uden en mailkonto.

## Kom i gang

```bash
cd survey
npm install
cp .env.example .env     # udfyld SMTP-oplysningerne, når du har dem
npm start
```

Åbn <http://localhost:3000>. Uden SMTP-oplysninger bliver mailen ikke sendt,
men gemt som en `.eml`-fil i `data/outbox/`, som du kan åbne i dit mailprogram
og se præcis, hvordan den kommer til at se ud.

## Sådan sender du rigtige mails

Udfyld `.env`:

```env
SMTP_HOST=smtp.ditdomæne.dk
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=ingen-svar@ditdomæne.dk
SMTP_PASS=dit-kodeord
MAIL_FROM=ingen-svar@ditdomæne.dk
MAIL_FROM_NAME=Dit firmanavn
```

Test opsætningen uden at udfylde skemaet:

```bash
npm run test:mail -- din@mail.dk
```

Værdier til de mest almindelige udbydere:

| Udbyder | SMTP_HOST | Port | Bemærkning |
| --- | --- | --- | --- |
| Gmail / Google Workspace | `smtp.gmail.com` | 587 | Kræver 2-faktor + en **appadgangskode** som `SMTP_PASS` |
| Microsoft 365 / Outlook | `smtp.office365.com` | 587 | `MAIL_FROM` skal være den samme postkasse som `SMTP_USER` |
| One.com | `send.one.com` | 465 (sæt `SMTP_SECURE=true`) | |
| Simply.com | `smtp.simply.com` | 587 | |
| Brevo, Mailgun, Postmark, Resend, SendGrid | se udbyderens SMTP-side | 587 | Anbefales ved mange mails |

Et par ting, der afgør, om mailen ender i indbakken frem for spam:

* Send fra dit **eget domæne** — ikke fra `@gmail.com` eller `@hotmail.com`.
* Få opsat **SPF** og **DKIM** på domænet (din mailudbyder har en vejledning).
* Brug en fast afsender som `ingen-svar@ditdomæne.dk` og en tydelig emnelinje.
* Skal du sende mange mails, så brug en rigtig mailtjeneste (Brevo, Mailgun,
  Postmark …) i stedet for en almindelig postkasse. De har alle SMTP, så
  opsætningen her er den samme.

## Lav skemaet om

Rediger `config/survey.json` og genstart serveren. Hvert spørgsmål har et `id`
(som bruges i CSV-filen og i mailen), en `type` og en `label`:

| Type | Bruges til |
| --- | --- |
| `text`, `textarea` | fritekst, kort eller lang |
| `email` | e-mailadresse (skemaet skal have præcis ét, og kvitteringen sendes hertil) |
| `tel`, `number`, `date` | telefon, tal, dato |
| `radio`, `select` | ét valg blandt `options` |
| `checkbox` | flere valg blandt `options`, evt. med `minSelected` / `maxSelected` |
| `rating` | skala fra 1 til `max`, med `minLabel` og `maxLabel` i enderne |
| `consent` | ét afkrydsningsfelt, fx samtykke |

`"required": true` gør feltet påkrævet. `emailField` øverst i filen peger på
det spørgsmål, kvitteringen skal sendes til, og `nameField` på navnet, der
bruges i tiltalen. Skriver du noget forkert i filen, siger serveren det med det
samme ved opstart i stedet for at starte med et halvt skema.

## Lav mailen om

* `templates/confirmation.html` — den pæne udgave
* `templates/confirmation.txt` — tekstudgaven (begge sendes med, så mailen også
  kan læses i programmer, der ikke viser HTML)
* Emnelinjen står under `"email"` i `config/survey.json`

I begge skabeloner kan du bruge:

| Felt | Indhold |
| --- | --- |
| `{{name}}` | navnet fra skemaet |
| `{{email}}` | modtagerens adresse |
| `{{date}}` | dato og klokkeslæt for besvarelsen |
| `{{surveyTitle}}` | skemaets titel |
| `{{senderName}}` | `MAIL_FROM_NAME` fra `.env` |
| `{{answersHtml}}` / `{{answersText}}` | hele svaroversigten |
| `{{answer.navn}}` | ét bestemt svar — brug spørgsmålets `id` |

## Se besvarelserne

Alle besvarelser ligger i `data/responses.jsonl` (én linje pr. besvarelse).
Sæt et `ADMIN_TOKEN` i `.env`, så kan du hente dem som regneark:

```
http://localhost:3000/admin/responses.csv?token=DIT_TOKEN
```

Filen bruger semikolon og BOM, så dansk Excel åbner den med kolonnerne på
plads. Uden `ADMIN_TOKEN` er `/admin`-adresserne slået helt fra.

## Sådan hænger det sammen

```
public/           siden i browseren (bygges ud fra config/survey.json)
server.js         ruterne: /api/survey, /api/submit, /admin/responses.csv
lib/config.js     indlæser og kontrollerer spørgeskemaet
lib/validate.js   tjekker en besvarelse på serveren (klienten tjekker også selv)
lib/mailer.js     bygger og sender kvitteringen
lib/storage.js    gemmer besvarelser og laver CSV
templates/        mailskabelonerne
data/             besvarelser og udgående mails (ligger uden for git)
```

Når et skema sendes ind, bliver svarene **gemt først** og mailen sendt bagefter.
Er mailserveren nede, er besvarelsen altså stadig gemt — brugeren får besked om,
at kvitteringen ikke kunne sendes, og CSV-filen viser "Nej" i kolonnen
*Kvittering sendt*.

Indbygget mod misbrug: skjult honningkrukkefelt, højst 5 indsendelser pr.
IP-adresse pr. 10 minutter (kan ændres i `.env`), grænse på indholdets størrelse,
og al indtastning valideres på serveren — også selvom den kom forbi browseren.
Linjeskift i e-mailfelter afvises, så ingen kan smugle ekstra mailmodtagere ind.

## Læg den på nettet

Alt, der skal til, er en server med Node 18 eller nyere: `npm ci && npm start`.
Sæt `PORT` og resten af `.env` som miljøvariabler hos udbyderen, og sæt
`TRUST_PROXY=true`, hvis siden kører bag en reverse proxy — ellers ser rate
limitten alle besøgende som den samme.

Husk, at `data/` skal ligge et sted, der overlever en genstart (en disk, der
følger med tjenesten). Ellers forsvinder besvarelserne, næste gang serveren
starter forfra. Har du mange besvarelser eller flere servere, er næste skridt at
gemme dem i en database i stedet — det er kun `lib/storage.js`, der skal skiftes.
