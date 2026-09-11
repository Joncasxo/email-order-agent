# Užsakymų agentas

Darbalaukio programa, kuri automatiškai peržiūri naujus pašto laiškus ir per Claude AI
atpažįsta klientų užsakymus — net jei jie persiųsti, paslėpti Excel/PDF prieduose ar nuotraukose.

## Paleidimas

Dukart spustelėkite **„Užsakymų agentas“** nuorodą darbalaukyje
(arba `Paleisti.bat` šiame aplanke).

## Pirmas paleidimas — ką reikia prijungti

Programai reikia dviejų prisijungimų (abu suvedami per **⚙ Nustatymai**, užtenka padaryti vieną kartą):

### 1. AI (Anthropic API raktas)

1. Atsidarykite **console.anthropic.com** ir susikurkite paskyrą (arba prisijunkite).
2. Skiltyje **Plans & Billing** papildykite kreditų (pvz., 10 $ — realus suvartojimas
   yra keli eurai per mėnesį; patogu įsijungti automatinį papildymą „auto-reload“).
3. Skiltyje **API Keys** spauskite **Create Key**, nukopijuokite raktą.
4. Programoje: **⚙ Nustatymai → Anthropic API raktas** — įklijuokite ir išsaugokite.
   Raktas saugomas užšifruotas (Windows DPAPI) tik šiame kompiuteryje.

### 2. Paštas (Microsoft Graph)

Programa laiškus skaito per Microsoft Graph — oficialią Microsoft 365 pašto sąsają
(veikia ir su „naujuoju Outlook“). Tam reikia vienkartinės programėlės registracijos
jūsų įmonės Microsoft aplinkoje:

1. Atsidarykite **portal.azure.com** ir prisijunkite savo **darbo** paskyra.
2. Paieškoje viršuje įveskite **App registrations** → spauskite **New registration**.
3. Užpildykite:
   - **Name:** pvz., `Uzsakymu agentas`;
   - **Supported account types:** palikite pirmą variantą („Accounts in this organizational directory only“);
   - **Redirect URI:** išskleidimo meniu pasirinkite **Public client/native (mobile & desktop)**
     ir įrašykite `http://localhost`.
   Spauskite **Register**.
4. Atsidariusiame puslapyje nusikopijuokite du kodus:
   - **Application (client) ID** → programos nustatymuose tai „Programėlės ID“;
   - **Directory (tenant) ID** → nustatymuose „Katalogo ID“.
5. Kairėje pasirinkite **API permissions → Add a permission → Microsoft Graph →
   Delegated permissions**, pažymėkite **Mail.Read** ir spauskite **Add permissions**.
   Jei jūsų įmonė reikalauja administratoriaus patvirtinimo — paspauskite
   „Grant admin consent“ (arba paprašykite IT administratoriaus patvirtinti prašymą).
6. Programoje: **⚙ Nustatymai → Pašto prisijungimas** — įrašykite abu ID ir savo pašto
   adresą, išsaugokite.
7. Paspaudus „Tikrinti dabar“ pirmą kartą atsidarys įprastas Microsoft prisijungimo
   langas — prisijunkite savo paštu. Toliau prisijungimas vyks automatiškai
   (žetonas saugomas užšifruotas faile `graph_token.bin`).

## Kaip veikia siuntėjų filtras

Įprastas atvejis — **užsakymai ateina tiesiai iš klientų**. Tada filtro laukas
paliekamas **tuščias** (tikrinami visi gauti laiškai) arba, jei norite taupyti,
įrašomi klientų domenai, pvz.: `@klientas.lt; @kitasklientas.lt` (atskirta
kabliataškiais). Filtras žiūri tik į tai, **kas laišką atsiuntė jums** — tinka
vardas, pašto dalis ar domenas.

Papildoma galimybė sudėtingesniam laiškų srautui (kai užsakymai jus pasiekia
per tarpininką):

- Jei užsakymus jums **persiunčia kolega** — į filtrą įrašykite ir jį. Laiškas
  praeis filtrą kaip kolegos, o **tikrąjį klientą** AI atpažins iš persiųsto
  laiško vidaus (From/Nuo antraščių ir parašo).
- Skaitoma ir visa susirašinėjimo grandinė laiško apačioje: jei tas pats užsakymas
  tęsiasi (Re:/FW:), programa atpažįsta, kad tai esamo užsakymo papildymas (🔔),
  o ne naujas užsakymas — dublikatų nebus.
- Abu keliai (tiesioginiai klientų laiškai ir kolegos persiuntimai) veikia kartu.

## Kaip veikia

- Kas 1 min. sekami tiekėjų atsakymai į užklausas (❗), kas N min. (nustatymuose) —
  pilnas naujų laiškų tikrinimas su AI.
- Kiekvieną naują laišką AI įvertina: ar tai užsakymas? Analizuojamas tekstas,
  **Excel priedai** (.xlsx/.xls/.ods), **PDF** ir **nuotraukos**. Spalvomis pažymėtos
  laiško vietos (pvz., raudonai pažymėta lentelės eilutė) AI perduodamos kaip žymės.
- Neaiškūs ar pavojingi prašymai (pvz., „tik pakoreguoti, ne užsakyti iš naujo“,
  neaiškūs kiekiai) pažymimi ⚠ ir raudonu įspėjimu detalėse.
- Užsakymai dedami į sąrašą su statusais ir filtrais (⏳ Neužbaigti / ✔ Užbaigti).
- **„Eksportuoti į Excel“** — sąrašas išsaugomas .xlsx faile.
- **„Rašyti tiekėjui“ / „Užklausos“** — užklausų tiekėjams kūrimas ir atsakymų sekimas
  pagal užsakymo kodą temoje.
- Kiekvienas laiškas analizuojamas tik vieną kartą (žymima `uzsakymai.db`).
- Klientų kontaktai automatiškai kaupiami darbalaukio faile `klientai.vcf` —
  jį galima importuoti į telefoną (Google Contacts arba tiesiogiai).

## Failai

| Failas | Paskirtis |
|---|---|
| `app.py` | Pagrindinė programa (grafinė sąsaja) |
| `graph_client.py` | Laiškų skaitymas per Microsoft Graph API |
| `outlook_client.py` | Atsarginis kelias per klasikinį Outlook (COM) |
| `extractor.py` | AI analizė per Claude API |
| `attachments.py` | Excel/PDF/ODS/nuotraukų paruošimas |
| `storage.py` | Vietinė SQLite duomenų bazė |
| `exporter.py` | Eksportas į Excel |
| `eksportuoti_kontaktus.py` | Klientų kontaktų (vCard) eksportas telefonui |
| `perziura.py` / `share.py` | Peržiūros režimas vadovui (tik skaitymas) |
| `config.json` | Nustatymai (raktai užšifruoti) |
| `uzsakymai.db` | Užsakymų ir apdorotų laiškų duomenų bazė |
| `zurnalas.txt` | Veiklos ir klaidų žurnalas |

## Pastabos

- Pašto šaltinis nustatomas `config.json` lauku `mail_source`: `graph` (rekomenduojama,
  veikia su nauju Outlook) arba `com` (klasikinis Outlook).
- Modelis: numatytas `claude-opus-5` (tiksliausias). Nustatymuose galima pasirinkti pigesnį
  `claude-sonnet-5` arba `claude-haiku-4-5`, jei kaštai svarbiau už tikslumą.
- Mažos nuotraukos (<15 KB, dažniausiai parašų logotipai) praleidžiamos.
- AI gali klysti — užsakymus prieš vykdymą peržiūrėkite (tam ir yra „naujas“ statusas,
  „Pasitikėjimas“ procentas bei ⚠ įspėjimai detalėse).
