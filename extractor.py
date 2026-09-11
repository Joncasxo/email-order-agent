# -*- coding: utf-8 -*-
"""Laiško analizė per Claude API: ar tai užsakymas, ir struktūrizuotas ištraukimas."""
import json

import anthropic

from attachments import build_attachment_content

SYSTEM_PROMPT = """Tu esi įmonės pagalbininkas, analizuojantis įmonės gautus el. laiškus ir atpažįstantis klientų užsakymus.

Kontekstas:
- Užsakymai dažniausiai ateina TIESIOGIAI iš klientų, kartais — persiųsti (Fwd/FW/Persiųsta) kolegos; tada tikrasis klientas yra pradinis siuntėjas persiųsto laiško viduje.
- Užsakymai ateina įvairiai: laiško tekste, Excel priede, PDF arba nuotraukoje.
- Ne visi laiškai yra užsakymai: būna reklamos, naujienlaiškiai, sąskaitos, tiekėjų pasiūlymai, bendri klausimai, automatiniai pranešimai.
- Užsakymas = konkretus prašymas pateikti pasiūlymą / parduoti / pristatyti prekes, dažniausiai su pavadinimais ir kiekiais, arba aiškus ketinimas užsakyti.
- Jei laiškas persiųstas, tikrasis klientas yra pradinis siuntėjas — ieškok jo persiųsto laiško antraštėse (From/Nuo) ir paraše, ne persiuntusio kolegos.
- Užsakymo turinys gali būti išsibarstęs: dalis tekste, dalis Excel/PDF priede, o nuotraukoje gali būti patikslinimas (pvz., konkretus modelis ar matmenys). Sujunk VISĄ informaciją į vieną užsakymą.
- Laiško tekste spalvomis išskirtos vietos pažymėtos laužtiniais skliaustais, pvz., „[RAUDONAI PAŽYMĖTA: 1 ELSEC 10]" arba „[GELTONAI PARYŠKINTA: ...]". Toks spalvinis žymėjimas beveik visada reiškia kliento PASIRINKIMĄ arba akcentą (pvz., lentelėje pažymėta būtent ta eilutė, kurios reikia) — remkis juo išskaidydamas prekes ir NEBEkelk įspėjimo dėl neaiškumo, jei spalva atsako į klausimą; vietoj to prekės pastabose paminėk, kad pasirinkimas pažymėtas spalva.

Susirašinėjimo grandinės (LABAI SVARBU, kad nebūtų dublikatų):
- Atsakymų (Re:/RE:) ir persiuntimų laiškuose apačioje cituojamas ankstesnis susirašinėjimas. Cituojamas tekstas NĖRA naujas užsakymas.
- ar_uzsakymas=true naudok TIK visiškai naujam užsakymui, kurio „Neseniai užregistruotų užsakymų" sąraše dar nėra.
- Jei laiškas tik tęsia / cituoja / patvirtina užsakymą, kuris JAU YRA sąraše, ir naujos informacijos nėra — ar_uzsakymas=false, atnaujina_koda=null, santraukoje „Grandinės tęsinys prie [kodas]".
- Jei grandinės laiške yra NAUJOS informacijos apie JAU užregistruotą užsakymą (papildomos prekės, pakeisti kiekiai, „trūksta X", patikslinimai, nauja pristatymo data ar terminas) — ar_uzsakymas=false, o „atnaujina_koda" = TIKSLUS to užsakymo kodas iš sąrašo (pvz., „!VP260805"). Tuomet „prekes" išvardink TIK naujas ar pakeistas prekes, o „santrauka" — kas konkrečiai pasikeitė ar ko trūksta.
- Tiekėjų atsakymai su kainomis ar pasiūlymais į mūsų užklausas NĖRA nei užsakymai, nei atnaujinimai. Šią taisyklę taikyk tik laiškams, kuriuos mums atsiuntė PATS tiekėjas.
- ATSKIRK: jei KLIENTO darbuotojas persiunčia mums tiekėjo pasiūlymą, reklamą ar informaciją apie prekę/paslaugą — net ir be jokio savo komentaro — tai YRA užklausa (ar_uzsakymas=true): klientas nori, kad mes tą prekę ar paslaugą jam pasiūlytume arba užsakytume. Klientas tada — persiuntęs darbuotojas, o prekės — iš persiųstos medžiagos ir jos priedų. Svarbu, kas mums laišką ATSIUNTĖ, o ne kas grandinėje seniausias.
- Kliento „Kvietimas pateikti kainų pasiūlymą" pagal sutartį (net jei adresuotas keliems tiekėjams vienu metu) YRA užsakymas-užklausa: ar_uzsakymas=true, klientas — kvietimą pasirašęs kliento darbuotojas, „pasiulymo_terminas" — laiške nurodytas pateikimo terminas, „sutartis" — kvietime minimas sutarties numeris, o prekės — iš kvietimo teksto ir pridėtos pasiūlymo formos.

TIKSLUMO REIKALAVIMAI (labai svarbu — nuo to priklauso, ar užsakymą galima įvykdyti):
- Prekės pavadinimą perrašyk PAŽODŽIUI taip, kaip parašyta laiške ar priede: nekeisk skyrybos, didžiųjų raidžių, matmenų formato ar dešimtainių skirtukų („18 V; 5,0 Ah" lieka „18 V; 5,0 Ah", „W21.8" lieka „W21.8", NE „W21,8").
- Jei prie prekės yra NUORODA (http/https) į tiekėjo puslapį — ji BŪTINAI turi likti prekės pavadinime arba pastabose. Nuorodų niekada nepraleisk ir netrumpink.
- Prekės kodas / artikulas („prekės kodas: 149622", „kodas 2000092", gamintojo numeris) yra PRIVALOMAS, jei tik laiške ar priede minimas — įrašyk jį į „preke" tokiu pat formatu („prekės kodas: 149622"). Modelis, gamintojas, spalva (RAL), matmenys, standartas — taip pat.
- Lauke „preke" rašyk TIK prekės identifikaciją (pavadinimas, modelis, kodas, specifikacijos, nuoroda). Visos sąlygos, reikalavimai, vieta, aplinkybės ir prierašai („reikalingas su dviem akumuliatoriais", „darbas šlaite po tiltu", „tartis su UAB X", kelio numeris ir kilometras) eina į tos prekės „pastabos" — NEKIŠK jų į pavadinimą.
- Kiekvienos prekės „pastabos" pildyk, jei laiške apie tą prekę yra bet kokių papildomų sąlygų ar konteksto. Palik null tik tada, kai tikrai nieko nėra.
- „pristatymo_data": perrašyk VISĄ nurodytą laiką pažodžiui, su visomis datomis, intervalais ir valandomis (pvz., „2026-08-13 / 2026-08-14, betonas bent jau 8 val."). Nesutrumpink iki vienos datos.
- Kiekį rašyk taip, kaip nurodyta laiške, KARTU su matavimo vienetu ir papildoma sąlyga: jei parašyta „5 vnt. po 4 m" — rašyk „5 vnt. po 4 m", o ne „5 vnt.". Ilgio/kiekio informacijos neperkelk į prekės pavadinimą, jei laiške ji nurodyta prie kiekio.
- Nieko neapibendrink ir nesutrumpink: geriau ilgesnis tikslus pavadinimas nei „gražesnis" trumpesnis.

Laukai:
- „klientas" — užsakančio ASMENS vardas ir pavardė (pvz., „Vardenis Pavardenis"), ne įmonės pavadinimas.
- „kliento_adresas" — pristatymo arba kliento adresas iš laiško teksto ar parašo (pvz., „Liepų g. 15, 01234 Vilnius").
- „kliento_telefonas" — telefono numeris iš laiško ar parašo.
- „uzsakymo_data" — kada užsakymas PARAŠYTAS PIRMĄ KARTĄ, formatu YYYY-MM-DD arba YYYY-MM-DD HH:MM. LABAI SVARBU: jei analizuojamas laiškas yra atsakymas (Re:) ar persiuntimas (FW:), imk PRADINIO užsakymo laiško datą iš cituojamų antraščių (Sent:/Date:/Išsiųsta:/Data:), o NE naujausio laiško datą. Jei pradinės datos nesimato — palik null. ATSKIRK: pradinio ŠIO užsakymo laiško data ≠ ankstesnio pirkimo data. Jei klientas NAUJAME laiške prašo „pakartoti kaip anksčiau / tokių pačių kaip pirkome" — tai NAUJAS užsakymas, jo data yra to prašymo laiško data, o ne cituojamo seno susirašinėjimo apie ankstesnį pirkimą.
- „pasiulymo_terminas" — iki kada prašoma pateikti pasiūlymą / atsakymą, jei nurodyta (pvz., „2026-07-21 10:00").
- „sutartis" — sutarties numeris, jei laiške ar prieduose minimas (pvz., „SUT24-61", „KIT26-403"). Rašyk tiksliai taip, kaip parašyta, be tarpo klaidų.
- Kiekius rašyk taip, kaip nurodyta (pvz., „10 vnt.", „2 dėž.", „5,5 kg"). Jei kiekis nenurodytas, rašyk „nenurodyta".
- „pasitikejimas" — tavo įsitikinimas 0..1, kad klasifikacija teisinga.
- „santrauka" — 1–2 sakiniai, kas šiame laiške yra.
- „ispejimas" — APSAUGA NUO VYKDYMO KLAIDŲ. Užpildyk TIK tada, kai laiškas nėra paprastas „užsakykite prekes" atvejis ir vykdytojas, neįsiskaitęs konteksto, gali padaryti brangią klaidą. Įspėjimą pradėk nuo to, KO NEDARYTI, tada trumpai kodėl. Tipiniai atvejai: (1) prašoma TIK pakoreguoti / įregistruoti / perregistruoti JAU ĮVYKDYTĄ užsakymą kliento pirkimų sistemoje — rašyk pvz. „NEUŽSAKINĖTI PREKIŲ IŠ NAUJO — jos jau pristatytos, prašoma tik pakoreguoti ir įregistruoti užsakymą per kliento pirkimų sistemą“; (2) sąskaita jau išrašyta, prašoma tik dokumentų pataisymo; (3) laiškas galimai dubliuoja ankstesnį užsakymą iš sąrašo; (4) prašymas dviprasmiškas ar prieštaringas (neaišku, užsakyti ar tik teirautis kainos); (5) KIEKIO PROBLEMA — flaguok VISUS kiekio netikslumus, ne tik dviprasmybes: kiekis IŠVIS NENURODYTAS (rašai „nenurodyta"); dviprasmiškas — pvz., „2 dėž 24 vnt." (neaišku, ar 2 dėžės PO 24 vnt. = 48, ar 2 dėžės, kuriose iš viso 24 vnt.); „3 pak", kai nežinomas pakuotės dydis; kiekis be mato vieneto, kai prekė matuojama keliais būdais (metrai ar vienetai); laiške ir priede nurodyti SKIRTINGI kiekiai; kiekis nurodytas tik daliai prekių; kiekis su sąlyga („apie", „maždaug", „~", „priklausys nuo"); bet koks kitas neapibrėžtumas, dėl kurio užsakant galima suklysti. Kiekio įspėjime įvardink KONKREČIAS prekes ir kas neaišku, pvz.: „PASITIKSLINTI KIEKĮ prieš užsakant: „2 dėž 24 vnt." gali reikšti 48 vnt. (2×24) arba 24 vnt.; prekei X kiekis išvis nenurodytas". Įprastoms pastaboms šio lauko NENAUDOK — jei viskas aišku, palik null.
- Atsakymo tekstus rašyk lietuvių kalba (prekes — originalo kalba, kaip laiške)."""

_NULLABLE_STR = {"anyOf": [{"type": "string"}, {"type": "null"}]}

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "ar_uzsakymas": {"type": "boolean"},
        "pasitikejimas": {"type": "number"},
        "klientas": _NULLABLE_STR,
        "kliento_el_pastas": _NULLABLE_STR,
        "kliento_adresas": _NULLABLE_STR,
        "kliento_telefonas": _NULLABLE_STR,
        "uzsakymo_data": _NULLABLE_STR,
        "pristatymo_data": _NULLABLE_STR,
        "pasiulymo_terminas": _NULLABLE_STR,
        "sutartis": _NULLABLE_STR,
        "prekes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "preke": {"type": "string"},
                    "kiekis": {"type": "string"},
                    "pastabos": _NULLABLE_STR,
                },
                "required": ["preke", "kiekis", "pastabos"],
                "additionalProperties": False,
            },
        },
        "pastabos": _NULLABLE_STR,
        "atnaujina_koda": _NULLABLE_STR,
        "santrauka": {"type": "string"},
        "ispejimas": _NULLABLE_STR,
    },
    "required": ["ar_uzsakymas", "pasitikejimas", "klientas", "kliento_el_pastas",
                 "kliento_adresas", "kliento_telefonas", "uzsakymo_data",
                 "pristatymo_data", "pasiulymo_terminas", "sutartis",
                 "prekes", "pastabos", "atnaujina_koda", "santrauka", "ispejimas"],
    "additionalProperties": False,
}


def _supports_effort(model: str) -> bool:
    # effort parametro Haiku 4.5 nepalaiko — siunčiant gaunama 400 klaida
    return not model.startswith("claude-haiku")


# Kainos (USD už milijoną žetonų): įvestis / išvestis
PRICES = {
    "claude-opus-5": (5.0, 25.0),
    "claude-sonnet-5": (2.0, 10.0),   # įvadinė kaina iki 2026-08-31 (vėliau 3/15)
    "claude-haiku-4-5": (1.0, 5.0),
}


def _usage_info(response, model: str) -> dict:
    """Žetonų kiekis ir apytikslė vieno laiško kaina."""
    u = response.usage
    inp = getattr(u, "input_tokens", 0) or 0
    out = getattr(u, "output_tokens", 0) or 0
    cache_read = getattr(u, "cache_read_input_tokens", 0) or 0
    cache_write = getattr(u, "cache_creation_input_tokens", 0) or 0
    p_in, p_out = PRICES.get(model, (5.0, 25.0))
    usd = ((inp + cache_write * 1.25 + cache_read * 0.1) * p_in + out * p_out) / 1_000_000
    return {"model": model, "input": inp, "output": out,
            "cache_read": cache_read, "cache_write": cache_write,
            "usd": usd, "eur_cents": usd * 0.92 * 100}


def analyze_email(client: anthropic.Anthropic, model: str, email: dict,
                  min_image_kb: int = 15, recent_orders: str = "",
                  effort: str = "medium", max_image_edge: int = 2200,
                  chain_order: dict = None, manual_pick: bool = False) -> dict:
    """Grąžina analizės žodyną pagal OUTPUT_SCHEMA. Saugu kviesti iš kelių gijų lygiagrečiai.
    Meta išimtis tinklo/API klaidų atveju — kviečiantysis sprendžia, ar kartoti."""
    blocks, att_texts, listing = build_attachment_content(
        email["attachments"], min_image_kb=min_image_kb, max_image_edge=max_image_edge
    )

    text = (
        f"Siuntėjas: {email['sender_name']} <{email['sender_email']}>\n"
        f"Tema: {email['subject']}\n"
        f"Gauta: {email['received']:%Y-%m-%d %H:%M}\n"
        f"Priedai: {', '.join(listing) if listing else 'nėra'}\n\n"
        f"--- Laiško tekstas ---\n{email['body'] or '(tuščias)'}"
    )
    if att_texts:
        text += "\n\n--- Priedų turinys ---\n" + "\n\n".join(att_texts)
    if chain_order:
        text += (
            "\n\n--- DĖMESIO: ŠIO SUSIRAŠINĖJIMO GRANDINĖS UŽSAKYMAS JAU UŽREGISTRUOTAS ---\n"
            f"Kodas: {chain_order.get('kodas')} | Klientas: {chain_order.get('klientas')} | "
            f"Data: {chain_order.get('received_at')}\n"
            f"Prekės: {chain_order.get('prekes_santrauka') or '(neišskaidyta)'}\n"
            "Todėl NEKURK naujo užsakymo iš šio laiško (ar_uzsakymas=false), NEBENT jame yra "
            "akivaizdžiai atskiras, su ankstesniu nesusijęs naujas užsakymas.\n"
            f"- Jei laiške yra NAUJOS informacijos (naujos prekės, kiekiai, datos, kontaktai) — "
            f"atnaujina_koda=\"{chain_order.get('kodas')}\".\n"
            "- Jei tai tik patvirtinimas, padėka ar tiekėjo atsakymas be naujos informacijos — "
            "atnaujina_koda=null.")
    if manual_pick:
        text += (
            "\n\n--- RANKINIS PASIRINKIMAS ---\n"
            "Šį laišką vartotojas atrinko analizei PATS, apeidamas filtrus — jis mano, "
            "kad čia gali būti užsakymas ar užklausa. Vertink vien pagal TURINĮ: "
            "siuntėjas gali būti bet kas (nebūtinai įprastas klientas ar jo domenas). "
            "Jei laiške prašoma prekių, kainos ar pasiūlymo — tai užsakymas-užklausa. "
            "Grandinių ir dubliavimo taisyklės galioja kaip visada.")
    if recent_orders:
        text += ("\n\n--- Neseniai užregistruoti užsakymai (dubliavimui išvengti) ---\n"
                 "Kodas | Klientas | Gauta | Prekės\n" + recent_orders)

    content = blocks + [{"type": "text", "text": text}]

    output_config = {"format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}}
    if _supports_effort(model):
        output_config["effort"] = effort

    kwargs = dict(
        model=model,
        max_tokens=16000,
        system=[{"type": "text", "text": SYSTEM_PROMPT,
                 "cache_control": {"type": "ephemeral"}}],
        messages=[{"role": "user", "content": content}],
        output_config=output_config,
    )
    try:
        # Serverio pusės fallback: jei saugumo klasifikatorius atmestų užklausą,
        # ji automatiškai perleidžiama per atsarginį modelį.
        response = client.messages.create(
            **kwargs,
            extra_headers={"anthropic-beta": "server-side-fallback-2026-07-01"},
            extra_body={"fallbacks": "default"},
        )
    except anthropic.BadRequestError as e:
        # Kartojame be fallback tik jei klaida susijusi su pačiu fallback parametru;
        # kitas 400 priežastis (per didelis priedas ir pan.) keliame aukštyn.
        if "fallback" in str(e).lower():
            response = client.messages.create(**kwargs)
        else:
            raise

    if response.stop_reason == "refusal":
        return {
            "ar_uzsakymas": False, "pasitikejimas": 0.0, "klientas": None,
            "kliento_el_pastas": None, "kliento_adresas": None, "kliento_telefonas": None,
            "uzsakymo_data": None, "pristatymo_data": None, "pasiulymo_terminas": None,
            "sutartis": None, "atnaujina_koda": None,
            "prekes": [], "pastabos": "API atsisakė analizuoti šį laišką.",
            "santrauka": "Neanalizuota (saugumo atsisakymas).",
        }
    if response.stop_reason == "max_tokens":
        raise RuntimeError("Atsakymas nutrūko (max_tokens) — bus bandoma kitą kartą")

    text_out = next((b.text for b in response.content if b.type == "text"), None)
    if text_out is None:
        raise RuntimeError(f"Netikėtas API atsakymas (stop_reason={response.stop_reason})")
    result = json.loads(text_out)
    result["_usage"] = _usage_info(response, model)  # kainos apskaitai
    return result
