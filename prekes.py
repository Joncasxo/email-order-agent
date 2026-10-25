# -*- coding: utf-8 -*-
"""Užsakymo duomenų tvarkymas po AI ištraukimo ir vienodas jų rodymas.

AI grąžina struktūrizuotus laukus, o šis modulis juos dar kartą sutvarko
kodu (be AI, 0 žetonų): nuima šiukšles, suvienodina vienetus, telefoną,
el. paštą, išmeta tuščias ir pasikartojančias prekių eilutes.

Prekės eilutė: {preke, kodas, kiekis, vienetas, pastabos, nuoroda}.
Seni įrašai turi tik {preke, kiekis, pastabos} (kiekis — laisvas tekstas) —
visos funkcijos čia su jais irgi veikia.
"""
import re

ITEM_KEYS = ("preke", "kodas", "kiekis", "vienetas", "pastabos", "nuoroda")

# Reikšmės, kurios iš tikro reiškia „nieko nėra"
_TUSCIA = {"", "-", "—", "–", "?", "n/a", "na", "none", "null", "nėra", "nera",
           "nenurodyta", "nenurodytas", "nežinoma", "nezinoma", "neaišku"}

# Vienetų suvienodinimas: rašyba laiškuose -> viena forma
_VIENETAI = {
    "vnt": "vnt.", "vnt.": "vnt.", "vienetai": "vnt.", "vienetų": "vnt.", "vienetas": "vnt.",
    "pcs": "vnt.", "pc": "vnt.", "pcs.": "vnt.", "st": "vnt.", "stk": "vnt.", "шт": "vnt.",
    "kompl": "kompl.", "kompl.": "kompl.", "komplektai": "kompl.", "komplektas": "kompl.",
    "set": "kompl.", "sets": "kompl.",
    "pak": "pak.", "pak.": "pak.", "pakuotė": "pak.", "pakuotės": "pak.", "pakuotes": "pak.",
    "dėž": "dėž.", "dėž.": "dėž.", "dez": "dėž.", "dez.": "dėž.", "dėžė": "dėž.",
    "dėžės": "dėž.", "dėžių": "dėž.", "box": "dėž.", "boxes": "dėž.",
    "rul": "rul.", "rul.": "rul.", "ritinys": "rul.", "ritiniai": "rul.", "roll": "rul.",
    "pad": "pad.", "pad.": "pad.", "padėklas": "pad.", "padėklai": "pad.", "pal": "pad.",
    "m": "m", "m.": "m", "metrai": "m", "metrų": "m", "metras": "m", "bėg. m": "m",
    "m2": "m²", "m²": "m²", "kv. m": "m²", "kv.m": "m²", "kvm": "m²",
    "m3": "m³", "m³": "m³", "kub. m": "m³", "kub.m": "m³",
    "kg": "kg", "kg.": "kg", "t": "t", "t.": "t", "tonos": "t",
    "l": "l", "l.": "l", "ltr": "l", "litrai": "l", "litrų": "l",
    "val": "val.", "val.": "val.", "h": "val.",
    "pora": "por.", "poros": "por.", "porų": "por.", "por.": "por.",
}


def tuscia(v) -> bool:
    return v is None or str(v).strip().lower() in _TUSCIA


def svarus(v, viena_eilute: bool = True):
    """Tekstas be perteklinių tarpų; „nieko nėra" reikšmės -> None."""
    if v is None:
        return None
    t = str(v).replace(" ", " ").strip()
    if viena_eilute:
        t = re.sub(r"\s*[\r\n]+\s*", "; ", t)
    t = re.sub(r"[ \t]{2,}", " ", t).strip(" ;,")
    return None if tuscia(t) else t


def vienetas(v):
    t = svarus(v)
    if not t:
        return None
    return _VIENETAI.get(t.lower(), _VIENETAI.get(t.lower().rstrip("."), t))


def skaicius_tekstu(v):
    """10 -> „10", 5.5 -> „5,5" (lietuviškas kablelis). Ne skaičius -> tekstas kaip yra."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if float(v).is_integer():
            return str(int(v))
        return f"{v:.3f}".rstrip("0").rstrip(".").replace(".", ",")
    return svarus(v)


def telefonas(v):
    t = svarus(v)
    if not t:
        return None
    # keli numeriai — tvarkome kiekvieną atskirai
    dalys = [d for d in re.split(r"[;,/]| ir ", t) if d.strip()]
    rez = []
    for d in dalys:
        sk = re.sub(r"[^\d+]", "", d)
        if sk.startswith("00"):
            sk = "+" + sk[2:]
        if re.fullmatch(r"8\d{8}", sk):          # 86xxxxxxx -> +3706xxxxxxx
            sk = "+370" + sk[1:]
        if re.fullmatch(r"\+370\d{8}", sk):
            rez.append(f"+370 {sk[4:7]} {sk[7:]}")
        else:
            rez.append(d.strip())
    return "; ".join(rez) or None


def el_pastas(v):
    t = svarus(v)
    if not t:
        return None
    t = t.replace("mailto:", "")
    rasti = re.findall(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", t)
    return "; ".join(dict.fromkeys(a.lower() for a in rasti)) if rasti else t


def asmuo(v):
    """Vardas Pavardė: tarpai sutvarkyti, VISOS DIDŽIOSIOS -> Pirma Didžioji."""
    t = svarus(v)
    if not t:
        return None
    if t.isupper() or t.islower():
        t = " ".join(w[:1].upper() + w[1:].lower() for w in t.split())
    return t


def prekes_eilute(it: dict):
    """Viena prekės eilutė iš AI arba rankos -> tvarkinga forma (arba None, jei tuščia)."""
    if not isinstance(it, dict):
        return None
    preke = svarus(it.get("preke"))
    kodas = svarus(it.get("kodas"))
    nuoroda = svarus(it.get("nuoroda"))
    pastabos = svarus(it.get("pastabos"), viena_eilute=False)
    vnt = vienetas(it.get("vienetas"))
    kiekis = skaicius_tekstu(it.get("kiekis"))
    if kiekis and not vnt:
        # „10 vnt." viename lauke (seni įrašai, ranka įvesta) -> 10 + vnt.
        m = re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*([^\d\s].*)", kiekis)
        if m and vienetas(m.group(2)) in _VIENETAI.values():
            kiekis, vnt = m.group(1).replace(".", ","), vienetas(m.group(2))
    if not (preke or kodas or nuoroda):
        return None
    return {"preke": preke or "", "kodas": kodas, "kiekis": kiekis or "",
            "vienetas": vnt, "pastabos": pastabos, "nuoroda": nuoroda}


def tvarkyti_prekes(items) -> list:
    """Sutvarko visas eilutes ir išmeta visiškai vienodus pasikartojimus
    (pvz., ta pati prekė ir laiško tekste, ir priede)."""
    rez, matyta = [], set()
    for it in items or []:
        e = prekes_eilute(it)
        if e is None:
            continue
        raktas = tuple((e.get(k) or "").lower() for k in ("preke", "kodas", "kiekis", "vienetas"))
        if raktas in matyta:
            continue
        matyta.add(raktas)
        rez.append(e)
    return rez


def tvarkyti_analize(a: dict) -> dict:
    """Visa AI analizė po ištraukimo — vienodas, švarus formatas."""
    if not isinstance(a, dict):
        return a
    for k in ("kliento_adresas", "kliento_imone", "sutartis", "pristatymo_data",
              "pasiulymo_terminas", "uzsakymo_data", "ispejimas", "atnaujina_koda"):
        if k in a:
            a[k] = svarus(a.get(k))
    if "pastabos" in a:
        a["pastabos"] = svarus(a.get("pastabos"), viena_eilute=False)
    if "klientas" in a:
        a["klientas"] = asmuo(a.get("klientas"))
    if "kliento_telefonas" in a:
        a["kliento_telefonas"] = telefonas(a.get("kliento_telefonas"))
    if "kliento_el_pastas" in a:
        a["kliento_el_pastas"] = el_pastas(a.get("kliento_el_pastas"))
    if a.get("sutartis"):
        a["sutartis"] = re.sub(r"\s*-\s*", "-", a["sutartis"]).upper()
    a["prekes"] = tvarkyti_prekes(a.get("prekes"))
    return a


# ---------------- rodymas ----------------

def kiekis_tekstu(it: dict) -> str:
    """„10 vnt." — kiekis kartu su vienetu (seni įrašai — kaip buvo)."""
    k = str(it.get("kiekis") or "").strip()
    v = (it.get("vienetas") or "").strip()
    if not k:
        return "nenurodyta" if not v else v
    return f"{k} {v}".strip()


def pilnas_pavadinimas(it: dict) -> str:
    """Prekė su kodu — kur reikia vienos eilutės (sąrašas, laiškas tiekėjui, Excel)."""
    p = (it.get("preke") or "").strip()
    kodas = (it.get("kodas") or "").strip()
    if kodas and kodas.lower() not in p.lower():
        p = f"{p} (kodas: {kodas})" if p else f"kodas: {kodas}"
    return p


def trumpas_sarasas(items, kiek: int = 4) -> str:
    """„Prekė (10 vnt.); Prekė2 (2 m) … (+3)" — užsakymų sąrašo stulpeliui."""
    items = [prekes_eilute(it) or it for it in (items or [])]
    s = "; ".join(f"{pilnas_pavadinimas(it)} ({kiekis_tekstu(it)})" for it in items[:kiek])
    if len(items) > kiek:
        s += f" … (+{len(items) - kiek})"
    return s
