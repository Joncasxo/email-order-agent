# -*- coding: utf-8 -*-
"""Laiškų skaitymas per Microsoft Graph API — alternatyva outlook_client (COM).

Skirtumai nuo COM kelio:
- skaito tiesiai iš serverio (laiškas matomas iškart, be vietinės sinchronizacijos);
- veikia ir be paleistos Outlook programos;
- laiško ID čia — Graph 'id' (KITOKS nei COM EntryID, todėl processed žymės
  tarp backend'ų nesuderinamos — perjungiant darbinę programą reikės migracijos).

Prisijungimas: MSAL su interaktyviu langu PIRMĄ kartą; toliau — tylus
atnaujinimas iš talpyklos, kuri saugoma DPAPI užšifruota (graph_token.bin).

Sąsaja atkartoja outlook_client.fetch_emails_since / load_email, kad app.py
galėtų perjungti šaltinį vienu config įrašu.
"""
import base64
import datetime
import os

import msal
import requests
import win32crypt

import config as _cfg_mod

SCOPES = ["Mail.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"
MAX_ATTACHMENT_MB = 20


def _prisijungimo_duomenys():
    """Client ID / Tenant ID / pašto adresas — iš config.json (⚙ Nustatymai).

    Laikoma nustatymuose, o ne kode, kad programa būtų universali: kiekvienas
    naudotojas įrašo SAVO Entra programėlės duomenis (žr. SKAITYK.md).
    """
    c = _cfg_mod.Config()
    client_id = (c.get("graph_client_id") or "").strip()
    tenant_id = (c.get("graph_tenant_id") or "").strip()
    login_hint = (c.get("graph_login_hint") or "").strip()
    if not client_id or not tenant_id:
        raise RuntimeError(
            "Neįvesti Microsoft Graph prisijungimo duomenys. Nustatymuose "
            "įrašykite programėlės (Client) ID ir katalogo (Tenant) ID — "
            "kaip juos gauti, žr. SKAITYK.md skyrių „Kaip prijungti savo paštą“.")
    return client_id, tenant_id, login_hint


def _login_hint() -> str:
    """Naudotojo pašto adresas mažosiomis (siųstų laiškų atpažinimui)."""
    try:
        return (_cfg_mod.Config().get("graph_login_hint") or "").strip().lower()
    except Exception:
        return ""

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_PATH = os.path.join(APP_DIR, "graph_token.bin")


# ---------------- žetonų talpykla (DPAPI) ----------------

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "rb") as f:
                blob = f.read()
            data = win32crypt.CryptUnprotectData(blob, None, None, None, 0)[1]
            cache.deserialize(data.decode("utf-8"))
        except Exception:
            pass  # sugadinta talpykla — prisijungsime iš naujo
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if not cache.has_state_changed:
        return
    try:
        blob = win32crypt.CryptProtectData(
            cache.serialize().encode("utf-8"), "graph_token", None, None, None, 0)
        with open(CACHE_PATH, "wb") as f:
            f.write(blob)
    except Exception:
        pass


def get_token(interactive_ok: bool = True) -> str:
    """Grąžina galiojantį prieigos žetoną. Pirmą kartą (arba atšaukus teises)
    atidaro Microsoft prisijungimo langą; toliau — tyliai iš talpyklos."""
    client_id, tenant_id, login_hint = _prisijungimo_duomenys()
    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id, authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache)
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result and interactive_ok:
        result = app.acquire_token_interactive(
            SCOPES, login_hint=login_hint or None, timeout=300)
    _save_cache(cache)
    if not result or "access_token" not in result:
        klaida = (result or {}).get("error_description", "nežinoma klaida")
        raise RuntimeError(f"Graph prisijungimas nepavyko: {klaida[:300]}")
    return result["access_token"]


# ---------------- HTML -> tekstas su spalvų žymėmis ----------------
# Klientai svarbias vietas (pvz., pasirinktą lentelės eilutę) dažnai žymi
# spalvomis. Verčiant į gryną tekstą spalvos dingsta, todėl AI jų nemato.
# Šis konverteris spalvotą tekstą apgaubia žyme, pvz., „[RAUDONAI PAŽYMĖTA: ...]",
# o lentelių eilutes/langelius atskiria naujomis eilutėmis ir tabuliacija.

import re as _re
from html.parser import HTMLParser as _HTMLParser

_CSS_VARDAI = {
    "red": (255, 0, 0), "darkred": (139, 0, 0), "crimson": (220, 20, 60),
    "firebrick": (178, 34, 34), "maroon": (128, 0, 0), "orangered": (255, 69, 0),
    "green": (0, 128, 0), "darkgreen": (0, 100, 0), "lime": (0, 255, 0),
    "blue": (0, 0, 255), "navy": (0, 0, 128), "orange": (255, 165, 0),
    "yellow": (255, 255, 0), "gold": (255, 215, 0), "purple": (128, 0, 128),
    "magenta": (255, 0, 255), "fuchsia": (255, 0, 255),
}


def _css_rgb(reiksme):
    s = str(reiksme).strip().lower().rstrip("!important").strip()
    if s in _CSS_VARDAI:
        return _CSS_VARDAI[s]
    m = _re.fullmatch(r"#([0-9a-f]{6})", s)
    if m:
        h = m.group(1)
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    m = _re.fullmatch(r"#([0-9a-f]{3})", s)
    if m:
        h = m.group(1)
        return int(h[0] * 2, 16), int(h[1] * 2, 16), int(h[2] * 2, 16)
    m = _re.fullmatch(r"rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", s)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    return None


def _spalvos_prieveiksmis(reiksme):
    """'#c00000' -> 'raudonai'; None — įprastoms (juoda/pilka/balta) spalvoms."""
    rgb = _css_rgb(reiksme)
    if rgb is None:
        return None
    r, g, b = rgb
    if max(rgb) - min(rgb) < 30:
        return None          # pilka / juoda / balta — ne žymėjimas
    if max(rgb) < 90:
        return None          # beveik juoda
    if r == max(rgb) and r > g + 50 and r > b + 50:
        return "oranžiškai" if g > 120 else "raudonai"
    if r > 180 and g > 160 and b < 120:
        return "geltonai"
    if g == max(rgb) and g > r + 40 and g > b + 40:
        return "žaliai"
    if b == max(rgb) and b > r + 40 and b > g + 40:
        return "mėlynai"
    if r > 120 and b > 120 and g < 100:
        return "violetiškai"
    return "spalvotai"


def _stiliaus_zyme(style):
    """Iš style atributo — žymės tekstas arba None."""
    if not style:
        return None
    m = _re.search(r"(?<![-\w])color\s*:\s*([^;]+)", style, _re.I)
    if m:
        adv = _spalvos_prieveiksmis(m.group(1))
        if adv:
            return f"{adv.upper()} PAŽYMĖTA"
    m = _re.search(r"background(?:-color)?\s*:\s*([^;]+)", style, _re.I)
    if m:
        adv = _spalvos_prieveiksmis(m.group(1))
        if adv:
            return f"{adv.upper()} PARYŠKINTA"
    return None


class _SpalvuTekstas(_HTMLParser):
    _NEMATOMI = {"style", "script", "head", "title"}
    _BLOKAI = {"p", "div", "tr", "table", "li", "ul", "ol",
               "h1", "h2", "h3", "h4", "blockquote"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.dalys = []
        self.stekas = []      # (tag, zyme|None)
        self.nematoma = 0
        self.a_gylis = 0
        self.run = None       # [zyme, [gabalai]]

    def _dabartine_zyme(self):
        if self.a_gylis:
            return None       # nuorodų mėlynumas — ne žymėjimas
        for _, z in reversed(self.stekas):
            if z:
                return z
        return None

    def _uzdaryti_run(self):
        if not self.run:
            return
        zyme, gabalai = self.run
        self.run = None
        tekstas = "".join(gabalai)
        svarus = " ".join(tekstas.split())
        # per ilgas „spalvotas" blokas (pvz., mėlyna citata) — žymė tik trukdytų
        if svarus and len(svarus) <= 400:
            self.dalys.append(f" [{zyme}: {svarus}] ")
        else:
            self.dalys.append(tekstas)

    def _prideti(self, tekstas):
        zyme = self._dabartine_zyme()
        if self.run and self.run[0] != zyme:
            self._uzdaryti_run()
        if zyme:
            if not self.run:
                self.run = [zyme, []]
            self.run[1].append(tekstas)
        else:
            self.dalys.append(tekstas)

    def handle_starttag(self, tag, attrs):
        if tag in self._NEMATOMI:
            self.nematoma += 1
            return
        if self.nematoma:
            return
        if tag == "br":
            self._prideti("\n")
            return
        if tag == "a":
            self.a_gylis += 1
        if tag in self._BLOKAI:
            self._uzdaryti_run()
            self.dalys.append("\n")
        if tag in ("td", "th"):
            self._prideti("\t")
        style = dict(attrs).get("style") or ""
        self.stekas.append((tag, _stiliaus_zyme(style)))

    def handle_startendtag(self, tag, attrs):
        if tag == "br" and not self.nematoma:
            self._prideti("\n")

    def handle_endtag(self, tag):
        if tag in self._NEMATOMI:
            self.nematoma = max(0, self.nematoma - 1)
            return
        if self.nematoma:
            return
        if tag == "a":
            self.a_gylis = max(0, self.a_gylis - 1)
        turejo_zyme = False
        for i in range(len(self.stekas) - 1, -1, -1):
            if self.stekas[i][0] == tag:
                turejo_zyme = any(z for _, z in self.stekas[i:])
                del self.stekas[i:]
                break
        if turejo_zyme:
            self._uzdaryti_run()
        if tag in self._BLOKAI:
            self.dalys.append("\n")

    def handle_data(self, data):
        if self.nematoma or not data:
            return
        self._prideti(data)

    def tekstas(self):
        self._uzdaryti_run()
        t = "".join(self.dalys)
        t = _re.sub(r"[ \t]*\n[ \t]*", "\n", t)
        t = _re.sub(r"\n{3,}", "\n\n", t)
        t = _re.sub(r"[ ]{2,}", " ", t)
        return t.strip()


def _html_i_teksta(html):
    try:
        p = _SpalvuTekstas()
        p.feed(html)
        p.close()
        return p.tekstas()
    except Exception:
        # atsarginis kelias — senas grubus valymas
        t = _re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", html,
                    flags=_re.S | _re.I)
        t = _re.sub(r"<[^>]+>", " ", t)
        return _re.sub(r"[ \t]+", " ", t)


# ---------------- laiškų skaitymas ----------------

def _sender_matches(sender_name: str, sender_email: str, filters) -> bool:
    if not filters:
        return True
    hay = f"{sender_name} <{sender_email}>".lower()
    for entry in filters:
        words = [w for w in str(entry).lower().split() if w]
        if words and all(w in hay for w in words):
            return True
    return False


def _msg_to_email(m: dict, attachments: list) -> dict:
    nuo = (m.get("from") or {}).get("emailAddress", {}) or {}
    gauta = m.get("receivedDateTime") or ""
    try:
        # Graph laikas UTC (…Z) — verčiame į vietinį, kad sutaptų su COM keliu
        dt = datetime.datetime.fromisoformat(gauta.replace("Z", "+00:00"))
        received = dt.astimezone().replace(tzinfo=None)
    except ValueError:
        received = datetime.datetime.now()
    body = m.get("body") or {}
    tekstas = body.get("content") or ""
    if (body.get("contentType") or "").lower() == "html":
        # spalvas išsaugantis vertimas: pažymėtos vietos tampa [RAUDONAI PAŽYMĖTA: ...]
        tekstas = _html_i_teksta(tekstas)
    return {
        "entry_id": m["id"],                    # Graph id (ne COM EntryID!)
        "store_id": "graph",
        "conversation_id": m.get("conversationId") or "",
        "subject": (m.get("subject") or "")[:500],
        "sender_name": nuo.get("name") or "",
        "sender_email": nuo.get("address") or "",
        "received": received,
        "body": tekstas[:30000],
        "attachments": attachments,
    }


def _fetch_attachments(headers, msg_id, on_skip=None, subject=""):
    out = []
    try:
        r = requests.get(f"{GRAPH}/me/messages/{msg_id}/attachments",
                         headers=headers, timeout=60)
        r.raise_for_status()
        for a in r.json().get("value", []):
            if a.get("@odata.type") != "#microsoft.graph.fileAttachment":
                continue  # įterpti laiškai/kalendoriai — praleidžiam
            if (a.get("size") or 0) > MAX_ATTACHMENT_MB * 1024 * 1024:
                continue
            data = base64.b64decode(a.get("contentBytes") or "")
            out.append((a.get("name") or "priedas", data, bool(a.get("isInline"))))
    except Exception as e:
        if on_skip:
            try:
                on_skip(subject, f"priedų klaida: {e}")
            except Exception:
                pass
    return out


def fetch_emails_since(since: datetime.datetime, max_count: int = 50,
                       skip=None, sender_filters=None, on_skip=None):
    """Kaip outlook_client.fetch_emails_since, tik per Graph API.

    Kūnas imamas HTML (be Prefer:text) — spalvų žymėjimą laiške išsaugo
    _html_i_teksta, kad AI matytų, kas pažymėta raudonai ir pan.
    """
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    # Graph filtruojame UTC laiku
    since_utc = since.astimezone().astimezone(datetime.timezone.utc)
    url = (f"{GRAPH}/me/mailFolders/inbox/messages"
           f"?$top=50&$orderby=receivedDateTime desc"
           f"&$select=subject,from,receivedDateTime,conversationId,body,hasAttachments"
           f"&$filter=receivedDateTime ge {since_utc:%Y-%m-%dT%H:%M:%SZ}")
    emails = []
    while url and len(emails) < max_count:
        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
        for m in data.get("value", []):
            nuo = (m.get("from") or {}).get("emailAddress", {}) or {}
            if not _sender_matches(nuo.get("name") or "", nuo.get("address") or "",
                                   sender_filters):
                continue
            if skip is not None and skip(m["id"]):
                continue
            atts = (_fetch_attachments(headers, m["id"], on_skip,
                                       m.get("subject") or "")
                    if m.get("hasAttachments") else [])
            emails.append(_msg_to_email(m, atts))
            if len(emails) >= max_count:
                break
        url = data.get("@odata.nextLink")
    return emails


def list_emails_since(since, max_count: int = 400, sender_filters=None):
    """LENGVAS sarasas rinkikliui: be teksto ir priedu turinio."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    since_utc = since.astimezone().astimezone(datetime.timezone.utc)
    url = (f"{GRAPH}/me/mailFolders/inbox/messages"
           f"?$top=50&$orderby=receivedDateTime desc"
           f"&$select=subject,from,receivedDateTime,hasAttachments"
           f"&$filter=receivedDateTime ge {since_utc:%Y-%m-%dT%H:%M:%SZ}")
    out = []
    while url and len(out) < max_count:
        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
        for m in data.get("value", []):
            nuo = (m.get("from") or {}).get("emailAddress", {}) or {}
            name, addr = nuo.get("name") or "", nuo.get("address") or ""
            if sender_filters and not _sender_matches(name, addr, sender_filters):
                continue
            gauta = m.get("receivedDateTime") or ""
            try:
                dt = datetime.datetime.fromisoformat(gauta.replace("Z", "+00:00"))
                received = dt.astimezone().replace(tzinfo=None)
            except ValueError:
                received = datetime.datetime.now()
            out.append({
                "entry_id": m["id"], "store_id": "graph",
                "subject": (m.get("subject") or "")[:500],
                "sender_name": name, "sender_email": addr,
                "received": received,
                "n_att": 1 if m.get("hasAttachments") else 0,
            })
            if len(out) >= max_count:
                break
        url = data.get("@odata.nextLink")
    return out


def list_sent_since(since, max_count: int = 150):
    """Išsiųstų laiškų sąrašas (temos + gavėjai) — užklausų sekimui."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    since_utc = since.astimezone().astimezone(datetime.timezone.utc)
    url = (f"{GRAPH}/me/mailFolders/sentitems/messages"
           f"?$top=50&$orderby=sentDateTime desc"
           f"&$select=subject,sentDateTime,toRecipients,ccRecipients"
           f"&$filter=sentDateTime ge {since_utc:%Y-%m-%dT%H:%M:%SZ}")
    out = []
    while url and len(out) < max_count:
        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
        for m in data.get("value", []):
            gavejai = []
            for lauka in ("toRecipients", "ccRecipients"):
                for g in (m.get(lauka) or []):
                    a = ((g.get("emailAddress") or {}).get("address") or "").strip()
                    if a:
                        gavejai.append(a)
            out.append({"entry_id": m["id"],
                        "subject": (m.get("subject") or "")[:500],
                        "sent": m.get("sentDateTime") or "",
                        "to": gavejai})
            if len(out) >= max_count:
                break
        url = data.get("@odata.nextLink")
    return out


def find_by_code(kodas: str, max_count: int = 25):
    """Visi laiškai (siųsti ir gauti), kurių temoje/tekste yra užsakymo kodas.
    Grąžina sąrašą su kryptimi, data, adresais ir webLink atidarymui."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    r = requests.get(
        f"{GRAPH}/me/messages?$search=\"{kodas}\""
        f"&$select=subject,from,toRecipients,receivedDateTime,webLink"
        f"&$top={max_count}",
        headers=headers, timeout=60)
    r.raise_for_status()
    out = []
    for m in r.json().get("value", []):
        nuo = ((m.get("from") or {}).get("emailAddress") or {})
        nuo_adr = (nuo.get("address") or "").lower()
        gavejai = [((g.get("emailAddress") or {}).get("address") or "")
                   for g in (m.get("toRecipients") or [])]
        issiustas = nuo_adr == _login_hint()
        gauta = m.get("receivedDateTime") or ""
        try:
            dt = datetime.datetime.fromisoformat(gauta.replace("Z", "+00:00"))
            kada = dt.astimezone().replace(tzinfo=None)
        except ValueError:
            kada = None
        out.append({
            "kryptis": "išsiųsta" if issiustas else "gauta",
            "kada": kada,
            "kas": ", ".join(a for a in gavejai if a) if issiustas
                   else (nuo.get("name") or nuo_adr),
            "subject": (m.get("subject") or "")[:200],
            "webLink": m.get("webLink"),
        })
    out.sort(key=lambda x: x["kada"] or datetime.datetime.min, reverse=True)
    return out


def _full_messages(ids, headers):
    """Paima nurodytų laiškų pilnus tekstus grandinės peržiūrai."""
    out = []
    for mid in ids:
        try:
            m = requests.get(
                f"{GRAPH}/me/messages/{mid}"
                "?$select=subject,from,toRecipients,receivedDateTime,body,webLink",
                headers=headers, timeout=30).json()
        except Exception:
            continue
        nuo = ((m.get("from") or {}).get("emailAddress") or {})
        gavejai = [((g.get("emailAddress") or {}).get("address") or "")
                   for g in (m.get("toRecipients") or [])]
        gauta = m.get("receivedDateTime") or ""
        try:
            dt = datetime.datetime.fromisoformat(gauta.replace("Z", "+00:00"))
            kada = dt.astimezone().replace(tzinfo=None)
        except ValueError:
            kada = None
        tekstas = ((m.get("body") or {}).get("content") or "").strip()
        out.append({
            "kada": kada,
            "nuo": nuo.get("address") or "",
            "nuo_vardas": nuo.get("name") or "",
            "kam": [a for a in gavejai if a],
            "subject": (m.get("subject") or "")[:200],
            "tekstas": tekstas[:8000],
            "webLink": m.get("webLink"),
        })
    out.sort(key=lambda x: x["kada"] or datetime.datetime.min)
    return out


def chain_messages(kodas: str, max_count: int = 20):
    """Grandinė pagal kodą temoje (užklausos tiekėjams) — su tekstais."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token,
               "Prefer": 'outlook.body-content-type="text"'}
    r = requests.get(
        f"{GRAPH}/me/messages?$search=\"{kodas}\"&$select=id&$top={max_count}",
        headers=headers, timeout=60)
    r.raise_for_status()
    return _full_messages([m["id"] for m in r.json().get("value", [])], headers)


def conversation_chain(conversation_id: str, max_count: int = 30):
    """Viso pokalbio (kliento susirašinėjimo) grandinė — su tekstais."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token,
               "Prefer": 'outlook.body-content-type="text"'}
    r = requests.get(
        f"{GRAPH}/me/messages?$filter=conversationId eq '{conversation_id}'"
        f"&$select=id&$top={max_count}",
        headers=headers, timeout=60)
    r.raise_for_status()
    return _full_messages([m["id"] for m in r.json().get("value", [])], headers)


def open_conversation(conversation_id: str, fallback_entry_id: str = None):
    """Atidaro VISĄ susirašinėjimo grandinę: naujausią pokalbio laišką
    Outlook web'e (ten gija rodoma pilna). Nepavykus — patį laišką."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    try:
        r = requests.get(
            f"{GRAPH}/me/messages"
            f"?$filter=conversationId eq '{conversation_id}'"
            f"&$select=webLink,receivedDateTime&$top=50",
            headers=headers, timeout=30)
        r.raise_for_status()
        msgs = r.json().get("value", [])
        if msgs:
            naujausias = max(msgs, key=lambda m: m.get("receivedDateTime") or "")
            link = naujausias.get("webLink")
            if link:
                os.startfile(link)
                return
    except Exception:
        pass
    if fallback_entry_id:
        open_email(fallback_entry_id)


def open_email(entry_id: str, store_id: str = "graph"):
    """Atidaro laiska narsykleje (Outlook web) per jo webLink."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    r = requests.get(f"{GRAPH}/me/messages/{entry_id}?$select=webLink",
                     headers=headers, timeout=30)
    r.raise_for_status()
    link = r.json().get("webLink")
    if link:
        os.startfile(link)


def get_attachments(entry_id: str, store_id: str = "graph"):
    """Laisko priedai [(vardas, baitai, inline)] — kaip outlook_client.get_attachments."""
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    return _fetch_attachments(headers, entry_id)


def load_email(entry_id: str, store_id: str = "graph"):
    """Pilnas vieno laiško paruošimas analizei (tekstas + priedai).

    Kūnas HTML — spalvų žymėjimą išsaugo _html_i_teksta (žr. fetch_emails_since).
    """
    token = get_token()
    headers = {"Authorization": "Bearer " + token}
    r = requests.get(f"{GRAPH}/me/messages/{entry_id}"
                     "?$select=subject,from,receivedDateTime,conversationId,body,hasAttachments",
                     headers=headers, timeout=60)
    r.raise_for_status()
    m = r.json()
    atts = (_fetch_attachments(headers, entry_id, None, m.get("subject") or "")
            if m.get("hasAttachments") else [])
    return _msg_to_email(m, atts)
