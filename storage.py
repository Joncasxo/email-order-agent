# -*- coding: utf-8 -*-
"""Vietinė SQLite duomenų bazė: apdoroti laiškai ir atpažinti užsakymai.

Užsakymo kodas sudaromas pagal vartotojo Excel tvarką:
KODO_PREFIKSAS + kliento inicialai + data YYMMDD, pvz. VP260702
(Vardenis Pavardenis, 2026-07-02).
Pasikartojantys tą pačią dieną gauna priedėlį (2), (3)...
"""
import datetime
import json
import os
import re
import sqlite3
import threading

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "uzsakymai.db")

_lock = threading.Lock()

_ORDER_COLUMNS = {
    "kodas": "TEXT",
    "ispejimas": "TEXT",
    "kliento_adresas": "TEXT",
    "kliento_telefonas": "TEXT",
    "pasiulymo_terminas": "TEXT",
    "mano_pastabos": "TEXT",
    "conversation_id": "TEXT",
    "priedai": "TEXT",
    "sutartis": "TEXT",
    "saskaita": "INTEGER DEFAULT 0",   # žymė: sąskaita išrašyta
    "atnaujinimai": "TEXT",            # grandinės papildymų žurnalas
    "atnaujinta": "INTEGER DEFAULT 0", # 🔔 — yra neperžiūrėtas atnaujinimas
    # Graph migracija (2026-08-27): senasis COM ConversationID išsaugomas čia,
    # o conversation_id perrašomas Graph formatu — kad grandinės saugiklis
    # atpažintų atsakymus, atėjusius jau per Graph API. NIEKAS netrinama.
    "conversation_id_com": "TEXT",
    # Tiekėjų užklausų sekimas (2026-09-01): kam išsiųsta užklausa su kodu temoje
    # ir ar kas nors atsakė (❗ sąraše, nusiimamas peržiūrėjus)
    "uzklausti_tiekejai": "TEXT",
    "tiekejas_atsake": "INTEGER DEFAULT 0",
}

_DAY_COLUMNS = {
    "pastabos": "TEXT",                # darbo eiga / užrašai
    "perkelimu": "INTEGER DEFAULT 0",  # kiek kartų perkeltas į kitą dieną
    "pradine_data": "TEXT",            # kada darbas užrašytas pirmą kartą
}

# --- Įmonei specifiniai nustatymai (pritaikykite savo įmonei) ---
# Užsakymo kodo priešdėlis (pvz., pagrindinio kliento inicialai).
KODO_PREFIKSAS = ""
# Prioritetinės sutarties numeris — šios sutarties užsakymų kodai prasideda „!".
# Tuščia reikšmė = prioritetinių sutarčių nėra.
PRIORITY_CONTRACT = ""


def _is_priority(analysis: dict) -> bool:
    if not PRIORITY_CONTRACT:
        return False
    sutartis = (analysis.get("sutartis") or "").upper().replace(" ", "")
    return sutartis == PRIORITY_CONTRACT.upper().replace(" ", "")


def _conn():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn


def _initials(name: str) -> str:
    words = [w for w in re.split(r"[\s,._-]+", (name or "").strip()) if w and w[0].isalpha()]
    ini = "".join(w[0] for w in words[:3]).upper()
    return ini or "X"


def _make_kodas(c, klientas: str, received: datetime.datetime, priority: bool = False) -> str:
    prefix = "!" if priority else ""
    base = f"{prefix}{KODO_PREFIKSAS}{_initials(klientas)}{received:%y%m%d}"
    kodas, n = base, 2
    while c.execute("SELECT 1 FROM orders WHERE kodas=?", (kodas,)).fetchone():
        kodas = f"{base}({n})"
        n += 1
    return kodas


def init_db():
    with _lock, _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS processed (
            entry_id TEXT PRIMARY KEY,
            subject TEXT,
            received_at TEXT,
            is_order INTEGER,
            processed_at TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id TEXT,
            store_id TEXT,
            kodas TEXT,
            received_at TEXT,
            sender TEXT,
            subject TEXT,
            klientas TEXT,
            kliento_el_pastas TEXT,
            kliento_adresas TEXT,
            kliento_telefonas TEXT,
            uzsakymo_data TEXT,
            pristatymo_data TEXT,
            pasiulymo_terminas TEXT,
            prekes_json TEXT,
            pastabos TEXT,
            mano_pastabos TEXT,
            santrauka TEXT,
            pasitikejimas REAL,
            statusas TEXT DEFAULT 'naujas',
            created_at TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS meta (
            k TEXT PRIMARY KEY, v TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS sutartys (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numeris TEXT UNIQUE,   -- sunormalizuotas, pvz. SUT25-114
            padalinys TEXT,
            adresas TEXT,
            asmenys TEXT,          -- kontaktai iš Word sąrašo
            pastaba TEXT,          -- pvz. NEBEGALIOJA
            eile INTEGER DEFAULT 0
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS dienos_darbai (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,            -- YYYY-MM-DD
            tekstas TEXT,
            kodas TEXT,           -- susijęs užsakymas (jei yra)
            auto INTEGER DEFAULT 0,
            atlikta INTEGER DEFAULT 0,
            created_at TEXT
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_darbai_data ON dienos_darbai(data)")
        # Migracija senoms duomenų bazėms: pridedame trūkstamus stulpelius
        existing = {r["name"] for r in c.execute("PRAGMA table_info(orders)")}
        for col, coltype in _ORDER_COLUMNS.items():
            if col not in existing:
                c.execute(f"ALTER TABLE orders ADD COLUMN {col} {coltype}")
        existing_day = {r["name"] for r in c.execute("PRAGMA table_info(dienos_darbai)")}
        for col, coltype in _DAY_COLUMNS.items():
            if col not in existing_day:
                c.execute(f"ALTER TABLE dienos_darbai ADD COLUMN {col} {coltype}")
        # Senų statusų perėjimas į naują schemą (užklausta/pasiūlyta/laimėta/užsakyta)
        c.execute("UPDATE orders SET statusas='' WHERE statusas='naujas'")
        c.execute("UPDATE orders SET statusas='užklausta' WHERE statusas='vykdomas'")
        c.execute("UPDATE orders SET statusas='užsakyta' WHERE statusas='įvykdytas'")
        # Užpildome kodus seniems įrašams
        for row in c.execute(
                "SELECT id, klientas, sender, received_at FROM orders "
                "WHERE kodas IS NULL OR kodas=''").fetchall():
            try:
                received = datetime.datetime.strptime(row["received_at"], "%Y-%m-%d %H:%M")
            except Exception:
                received = datetime.datetime.now()
            kodas = _make_kodas(c, row["klientas"] or row["sender"] or "", received)
            c.execute("UPDATE orders SET kodas=? WHERE id=?", (kodas, row["id"]))


def get_meta(key, default=None):
    with _lock, _conn() as c:
        row = c.execute("SELECT v FROM meta WHERE k=?", (key,)).fetchone()
        return row["v"] if row else default


def set_meta(key, value):
    with _lock, _conn() as c:
        c.execute("INSERT INTO meta(k, v) VALUES(?, ?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
                  (key, str(value)))


def is_processed(entry_id: str) -> bool:
    with _lock, _conn() as c:
        return c.execute("SELECT 1 FROM processed WHERE entry_id=?", (entry_id,)).fetchone() is not None


def has_order_entry(entry_id: str) -> bool:
    """Ar iš šio laiško jau sukurtas užsakymas (pakartotinei patikrai:
    tokie laiškai praleidžiami, kad neatsirastų dublikatų)."""
    with _lock, _conn() as c:
        return c.execute("SELECT 1 FROM orders WHERE entry_id=?", (entry_id,)).fetchone() is not None


def _order_datetime(analysis: dict, email: dict) -> datetime.datetime:
    """Užsakymo data: AI ištraukta pradinio laiško data (grandinėse ji ankstesnė
    už atsakymo gavimo laiką); jei jos nėra ar ji nepatikima — laiško gavimo laikas."""
    raw = (analysis.get("uzsakymo_data") or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?", raw)
    if m:
        try:
            dt = datetime.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                                   int(m.group(4) or 12), int(m.group(5) or 0))
            now = datetime.datetime.now()
            # saugikliai: ne ateitis, ne vėlesnė už gavimą, ne senesnė nei metai
            if (now - datetime.timedelta(days=365) < dt
                    <= email["received"] + datetime.timedelta(days=1)):
                return dt
        except ValueError:
            pass
    return email["received"]


def record_processed(email: dict, analysis: dict, is_order: bool, priedai: str = ""):
    """Vienoje transakcijoje įrašo ir užsakymą (jei yra), ir apdorojimo žymą —
    kad nutrūkus darbui tarp jų neliktų „apdorotas, bet užsakymas pamestas"."""
    with _lock, _conn() as c:
        kodas = None
        if is_order:
            order_dt = _order_datetime(analysis, email)
            kodas = _make_kodas(c, analysis.get("klientas") or email["sender_name"],
                                order_dt, priority=_is_priority(analysis))
            c.execute("""INSERT INTO orders
                (entry_id, store_id, kodas, conversation_id, priedai, sutartis,
                 received_at, sender, subject,
                 klientas, kliento_el_pastas, kliento_adresas, kliento_telefonas,
                 uzsakymo_data, pristatymo_data, pasiulymo_terminas,
                 prekes_json, pastabos, mano_pastabos, santrauka, ispejimas,
                 pasitikejimas, statusas, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, '', ?)""",
                (email["entry_id"], email["store_id"], kodas,
                 email.get("conversation_id") or None, priedai or None,
                 analysis.get("sutartis"),
                 f"{order_dt:%Y-%m-%d %H:%M}",
                 f"{email['sender_name']} <{email['sender_email']}>", email["subject"],
                 analysis.get("klientas"), analysis.get("kliento_el_pastas"),
                 analysis.get("kliento_adresas"), analysis.get("kliento_telefonas"),
                 analysis.get("uzsakymo_data"), analysis.get("pristatymo_data"),
                 analysis.get("pasiulymo_terminas"),
                 json.dumps(analysis.get("prekes", []), ensure_ascii=False),
                 analysis.get("pastabos"), analysis.get("santrauka"),
                 analysis.get("ispejimas"),
                 analysis.get("pasitikejimas"),
                 datetime.datetime.now().isoformat(timespec="seconds")))
        c.execute("INSERT OR REPLACE INTO processed VALUES (?, ?, ?, ?, ?)",
                  (email["entry_id"], email["subject"], str(email["received"]),
                   int(is_order), datetime.datetime.now().isoformat(timespec="seconds")))
        return kodas


def add_manual_order(data: dict) -> str:
    """Rankiniu būdu sukurtas užsakymas — be laiško. Kodas generuojamas tuo
    pačiu mechanizmu kaip AI atpažintiems (KODO_PREFIKSAS + inicialai + data, „!" jei
    prioritetinė sutartis). data["prekes"] — sąrašas {preke, kiekis, pastabos}."""
    import uuid
    with _lock, _conn() as c:
        try:
            d = datetime.datetime.strptime(
                (data.get("uzsakymo_data") or "").strip(), "%Y-%m-%d")
            if d.date() == datetime.date.today():
                order_dt = datetime.datetime.now()
            else:
                order_dt = d.replace(hour=12)
        except ValueError:
            order_dt = datetime.datetime.now()
        kodas = _make_kodas(c, data.get("klientas") or "?", order_dt,
                            priority=_is_priority(data))
        c.execute("""INSERT INTO orders
            (entry_id, store_id, kodas, conversation_id, priedai, sutartis,
             received_at, sender, subject,
             klientas, kliento_el_pastas, kliento_adresas, kliento_telefonas,
             uzsakymo_data, pristatymo_data, pasiulymo_terminas,
             prekes_json, pastabos, mano_pastabos, santrauka,
             pasitikejimas, statusas, created_at)
            VALUES (?, ?, ?, NULL, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, '', ?)""",
            (f"manual-{uuid.uuid4()}", "manual", kodas,
             (data.get("sutartis") or "").strip() or None,
             f"{order_dt:%Y-%m-%d %H:%M}",
             "rankinis įrašas", "(rankinis užsakymas)",
             (data.get("klientas") or "").strip(),
             (data.get("kliento_el_pastas") or "").strip() or None,
             (data.get("kliento_adresas") or "").strip() or None,
             (data.get("kliento_telefonas") or "").strip() or None,
             f"{order_dt:%Y-%m-%d}",
             (data.get("pristatymo_data") or "").strip() or None,
             (data.get("pasiulymo_terminas") or "").strip() or None,
             json.dumps(data.get("prekes") or [], ensure_ascii=False),
             None,
             (data.get("mano_pastabos") or "").strip(),
             "Rankiniu būdu sukurtas užsakymas.",
             datetime.datetime.now().isoformat(timespec="seconds")))
        return kodas


def mark_processed_simple(email: dict, note: str = "grandinės dalis"):
    """Pažymi laišką apdorotu be analizės (pvz., senesni tos pačios grandinės laiškai)."""
    with _lock, _conn() as c:
        c.execute("INSERT OR REPLACE INTO processed VALUES (?, ?, ?, 0, ?)",
                  (email["entry_id"], f"[{note}] {email['subject']}",
                   str(email["received"]),
                   datetime.datetime.now().isoformat(timespec="seconds")))


def unmark_processed(entry_id: str):
    """Nuima „apdorota" žymę — laišką bus galima analizuoti iš naujo
    (rankinis pasirinkimas; įprastas tikrinimas jo pats neims, nes žymė
    grąžinama iškart po pakartotinės analizės)."""
    with _lock, _conn() as c:
        c.execute("DELETE FROM processed WHERE entry_id=?", (entry_id,))


def order_by_conversation(conversation_id: str):
    """Deterministinis dublikatų saugiklis: ar ši Outlook susirašinėjimo grandinė
    JAU turi užregistruotą užsakymą? Veikia nepriklausomai nuo praėjusio laiko
    (14/45 d. langas čia nesvarbus) ir nekainuoja nė vieno AI tokeno."""
    if not conversation_id:
        return None
    with _lock, _conn() as c:
        row = c.execute(
            "SELECT kodas, klientas, received_at, prekes_json FROM orders "
            "WHERE conversation_id=? ORDER BY id LIMIT 1", (conversation_id,)).fetchone()
        if row is None:
            return None
        o = dict(row)
    try:
        items = json.loads(o.get("prekes_json") or "[]")
    except Exception:
        items = []
    o["prekes_santrauka"] = "; ".join(
        f"{it.get('preke')} ({it.get('kiekis')})" for it in items[:6])
    return o


def recent_orders_summary(days: int = 45, limit: int = 15) -> str:
    """Kompaktiška neseniai užregistruotų užsakymų santrauka AI kontekstui
    (kad grandinės laiškai nebūtų registruojami kaip nauji užsakymai)."""
    cutoff = (datetime.datetime.now() - datetime.timedelta(days=days)).strftime("%Y-%m-%d")
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT kodas, klientas, received_at, prekes_json FROM orders "
            "WHERE received_at >= ? ORDER BY received_at DESC LIMIT ?",
            (cutoff, limit)).fetchall()
    lines = []
    for r in rows:
        try:
            items = json.loads(r["prekes_json"] or "[]")
        except Exception:
            items = []
        prekes = "; ".join(f"{it.get('preke')} ({it.get('kiekis')})" for it in items[:4])
        if len(items) > 4:
            prekes += f" … (+{len(items) - 4})"
        lines.append(f"{r['kodas']} | {r['klientas'] or '?'} | {r['received_at']} | {prekes}")
    return "\n".join(lines)


def list_orders(status_filter=None, date_from: str = None, date_to: str = None,
                search: str = None):
    """Užsakymų sąrašas su filtrais. Datos formatu YYYY-MM-DD (imtinai)."""
    sql = "SELECT * FROM orders WHERE 1=1"
    params = []
    if status_filter == "nauji":
        sql += " AND IFNULL(statusas,'')=''"
    elif status_filter == "uzbaigti":
        # užbaigtas = (užsakyta IR sąskaita PILNAI išrašyta) ARBA nelaimėta;
        # dalinė sąskaita (2) nesiskaito
        sql += (" AND ((statusas='užsakyta' AND IFNULL(saskaita,0)=1)"
                " OR statusas='nelaimėta')")
    elif status_filter == "neuzbaigti":
        sql += (" AND NOT ((statusas='užsakyta' AND IFNULL(saskaita,0)=1)"
                " OR statusas='nelaimėta')")
    elif status_filter:
        sql += " AND statusas=?"
        params.append(status_filter)
    if date_from:
        sql += " AND substr(received_at, 1, 10) >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND substr(received_at, 1, 10) <= ?"
        params.append(date_to)
    if search:
        like = f"%{search.lower()}%"
        sql += (" AND (lower(IFNULL(kodas,'')) LIKE ? OR lower(IFNULL(klientas,'')) LIKE ?"
                " OR lower(IFNULL(subject,'')) LIKE ? OR lower(IFNULL(prekes_json,'')) LIKE ?"
                " OR lower(IFNULL(mano_pastabos,'')) LIKE ?)")
        params.extend([like] * 5)
    sql += " ORDER BY received_at DESC"
    with _lock, _conn() as c:
        return [dict(r) for r in c.execute(sql, params).fetchall()]


def orders_date_range():
    """Anksčiausio ir vėliausio užsakymo datos (YYYY-MM-DD)."""
    with _lock, _conn() as c:
        row = c.execute("SELECT MIN(substr(received_at,1,10)) a, "
                        "MAX(substr(received_at,1,10)) b FROM orders").fetchone()
        return (row["a"], row["b"]) if row else (None, None)


def get_order(order_id: int):
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        return dict(row) if row else None


def set_order_status(order_id: int, status: str):
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET statusas=? WHERE id=?", (status, order_id))


def set_order_priedai(order_id: int, priedai: str):
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET priedai=? WHERE id=?", (priedai or None, order_id))


# ---------------- Dienos darbai ----------------

def _add_task(c, data: str, tekstas: str, kodas=None, auto=0):
    c.execute("""INSERT INTO dienos_darbai (data, tekstas, kodas, auto, atlikta, created_at)
                 VALUES (?, ?, ?, ?, 0, ?)""",
              (data, tekstas, kodas, auto,
               datetime.datetime.now().isoformat(timespec="seconds")))
    return True


def carry_over_tasks(today: str = None) -> int:
    """Neatliktus ankstesnių dienų darbus perkelia į šiandieną kartu su užrašais.
    Grąžina, kiek darbų perkelta."""
    today = today or datetime.date.today().isoformat()
    with _lock, _conn() as c:
        rows = c.execute("SELECT id, data, pradine_data, perkelimu FROM dienos_darbai "
                         "WHERE atlikta=0 AND data < ?", (today,)).fetchall()
        for r in rows:
            c.execute("UPDATE dienos_darbai SET data=?, perkelimu=?, pradine_data=? WHERE id=?",
                      (today, (r["perkelimu"] or 0) + 1,
                       r["pradine_data"] or r["data"], r["id"]))
        return len(rows)


def set_task_notes(task_id: int, notes: str):
    with _lock, _conn() as c:
        c.execute("UPDATE dienos_darbai SET pastabos=? WHERE id=?", (notes, task_id))


def get_task(task_id: int):
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM dienos_darbai WHERE id=?", (task_id,)).fetchone()
        return dict(row) if row else None


def all_day_tasks():
    """Visi darbai (eksportui)."""
    with _lock, _conn() as c:
        rows = c.execute("SELECT * FROM dienos_darbai ORDER BY data DESC, id").fetchall()
        return [dict(r) for r in rows]


def clear_auto_day_tasks():
    """Pašalina anksčiau automatiškai keltus užsakymų įrašus — dienos darbai
    yra atskiras, su užsakymais nesusijęs sąrašas."""
    with _lock, _conn() as c:
        c.execute("DELETE FROM dienos_darbai WHERE auto=1")


def list_day_tasks(data: str = None):
    data = data or datetime.date.today().isoformat()
    with _lock, _conn() as c:
        rows = c.execute("SELECT * FROM dienos_darbai WHERE data=? "
                         "ORDER BY atlikta, auto DESC, id", (data,)).fetchall()
        return [dict(r) for r in rows]


def list_day_tasks_history(data: str):
    """Praeitos dienos vaizdas: tos dienos darbai (ir atlikti) + darbai, kurie
    tą dieną dar tęsėsi, bet vėliau buvo perkelti (perkeltas_veliau=1)."""
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT *, (data > ?) AS perkeltas_veliau FROM dienos_darbai "
            "WHERE data=? OR (pradine_data IS NOT NULL AND pradine_data <= ? AND data > ?) "
            "ORDER BY perkeltas_veliau, atlikta, auto DESC, id",
            (data, data, data, data)).fetchall()
        return [dict(r) for r in rows]


def list_task_dates(limit: int = 60):
    with _lock, _conn() as c:
        rows = c.execute("SELECT DISTINCT data FROM dienos_darbai "
                         "ORDER BY data DESC LIMIT ?", (limit,)).fetchall()
        return [r["data"] for r in rows]


def add_day_task(tekstas: str, data: str = None):
    data = data or datetime.date.today().isoformat()
    with _lock, _conn() as c:
        _add_task(c, data, tekstas.strip()[:300], None, auto=0)


def set_task_done(task_id: int, done: bool):
    with _lock, _conn() as c:
        c.execute("UPDATE dienos_darbai SET atlikta=? WHERE id=?", (1 if done else 0, task_id))


def delete_day_task(task_id: int):
    with _lock, _conn() as c:
        c.execute("DELETE FROM dienos_darbai WHERE id=?", (task_id,))


def apply_order_update(kodas: str, analysis: dict, email: dict) -> bool:
    """Grandinės papildymas: prie ESAMO užsakymo prideda naujas prekes ir
    įrašą į atnaujinimų žurnalą, uždega 🔔. Data ir kodas nekeičiami."""
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM orders WHERE kodas=?", (kodas,)).fetchone()
        if row is None:  # tolerancija: be šauktuko / kitoks raidžių dydis
            row = c.execute(
                "SELECT * FROM orders WHERE upper(replace(IFNULL(kodas,''),'!',''))=?",
                (kodas.upper().replace("!", ""),)).fetchone()
        if row is None:
            return False
        o = dict(row)
        entry = f"{email['received']:%m-%d %H:%M} — {analysis.get('santrauka') or 'papildymas'}"

        try:
            items = json.loads(o.get("prekes_json") or "[]")
        except Exception:
            items = []
        have = {(it.get("preke") or "").strip().lower() for it in items}
        added = 0
        for it in analysis.get("prekes") or []:
            if (it.get("preke") or "").strip().lower() not in have:
                items.append(it)
                added += 1
        if added:
            entry += f" (pridėta prekių: {added})"

        log = (o.get("atnaujinimai") or "").rstrip()
        log = (log + "\n" if log else "") + entry

        sets = {"prekes_json": json.dumps(items, ensure_ascii=False),
                "atnaujinimai": log, "atnaujinta": 1}
        for field in ("pristatymo_data", "pasiulymo_terminas",
                      "kliento_telefonas", "kliento_adresas"):
            if analysis.get(field):
                sets[field] = analysis[field]

        # AI įspėjimas apie neaiškų/pavojingą prašymą — rodomas raudonai
        if analysis.get("ispejimas"):
            sets["ispejimas"] = analysis["ispejimas"]
            sets["atnaujinimai"] += "\n   ⚠ " + analysis["ispejimas"]

        # Jei užsakymas jau buvo UŽBAIGTAS, grandinės papildymas grąžina jį
        # į neužbaigtus — atsirado naujo darbo:
        #  - užsakyta + sąskaita pilnai -> sąskaita tampa DALINE (nauja dalis
        #    dar neapmokestinta);
        #  - nelaimėta -> statusas nuimamas (grandinė atgijo, reikia peržiūrėti).
        if (o.get("statusas") == "užsakyta" and int(o.get("saskaita") or 0) == 1):
            sets["saskaita"] = 2
            sets["atnaujinimai"] += ("\n   ↩ buvo užbaigtas — sąskaita pažymėta "
                                     "daline, užsakymas grįžo į neužbaigtus")
        elif o.get("statusas") == "nelaimėta":
            sets["statusas"] = ""
            sets["atnaujinimai"] += ("\n   ↩ buvo nelaimėtas — grandinė atgijo, "
                                     "statusas nuimtas (grįžo į neužbaigtus)")

        c.execute("UPDATE orders SET " + ", ".join(f"{k}=?" for k in sets) + " WHERE id=?",
                  (*sets.values(), o["id"]))
        return True


def set_order_seen(order_id: int):
    """Nuima 🔔 (vartotojas peržiūrėjo atnaujinimą)."""
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET atnaujinta=0 WHERE id=?", (order_id,))


# ---------------- tiekėjų užklausų sekimas ----------------

def orders_inquiry_tracking(days: int = 60):
    """Užsakymai, kurių kodų ieškome siųstų/gautų laiškų temose."""
    riba = (datetime.date.today() - datetime.timedelta(days=days)).isoformat()
    with _lock, _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT id, kodas, statusas, tiekejas_atsake, uzklausti_tiekejai "
            "FROM orders WHERE substr(received_at,1,10) >= ?", (riba,))]


def inquiry_sent(order_id: int, recipients: list, set_uzklausta: bool = True):
    """Išsiųsta užklausa su kodu temoje: prirašo gavėjus prie užsakymo (be dublių)
    ir, jei statusas dar tuščias, pažymi „užklausta"."""
    with _lock, _conn() as c:
        r = c.execute("SELECT uzklausti_tiekejai, statusas FROM orders WHERE id=?",
                      (order_id,)).fetchone()
        if r is None:
            return
        esami = [t.strip() for t in (r["uzklausti_tiekejai"] or "").split(";") if t.strip()]
        zemi = {t.split(" ")[0].lower() for t in esami}
        for a in recipients:
            a = (a or "").strip()
            if a and a.lower() not in zemi:
                esami.append(a)
                zemi.add(a.lower())
        sets = {"uzklausti_tiekejai": "; ".join(esami)}
        if set_uzklausta and not (r["statusas"] or ""):
            sets["statusas"] = "užklausta"
        c.execute(f"UPDATE orders SET {', '.join(k + '=?' for k in sets)} WHERE id=?",
                  (*sets.values(), order_id))


def supplier_replied(order_id: int, sender_email: str):
    """Gautas laiškas su užsakymo kodu temoje — uždeda ❗ ir pažymi atsakiusį tiekėją."""
    with _lock, _conn() as c:
        r = c.execute("SELECT uzklausti_tiekejai FROM orders WHERE id=?",
                      (order_id,)).fetchone()
        if r is None:
            return
        s = (sender_email or "").strip()
        esami = [t.strip() for t in (r["uzklausti_tiekejai"] or "").split(";") if t.strip()]
        rasta = False
        for i, t in enumerate(esami):
            if s and t.split(" ")[0].lower() == s.lower():
                if "✓" not in t:
                    esami[i] = t + " ✓"
                rasta = True
        if s and not rasta:
            esami.append(f"{s} ✓ (atsakė)")
        c.execute("UPDATE orders SET tiekejas_atsake=1, uzklausti_tiekejai=? WHERE id=?",
                  ("; ".join(esami), order_id))


def clear_supplier_replied(order_id: int):
    """Nuima ❗ (vartotojas peržiūrėjo tiekėjo atsakymą)."""
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET tiekejas_atsake=0 WHERE id=?", (order_id,))


def set_order_invoice(order_id: int, value: int):
    """Sąskaitos žymė: 0 — nėra, 1 — išrašyta, 2 — dalinai išrašyta."""
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET saskaita=? WHERE id=?", (int(value), order_id))


def set_order_notes(order_id: int, notes: str):
    with _lock, _conn() as c:
        c.execute("UPDATE orders SET mano_pastabos=? WHERE id=?", (notes, order_id))


def delete_order(order_id: int):
    with _lock, _conn() as c:
        c.execute("DELETE FROM orders WHERE id=?", (order_id,))


# ---------------- Sutartys ----------------

def save_sutartys(rows: list) -> int:
    """Įkelia Word sąrašą: naujus numerius prideda, esamų padalinį/adresą atnaujina.
    Ranka pildytų žmonių ir statuso NEPERRAŠO (Word reikšmė imama tik jei DB tuščia)."""
    with _lock, _conn() as c:
        for i, r in enumerate(rows):
            row = c.execute("SELECT * FROM sutartys WHERE numeris=?",
                            (r["numeris"],)).fetchone()
            if row is None:
                c.execute("INSERT INTO sutartys "
                          "(numeris, padalinys, adresas, asmenys, pastaba, eile) "
                          "VALUES (?,?,?,?,?,?)",
                          (r["numeris"], r.get("padalinys") or "", r.get("adresas") or "",
                           r.get("asmenys") or "", r.get("pastaba") or "", i))
            else:
                c.execute("UPDATE sutartys SET padalinys=?, adresas=?, asmenys=?, "
                          "pastaba=?, eile=? WHERE id=?",
                          (r.get("padalinys") or row["padalinys"],
                           r.get("adresas") or row["adresas"],
                           row["asmenys"] or r.get("asmenys") or "",
                           row["pastaba"] or r.get("pastaba") or "",
                           i, row["id"]))
        return len(rows)


def save_sutartis(data: dict) -> int:
    """Prideda arba atnaujina VIENĄ sutartį (rankinis redagavimas lange).
    Jei data turi id — atnaujina; kitaip ieško pagal numerį arba prideda gale."""
    with _lock, _conn() as c:
        fields = (data.get("numeris") or "", data.get("padalinys") or "",
                  data.get("adresas") or "", data.get("asmenys") or "",
                  data.get("pastaba") or "")
        sid = data.get("id")
        if not sid:
            row = c.execute("SELECT id FROM sutartys WHERE numeris=?",
                            (fields[0],)).fetchone()
            sid = row["id"] if row else None
        if sid:
            c.execute("UPDATE sutartys SET numeris=?, padalinys=?, adresas=?, "
                      "asmenys=?, pastaba=? WHERE id=?", (*fields, sid))
            return sid
        eile = c.execute("SELECT IFNULL(MAX(eile), -1) + 1 FROM sutartys").fetchone()[0]
        cur = c.execute("INSERT INTO sutartys "
                        "(numeris, padalinys, adresas, asmenys, pastaba, eile) "
                        "VALUES (?,?,?,?,?,?)", (*fields, eile))
        return cur.lastrowid


def delete_sutartis(sid: int):
    with _lock, _conn() as c:
        c.execute("DELETE FROM sutartys WHERE id=?", (sid,))


def sutartys_su_uzsakymais():
    """Sutarčių sąrašas su prie kiekvienos priskirtais užsakymais (žmonėmis).

    Priskyrimas — vien kodu, be AI: 1) užsakymo „sutartis" laukas (AI jį ištraukia
    ir taip, kiekvienos analizės metu), 2) sutarties numeris laiško temoje,
    3) kliento vardas/pavardė jau įrašyta sutarties kontaktuose,
    4) užsakymo pristatymo ADRESAS sutampa su sutarties padaliniu/adresu
       (bent 2 žodžių kamienai, kad vien miesto nepakaktų klaidingam priskyrimui).
    Grąžina (sutartys, be_sutarties): sutartys turi raktą „uzsakymai",
    sąrašo gale — pseudo įrašai numeriams, kurių nėra Word sąraše.
    """
    from sutartys_docx import find_numbers
    with _lock, _conn() as c:
        suts = [dict(r) for r in c.execute(
            "SELECT * FROM sutartys ORDER BY eile, id").fetchall()]
        orders = [dict(r) for r in c.execute(
            "SELECT id, kodas, klientas, received_at, statusas, saskaita, "
            "sutartis, subject, kliento_adresas "
            "FROM orders ORDER BY received_at DESC").fetchall()]

    by_num = {}
    for s in suts:
        s["uzsakymai"] = []
        by_num[s["numeris"]] = s

    # kontaktų žodžiai -> sutartis (kliento pavardei atpažinti)
    word_map = {}
    for s in suts:
        for w in re.split(r"[\s,;]+", (s.get("asmenys") or "").lower()):
            if len(w) > 2:
                word_map.setdefault(w, set()).add(s["numeris"])

    # adresų kamienai (4 taisyklei): „Liepų g. 15, Vilnius" -> {liepu, vilni}
    _lt = str.maketrans("ąčęėįšųūž", "aceeisuuz")

    def _kamienai(tekstas):
        t = (tekstas or "").lower().translate(_lt)
        return {w[:5] for w in re.split(r"[^a-z0-9]+", t) if len(w) >= 4}

    for s in suts:
        s["_kamienai"] = _kamienai(f"{s.get('padalinys') or ''} {s.get('adresas') or ''}")

    be_sutarties = []
    for o in orders:
        nums = find_numbers(o.get("sutartis") or "") or find_numbers(o.get("subject") or "")
        num = next((n for n in nums if n in by_num), nums[0] if nums else None)
        if not num:
            words = [w for w in re.split(r"[\s,;]+", (o.get("klientas") or "").lower())
                     if len(w) > 2]
            # būtinai turi sutapti PAVARDĖ (paskutinis žodis) — vien vardo nepakanka,
            # nes tas pats vardas gali būti keliose sutartyse pas skirtingus žmones
            cands = word_map.get(words[-1]) if words else None
            if cands:
                for w in words[:-1]:
                    f = word_map.get(w)
                    if f and (cands & f):
                        cands = cands & f
            if cands and len(cands) == 1:
                num = next(iter(cands))
        if not num:
            # 4) pagal pristatymo adresą: reikia bent 2 sutampančių žodžių kamienų
            # (gatvė + miestas), vien miesto nepakanka; laimi VIENINTELIS geriausias
            okam = _kamienai(o.get("kliento_adresas"))
            if okam:
                geriausi, max_bendru = [], 1
                for s in suts:
                    bendri = len(okam & (s.get("_kamienai") or set()))
                    if bendri > max_bendru:
                        geriausi, max_bendru = [s["numeris"]], bendri
                    elif bendri == max_bendru and bendri > 1:
                        geriausi.append(s["numeris"])
                if len(geriausi) == 1:
                    num = geriausi[0]
        if num and num in by_num:
            by_num[num]["uzsakymai"].append(o)
        elif num:  # numeris yra, bet jo nėra Word sąraše (pvz. kitų serijų sutartys)
            extra = by_num.get(num)
            if extra is None:
                extra = {"numeris": num, "padalinys": "(nėra sąraše)", "adresas": "",
                         "asmenys": "", "pastaba": "", "uzsakymai": []}
                by_num[num] = extra
                suts.append(extra)
            extra["uzsakymai"].append(o)
        else:
            be_sutarties.append(o)
    return suts, be_sutarties
