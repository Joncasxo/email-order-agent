# -*- coding: utf-8 -*-
"""Sukuria/atnaujina vCard failą iš užsakymų DB klientų (vardas + telefonas + el. paštas).

Naudojama dviem būdais:
  * programa (app.py) kviečia eksportuoti() po kiekvieno patikrinimo, kuriame
    atsirado naujų užsakymų ar papildymų — darbalaukio failas visada šviežias;
  * galima paleisti ir ranka: python eksportuoti_kontaktus.py
"""
import os
import re
import sqlite3

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uzsakymai.db")
ISVESTIS = os.path.join(os.path.expanduser("~"), "Desktop", "klientai.vcf")

# Domenas -> organizacijos pavadinimas vCard kortelėje (pritaikykite savo įmonei), pvz.:
#     "example.lt": "UAB Pavyzdys",
ORG_PAGAL_DOMENA = {}


def normalizuoti_numeri(n):
    s = re.sub(r"[^\d+]", "", str(n))
    if not s:
        return None
    if s.startswith("+"):
        return s if len(s) >= 11 else None
    if s.startswith("370") and len(s) == 11:
        return "+" + s
    if s.startswith("8") and len(s) == 9:
        return "+370" + s[1:]
    if s.startswith("6") and len(s) == 8:
        return "+370" + s
    return s if len(s) >= 8 else None


def eksportuoti(db=DB, isvestis=ISVESTIS):
    """Perrašo vCard failą iš DB. Grąžina (kontaktų, iš jų su telefonu)."""
    conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    klientai = {}
    try:
        for r in conn.execute("""SELECT klientas, kliento_telefonas, kliento_el_pastas
                                 FROM orders
                                 WHERE IFNULL(klientas,'') <> ''
                                 ORDER BY received_at"""):
            vardas = " ".join(str(r["klientas"]).split()).strip()
            if not vardas or len(vardas) < 4:
                continue
            irasas = klientai.setdefault(vardas.lower(),
                                         {"vardas": vardas, "tel": [], "email": None})
            if r["kliento_el_pastas"] and "@" in str(r["kliento_el_pastas"]):
                irasas["email"] = str(r["kliento_el_pastas"]).strip().lower()
            for z in re.split(r"[;,/]| ir ", str(r["kliento_telefonas"] or "")):
                num = normalizuoti_numeri(z)
                if num and num not in irasas["tel"]:
                    irasas["tel"].append(num)
    finally:
        conn.close()

    kortos = []
    su_tel = 0
    for irasas in sorted(klientai.values(), key=lambda x: x["vardas"].lower()):
        if not irasas["tel"] and not irasas["email"]:
            continue
        vardas = irasas["vardas"]
        dalys = vardas.split(" ", 1)
        org = ""
        if irasas["email"]:
            org = ORG_PAGAL_DOMENA.get(irasas["email"].split("@")[-1], "")
        eilutes = ["BEGIN:VCARD", "VERSION:3.0",
                   f"FN:{vardas}",
                   f"N:{dalys[1] if len(dalys) > 1 else ''};{dalys[0]};;;"]
        for t in irasas["tel"]:
            eilutes.append(f"TEL;TYPE=CELL:{t}")
        if irasas["email"]:
            eilutes.append(f"EMAIL;TYPE=WORK:{irasas['email']}")
        if org:
            eilutes.append(f"ORG:{org}")
        eilutes.append("END:VCARD")
        kortos.append("\n".join(eilutes))
        if irasas["tel"]:
            su_tel += 1

    with open(isvestis, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(kortos) + "\n")
    return len(kortos), su_tel


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    viso, su_tel = eksportuoti()
    print(f"Kontaktų faile: {viso} (su telefonu: {su_tel})")
    print(f"Failas: {ISVESTIS}")
