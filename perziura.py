# -*- coding: utf-8 -*-
"""Užsakymų PERŽIŪRA (vadovui) — tik skaitymo langas.

Rodo užsakymus iš uzsakymai.db šiame aplanke ir automatiškai atsinaujina,
kai OneDrive atsiunčia naujesnę duomenų bazės kopiją. Nereikia jokių
papildomų bibliotekų — užtenka paprasto Python (su tkinter).
"""
import datetime
import json
import os
import re
import sqlite3
import sys
import tkinter as tk
from tkinter import ttk

if getattr(sys, "frozen", False):  # paleista kaip PyInstaller .exe
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

SHARE_FOLDER_NAME = "Užsakymų stebėjimas"
DB_NAME = "uzsakymai.db"
REFRESH_MS = 3000


def _log(msg):
    try:
        with open(os.path.join(APP_DIR, "perziura_log.txt"), "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%H:%M:%S} {msg}\n")
    except Exception:
        pass


def _candidates():
    """Visos vietos, kur gali būti duomenų failas."""
    found = [os.path.join(APP_DIR, DB_NAME)]

    hint = os.path.join(APP_DIR, "duomenu_kelias.txt")
    if os.path.exists(hint):
        try:
            with open(hint, "r", encoding="utf-8-sig") as f:
                p = f.read().strip().strip('"')
            if p:
                found.append(os.path.join(p, DB_NAME) if os.path.isdir(p) else p)
        except Exception:
            pass

    home = os.path.expanduser("~")
    # Paštu atsiųstos kopijos: Downloads\uzsakymai.db, uzsakymai (1).db ir pan.
    for folder in (os.path.join(home, "Downloads"), os.path.join(home, "Desktop")):
        try:
            for name in os.listdir(folder):
                low = name.lower()
                if low.startswith("uzsakymai") and low.endswith(".db"):
                    found.append(os.path.join(folder, name))
        except OSError:
            pass

    # OneDrive aplankai (jei kada nors bus bendrinamas aplankas)
    try:
        roots = [os.path.join(home, d) for d in os.listdir(home)
                 if d.lower().startswith("onedrive")]
    except OSError:
        roots = []
    for root in roots:
        found.append(os.path.join(root, SHARE_FOLDER_NAME, DB_NAME))
    return found


def _find_db():
    """Grąžina NAUJAUSIĄ rastą duomenų failą (kad ką tik atsiųsta kopija
    būtų paimta automatiškai, nieko neperrašinėjant)."""
    best, best_mtime = None, -1.0
    for path in _candidates():
        try:
            if os.path.exists(path):
                m = os.path.getmtime(path)
                if m > best_mtime:
                    best, best_mtime = path, m
        except OSError:
            continue
    return best or os.path.join(APP_DIR, DB_NAME)


try:
    DB_PATH = _find_db()
    _log(f"start: DB={DB_PATH} exists={os.path.exists(DB_PATH)}")
except Exception:
    import traceback
    _log("start KLAIDA:\n" + traceback.format_exc())
    DB_PATH = os.path.join(APP_DIR, DB_NAME)

STATUS_COLORS = {
    "p_užklausta": "#ffffff",
    "p_pasiūlyta": "#cfe2ff",
    "p_laimėta": "#e6d3f2",
    "p_nelaimėta": "#e8c4c4",
    "p_užsakyta": "#c8f0c8",
    "n_užklausta": "#ffffff",
    "n_užsakyta": "#d1ecf1",
}

SASKAITOS_ZENKLAI = {("n", 1): "✅", ("n", 2): "◐",
                     ("p", 1): "✅", ("p", 2): "🟠"}


def _saskaitos_zenklas(o):
    s = o.get("saskaita") or 0
    if not s:
        return ""
    g = "p" if (o.get("kodas") or "").startswith("!") else "n"
    return SASKAITOS_ZENKLAI.get((g, int(s)), "")


TERMINO_SPALVA = "#c00000"          # „!" užsakymo terminas dar galioja
TERMINO_SPALVA_PRAEJES = "#8a8a8a"  # terminas jau praėjęs

_TERM_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{1,2})[:.](\d{2}))?")


def _terminas(o):
    """(tekstas, iki) — kada pateikti pasiūlymą; kaip pagrindinėje programoje."""
    raw = (o.get("pasiulymo_terminas") or "").strip()
    prior = (o.get("kodas") or "").startswith("!")
    if not raw and not prior:
        raw = (o.get("pristatymo_data") or "").strip()
    if not raw:
        return "", None
    m = _TERM_RE.search(raw)
    if not m:
        return raw[:24], None
    y, mo, d, hh, mm = m.groups()
    try:
        if hh is not None:
            return (f"{y}-{mo}-{d} {int(hh):02d}:{mm}",
                    datetime.datetime(int(y), int(mo), int(d), int(hh), int(mm)))
        return f"{y}-{mo}-{d}", datetime.datetime(int(y), int(mo), int(d), 23, 59)
    except ValueError:
        return raw[:24], None


def load_orders(status_filter=None):
    if not os.path.exists(DB_PATH):
        return []
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True, timeout=5)
    conn.row_factory = sqlite3.Row
    try:
        if status_filter == "nauji":
            rows = conn.execute("SELECT * FROM orders WHERE IFNULL(statusas,'')='' "
                                "ORDER BY received_at DESC").fetchall()
        elif status_filter:
            rows = conn.execute("SELECT * FROM orders WHERE statusas=? "
                                "ORDER BY received_at DESC", (status_filter,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM orders ORDER BY received_at DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


class Viewer:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.last_mtime = 0.0
        root.title("Užsakymai — peržiūra (tik skaitymui)")
        root.geometry("1200x640")

        filter_bar = ttk.Frame(root, padding=(8, 6))
        filter_bar.pack(fill="x")
        ttk.Label(filter_bar, text="Rodyti:").pack(side="left")
        self.filter_var = tk.StringVar(value="visi")
        for val, text in [("visi", "Visus"), ("nauji", "Naujus"),
                          ("užklausta", "Užklausta"), ("pasiūlyta", "Pasiūlyta"),
                          ("laimėta", "Laimėta"), ("nelaimėta", "Nelaimėta"),
                          ("užsakyta", "Užsakyta")]:
            ttk.Radiobutton(filter_bar, text=text, value=val, variable=self.filter_var,
                            command=self.refresh).pack(side="left", padx=4)

        main = ttk.PanedWindow(root, orient="horizontal")
        main.pack(fill="both", expand=True, padx=8, pady=6)

        left = ttk.Frame(main)
        main.add(left, weight=3)
        cols = ("atn", "kodas", "terminas", "gauta", "klientas", "prekes",
                "priedai", "statusas", "saskaita")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for col, text, w, stretch in [
                ("atn", "🔔", 36, False),
                ("kodas", "Kodas", 110, False),
                ("terminas", "⏰ Pateikti iki", 125, False),
                ("gauta", "Data", 115, False),
                ("klientas", "Klientas", 165, False), ("prekes", "Prekės", 250, True),
                ("priedai", "📎", 45, False), ("statusas", "Statusas", 90, False),
                ("saskaita", "Sąsk.", 50, False)]:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=w, stretch=stretch,
                             anchor="center" if col in ("atn", "priedai", "saskaita") else "w")
        for tag, color in STATUS_COLORS.items():
            self.tree.tag_configure(tag, background=color)
            self.tree.tag_configure(f"{tag}_t", background=color,
                                    foreground=TERMINO_SPALVA)
            self.tree.tag_configure(f"{tag}_p", background=color,
                                    foreground=TERMINO_SPALVA_PRAEJES)
        self.tree.tag_configure("t", foreground=TERMINO_SPALVA)
        self.tree.tag_configure("p", foreground=TERMINO_SPALVA_PRAEJES)
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)

        right = ttk.Frame(main)
        main.add(right, weight=2)
        self.detail = tk.Text(right, wrap="word", state="disabled", font=("Segoe UI", 10))
        self.detail.pack(fill="both", expand=True)

        status_bar = ttk.Frame(root, padding=(8, 3))
        status_bar.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="")
        ttk.Label(status_bar, textvariable=self.status_var, foreground="#555").pack(side="left")

        self.refresh()
        self.root.after(REFRESH_MS, self.poll)

    def poll(self):
        global DB_PATH
        try:
            found = _find_db()  # ar neatsirado naujesnės kopijos (pvz., Downloads)
            if found != DB_PATH and os.path.exists(found):
                DB_PATH = found
                self.last_mtime = 0  # priverstinai perskaitome naują failą
                _log(f"perjungta i naujesni faila: {found}")
            mtime = os.path.getmtime(DB_PATH) if os.path.exists(DB_PATH) else 0
            if mtime != self.last_mtime:
                self.refresh()
        except Exception:
            pass
        self.root.after(REFRESH_MS, self.poll)

    def refresh(self):
        try:
            self.last_mtime = os.path.getmtime(DB_PATH) if os.path.exists(DB_PATH) else 0
            status = None if self.filter_var.get() == "visi" else self.filter_var.get()
            orders = load_orders(status)
        except Exception:
            return  # OneDrive dar siunčia failą — bandysime kitą kartą
        selected = self.tree.selection()
        sel_id = selected[0] if selected else None
        for item in self.tree.get_children():
            self.tree.delete(item)
        for o in orders:
            try:
                items = json.loads(o.get("prekes_json") or "[]")
            except Exception:
                items = []
            prekes = "; ".join(f"{it.get('preke')} ({it.get('kiekis')})" for it in items[:4])
            if len(items) > 4:
                prekes += f" … (+{len(items) - 4})"
            n_atts = len([s for s in (o.get("priedai") or "").split(";") if s.strip()])
            status_o = o.get("statusas") or ""
            prior = (o.get("kodas") or "").startswith("!")
            group = "p" if prior else "n"
            term_txt, term_iki = _terminas(o)
            zyme = ""
            if prior and term_txt:
                zyme = "p" if (term_iki and term_iki < datetime.datetime.now()) else "t"
            base = f"{group}_{status_o}" if status_o else ""
            if zyme:
                tags = (f"{base}_{zyme}",) if base else (zyme,)
            else:
                tags = (base,) if base else ()
            self.tree.insert("", "end", iid=str(o["id"]), tags=tags, values=(
                "🔔" if o.get("atnaujinta") else "",
                o.get("kodas") or "", term_txt, o.get("received_at"),
                o.get("klientas") or "?",
                prekes, f"📎{n_atts}" if n_atts else "", status_o or "—",
                _saskaitos_zenklas(o)))
        if sel_id and self.tree.exists(sel_id):
            self.tree.selection_set(sel_id)
        if os.path.exists(DB_PATH):
            try:
                stamp = datetime.datetime.fromtimestamp(
                    os.path.getmtime(DB_PATH)).strftime("%m-%d %H:%M")
            except OSError:
                stamp = "—"
            self.status_var.set(
                f"Užsakymų: {len(orders)} · Duomenys iš {stamp} "
                f"({os.path.basename(os.path.dirname(DB_PATH))}) · "
                f"tikrinama automatiškai kas 3 s")
        else:
            self.status_var.set(
                "Duomenų nerasta. Įdėkite failą „uzsakymai.db“ į šį aplanką "
                "(arba tiesiog išsaugokite jį iš laiško į Downloads) — "
                "programa jį pasiims pati.")

    def show_detail(self, _event=None):
        sel = self.tree.selection()
        self.detail.config(state="normal")
        self.detail.delete("1.0", "end")
        if sel:
            conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True, timeout=5)
            conn.row_factory = sqlite3.Row
            row = conn.execute("SELECT * FROM orders WHERE id=?", (sel[0],)).fetchone()
            conn.close()
            if row:
                o = dict(row)
                try:
                    items = json.loads(o.get("prekes_json") or "[]")
                except Exception:
                    items = []
                lines = [
                    f"Kodas:  {o.get('kodas') or '—'}",
                    f"Klientas:  {o.get('klientas') or '—'}",
                    f"Telefonas:  {o.get('kliento_telefonas') or '—'}",
                    f"Adresas:  {o.get('kliento_adresas') or '—'}",
                    f"El. paštas:  {o.get('kliento_el_pastas') or '—'}",
                    f"Užsakymo data:  {o.get('received_at')}",
                    f"Pasiūlyti iki:  {o.get('pasiulymo_terminas') or '—'}",
                    f"Sutartis:  {o.get('sutartis') or '—'}",
                    f"Priedai:  {o.get('priedai') or 'nėra'}",
                    f"Statusas:  {o.get('statusas') or '—'}",
                    f"Sąskaita:  "
                    f"{ {1: '✅ išrašyta', 2: '◐ dalinai išrašyta'}.get(int(o.get('saskaita') or 0), 'neišrašyta') }",
                    "",
                    "PREKĖS:",
                ]
                for i, it in enumerate(items, 1):
                    line = f"  {i}. {it.get('preke')} — {it.get('kiekis')}"
                    if it.get("pastabos"):
                        line += f"   [{it['pastabos']}]"
                    lines.append(line)
                if o.get("atnaujinimai"):
                    lines += ["", "🔔 ATNAUJINIMAI (grandinės papildymai):"]
                    lines += ["   " + ln for ln in o["atnaujinimai"].splitlines()]
                if o.get("mano_pastabos"):
                    lines += ["", f"Pastabos: {o['mano_pastabos']}"]
                if o.get("santrauka"):
                    lines += ["", f"Santrauka: {o['santrauka']}"]
                lines += ["", f"Laiškas: „{o.get('subject')}“", f"Siuntėjas: {o.get('sender')}"]
                self.detail.insert("1.0", "\n".join(lines))
        self.detail.config(state="disabled")


def main():
    try:
        root = tk.Tk()
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        Viewer(root)
        _log("langas sukurtas, paleidziamas mainloop")
        root.mainloop()
        _log("mainloop baigesi")
    except Exception:
        import traceback
        _log("main KLAIDA:\n" + traceback.format_exc())
        try:
            from tkinter import messagebox
            messagebox.showerror("Klaida", traceback.format_exc()[:1500])
        except Exception:
            pass


if __name__ == "__main__":
    main()
