# -*- coding: utf-8 -*-
"""Užsakymų agentas — darbalaukio programa.

Periodiškai (arba paspaudus mygtuką) peržiūri naujus Outlook laiškus,
per Claude API atpažįsta užsakymus (tekstas, Excel, PDF, nuotraukos)
ir sudeda juos į sąrašą su kodais (priešdėlis + inicialai + data), pastabomis
ir eksportu į Excel.
"""
import datetime
import json
import os
import queue
import re
import shutil
import sqlite3
import tempfile
import threading
import traceback
import webbrowser
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed, wait as futures_wait
from tkinter import filedialog, messagebox, ttk

import anthropic

import attachments as attachments_mod
import config as cfg_mod
import exporter
import html_perziura
import extractor
import graph_client
import outlook_client
import prekes as prekes_mod
import share
import storage

# Moderni išvaizda (ttkbootstrap). Jei neįdiegta — programa veikia su įprasta tema.
try:
    import ttkbootstrap as tb
    from ttkbootstrap.widgets import DateEntry as _TbDateEntry
    TB = True
except Exception:
    tb = None
    _TbDateEntry = None
    TB = False
TEMA = "cosmo"

MODELS = ["claude-opus-5", "claude-sonnet-5", "claude-haiku-4-5"]

# Statusai: prioritetinės sutarties („!" kodai) turi pilną eigą, kiti — tik užklausta/užsakyta.
PRIORITY_STATUSES = ["užklausta", "pasiūlyta", "laimėta", "nelaimėta", "užsakyta"]
NORMAL_STATUSES = ["užklausta", "užsakyta"]
STATUS_COLORS = {
    "p_užklausta": "#ffffff",   # ! balta (vartotojo pageidavimu — nesispalvina)
    "p_pasiūlyta": "#cfe2ff",   # ! mėlyna
    "p_laimėta":  "#e6d3f2",    # ! violetinė
    "p_nelaimėta": "#e8c4c4",   # ! rausvai pilka — pralaimėtas pasiūlymas
    "p_užsakyta": "#c8f0c8",    # ! žalia
    "n_užklausta": "#ffffff",   # balta (vartotojo pageidavimu — nesispalvina)
    "n_užsakyta": "#d1ecf1",    # žalsvai melsva
}
TEMP_ATT_ROOT = os.path.join(tempfile.gettempdir(), "uzsakymu_priedai")

# „!" užsakymų (prioritetinės sutarties) pasiūlymo terminas rodomas raudonai —
# jį pramiegoti brangiausia, todėl matomas iškart sąraše, be jokių paspaudimų.
TERMINO_SPALVA = "#c00000"
TERMINO_SPALVA_PRAEJES = "#8a8a8a"   # terminas jau praėjęs

# Kiek parų atgal visada peržvelgiame Outlook, nesvarbu, kada buvo paskutinis
# tikrinimas. Laiškas į vietinį Outlook neretai nusėda vėliau nei jo laiko žyma
# (ypač su priedais); be šio saugiklio jis liktų senesnis už tikrinimo langą ir
# nebūtų paimtas NIEKADA. AI kaštų tai nekainuoja — jau apdoroti laiškai
# praleidžiami dar prieš skaitant priedus (processed lentelė).
SAUGOS_LANGO_DIENOS = 2


def _is_priority_order(o: dict) -> bool:
    return (o.get("kodas") or "").startswith("!")


# Sąskaitos žymė (saskaita): 0 — nėra, 1 — išrašyta, 2 — dalinai.
# Išrašyta — visiems vienodai ✅; skiriasi tik DALINAI ženklas, kad matytųsi,
# ar dalinė sąskaita priklauso „!" (prioritetinės sutarties) užsakymui.
SASKAITOS_ZENKLAI = {("n", 1): "✅", ("n", 2): "◐",
                     ("p", 1): "✅", ("p", 2): "🟠"}
SASKAITOS_PASIRINKIMAI = ["—", "dalinai", "išrašyta"]
_SASK_I_REIKSME = {"—": 0, "dalinai": 2, "išrašyta": 1}
_SASK_I_TEKSTA = {0: "—", 1: "išrašyta", 2: "dalinai"}


def _saskaitos_zenklas(o: dict) -> str:
    s = o.get("saskaita") or 0
    if not s:
        return ""
    g = "p" if _is_priority_order(o) else "n"
    return SASKAITOS_ZENKLAI.get((g, int(s)), "")


def _be_diakritiku(t: str) -> str:
    """Mazosios be lietuvisku zenklu: „Uzsakymas" == „užsakymas" (raktazodziams)."""
    import unicodedata
    t = unicodedata.normalize("NFKD", t or "")
    return "".join(c for c in t if not unicodedata.combining(c)).lower()


_TERM_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{1,2})[:.](\d{2}))?")


def _terminas(o: dict):
    """Grąžina (tekstas, iki) — kada pateikti pasiūlymą ir kelintą valandą.

    AI terminą ištraukia laisvu tekstu („2026-08-10 10.00 val.", „2026-08-12 10:00
    (kaip nurodyta užklausoje)"), todėl datą ir laiką išsitraukiame reguliariuoju
    reiškiniu; neatpažinus formato rodome pirmus simbolius kaip yra."""
    raw = (o.get("pasiulymo_terminas") or "").strip()
    if not raw and not _is_priority_order(o):
        raw = (o.get("pristatymo_data") or "").strip()
    if not raw:
        return "", None
    m = _TERM_RE.search(raw)
    if not m:
        return raw[:24], None
    y, mo, d, hh, mm = m.groups()
    if hh is not None:
        try:
            iki = datetime.datetime(int(y), int(mo), int(d), int(hh), int(mm))
        except ValueError:
            return raw[:24], None
        return f"{y}-{mo}-{d} {int(hh):02d}:{mm}", iki
    try:
        iki = datetime.datetime(int(y), int(mo), int(d), 23, 59)
    except ValueError:
        return raw[:24], None
    return f"{y}-{mo}-{d}", iki



# ---------------- Bendri UI pagalbininkai ----------------

def _bs(kw: dict, style):
    """ttkbootstrap spalvinis stilius (jei tema įdiegta) — kitaip be jo."""
    if TB and style:
        kw["bootstyle"] = style
    return kw


def _tkkw() -> dict:
    """Klasikiniams tk valdikliams: kad tema neperrašytų mūsų spalvų."""
    return {"autostyle": False} if TB else {}


def _btn(parent, text, command, style=None, **kw):
    return ttk.Button(parent, text=text, command=command, **_bs(kw, style))


def _date_field(parent, var, width=11):
    """Datos laukas su kalendoriumi (jei yra ttkbootstrap), kitaip — paprastas laukas."""
    if TB:
        try:
            de = _TbDateEntry(parent, dateformat="%Y-%m-%d", width=width, firstweekday=0)
            de.entry.delete(0, "end")
            de.entry.configure(textvariable=var)
            return de
        except Exception:
            pass
    return ttk.Entry(parent, textvariable=var, width=width)


class Tooltip:
    """Paaiškinimas užvedus pelę. Su tekstu — automatinis; be teksto — show()/hide() ranka."""

    def __init__(self, widget, text: str = "", delay: int = 450):
        self.widget, self.text, self.delay = widget, text, delay
        self.win = None
        self._after = None
        if text:
            widget.bind("<Enter>", lambda e: self._schedule(), add="+")
            widget.bind("<Leave>", lambda e: self.hide(), add="+")
            widget.bind("<ButtonPress>", lambda e: self.hide(), add="+")

    def _schedule(self):
        self._cancel()
        self._after = self.widget.after(self.delay, lambda: self.show(
            self.widget.winfo_pointerx() + 14, self.widget.winfo_pointery() + 18, self.text))

    def _cancel(self):
        if self._after:
            try:
                self.widget.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def show(self, x, y, text):
        self._close()
        if not text:
            return
        try:
            tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)
            lbl = tk.Label(tw, text=text, justify="left", padx=9, pady=6,
                           font=("Segoe UI", 9), wraplength=460, **_tkkw())
            lbl.configure(background="#1f2937", foreground="#f9fafb")
            lbl.pack()
            tw.wm_geometry(f"+{x}+{y}")
            self.win = tw
        except tk.TclError:
            self.win = None

    def _close(self):
        if self.win is not None:
            try:
                self.win.destroy()
            except tk.TclError:
                pass
            self.win = None

    def hide(self):
        self._cancel()
        self._close()


UNIT_CHOICES = ["vnt.", "kompl.", "pak.", "dėž.", "m", "m²", "m³", "kg", "l", "t",
                "rul.", "pad.", "por.", "val."]


class ItemsEditor(ttk.Frame):
    """Prekių lentelė, kurią galima taisyti tiesiai langeliuose.

    Dukart spustelėjus langelį atsiranda laukas: Enter — išsaugo, Esc — atšaukia,
    Tab / Shift+Tab — į kitą / ankstesnį langelį. Eilutės be kiekio pažymimos
    oranžiškai — kad iš karto matytųsi, ką reikia pasitikslinti."""

    COLS = ("nr", "preke", "kodas", "kiekis", "vienetas", "pastabos", "nuoroda")
    EDIT_COLS = ("preke", "kodas", "kiekis", "vienetas", "pastabos", "nuoroda")
    HEAD = {"nr": "#", "preke": "Prekė", "kodas": "Kodas", "kiekis": "Kiekis",
            "vienetas": "Vnt.", "pastabos": "Pastabos", "nuoroda": "Nuoroda"}
    WIDTH = {"nr": (32, False), "preke": (200, True), "kodas": (90, False),
             "kiekis": (60, False), "vienetas": (52, False), "pastabos": (140, True),
             "nuoroda": (70, False)}

    def __init__(self, master, on_change=None, height=6):
        super().__init__(master)
        self.on_change = on_change
        self.items = []
        self._editor = None
        self._editable = True
        self._tip_key = None

        fr = ttk.Frame(self)
        fr.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(fr, columns=self.COLS, show="headings", height=height,
                                 selectmode="extended")
        for c in self.COLS:
            w, stretch = self.WIDTH[c]
            self.tree.heading(c, text=self.HEAD[c])
            self.tree.column(c, width=w, minwidth=30, stretch=stretch,
                             anchor="center" if c in ("nr", "kiekis", "vienetas", "nuoroda")
                             else "w")
        self.tree.tag_configure("truksta", background="#fff1d6")
        vsb = ttk.Scrollbar(fr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(4, 0))
        self.btn_add = _btn(bar, "➕ Pridėti prekę", self.add_row, "success-outline")
        self.btn_add.pack(side="left")
        self.btn_del = _btn(bar, "🗑 Pašalinti pažymėtas", self.delete_selected,
                            "secondary-outline")
        self.btn_del.pack(side="left", padx=(6, 0))
        self.hint = ttk.Label(bar, text="Dukart spustelėkite langelį, kad pataisytumėte",
                              foreground="#6b7280")
        self.hint.pack(side="right")

        self.tip = Tooltip(self.tree)
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Return>", self._on_enter)
        self.tree.bind("<F2>", self._on_enter)
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Motion>", self._on_motion)
        self.tree.bind("<Leave>", lambda e: self._hide_tip())
        self.tree.bind("<MouseWheel>", lambda e: (self.commit_edit(), self._hide_tip()), add="+")

    # --- duomenys ---

    def set_items(self, items, editable: bool = True):
        self.cancel_edit()
        norm = []
        for it in items or []:
            e = prekes_mod.prekes_eilute(it)
            if e is not None:
                norm.append(e)
        self.items = norm
        self._editable = editable
        state = "normal" if editable else "disabled"
        self.btn_add.config(state=state)
        self.btn_del.config(state=state)
        self.hint.config(text="Dukart spustelėkite langelį, kad pataisytumėte" if editable else "")
        self.render()

    def get_items(self) -> list:
        self.commit_edit()
        return [dict(it) for it in self.items]

    def render(self, select: int = None):
        self.tree.delete(*self.tree.get_children())
        for i, it in enumerate(self.items):
            kiekis = str(it.get("kiekis") or "")
            self.tree.insert("", "end", iid=str(i),
                             tags=("truksta",) if not kiekis.strip() else (),
                             values=(i + 1, it.get("preke") or "", it.get("kodas") or "",
                                     kiekis, it.get("vienetas") or "",
                                     it.get("pastabos") or "",
                                     "🔗" if it.get("nuoroda") else ""))
        if select is not None and 0 <= select < len(self.items):
            self.tree.selection_set(str(select))
            self.tree.focus(str(select))
            self.tree.see(str(select))

    def _changed(self):
        if self.on_change:
            self.on_change()

    # --- veiksmai ---

    def add_row(self):
        if not self._editable:
            return
        self.commit_edit()
        self.items.append({"preke": "", "kodas": None, "kiekis": "", "vienetas": None,
                           "pastabos": None, "nuoroda": None})
        idx = len(self.items) - 1
        self.render(select=idx)
        self._changed()
        self.edit_cell(str(idx), "preke")

    def delete_selected(self):
        if not self._editable:
            return
        self.commit_edit()
        sel = sorted((int(i) for i in self.tree.selection()), reverse=True)
        if not sel:
            messagebox.showinfo("Pašalinti", "Pirma pažymėkite prekės eilutę.", parent=self)
            return
        for i in sel:
            if i < len(self.items):
                del self.items[i]
        self.render(select=min(sel[-1], len(self.items) - 1) if self.items else None)
        self._changed()

    # --- redagavimas langelyje ---

    def edit_cell(self, iid: str, col: str):
        if not self._editable or col not in self.EDIT_COLS:
            return
        self.commit_edit()
        self.tree.see(iid)
        self.tree.update_idletasks()
        bbox = self.tree.bbox(iid, col)
        if not bbox:
            return
        x, y, w, h = bbox
        idx = int(iid)
        val = str(self.items[idx].get(col) or "")
        if col == "vienetas":
            ed = ttk.Combobox(self.tree, values=UNIT_CHOICES)
            ed.bind("<<ComboboxSelected>>", lambda e: self.commit_edit())
        else:
            ed = ttk.Entry(self.tree)
        ed.insert(0, val)
        ed.select_range(0, "end")
        ed.icursor("end")
        # siauri stulpeliai (kodas, nuoroda) redaguojant praplatinami
        min_w = {"nuoroda": 320, "kodas": 140, "vienetas": 90, "kiekis": 80}.get(col, w)
        ed.place(x=x, y=y, width=max(w, min_w), height=h)
        ed.focus_set()
        self._editor = (ed, idx, col)
        ed.bind("<Return>", lambda e: (self.commit_edit(), "break")[1])
        ed.bind("<KP_Enter>", lambda e: (self.commit_edit(), "break")[1])
        ed.bind("<Escape>", lambda e: (self.cancel_edit(), "break")[1])
        ed.bind("<Tab>", lambda e: (self._move(1), "break")[1])
        ed.bind("<Shift-Tab>", lambda e: (self._move(-1), "break")[1])
        ed.bind("<FocusOut>", lambda e: self.after(120, self._focus_check))

    def _focus_check(self):
        """Fokusas paliko lauką (bet ne į jo išskleidžiamą sąrašą) — išsaugome."""
        if not self._editor:
            return
        ed = self._editor[0]
        try:
            f = str(self.tk.call("focus"))
        except tk.TclError:
            f = ""
        if f and f.startswith(str(ed)):
            return
        self.commit_edit()

    def commit_edit(self):
        if not self._editor:
            return
        ed, idx, col = self._editor
        self._editor = None
        try:
            val = ed.get().strip()
            ed.destroy()
        except tk.TclError:
            return
        if idx >= len(self.items):
            return
        old = str(self.items[idx].get(col) or "")
        if old != val:
            self.items[idx][col] = val if col in ("preke", "kiekis") else (val or None)
            self.render(select=idx)
            self._changed()

    def cancel_edit(self):
        if self._editor:
            ed = self._editor[0]
            self._editor = None
            try:
                ed.destroy()
            except tk.TclError:
                pass

    def _move(self, step: int):
        if not self._editor:
            return
        _ed, idx, col = self._editor
        self.commit_edit()
        cols = list(self.EDIT_COLS)
        pos = cols.index(col) + step
        if pos >= len(cols):
            pos, idx = 0, idx + 1
        elif pos < 0:
            pos, idx = len(cols) - 1, idx - 1
        if 0 <= idx < len(self.items):
            self.edit_cell(str(idx), cols[pos])

    def _on_double(self, e):
        if not self._editable:
            return
        region = self.tree.identify_region(e.x, e.y)
        if region == "nothing":
            self.add_row()
            return
        if region != "cell":
            return
        iid = self.tree.identify_row(e.y)
        col = self.COLS[int(self.tree.identify_column(e.x)[1:]) - 1]
        self.edit_cell(iid, "preke" if col == "nr" else col)

    def _on_enter(self, _e=None):
        iid = self.tree.focus()
        if iid:
            self.edit_cell(iid, "preke")
        return "break"

    def _on_right_click(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        if iid not in self.tree.selection():
            self.tree.selection_set(iid)
        idx = int(iid)
        m = tk.Menu(self, tearoff=0)
        if self._editable:
            m.add_command(label="✎ Taisyti", command=lambda: self.edit_cell(iid, "preke"))
        link = self.items[idx].get("nuoroda") if idx < len(self.items) else None
        if link:
            m.add_command(label="🔗 Atidaryti nuorodą", command=lambda: webbrowser.open(link))
        if self._editable:
            m.add_separator()
            m.add_command(label="🗑 Pašalinti eilutę", command=self.delete_selected)
        m.tk_popup(e.x_root, e.y_root)

    # --- pilnas tekstas užvedus pelę ---

    def _on_motion(self, e):
        if self.tree.identify_region(e.x, e.y) != "cell":
            self._hide_tip()
            return
        iid = self.tree.identify_row(e.y)
        col = self.COLS[int(self.tree.identify_column(e.x)[1:]) - 1]
        key = (iid, col)
        if key == self._tip_key:
            return
        self._hide_tip()
        if not iid or col not in ("preke", "kodas", "pastabos", "nuoroda"):
            return
        val = str(self.items[int(iid)].get(col) or "") if int(iid) < len(self.items) else ""
        if col == "nuoroda" or len(val) > 28:
            self._tip_key = key
            self.tip.show(e.x_root + 14, e.y_root + 18, val)

    def _hide_tip(self):
        self._tip_key = None
        self.tip.hide()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = cfg_mod.Config()
        storage.init_db()

        self.ui_queue = queue.Queue()
        self.check_now_event = threading.Event()
        self._day_win = None
        self._day_current = None
        self._sut_win = None            # sutarčių langas
        self._pick_win = None           # rankinio laiškų pasirinkimo langas
        self._day_notes_after = None    # atidėto automatinio saugojimo laikmatis
        self._notes_order_id = None     # kurio užsakymo pastabos atvertos
        self._notes_after = None
        self._analysis_lock = threading.Lock()  # kad importas ir tikrinimas nesusidurtų
        storage.carry_over_tasks()  # paleidus: vakarykščiai neatlikti darbai -> šiandien
        self._carry_date = datetime.date.today()
        shutil.rmtree(TEMP_ATT_ROOT, ignore_errors=True)  # senų laikinų priedų valymas

        root.title("Užsakymų agentas — Outlook + Claude")
        # lango dydis — pagal ekraną (kad tilptų ir nešiojamame kompiuteryje)
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h = min(1600, sw - 60), min(900, sh - 100)
        root.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        root.minsize(1000, 600)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._refresh_orders()

        self.worker = threading.Thread(target=self._worker_loop, daemon=True)
        self.worker.start()
        self.root.after(200, self._poll_queue)

        if not self.cfg.get_api_key():
            self.root.after(400, self._open_settings)

    def _mail(self):
        """Pašto šaltinis pagal nustatymą: Graph API arba klasikinis Outlook (COM)."""
        return graph_client if self.cfg["mail_source"] == "graph" else outlook_client

    # ---------------- UI ----------------

    # Filtrai „Rodyti": vidinis kodas -> rodomas tekstas
    FILTRAI = [("visi", "Visi"), ("neuzbaigti", "⏳ Neužbaigti"),
               ("nauji", "🆕 Nauji (be statuso)"), ("užklausta", "Užklausta"),
               ("pasiūlyta", "Pasiūlyta"), ("laimėta", "Laimėta"),
               ("nelaimėta", "Nelaimėta"), ("užsakyta", "Užsakyta"),
               ("uzbaigti", "✔ Užbaigti")]
    LAIKOTARPIAI = [("Visas laikas", None), ("Šiandien", 0), ("Paskutinės 7 d.", 7),
                    ("Paskutinės 30 d.", 30), ("Šis mėnuo", -1), ("Pasirinktos datos", "x")]
    # Kortelės laukai: (raktas, pavadinimas, eilutė, stulpelis, per visą plotį)
    KORTELES_LAUKAI = [
        ("klientas", "Klientas", 0, 0, False), ("kliento_imone", "Įmonė", 0, 1, False),
        ("kliento_telefonas", "Telefonas", 1, 0, False),
        ("kliento_el_pastas", "El. paštas", 1, 1, False),
        ("kliento_adresas", "Adresas", 2, 0, True),
        ("sutartis", "Sutartis", 3, 0, False),
        ("pasiulymo_terminas", "Pasiūlyti iki", 3, 1, False),
        ("pristatymo_data", "Pristatymas", 4, 0, True),
    ]
    STULPELIU_PAAISKINIMAI = {
        "atn": "⚠  AI įspėjimas — būtinai perskaitykite prieš vykdant\n"
               "🔔  grandinėje atėjo naujas laiškas (papildymas)\n"
               "❗  atsakė tiekėjas į užklausą\n\n🔔 ir ❗ nusiima peržiūrėjus užsakymą.",
        "terminas": "Iki kada pateikti pasiūlymą.\nRaudonai — artėja, pilkai — jau praėjo.",
        "priedai": "Kiek priedų (be parašo logotipų) turi laiškas.",
        "saskaita": "✅  sąskaita išrašyta\n◐ / 🟠  išrašyta dalinai",
        "statusas": "Užsakymo eiga. Keičiama dešinėje, laukelyje „Statusas“.",
    }

    def _build_ui(self):
        self._card_id = None          # kuris užsakymas atvertas kortelėje
        self._card_dirty = False      # ar kortelėje yra neišsaugotų pataisymų
        self._card_loading = False
        self._card_snapshot = None
        self._period_prog = False     # laikotarpį nustato programa (ne žmogus)
        self._refresh_after = None
        self._tree_tip_key = None

        # --- Viršutinė juosta: tik dažniausi veiksmai, kiti — „Daugiau" ---
        toolbar = ttk.Frame(self.root, padding=(10, 8, 10, 4))
        toolbar.pack(fill="x")

        self.btn_check = _btn(toolbar, "🔄  Tikrinti paštą", self._check_now, "primary")
        self.btn_check.pack(side="left")
        Tooltip(self.btn_check, "Patikrinti naujus laiškus dabar")

        self.auto_var = tk.BooleanVar(value=bool(self.cfg["auto_check"]))
        self.lbl_interval = ttk.Checkbutton(
            toolbar, text=f"Automatiškai kas {self.cfg['check_interval_min']} min.",
            variable=self.auto_var, command=self._toggle_auto,
            **_bs({}, "round-toggle"))
        self.lbl_interval.pack(side="left", padx=(12, 0))

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=12)

        b = _btn(toolbar, "➕  Naujas užsakymas", self._new_order_dialog, "success-outline")
        b.pack(side="left")
        Tooltip(b, "Įvesti užsakymą ranka (be laiško)")
        b = _btn(toolbar, "📅  Dienos darbai", self._open_day_tasks, "secondary-outline")
        b.pack(side="left", padx=(6, 0))
        b = _btn(toolbar, "📑  Sutartys", self._open_contracts, "secondary-outline")
        b.pack(side="left", padx=(6, 0))

        more = ttk.Menubutton(toolbar, text="⋯  Daugiau", **_bs({}, "secondary-outline"))
        more.pack(side="left", padx=(6, 0))
        m = tk.Menu(more, tearoff=0)
        m.add_command(label="📥  Importuoti senesnius laiškus…", command=self._import_history)
        m.add_command(label="🖐  Rinktis laiškus analizei…", command=self._open_picker)
        m.add_separator()
        m.add_command(label="📊  Eksportuoti matomus į Excel…",
                      command=lambda: self._export(visible_only=True))
        m.add_command(label="📊  Eksportuoti visus į Excel…", command=self._export)
        m.add_separator()
        m.add_command(label="📤  Siųsti peržiūrą vadovui…", command=self._siusti_perziura)
        more["menu"] = m

        b = _btn(toolbar, "⚙  Nustatymai", self._open_settings, "secondary-link")
        b.pack(side="right")

        # --- Viena filtrų juosta ---
        fbar = ttk.Frame(self.root, padding=(10, 4, 10, 6))
        fbar.pack(fill="x")

        ttk.Label(fbar, text="Rodyti").pack(side="left")
        self.filter_var = tk.StringVar(value="visi")
        self.filter_combo = ttk.Combobox(fbar, state="readonly", width=24)
        self.filter_combo.pack(side="left", padx=(6, 14))
        self.filter_combo.bind("<<ComboboxSelected>>", self._on_filter_pick)

        ttk.Label(fbar, text="Laikotarpis").pack(side="left")
        self.period_combo = ttk.Combobox(fbar, state="readonly", width=17,
                                         values=[t for t, _ in self.LAIKOTARPIAI])
        self.period_combo.set("Visas laikas")
        self.period_combo.pack(side="left", padx=(6, 6))
        self.period_combo.bind("<<ComboboxSelected>>", self._on_period_pick)
        self.date_from = tk.StringVar()
        self.date_to = tk.StringVar()
        _date_field(fbar, self.date_from).pack(side="left")
        ttk.Label(fbar, text="–").pack(side="left", padx=3)
        _date_field(fbar, self.date_to).pack(side="left")
        self.date_from.trace_add("write", self._on_date_typed)
        self.date_to.trace_add("write", self._on_date_typed)

        ttk.Label(fbar, text="🔍").pack(side="left", padx=(16, 4))
        self.search_var = tk.StringVar()
        se = ttk.Entry(fbar, textvariable=self.search_var, width=26)
        se.pack(side="left")
        Tooltip(se, "Ieško pagal kodą, klientą, laiško temą, prekes ir jūsų pastabas.\n"
                    "Rezultatai rodomi iškart rašant.")
        self.search_var.trace_add("write", lambda *a: self._refresh_later())
        se.bind("<Escape>", lambda e: self.search_var.set(""))
        clr = _btn(fbar, "✕", lambda: self.search_var.set(""), "secondary-link", width=3)
        clr.pack(side="left")
        Tooltip(clr, "Išvalyti paiešką")

        self.count_var = tk.StringVar()
        ttk.Label(fbar, textvariable=self.count_var, foreground="#6b7280").pack(side="right")

        # --- Būsenos juosta (apačioje) ---
        status_bar = ttk.Frame(self.root, padding=(10, 3))
        status_bar.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Paruošta.")
        ttk.Label(status_bar, textvariable=self.status_var, foreground="#555").pack(side="left")
        ttk.Label(status_bar, foreground="#8a8a8a",
                  text="⚠ AI įspėjimas   🔔 naujas laiškas   ❗ atsakė tiekėjas   "
                       "📎 priedai   ✅ sąskaita").pack(side="right")

        # --- Pagrindinė dalis: sąrašas kairėje, kortelė dešinėje ---
        main = ttk.PanedWindow(self.root, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        left = ttk.Frame(main)
        main.add(left, weight=3)

        cols = ("atn", "kodas", "terminas", "gauta", "klientas", "prekes",
                "priedai", "statusas", "saskaita")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("atn", text="🔔")
        self.tree.column("atn", width=40, stretch=False, anchor="center")
        self.tree.heading("kodas", text="Kodas")
        self.tree.heading("terminas", text="⏰ Pateikti iki")
        self.tree.heading("gauta", text="Data")
        self.tree.heading("klientas", text="Klientas")
        self.tree.heading("prekes", text="Prekės")
        self.tree.heading("priedai", text="📎")
        self.tree.heading("statusas", text="Statusas")
        self.tree.heading("saskaita", text="Sąsk.")
        self.tree.column("kodas", width=105, minwidth=90, stretch=False)
        self.tree.column("terminas", width=140, minwidth=105, stretch=False)
        self.tree.column("gauta", width=140, minwidth=90, stretch=False)
        self.tree.column("klientas", width=140, minwidth=100, stretch=False)
        self.tree.column("prekes", width=170, minwidth=120)
        self.tree.column("priedai", width=42, stretch=False, anchor="center")
        self.tree.column("statusas", width=85, minwidth=70, stretch=False)
        self.tree.column("saskaita", width=50, stretch=False, anchor="center")
        for tag, color in STATUS_COLORS.items():
            self.tree.tag_configure(tag, background=color)
            # „_t" — tas pats fonas, bet raudonas tekstas (artėja pasiūlymo terminas)
            self.tree.tag_configure(f"{tag}_t", background=color,
                                    foreground=TERMINO_SPALVA)
            self.tree.tag_configure(f"{tag}_p", background=color,
                                    foreground=TERMINO_SPALVA_PRAEJES)
        self.tree.tag_configure("t", foreground=TERMINO_SPALVA)
        self.tree.tag_configure("p", foreground=TERMINO_SPALVA_PRAEJES)
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.tree.bind("<Leave>", lambda e: self._tree_tip_hide())
        self._tree_tip = Tooltip(self.tree)

        right = ttk.Frame(main, padding=(10, 0, 0, 0))
        main.add(right, weight=2)
        self._build_card(right)

        # Pradinė riba: sąrašas ~56 %, kortelė ~44 % lango pločio
        # (kitaip platus sąrašas suspaudžia kortelę iki plyšio)
        def _pradine_riba(e=None):
            w = main.winfo_width()
            if w < 400:          # langas dar neatsidarė — palaukiame
                return
            main.unbind("<Configure>")
            try:
                main.sashpos(0, int(w * 0.56))
            except tk.TclError:
                pass
        main.bind("<Configure>", _pradine_riba)

        self.root.bind_all("<Control-s>", lambda e: self._save_card())
        self.root.bind_all("<Control-S>", lambda e: self._save_card())

    def _build_card(self, right):
        """Dešinė pusė: užsakymo kortelė, kurią galima taisyti."""
        self.card_empty = ttk.Label(right, anchor="center", justify="center",
                                    foreground="#8a8a8a", font=("Segoe UI", 11),
                                    text="Pasirinkite užsakymą sąraše kairėje.\n\n"
                                         "Čia matysite ir galėsite pataisyti\n"
                                         "kliento duomenis ir prekes.")
        self.card = ttk.Frame(right)

        # Antraštė: kodas + trumpa informacija; šalinimas — atokiau nuo kasdienių mygtukų
        hdr = ttk.Frame(self.card)
        hdr.pack(fill="x")
        self.lbl_kodas = ttk.Label(hdr, font=("Segoe UI", 15, "bold"))
        self.lbl_kodas.pack(side="left")
        self.lbl_info = ttk.Label(hdr, foreground="#6b7280")
        self.lbl_info.pack(side="left", padx=(10, 0), pady=(5, 0))
        b = _btn(hdr, "🗑", self._delete_order, "danger-link", width=3)
        b.pack(side="right")
        Tooltip(b, "Šalinti užsakymą iš sąrašo (laiškas pašte liks)")

        # AI įspėjimas — ryški raudona juosta, rodoma tik kai yra
        self.warn_lbl = tk.Label(self.card, justify="left", anchor="w", padx=10, pady=6,
                                 font=("Segoe UI", 10, "bold"), wraplength=520, **_tkkw())
        self.warn_lbl.configure(background="#fde2e2", foreground="#9b1c1c")

        # Veiksmai su laišku + statusas
        acts = ttk.Frame(self.card)
        acts.pack(fill="x", pady=(8, 6))
        self._acts = acts
        for text, cmd, tip in [
                ("✉ Laiškas", self._open_email, "Atidaryti pradinį laišką"),
                ("📎 Priedai", self._open_attachments, "Atsisiųsti ir atidaryti priedus"),
                ("📨 Rašyti tiekėjui", self._write_supplier,
                 "Paruošti užklausos laišką tiekėjui su šio užsakymo prekėmis"),
                ("📮 Užklausos", self._inquiries_window, "Kam išsiųsta užklausa ir kas atsakė")]:
            b = _btn(acts, text, cmd, "secondary-outline")
            b.pack(side="left", padx=(0, 4))
            Tooltip(b, tip)
        stat = ttk.Frame(self.card)
        stat.pack(fill="x", pady=(0, 6))
        ttk.Label(stat, text="Statusas").pack(side="left")
        self.status_combo = ttk.Combobox(stat, state="disabled", width=12)
        self.status_combo.pack(side="left", padx=(6, 14))
        self.status_combo.bind("<<ComboboxSelected>>", self._on_status_change)
        ttk.Label(stat, text="Sąskaita").pack(side="left")
        self.invoice_combo = ttk.Combobox(stat, state="disabled", width=10,
                                          values=SASKAITOS_PASIRINKIMAI)
        self.invoice_combo.pack(side="left", padx=(6, 0))
        self.invoice_combo.bind("<<ComboboxSelected>>", self._on_invoice_change)
        ttk.Label(stat, text="(išsaugoma iškart)", foreground="#8a8a8a").pack(side="left",
                                                                              padx=(8, 0))

        # Vertikaliai: viršuje duomenys + prekės, apačioje pastabos (riba tempiama)
        split = ttk.PanedWindow(self.card, orient="vertical")
        split.pack(fill="both", expand=True)
        top = ttk.Frame(split)
        split.add(top, weight=4)

        form = ttk.Labelframe(top, text=" Kliento duomenys ", padding=(10, 6))
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)
        self.card_vars = {}
        for key, label, r, c, wide in self.KORTELES_LAUKAI:
            ttk.Label(form, text=label).grid(row=r, column=c * 2, sticky="w",
                                             padx=(0 if c == 0 else 12, 6), pady=2)
            var = tk.StringVar()
            ent = ttk.Entry(form, textvariable=var)
            ent.grid(row=r, column=c * 2 + 1, columnspan=3 if wide else 1,
                     sticky="we", pady=2)
            var.trace_add("write", lambda *a: self._mark_dirty())
            self.card_vars[key] = var
            if key == "sutartis":
                Tooltip(ent, "Užsakymo kodas nesikeičia, net jei pataisote sutartį.")
            elif key == "pasiulymo_terminas":
                Tooltip(ent, "Formatas: 2026-09-01 12:00 (arba tik data)")
        self.lbl_tiekejai = ttk.Label(form, foreground="#6b7280", wraplength=520,
                                      justify="left")
        self.lbl_tiekejai.grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self.lbl_tiekejai.grid_remove()

        items_fr = ttk.Labelframe(top, text=" Prekės ", padding=(8, 6))
        items_fr.pack(fill="both", expand=True, pady=(8, 0))
        self.items_editor = ItemsEditor(items_fr, on_change=self._mark_dirty, height=6)
        self.items_editor.pack(fill="both", expand=True)

        savebar = ttk.Frame(top)
        savebar.pack(side="bottom", fill="x", pady=(6, 4), before=items_fr)
        self.btn_save = _btn(savebar, "💾  Išsaugoti pakeitimus", self._save_card, "success")
        self.btn_save.pack(side="left")
        Tooltip(self.btn_save, "Arba Ctrl+S")
        self.btn_revert = _btn(savebar, "↶ Atšaukti", self._revert_card, "secondary-link")
        self.btn_revert.pack(side="left", padx=(6, 0))
        self.dirty_var = tk.StringVar()
        ttk.Label(savebar, textvariable=self.dirty_var, foreground="#c2410c").pack(
            side="left", padx=(10, 0))

        # Apačioje skirtukai: mano pastabos | kas buvo laiške
        nb = ttk.Notebook(split)
        split.add(nb, weight=1)
        notes_tab = ttk.Frame(nb, padding=(0, 6, 0, 0))
        nb.add(notes_tab, text="  📝 Mano pastabos  ")
        ttk.Label(notes_tab, text="Priminimai sau — išsisaugo patys, nieko spausti nereikia.",
                  foreground="#8a8a8a").pack(anchor="w")
        self.notes = tk.Text(notes_tab, wrap="word", height=4, font=("Segoe UI", 10),
                             relief="flat", padx=8, pady=6, **_tkkw())
        self.notes.configure(background="#fffbe6", foreground="#1f2937",
                             insertbackground="#1f2937")
        self.notes.pack(fill="both", expand=True, pady=(2, 0))
        self.notes.bind("<KeyRelease>", self._notes_changed)
        self.notes.bind("<FocusOut>", lambda e: self._save_notes())

        info_tab = ttk.Frame(nb, padding=(0, 6, 0, 0))
        nb.add(info_tab, text="  ✉ Iš laiško  ")
        self.info_text = tk.Text(info_tab, wrap="word", height=4, font=("Segoe UI", 10),
                                 relief="flat", padx=8, pady=6, state="disabled", **_tkkw())
        self.info_text.configure(background="#f8fafc", foreground="#1f2937")
        self.info_text.pack(fill="both", expand=True)

        self._set_clean()
        self._show_card(False)

    def _show_card(self, show: bool):
        if show:
            self.card_empty.pack_forget()
            self.card.pack(fill="both", expand=True)
        else:
            self.card.pack_forget()
            self.card_empty.pack(fill="both", expand=True)

    # --- sąrašo pagalbininkai ---

    def _on_tree_motion(self, e):
        """Paaiškinimai: stulpelių antraštės ir 🔔 stulpelio ženkliukai."""
        region = self.tree.identify_region(e.x, e.y)
        col_id = self.tree.identify_column(e.x)
        try:
            col = self.tree["columns"][int(col_id[1:]) - 1]
        except (ValueError, IndexError):
            col = None
        text, key = "", None
        if region == "heading" and col in self.STULPELIU_PAAISKINIMAI:
            key, text = ("h", col), self.STULPELIU_PAAISKINIMAI[col]
        elif region == "cell" and col == "atn":
            iid = self.tree.identify_row(e.y)
            val = self.tree.set(iid, "atn") if iid else ""
            dalys = []
            if "⚠" in val:
                dalys.append("⚠ AI įspėjimas — atidarykite ir perskaitykite")
            if "🔔" in val:
                dalys.append("🔔 grandinėje atėjo naujas laiškas")
            if "❗" in val:
                dalys.append("❗ atsakė tiekėjas")
            key, text = ("c", iid), "\n".join(dalys)
        elif region == "cell" and col == "prekes":
            iid = self.tree.identify_row(e.y)
            val = self.tree.set(iid, "prekes") if iid else ""
            if len(val) > 40:
                key, text = ("p", iid), val.replace("; ", "\n")
        if key == self._tree_tip_key:
            return
        self._tree_tip_hide()
        if text:
            self._tree_tip_key = key
            self._tree_tip.show(e.x_root + 14, e.y_root + 18, text)

    def _tree_tip_hide(self):
        self._tree_tip_key = None
        self._tree_tip.hide()

    def _on_tree_right_click(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="✉  Atidaryti laišką", command=self._open_email)
        m.add_command(label="📎  Priedai", command=self._open_attachments)
        m.add_command(label="📨  Rašyti tiekėjui", command=self._write_supplier)
        m.add_command(label="📮  Užklausos", command=self._inquiries_window)
        m.add_separator()
        m.add_command(label="🗑  Šalinti užsakymą…", command=self._delete_order)
        m.tk_popup(e.x_root, e.y_root)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "status":
                    self.status_var.set(payload)
                elif kind == "refresh":
                    self._refresh_orders()
                    if self._sut_win is not None and self._sut_win.winfo_exists():
                        self._sut_refresh()
                elif kind == "check_state":
                    self.btn_check.config(state="normal" if not payload else "disabled")
                elif kind == "bell":
                    self.root.bell()
                elif kind == "error":
                    messagebox.showerror("Klaida", payload)
        except queue.Empty:
            pass
        # jei programa veikia per naktį — pasikeitus datai darbai persikelia patys
        today = datetime.date.today()
        if today != self._carry_date:
            self._carry_date = today
            storage.carry_over_tasks()
            if self._day_win is not None and self._day_win.winfo_exists():
                self.day_var.set(today.isoformat())
                self._day_refresh()
        self.root.after(200, self._poll_queue)

    # --- filtrai ---

    def _refresh_later(self, ms: int = 300):
        """Paieška / datos rašant: sąrašas atsinaujina, kai trumpam nustojama rašyti."""
        if self._refresh_after:
            self.root.after_cancel(self._refresh_after)
        self._refresh_after = self.root.after(ms, self._refresh_orders)

    def _on_filter_pick(self, _e=None):
        i = self.filter_combo.current()
        if 0 <= i < len(self.FILTRAI):
            self.filter_var.set(self.FILTRAI[i][0])
        self._refresh_orders()

    def _on_period_pick(self, _e=None):
        txt = self.period_combo.get()
        days = dict(self.LAIKOTARPIAI).get(txt)
        if days == "x":   # „Pasirinktos datos" — žmogus įrašys / pasirinks pats
            return
        self._set_period(days)

    def _on_date_typed(self, *_a):
        if self._period_prog:
            return
        self.period_combo.set("Pasirinktos datos")
        self._refresh_later(500)

    def _set_period(self, days):
        """Greitieji laikotarpiai: šiandien / N dienų / šis mėnuo / visi."""
        today = datetime.date.today()
        if days is None:
            a, b = "", ""
        elif days == -1:
            a, b = today.replace(day=1).isoformat(), today.isoformat()
        elif days == 0:
            a, b = today.isoformat(), today.isoformat()
        else:
            a, b = (today - datetime.timedelta(days=days - 1)).isoformat(), today.isoformat()
        self._period_prog = True
        try:
            self.date_from.set(a)
            self.date_to.set(b)
        finally:
            self._period_prog = False
        self._refresh_orders()

    @staticmethod
    def _clean_date(var):
        v = (var.get() or "").strip()
        if not v:
            return None
        v = v.replace(".", "-").replace("/", "-")
        try:
            datetime.date.fromisoformat(v)
            return v
        except ValueError:
            return None

    def _current_orders(self, status_override=None):
        """Užsakymai pagal dabar pasirinktus filtrus (statusas, datos, paieška)."""
        sel_filter = status_override or (
            self.filter_var.get() if hasattr(self, "filter_var") else "visi")
        status = None if sel_filter == "visi" else sel_filter
        return storage.list_orders(status, self._clean_date(self.date_from),
                                   self._clean_date(self.date_to),
                                   (self.search_var.get() or "").strip() or None)

    def _refresh_orders(self):
        self._refresh_after = None
        selected = self.tree.selection()
        sel_id = selected[0] if selected else None
        for item in self.tree.get_children():
            self.tree.delete(item)
        orders = self._current_orders()
        for o in orders:
            try:
                items = json.loads(o.get("prekes_json") or "[]")
            except Exception:
                items = []
            prekes = prekes_mod.trumpas_sarasas(items, 4)
            n_atts = len([s for s in (o.get("priedai") or "").split(";") if s.strip()])
            status = o.get("statusas") or ""
            group = "p" if _is_priority_order(o) else "n"
            term_txt, term_iki = _terminas(o)
            # Raudonai — tik „!" užsakymai su dar nepraėjusiu terminu; praėjusius
            # nuspalviname pilkai, kad nemirgėtų raudona dėl seniai baigtų.
            zyme = ""
            if _is_priority_order(o) and term_txt:
                zyme = "p" if (term_iki and term_iki < datetime.datetime.now()) else "t"
            base = f"{group}_{status}" if status else ""
            if zyme:
                tags = (f"{base}_{zyme}",) if base else (zyme,)
            else:
                tags = (base,) if base else ()
            self.tree.insert("", "end", iid=str(o["id"]), tags=tags, values=(
                ("⚠" if o.get("ispejimas") else "")
                + ("🔔" if o.get("atnaujinta") else "")
                + ("❗" if o.get("tiekejas_atsake") else ""),
                o.get("kodas") or "", term_txt, o.get("received_at"),
                o.get("klientas") or "?",
                prekes, f"📎{n_atts}" if n_atts else "", status or "—",
                _saskaitos_zenklas(o)))
        if sel_id and self.tree.exists(sel_id):
            self.tree.selection_set(sel_id)

        # „Rodyti" sąrašas su kiekiais — iškart matosi, kiek kur yra
        values, cur = [], 0
        for i, (code, text) in enumerate(self.FILTRAI):
            n = len(orders) if code == self.filter_var.get() else len(
                self._current_orders(status_override=code))
            values.append(f"{text}  ({n})")
            if code == self.filter_var.get():
                cur = i
        self.filter_combo.config(values=values)
        self.filter_combo.current(cur)

        df, dt = (self.date_from.get() or "").strip(), (self.date_to.get() or "").strip()
        blogos = [d for d, var in ((df, self.date_from), (dt, self.date_to))
                  if d and not self._clean_date(var)]
        self.count_var.set(f"Rodoma: {len(orders)}")
        if blogos:
            self.status_var.set(f"Data neatpažinta: {', '.join(blogos)} "
                                "(rašykite 2026-10-25 arba pasirinkite kalendoriuje)")
        elif hasattr(self, "status_var"):
            period = f" ({df or '…'} – {dt or '…'})" if (df or dt) else ""
            self.status_var.set(f"Rodoma užsakymų: {len(orders)}{period}")

    def _selected_order(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return storage.get_order(int(sel[0]))

    # --- užsakymo kortelė ---

    def _on_select(self, _event=None):
        sel = self.tree.selection()
        new_id = int(sel[0]) if sel else None
        if self._card_dirty and self._card_id is not None:
            if new_id is None or new_id == self._card_id:
                return   # sąrašas persikrovė — pataisymų nenumetame
            kodas = self.lbl_kodas.cget("text")
            ans = messagebox.askyesnocancel(
                "Neišsaugoti pakeitimai",
                f"Užsakyme {kodas} yra neišsaugotų pataisymų.\n\n"
                "Išsaugoti juos prieš pereinant prie kito užsakymo?")
            if ans is None or (ans and not self._save_card()):
                if self.tree.exists(str(self._card_id)):
                    self.tree.selection_set(str(self._card_id))
                return
            if ans is False:
                self._set_clean()
        self._save_notes()  # ankstesnio užsakymo pastabos — prieš persijungiant
        self._load_card(new_id)

    def _card_snap(self, o: dict):
        return tuple(o.get(k) for k in storage.REDAGUOJAMI_LAUKAI) + (o.get("prekes_json"),)

    def _load_card(self, order_id):
        o = storage.get_order(order_id) if order_id is not None else None
        self._card_id = o["id"] if o else None
        self._notes_order_id = self._card_id
        if not o:
            self.status_combo.set("")
            self.status_combo.config(state="disabled")
            self.invoice_combo.set("")
            self.invoice_combo.config(state="disabled")
            self.notes.delete("1.0", "end")
            self.notes.edit_modified(False)
            self._set_clean()
            self._show_card(False)
            return
        self._show_card(True)
        self._card_loading = True
        try:
            values = PRIORITY_STATUSES if _is_priority_order(o) else NORMAL_STATUSES
            self.status_combo.config(state="readonly", values=["—"] + values)
            self.status_combo.set(o.get("statusas") or "—")
            self.invoice_combo.config(state="readonly")
            self.invoice_combo.set(_SASK_I_TEKSTA.get(int(o.get("saskaita") or 0), "—"))

            self.lbl_kodas.config(text=o.get("kodas") or "—")
            rankinis = (o.get("store_id") or "") == "manual"
            info = [f"gauta {o.get('received_at') or '—'}"]
            if rankinis:
                info.append("įvesta ranka")
            elif o.get("pasitikejimas") is not None:
                info.append(f"AI tikrumas {round((o.get('pasitikejimas') or 0) * 100)}%")
            self.lbl_info.config(text="  ·  ".join(info))

            warn = (o.get("ispejimas") or "").strip()
            if warn:
                self.warn_lbl.config(text=f"⚠  {warn}")
                self.warn_lbl.pack(fill="x", pady=(6, 0), before=self._acts)
            else:
                self.warn_lbl.pack_forget()

            for key, var in self.card_vars.items():
                var.set(o.get(key) or "")
            tiek = (o.get("uzklausti_tiekejai") or "").strip()
            self.lbl_tiekejai.config(text=f"Užklausti tiekėjai: {tiek}" if tiek else "")
            if tiek:
                self.lbl_tiekejai.grid()
            else:
                self.lbl_tiekejai.grid_remove()

            try:
                items = json.loads(o.get("prekes_json") or "[]")
            except Exception:
                items = []
            self.items_editor.set_items(items, editable=True)

            lines = []
            if o.get("santrauka"):
                lines += [f"Santrauka: {o['santrauka']}", ""]
            if o.get("pastabos"):
                lines += [f"Pastabos iš laiško: {o['pastabos']}", ""]
            if o.get("atnaujinimai"):
                lines += ["🔔 Atnaujinimai (grandinės papildymai):"]
                lines += ["   " + ln for ln in o["atnaujinimai"].splitlines()] + [""]
            lines += [f"Laiškas: „{o.get('subject')}“", f"Siuntėjas: {o.get('sender')}",
                      f"Priedai: {o.get('priedai') or 'nėra'}"]
            self.info_text.config(state="normal")
            self.info_text.delete("1.0", "end")
            self.info_text.insert("1.0", "\n".join(lines))
            self.info_text.config(state="disabled")

            self.notes.delete("1.0", "end")
            self.notes.insert("1.0", o.get("mano_pastabos") or "")
            self.notes.edit_modified(False)
            self._card_snapshot = self._card_snap(o)
        finally:
            self._card_loading = False
        self._set_clean()

        if o.get("atnaujinta") or o.get("tiekejas_atsake"):
            # peržiūrėjus — 🔔 ir ❗ nuimami (⚠ lieka visada)
            if o.get("atnaujinta"):
                storage.set_order_seen(o["id"])
            if o.get("tiekejas_atsake"):
                storage.clear_supplier_replied(o["id"])
            try:
                self.tree.set(str(o["id"]), "atn", "⚠" if warn else "")
            except tk.TclError:
                pass

    def _mark_dirty(self):
        if self._card_loading or self._card_id is None:
            return
        if not self._card_dirty:
            self._card_dirty = True
            self.btn_save.config(state="normal")
            self.btn_revert.config(state="normal")
            self.dirty_var.set("●  Yra neišsaugotų pakeitimų")

    def _set_clean(self):
        self._card_dirty = False
        self.btn_save.config(state="disabled")
        self.btn_revert.config(state="disabled")
        self.dirty_var.set("")

    def _revert_card(self):
        if self._card_id is not None:
            self._set_clean()
            self._load_card(self._card_id)
            self.status_var.set("Pakeitimai atšaukti.")

    def _save_card(self) -> bool:
        """Išsaugo kortelės pataisymus. Grąžina False, jei neišsaugota."""
        if self._card_id is None or not self._card_dirty:
            return True
        self.items_editor.commit_edit()
        o = storage.get_order(self._card_id)
        if o is None:
            messagebox.showerror("Klaida", "Šis užsakymas jau pašalintas.")
            self._set_clean()
            self._load_card(None)
            return False
        if self._card_snap(o) != self._card_snapshot:
            if not messagebox.askyesno(
                    "Užsakymas atsinaujino",
                    "Kol taisėte, šis užsakymas atsinaujino iš pašto (atėjo papildymas).\n\n"
                    "Taip — išsaugoti jūsų pataisymus (papildymas bus perrašytas).\n"
                    "Ne — nieko nesaugoti ir parodyti naujausią versiją.",
                    icon="warning"):
                self._set_clean()
                self._load_card(self._card_id)
                return False
        fields = {k: v.get() for k, v in self.card_vars.items()}
        storage.update_order_fields(self._card_id, fields, self.items_editor.get_items())
        kodas = o.get("kodas")
        self._set_clean()
        self._refresh_orders()
        self._load_card(self._card_id)
        self._publish_share()
        self.status_var.set(f"✔ {kodas}: pakeitimai išsaugoti.")
        return True

    # ---------------- Veiksmai ----------------

    def _check_now(self):
        if not self.cfg.get_api_key():
            messagebox.showwarning("Trūksta API rakto",
                                   "Pirmiausia nustatymuose įveskite Anthropic API raktą.")
            return
        self.check_now_event.set()

    def _toggle_auto(self):
        self.cfg["auto_check"] = self.auto_var.get()
        self.cfg.save()

    def _open_email(self):
        o = self._selected_order()
        if not o:
            return
        if (o.get("store_id") or "") == "manual":
            messagebox.showinfo("Rankinis užsakymas",
                                "Šis užsakymas sukurtas ranka — laiško nėra.")
            return
        try:
            if (o.get("store_id") or "") == "graph":
                # pradinis laiškas naršyklėje — ten galima iškart atrašyti klientui
                graph_client.open_email(o["entry_id"])
            else:
                outlook_client.open_email(o["entry_id"], o["store_id"])
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko atidaryti laiško:\n{e}")

    def _attachments_dir(self, kodas: str) -> str:
        safe = re.sub(r'[<>:"/\\|?*]', "_", kodas or "be_kodo")
        return os.path.join(TEMP_ATT_ROOT, safe)

    def _relevant_attachments(self, attachments):
        """Priedai be parašų šiukšlių (logotipų, mažų nuotraukų)."""
        min_kb = int(self.cfg["min_image_kb"])
        return [(n, d, i) for n, d, i in (attachments or [])
                if not attachments_mod.is_noise_attachment(n, d, i, min_kb)]

    def _open_attachments(self):
        """Paima priedus iš Outlook, iškleidžia į laikiną vietą ir atidaro.
        Diske nuolat niekas nesaugoma — laikina vieta išvaloma paleidžiant programą."""
        o = self._selected_order()
        if not o:
            return
        self.status_var.set("Imami priedai iš Outlook…")
        self.root.update_idletasks()
        try:
            atts = self._relevant_attachments(
                graph_client.get_attachments(o["entry_id"])
                if (o.get("store_id") or "") == "graph"
                else outlook_client.get_attachments(o["entry_id"], o["store_id"]))
        except Exception as e:
            self.status_var.set("Paruošta.")
            messagebox.showerror("Klaida", f"Nepavyko paimti priedų iš Outlook:\n{e}")
            return
        if not atts:
            self.status_var.set("Paruošta.")
            messagebox.showinfo("Priedai", "Šis laiškas priedų (be parašo logotipų) neturi.")
            return
        if not o.get("priedai"):
            storage.set_order_priedai(o["id"], "; ".join(n for n, _, _ in atts))
            self._refresh_orders()
        folder = self._attachments_dir(o.get("kodas") or "")
        shutil.rmtree(folder, ignore_errors=True)
        paths = []
        for idx, (name, data, _inline) in enumerate(atts, 1):
            os.makedirs(folder, exist_ok=True)
            safe = re.sub(r'[<>:"/\\|?*]', "_", name)[:120]
            path = os.path.join(folder, safe)
            if os.path.exists(path):
                path = os.path.join(folder, f"{idx}_{safe}")
            try:
                with open(path, "wb") as f:
                    f.write(data)
                paths.append(path)
            except Exception:
                pass
        self.status_var.set(f"Priedai paruošti: {len(paths)}")
        if len(paths) == 1:
            os.startfile(paths[0])
        elif paths:
            os.startfile(folder)

    def _publish_share(self):
        """Atnaujina bendrinamą aplanką vadovo stebėjimui (jei nustatytas)."""
        share_dir = (self.cfg.get("share_dir") or "").strip()
        if not share_dir:
            return
        try:
            share.publish(share_dir)
        except Exception as e:
            self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} "
                      f"Stebėjimo sinchronizacija nepavyko: {e}")

    def _on_status_change(self, _event=None):
        o = self._selected_order()
        if not o:
            return
        val = self.status_combo.get()
        storage.set_order_status(o["id"], "" if val == "—" else val)
        self._refresh_orders()
        self._publish_share()

    def _on_invoice_change(self, _event=None):
        o = self._selected_order()
        if not o:
            return
        val = _SASK_I_REIKSME.get(self.invoice_combo.get(), 0)
        storage.set_order_invoice(o["id"], val)
        tekstai = {0: "žymė nuimta", 1: "išrašyta", 2: "dalinai išrašyta"}
        self.status_var.set(f"{o.get('kodas')}: sąskaita — {tekstai[val]}.")
        self._refresh_orders()
        self._publish_share()

    def _notes_changed(self, _event=None):
        """Rašant pastabas — išsaugoma pačiai, nustojus rašyti."""
        if self._notes_after:
            self.root.after_cancel(self._notes_after)
        self._notes_after = self.root.after(1500, self._save_notes)

    def _save_notes(self):
        """Automatinis pastabų išsaugojimas: rašant, persijungiant, uždarant."""
        if self._notes_after:
            self.root.after_cancel(self._notes_after)
            self._notes_after = None
        if self._notes_order_id is None:
            return
        try:
            if not self.notes.edit_modified():
                return
            text = self.notes.get("1.0", "end-1c").strip()
        except tk.TclError:
            return
        storage.set_order_notes(self._notes_order_id, text)
        self.notes.edit_modified(False)
        self.status_var.set("Pastabos išsaugotos.")
        self._publish_share()

    def _on_close(self):
        """Uždarant programą — neišsaugotos pastabos ir darbų eiga įrašomos."""
        if self._card_dirty:
            ans = messagebox.askyesnocancel(
                "Neišsaugoti pakeitimai",
                "Užsakymo kortelėje yra neišsaugotų pataisymų.\n\nIšsaugoti prieš uždarant?")
            if ans is None or (ans and not self._save_card()):
                return
        try:
            self._save_notes()
            if self._day_win is not None and self._day_win.winfo_exists():
                self._day_save_notes()
        except Exception:
            pass
        self.root.destroy()

    def _new_order_dialog(self):
        """Rankinis užsakymo sukūrimas: kodas generuojasi pats iš vardo ir datos,
        prekės — ta pati lentelė kaip kortelėje."""
        win = tk.Toplevel(self.root)
        win.title("Naujas užsakymas (rankinis)")
        win.grab_set()
        win.minsize(860, 0)
        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)
        # mygtukai — apačioje visada matomi, net jei langas mažas
        mygt = ttk.Frame(frm)
        mygt.pack(side="bottom", fill="x", pady=(12, 0))

        kodo_var = tk.StringVar(value="Kodas: —")
        ttk.Label(frm, textvariable=kodo_var, font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(frm, text="Kodas susidaro pats iš kliento vardo ir datos.",
                  foreground="#6b7280").pack(anchor="w", pady=(0, 8))

        form = ttk.Labelframe(frm, text=" Kliento duomenys ", padding=(10, 6))
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        form.columnconfigure(3, weight=1)
        v = {}
        laukai = [  # (raktas, pavadinimas, eilutė, stulpelis, per visą plotį)
            ("klientas", "Klientas *", 0, 0, False), ("kliento_imone", "Įmonė", 0, 1, False),
            ("kliento_telefonas", "Telefonas", 1, 0, False),
            ("kliento_el_pastas", "El. paštas", 1, 1, False),
            ("kliento_adresas", "Adresas", 2, 0, True),
            ("uzsakymo_data", "Užsakymo data", 3, 0, False),
            ("sutartis", "Sutartis", 3, 1, False),
            ("pasiulymo_terminas", "Pasiūlyti iki", 4, 0, False),
            ("pristatymo_data", "Pristatymas", 4, 1, False),
        ]
        first = None
        for key, label, r, c, wide in laukai:
            ttk.Label(form, text=label).grid(row=r, column=c * 2, sticky="w",
                                             padx=(0 if c == 0 else 12, 6), pady=2)
            v[key] = tk.StringVar()
            if key == "uzsakymo_data":
                ent = _date_field(form, v[key])
            else:
                ent = ttk.Entry(form, textvariable=v[key])
            ent.grid(row=r, column=c * 2 + 1, columnspan=3 if wide else 1, sticky="we", pady=2)
            if key == "sutartis":
                Tooltip(ent, "Prioritetinės sutarties užsakymo kodas gaus „!“")
            elif key == "pasiulymo_terminas":
                Tooltip(ent, "Formatas: 2026-09-01 12:00 (arba tik data)")
            first = first or ent
        v["uzsakymo_data"].set(datetime.date.today().isoformat())

        items_fr = ttk.Labelframe(frm, text=" Prekės ", padding=(8, 6))
        items_fr.pack(fill="both", expand=True, pady=(8, 0))
        editor = ItemsEditor(items_fr, height=5)
        editor.pack(fill="both", expand=True)
        editor.set_items([], editable=True)

        ttk.Label(frm, text="Mano pastabos").pack(anchor="w", pady=(8, 2))
        past_txt = tk.Text(frm, height=3, font=("Segoe UI", 10), relief="flat",
                           padx=8, pady=6, **_tkkw())
        past_txt.configure(background="#fffbe6", foreground="#1f2937")
        past_txt.pack(fill="x")

        def perziureti_koda(*_a):
            try:
                d = datetime.datetime.strptime(
                    v["uzsakymo_data"].get().strip(), "%Y-%m-%d")
            except ValueError:
                d = datetime.datetime.now()
            prior = "!" if storage._is_priority(
                {"sutartis": v["sutartis"].get()}) else ""
            kodo_var.set(f"Kodas: {prior}{storage.KODO_PREFIKSAS}"
                         f"{storage._initials(v['klientas'].get())}{d:%y%m%d}")
        for lk in ("klientas", "uzsakymo_data", "sutartis"):
            v[lk].trace_add("write", perziureti_koda)
        perziureti_koda()

        def issaugoti():
            if not v["klientas"].get().strip():
                messagebox.showwarning("Trūksta", "Įrašykite kliento vardą — iš jo "
                                       "kuriamas kodas.", parent=win)
                return
            data = {k: var.get() for k, var in v.items()}
            data["prekes"] = editor.get_items()
            data["mano_pastabos"] = past_txt.get("1.0", "end-1c").strip()
            kodas = storage.add_manual_order(data)
            win.destroy()
            self._refresh_orders()
            self._publish_share()
            self.status_var.set(f"✔ Sukurtas rankinis užsakymas {kodas}.")

        _btn(mygt, "Atšaukti", win.destroy, "secondary-link").pack(side="right")
        _btn(mygt, "💾  Sukurti užsakymą", issaugoti, "success").pack(side="right", padx=6)
        if first is not None:
            first.focus_set()

    def _write_supplier(self):
        """Sukuria Outlook JUODRAŠTĮ tiekėjui: tema „Užklausa <kodas>", tekste —
        prekės. Gavėją įrašote ir siunčiate patys; programa pagal kodą temoje
        pati atseka, kam išsiųsta ir kas atsakė."""
        o = self._selected_order()
        if not o:
            messagebox.showinfo("Pasirinkite", "Pirma pasirinkite užsakymą sąraše.")
            return
        try:
            items = json.loads(o.get("prekes_json") or "[]")
        except Exception:
            items = []
        eilutes = []
        for i, p in enumerate(items, 1):
            eil = f"{i}. {prekes_mod.pilnas_pavadinimas(p)} — {prekes_mod.kiekis_tekstu(p)}"
            if p.get("nuoroda"):
                eil += f"\n   {p['nuoroda']}"
            eilutes.append(eil)
        subject = f"Užklausa {o.get('kodas')}"
        body = ("Sveiki,\n\n"
                "reikėtų pasiūlymo:\n\n"
                + "\n".join(eilutes) + "\n")
        try:
            import urllib.parse
            mailto = (f"mailto:?subject={urllib.parse.quote(subject)}"
                      f"&body={urllib.parse.quote(body)}")
            if len(mailto) < 1900:
                # atidaro NUMATYTĄJĄ pašto programą (naują Outlook, jei jis numatytas)
                os.startfile(mailto)
            else:
                # per ilgas mailto nuorodai (daug prekių) — klasikinio Outlook langas
                outlook_client.create_draft_with_attachment(subject, body, None,
                                                            display=True)
            self.status_var.set(f"Atidarytas laiškas „{subject}“ — "
                                "įrašykite gavėją ir siųskite.")
        except Exception as e:
            messagebox.showerror("Klaida", f"Laiško sukurti nepavyko:\n{e}")

    def _inquiries_window(self):
        """Pasirinkto užsakymo užklausų langas: tiekėjų sąrašas su būsena;
        dvigubas paspaudimas atidaro TO tiekėjo susirašinėjimo grandinę."""
        o = self._selected_order()
        if not o:
            messagebox.showinfo("Pasirinkite", "Pirma pasirinkite užsakymą sąraše.")
            return
        if self.cfg["mail_source"] != "graph":
            messagebox.showinfo("Užklausos", "Veikia tik Graph režimu.")
            return
        kodas = o.get("kodas") or ""
        self.status_var.set(f"Ieškoma laiškų su {kodas}…")
        self.root.update_idletasks()
        try:
            laiskai = graph_client.find_by_code(kodas)
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko gauti laiškų:\n{e}")
            return
        self.status_var.set("Paruošta.")
        if not laiskai:
            messagebox.showinfo("Užklausos", f"Laiškų su kodu {kodas} temoje nerasta.\n"
                                "(Užklausa dar neišsiųsta arba kodas temoje pakeistas.)")
            return

        # grupuojame pagal tiekėją: siųstiems — gavėjai, gautiems — siuntėjas
        grupes = {}
        for m in laiskai:
            adresai = ([a.strip() for a in m["kas"].split(",")]
                       if m["kryptis"] == "išsiųsta" else [m["kas"]])
            for a in adresai:
                if not a:
                    continue
                g = grupes.setdefault(a.lower(), {"kas": a, "n": 0, "paskutinis": None,
                                                  "atsake": False, "link": None,
                                                  "link_gauto": None})
                g["n"] += 1
                if m["kada"] and (g["paskutinis"] is None or m["kada"] > g["paskutinis"]):
                    g["paskutinis"] = m["kada"]
                    g["link"] = m["webLink"]
                if m["kryptis"] == "gauta":
                    g["atsake"] = True
                    if g["link_gauto"] is None:
                        g["link_gauto"] = m["webLink"]

        win = tk.Toplevel(self.root)
        win.title(f"Užklausos — {kodas}")
        win.geometry("640x320")
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=f"{kodas}: rasta laiškų — {len(laiskai)}. "
                  "Dvigubas paspaudimas atidaro tiekėjo grandinę naršyklėje.",
                  foreground="#666").pack(anchor="w", pady=(0, 6))
        tree = ttk.Treeview(frm, columns=("kas", "n", "pask", "busena"),
                            show="headings", selectmode="browse")
        for c, t, w in (("kas", "Tiekėjas", 280), ("n", "Laiškų", 60),
                        ("pask", "Paskutinis", 120), ("busena", "Būsena", 120)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="center" if c == "n" else "w")
        tree.pack(fill="both", expand=True)
        nuorodos = {}
        for g in sorted(grupes.values(), key=lambda x: x["paskutinis"] or
                        datetime.datetime.min, reverse=True):
            iid = tree.insert("", "end", values=(
                g["kas"], g["n"],
                f"{g['paskutinis']:%m-%d %H:%M}" if g["paskutinis"] else "—",
                "✅ atsakė" if g["atsake"] else "⏳ laukiama"))
            # atidaroma NAUJAUSIA TIEKĖJO žinutė (ne mūsų); kol neatsakė —
            # mūsų išsiųsta užklausa
            nuorodos[iid] = g["link_gauto"] or g["link"]

        def atidaryti(_e=None):
            sel = tree.selection()
            if sel and nuorodos.get(sel[0]):
                os.startfile(nuorodos[sel[0]])
        tree.bind("<Double-1>", atidaryti)
        ttk.Button(frm, text="✉ Atidaryti naujausią atsakymą",
                   command=atidaryti).pack(anchor="e", pady=(8, 0))

    def _chain_window(self, kodas: str, tiekejas: str = None):
        """Visa grandinė su tiekėju — laiškų tekstai pačioje programoje."""
        self.status_var.set(f"Kraunama {kodas} grandinė…")
        self.root.update_idletasks()
        try:
            visi = graph_client.chain_messages(kodas)
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko įkelti grandinės:\n{e}")
            return
        self.status_var.set("Paruošta.")
        t = (tiekejas or "").lower()
        if t:
            zinutes = [m for m in visi
                       if m["nuo"].lower() == t
                       or any(a.lower() == t for a in m["kam"])]
        else:
            zinutes = visi
        if not zinutes:
            messagebox.showinfo("Grandinė", "Laiškų su šiuo tiekėju nerasta.")
            return
        self._rodyti_grandine(
            f"Grandinė — {kodas}" + (f" · {tiekejas}" if tiekejas else ""), zinutes)

    def _rodyti_grandine(self, pavadinimas: str, zinutes: list):
        """Bendras grandinės langas: laiškai chronologiškai su pilnais tekstais."""
        win = tk.Toplevel(self.root)
        win.title(pavadinimas)
        win.geometry("860x600")
        frm = ttk.Frame(win, padding=8)
        frm.pack(fill="both", expand=True)
        txt = tk.Text(frm, wrap="word", font=("Segoe UI", 10))
        vsb = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        txt.tag_configure("galva_mes", foreground="#1f3a5f",
                          font=("Segoe UI", 10, "bold"))
        txt.tag_configure("galva_jie", foreground="#7a3b00",
                          font=("Segoe UI", 10, "bold"))
        txt.tag_configure("atskyrimas", foreground="#bbb")
        savas = graph_client.LOGIN_HINT.lower()
        for m in zinutes:
            mes = m["nuo"].lower() == savas
            galva = (f"{m['kada']:%m-%d %H:%M}  "
                     + ("JŪS → " + ", ".join(m["kam"][:3]) if mes
                        else (m["nuo_vardas"] or m["nuo"])))
            txt.insert("end", "─" * 90 + "\n", "atskyrimas")
            txt.insert("end", galva + "\n\n", "galva_mes" if mes else "galva_jie")
            txt.insert("end", (m["tekstas"] or "(tuščias)").strip() + "\n\n")
        txt.config(state="disabled")
        apacia = ttk.Frame(win, padding=(8, 4))
        apacia.pack(fill="x")
        paskutine = zinutes[-1]
        if paskutine.get("webLink"):
            ttk.Button(apacia, text="🌐 Atidaryti naršyklėje",
                       command=lambda: os.startfile(paskutine["webLink"])
                       ).pack(side="right")

    def _track_inquiries(self, since):
        """Užklausų sekimas pagal kodą temoje (tik Graph režimu):
        - išsiųstas laiškas su kodu -> gavėjai prirašomi, statusas ''→užklausta;
        - gautas laiškas su kodu -> ❗ ir tiekėjo žyma ✓."""
        if self.cfg["mail_source"] != "graph":
            return 0
        orders = storage.orders_inquiry_tracking()
        if not orders:
            return 0
        pakeista = 0
        siusti = graph_client.list_sent_since(since)
        for m in siusti:
            if storage.is_processed(m["entry_id"]):
                continue
            for o in orders:
                if o["kodas"] and o["kodas"] in m["subject"]:
                    storage.inquiry_sent(o["id"], m["to"])
                    storage.mark_processed_simple(
                        {"entry_id": m["entry_id"], "subject": m["subject"],
                         "received": m["sent"]}, note=f"užklausa {o['kodas']}")
                    self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Užklausa "
                              f"{o['kodas']} išsiųsta: {', '.join(m['to'][:4])}")
                    pakeista += 1
        gauti = graph_client.list_emails_since(since, max_count=200, sender_filters=None)
        for m in gauti:
            if storage.is_processed(m["entry_id"]):
                continue
            for o in orders:
                if o["kodas"] and o["kodas"] in m["subject"]:
                    storage.supplier_replied(o["id"], m.get("sender_email") or "")
                    storage.mark_processed_simple(
                        {"entry_id": m["entry_id"], "subject": m["subject"],
                         "received": m["received"]},
                        note=f"tiekėjo atsakymas {o['kodas']}")
                    self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} ❗ Tiekėjas "
                              f"atsakė dėl {o['kodas']}: {m.get('sender_email')}")
                    pakeista += 1
        return pakeista

    def _delete_order(self):
        o = self._selected_order()
        if not o:
            return
        if messagebox.askyesno("Šalinti", "Pašalinti šį užsakymą iš sąrašo?\n"
                               "(Pats laiškas Outlook'e liks nepaliestas.)"):
            if self._notes_order_id == o["id"]:
                self._notes_order_id = None  # kad autosave nerašytų į ištrintą
                self.notes.edit_modified(False)
            if self._card_id == o["id"]:
                self._set_clean()
                self._load_card(None)
            storage.delete_order(o["id"])
            self._refresh_orders()
            self._publish_share()

    def _export(self, visible_only: bool = False):
        orders = self._current_orders() if visible_only else storage.list_orders()
        if not orders:
            messagebox.showinfo("Eksportas", "Pagal pasirinktus filtrus užsakymų nėra.")
            return
        df = (self.date_from.get() or "").strip()
        dt = (self.date_to.get() or "").strip()
        if visible_only and (df or dt):
            name = f"uzsakymai_{df or 'pradzia'}_iki_{dt or 'dabar'}.xlsx"
        else:
            name = f"uzsakymai_{datetime.date.today():%Y-%m-%d}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=name)
        if not path:
            return
        try:
            exporter.export_orders(orders, path, day_tasks=storage.all_day_tasks())
            self.status_var.set(f"Eksportuota užsakymų: {len(orders)} → {path}")
        except PermissionError:
            messagebox.showerror("Klaida", "Failas atidarytas Excel programoje — uždarykite jį ir bandykite dar kartą.")
        except Exception as e:
            messagebox.showerror("Klaida", f"Eksportuoti nepavyko:\n{e}")

    def _siusti_perziura(self):
        """Pasirinkimas, kokiu pavidalu siųsti tik skaitymo peržiūrą."""
        win = tk.Toplevel(self.root)
        win.title("Siųsti peržiūrą")
        win.grab_set()
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Kokiu pavidalu paruošti peržiūrą (tik skaitymui)?",
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=0,
                                                      columnspan=2, sticky="w")
        ttk.Label(frm, foreground="#666", justify="left", text=(
            "🌐 HTML — vienas mažas failas, atsidaro naršyklėje (ir telefone).\n"
            "🖥 Programa — atrodo lygiai kaip šis langas; ZIP su .exe ir\n"
            "     duomenimis (~13 MB, tik Windows; pirmą kartą paleidžiant\n"
            "     Windows gali paklausti „Ar tikrai paleisti?“).")
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 10))

        def html():
            win.destroy()
            self._html_perziura()

        def zipas():
            win.destroy()
            self._zip_perziura()

        ttk.Button(frm, text="🌐 HTML failas", command=html).grid(row=2, column=0, padx=(0, 6))
        ttk.Button(frm, text="🖥 Programa (ZIP)", command=zipas).grid(row=2, column=1)

    def _zip_perziura(self):
        """ZIP su tik skaitymo programa (.exe) ir šviežia duomenų kopija."""
        exe = os.path.join(cfg_mod.APP_DIR, "Uzsakymu perziura.exe")
        if not os.path.exists(exe):
            messagebox.showerror("Klaida", "Nerastas failas „Uzsakymu perziura.exe“ "
                                 "programos aplanke.")
            return
        name = f"uzsakymu_perziura_{datetime.date.today():%Y-%m-%d}.zip"
        path = filedialog.asksaveasfilename(
            defaultextension=".zip", filetypes=[("ZIP", "*.zip")], initialfile=name)
        if not path:
            return
        self.status_var.set("Ruošiamas peržiūros paketas…")
        self.root.update_idletasks()
        try:
            import zipfile
            tmp_db = os.path.join(tempfile.gettempdir(), "uzsakymai_kopija.db")
            src = sqlite3.connect(f"file:{storage.DB_PATH}?mode=ro", uri=True)
            dst = sqlite3.connect(tmp_db)
            src.backup(dst)   # vientisa kopija, net jei tuo metu rašoma
            dst.close()
            src.close()
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(exe, "Uzsakymu perziura.exe")
                z.write(tmp_db, "uzsakymai.db")
            os.remove(tmp_db)
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko paruošti:\n{e}")
            return
        dydis = os.path.getsize(path) // (1024 * 1024)
        self.status_var.set(f"Peržiūros paketas paruoštas ({dydis} MB): {path}")
        if messagebox.askyesno("Siųsti", "Sukurti Outlook juodraštį su šiuo ZIP?\n"
                               "(Gavėją įrašysite ir išsiųsite patys.)"):
            try:
                outlook_client.create_draft_with_attachment(
                    f"Užsakymų peržiūra {datetime.date.today():%Y-%m-%d}",
                    "Sveiki,\n\nprisegta užsakymų peržiūros programa: išskleiskite ZIP "
                    "į bet kurį aplanką ir paleiskite „Uzsakymu perziura.exe“.\n"
                    "Jei Windows paklaus — „More info“ → „Run anyway“.\n"
                    "Nieko diegti nereikia, redaguoti nieko negalima — tik peržiūra.\n",
                    path)
                self.status_var.set("Juodraštis sukurtas — rasite Outlook „Juodraščiai“.")
            except Exception as e:
                messagebox.showerror("Klaida", f"Juodraščio sukurti nepavyko:\n{e}")

    def _html_perziura(self):
        """Vieno failo interaktyvi peržiūra naršyklei — galima siųsti bet kam,
        gavėjui nieko diegti nereikia."""
        name = f"uzsakymai_perziura_{datetime.date.today():%Y-%m-%d}.html"
        path = filedialog.asksaveasfilename(
            defaultextension=".html", filetypes=[("HTML", "*.html")], initialfile=name)
        if not path:
            return
        try:
            kiek = html_perziura.generuoti(path)
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko sugeneruoti:\n{e}")
            return
        self.status_var.set(f"HTML peržiūra sukurta: {kiek} užsakymų → {path}")
        try:
            os.startfile(path)
        except Exception:
            pass
        if messagebox.askyesno("Siųsti", "Sukurti Outlook juodraštį su šiuo failu?\n"
                               "(Gavėją įrašysite ir išsiųsite patys.)"):
            try:
                outlook_client.create_draft_with_attachment(
                    f"Užsakymų peržiūra {datetime.date.today():%Y-%m-%d}",
                    "Sveiki,\n\nprisegta užsakymų peržiūra — atidarykite failą naršyklėje "
                    "(dukart spustelėkite). Veikia be interneto, nieko diegti nereikia.\n",
                    path)
                self.status_var.set("Juodraštis sukurtas — rasite Outlook „Juodraščiai“.")
            except Exception as e:
                messagebox.showerror("Klaida", f"Juodraščio sukurti nepavyko:\n{e}")

    # ---------------- Istorijos importas ----------------

    def _import_history(self):
        """Analizuoja senesnius laiškus, kurių programa dar nematė."""
        if not self.cfg.get_api_key():
            messagebox.showwarning("Trūksta API rakto", "Pirmiausia įveskite API raktą.")
            return
        win = tk.Toplevel(self.root)
        win.title("Importuoti senesnius užsakymus")
        win.grab_set()
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)

        today = datetime.date.today()
        ttk.Label(frm, text="Nuo dienos (imama nuo 7:30 iki dabar):").grid(
            row=0, column=0, sticky="w")
        v_from = tk.StringVar(value=today.isoformat())
        ttk.Entry(frm, textvariable=v_from, width=14).grid(row=0, column=1, sticky="w", pady=3)

        quick = ttk.Frame(frm)
        quick.grid(row=1, column=0, columnspan=3, sticky="w")
        for text, back in [("Šiandien", 0), ("Vakar", 1), ("Užvakar", 2), ("-3 d.", 3)]:
            ttk.Button(quick, text=text, width=9,
                       command=lambda b=back: v_from.set(
                           (today - datetime.timedelta(days=b)).isoformat())
                       ).pack(side="left", padx=2)

        info = ttk.Label(frm, text="Tikrinami tik laiškai, kurių programa dar nėra mačiusi\n"
                                   "(patikrintus ji žymisi duomenų bazėje). Pirma paspauskite\n"
                                   "„Suskaičiuoti“ — pamatysite kiekį ir apytikslę kainą.",
                         foreground="#666", justify="left")
        info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 4))

        state = {"emails": None}

        def parse(var):
            try:
                return datetime.date.fromisoformat(
                    var.get().strip().replace(".", "-").replace("/", "-"))
            except ValueError:
                return None

        def count():
            d1 = parse(v_from)
            if not d1 or d1 > today:
                info.config(text="Neteisinga data (formatas: 2026-08-05).")
                return
            since = datetime.datetime.combine(d1, datetime.time(7, 30))
            info.config(text=f"Skaičiuojama: nuo {since:%m-%d %H:%M} iki dabar…")
            win.update_idletasks()
            try:
                emails = self._mail().fetch_emails_since(
                    since, max_count=1000, skip=storage.is_processed,
                    sender_filters=self.cfg["sender_filters"],
                    on_skip=lambda subj, err: self._log(
                        f"{datetime.datetime.now():%Y-%m-%d %H:%M} Laiško nepavyko "
                        f"perskaityti (importas): „{subj}“ — {err}"))
            except Exception as e:
                info.config(text=f"Klaida skaitant Outlook: {e}")
                return
            state["emails"] = emails
            if not emails:
                info.config(text=f"Nuo {since:%m-%d %H:%M} nepatikrintų laiškų nėra —\n"
                                 "viskas jau peržiūrėta.")
                btn_go.config(state="disabled")
                return
            info.config(text=f"Nuo {since:%m-%d 7:30} iki dabar rasta laiškų: {len(emails)}\n"
                             f"Apytikslė kaina: ~{len(emails) * 0.9:.0f} ct "
                             f"(~{len(emails) * 0.009:.2f} €)\n"
                             f"Užsakymais taps tik dalis jų.")
            btn_go.config(state="normal")

        def go():
            emails = state.get("emails")
            if not emails:
                return
            win.destroy()
            threading.Thread(target=self._run_import, args=(emails,), daemon=True).start()

        ttk.Button(frm, text="Suskaičiuoti", command=count).grid(row=4, column=0, pady=(10, 0))
        btn_go = ttk.Button(frm, text="📥 Analizuoti", command=go, state="disabled")
        btn_go.grid(row=4, column=1, sticky="w", pady=(10, 0))
        ttk.Button(frm, text="Atšaukti", command=win.destroy).grid(row=4, column=2, pady=(10, 0))

    def _run_import(self, emails):
        """Ta pati analizė kaip tikrinant paštą, tik nurodytiems laiškams."""
        with self._analysis_lock:
            self.ui_queue.put(("check_state", True))
            try:
                found, done, spent, _failed, _sk, updated = self._analyze_emails(
                    emails, "Importuojama")
                upd_note = f", 🔔 papildyta: {updated}" if updated else ""
                self.ui_queue.put(("status",
                    f"✔ Importas baigtas: rasta užsakymų {found} iš {done} laiškų"
                    f"{upd_note}, kaina {spent:.1f} ct"))
                self._publish_share()
                self.ui_queue.put(("refresh", None))
            except Exception:
                self._log(traceback.format_exc(limit=3))
                self.ui_queue.put(("status", "Importas nutrūko — žr. zurnalas.txt"))
            finally:
                self.ui_queue.put(("check_state", False))

    # ---------------- Rankinis laiškų pasirinkimas ----------------

    def _open_picker(self):
        """Laiškų sąrašas su varnelėmis: analizuojami TIK pažymėti.

        Sąrašas paimamas be laiško teksto ir be priedų (greita, nemokama);
        priedai skaitomi tik pažymėtiems, prieš pat analizę."""
        if not self.cfg.get_api_key():
            messagebox.showwarning("Trūksta API rakto", "Pirmiausia įveskite API raktą.")
            return
        if getattr(self, "_pick_win", None) is not None and self._pick_win.winfo_exists():
            self._pick_win.lift()
            return

        win = tk.Toplevel(self.root)
        self._pick_win = win
        win.title("Rinktis laiškus analizei")
        win.geometry("1100x620")
        win.minsize(820, 420)

        self._pick_rows = {}      # iid -> laiško info
        self._pick_checked = set()

        top = ttk.Frame(win, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Nuo:").pack(side="left")
        self._pick_from = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(top, textvariable=self._pick_from, width=13).pack(side="left", padx=(4, 6))
        for text, back in (("Šiandien", 0), ("Vakar", 1), ("3 d.", 2), ("7 d.", 6)):
            ttk.Button(top, text=text, width=8,
                       command=lambda bb=back: (self._pick_from.set(
                           (datetime.date.today() - datetime.timedelta(days=bb)).isoformat()),
                           self._pick_load())).pack(side="left", padx=1)
        self._pick_only_new = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="tik dar neanalizuoti", variable=self._pick_only_new,
                        command=self._pick_load).pack(side="left", padx=(12, 0))
        self._pick_use_filter = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="tik pagal siuntėjų filtrą", variable=self._pick_use_filter,
                        command=self._pick_load).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="🔄 Užkrauti", command=self._pick_load).pack(side="left", padx=(12, 0))
        self._pick_info = ttk.Label(top, text="", foreground="#666")
        self._pick_info.pack(side="left", padx=(12, 0))

        body = ttk.Frame(win, padding=(10, 0))
        body.pack(fill="both", expand=True)
        cols = ("chk", "gauta", "siuntejas", "tema", "priedai", "busena")
        self._pick_tree = ttk.Treeview(body, columns=cols, show="headings",
                                       selectmode="extended")
        for c, t, w, mw in (("chk", "✔", 34, 34), ("gauta", "Gauta", 125, 110),
                            ("siuntejas", "Siuntėjas", 210, 120),
                            ("tema", "Tema", 430, 160),
                            ("priedai", "📎", 42, 42), ("busena", "Būsena", 130, 90)):
            self._pick_tree.heading(c, text=t)
            self._pick_tree.column(c, width=w, minwidth=mw, stretch=(c == "tema"),
                                   anchor="center" if c in ("chk", "priedai") else "w")
        self._pick_tree.tag_configure("apdorotas", foreground="#999")
        self._pick_tree.tag_configure("pazymeta", background="#d6ecff")
        self._pick_tree.tag_configure("apdorotas_pazymeta", background="#d6ecff",
                                      foreground="#999")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self._pick_tree.yview)
        self._pick_tree.configure(yscrollcommand=vsb.set)
        self._pick_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self._pick_tree.bind("<Button-1>", self._pick_click)
        self._pick_tree.bind("<space>", lambda e: self._pick_toggle_selected())

        bot = ttk.Frame(win, padding=(10, 8))
        bot.pack(fill="x")
        ttk.Button(bot, text="Žymėti visus",
                   command=lambda: self._pick_all(True)).pack(side="left")
        ttk.Button(bot, text="Nuimti",
                   command=lambda: self._pick_all(False)).pack(side="left", padx=4)
        ttk.Label(bot, text="  (paspaudimas ant eilutės arba tarpo klavišas — pažymi)",
                  foreground="#888").pack(side="left")
        self._pick_go = ttk.Button(bot, text="🔍 Analizuoti pažymėtus",
                                   command=self._pick_run, state="disabled")
        self._pick_go.pack(side="right")
        self._pick_sum = ttk.Label(bot, text="Pažymėta: 0", foreground="#666")
        self._pick_sum.pack(side="right", padx=10)

        self._pick_load()

    def _pick_load(self):
        try:
            d = datetime.date.fromisoformat(
                self._pick_from.get().strip().replace(".", "-").replace("/", "-"))
        except ValueError:
            self._pick_info.config(text="Neteisinga data (pvz., 2026-08-13).")
            return
        since = datetime.datetime.combine(d, datetime.time(0, 0))
        self._pick_info.config(text="Skaitoma iš Outlook…")
        self._pick_win.update_idletasks()
        try:
            laiskai = self._mail().list_emails_since(
                since, max_count=400,
                sender_filters=(self.cfg["sender_filters"]
                                if self._pick_use_filter.get() else None))
        except Exception as e:
            self._pick_info.config(text=f"Klaida skaitant Outlook: {e}")
            return
        self._pick_tree.delete(*self._pick_tree.get_children())
        self._pick_rows.clear()
        self._pick_checked.clear()
        n_apd = 0
        for e in laiskai:
            apdorotas = storage.is_processed(e["entry_id"])
            if apdorotas:
                n_apd += 1
                if self._pick_only_new.get():
                    continue
            iid = self._pick_tree.insert(
                "", "end", tags=("apdorotas",) if apdorotas else (),
                values=("☐", f"{e['received']:%m-%d %H:%M}",
                        e["sender_name"] or e["sender_email"], e["subject"],
                        f"📎{e['n_att']}" if e["n_att"] else "",
                        "jau analizuotas" if apdorotas else "naujas"))
            e["apdorotas"] = apdorotas
            self._pick_rows[iid] = e
        self._pick_info.config(
            text=f"Nuo {since:%m-%d} rasta: {len(laiskai)} · rodoma: {len(self._pick_rows)}"
                 + (f" · jau analizuotų: {n_apd}" if n_apd else ""))
        self._pick_update_sum()

    def _pick_click(self, event):
        iid = self._pick_tree.identify_row(event.y)
        if iid:
            self._pick_toggle(iid)

    def _pick_toggle_selected(self):
        for iid in self._pick_tree.selection():
            self._pick_toggle(iid)
        return "break"

    def _pick_toggle(self, iid):
        e = self._pick_rows.get(iid)
        if e is None:
            return
        apd = bool(e.get("apdorotas"))
        if iid in self._pick_checked:
            self._pick_checked.discard(iid)
            self._pick_tree.set(iid, "chk", "☐")
            self._pick_tree.item(iid, tags=("apdorotas",) if apd else ())
        else:
            self._pick_checked.add(iid)
            self._pick_tree.set(iid, "chk", "✔")
            self._pick_tree.item(iid, tags=("apdorotas_pazymeta",) if apd else ("pazymeta",))
        self._pick_update_sum()

    def _pick_all(self, on: bool):
        for iid in self._pick_tree.get_children():
            if (iid in self._pick_checked) != on:
                self._pick_toggle(iid)

    def _pick_update_sum(self):
        n = len(self._pick_checked)
        if not n:
            self._pick_sum.config(text="Pažymėta: 0")
            self._pick_go.config(state="disabled")
            return
        pak = sum(1 for i in self._pick_checked if self._pick_rows[i].get("apdorotas"))
        note = f", pakartotinai {pak}" if pak else ""
        self._pick_sum.config(text=f"Pažymėta: {n}{note} · apytiksliai ~{n * 2.0:.0f} ct")
        self._pick_go.config(state="normal")

    def _pick_run(self):
        pasirinkti = [self._pick_rows[i] for i in self._pick_checked]
        if not pasirinkti:
            return
        pak = [e for e in pasirinkti if e.get("apdorotas")]
        klausimas = (f"Analizuoti pažymėtus laiškus: {len(pasirinkti)}?\n"
                     f"Apytikslė kaina: ~{len(pasirinkti) * 2.0:.0f} ct.")
        if pak:
            klausimas += (f"\n\n{len(pak)} iš jų jau buvo analizuoti — žymė bus nuimta "
                          "ir laiškas apdorotas iš naujo.")
        if not messagebox.askyesno("Analizuoti", klausimas):
            return
        self._pick_win.destroy()
        self._pick_win = None
        threading.Thread(target=self._run_picked, args=(pasirinkti,), daemon=True).start()

    def _run_picked(self, pasirinkti):
        """Įkelia pažymėtų laiškų tekstą su priedais ir paleidžia tą pačią analizę."""
        with self._analysis_lock:
            self.ui_queue.put(("check_state", True))
            try:
                emails = []
                for i, e in enumerate(pasirinkti, 1):
                    self.ui_queue.put(("status",
                        f"Ruošiami pasirinkti laiškai ({i}/{len(pasirinkti)})…"))
                    try:
                        if e.get("store_id") == "graph":
                            pilnas = graph_client.load_email(e["entry_id"])
                        else:
                            pilnas = outlook_client.load_email(e["entry_id"], e["store_id"])
                    except Exception as err:
                        self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Nepavyko įkelti "
                                  f"pasirinkto laiško „{e['subject'][:50]}“ — {err}")
                        continue
                    if e.get("apdorotas"):
                        storage.unmark_processed(e["entry_id"])
                    emails.append(pilnas)
                if not emails:
                    self.ui_queue.put(("status", "Nepavyko įkelti nė vieno pasirinkto laiško."))
                    return
                self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Rankinis pasirinkimas: "
                          f"analizuojami {len(emails)} laiškai")
                found, done, spent, _f, _sk, updated = self._analyze_emails(
                    emails, "Analizuojami pasirinkti laiškai",
                    keyword_filter=False, manual_pick=True)
                upd = f", 🔔 papildyta: {updated}" if updated else ""
                self.ui_queue.put(("status",
                    f"✔ Pasirinkti laiškai: rasta užsakymų {found} iš {done}"
                    f"{upd}, kaina {spent:.1f} ct"))
                self._publish_share()
                self.ui_queue.put(("refresh", None))
            except Exception:
                self._log(traceback.format_exc(limit=3))
                self.ui_queue.put(("status", "Analizė nutrūko — žr. zurnalas.txt"))
            finally:
                self.ui_queue.put(("check_state", False))

    # ---------------- Dienos darbai ----------------

    def _open_day_tasks(self):
        storage.carry_over_tasks()  # neatlikti darbai visada „gyvena" šiandienoje
        if getattr(self, "_day_win", None) is not None and self._day_win.winfo_exists():
            self._day_win.lift()
            self._day_refresh()
            return

        win = tk.Toplevel(self.root)
        self._day_win = win
        self._day_current = None      # kurio darbo užrašai atverti
        win.title("Dienos darbai")
        win.geometry("1080x600")
        win.minsize(820, 460)

        win.protocol("WM_DELETE_WINDOW", self._day_close)

        top = ttk.Frame(win, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="Diena:").pack(side="left")
        self.day_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Button(top, text="◀", width=3,
                   command=lambda: self._day_step(-1)).pack(side="left", padx=(4, 0))
        self.day_combo = ttk.Combobox(top, textvariable=self.day_var, width=14, state="normal")
        self.day_combo.pack(side="left", padx=2)
        self.day_combo.bind("<<ComboboxSelected>>", lambda e: self._day_refresh())
        self.day_combo.bind("<Return>", lambda e: self._day_refresh())
        ttk.Button(top, text="▶", width=3,
                   command=lambda: self._day_step(1)).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Šiandien", command=self._day_today).pack(side="left")
        ttk.Button(top, text="🔄 Atnaujinti", command=self._day_refresh).pack(side="left", padx=6)
        self.day_info = ttk.Label(top, text="", foreground="#666")
        self.day_info.pack(side="right")

        mid = ttk.PanedWindow(win, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=10)

        left = ttk.Frame(mid)
        mid.add(left, weight=3)
        cols = ("atlikta", "darbas", "perkelta", "uzrasai")
        self.day_tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        self.day_tree.heading("atlikta", text="✔")
        self.day_tree.heading("darbas", text="Darbas")
        self.day_tree.heading("perkelta", text="↻")
        self.day_tree.heading("uzrasai", text="📝")
        self.day_tree.column("atlikta", width=34, stretch=False, anchor="center")
        self.day_tree.column("darbas", width=420)
        self.day_tree.column("perkelta", width=40, stretch=False, anchor="center")
        self.day_tree.column("uzrasai", width=36, stretch=False, anchor="center")
        self.day_tree.tag_configure("done", background="#e8e8e8", foreground="#777")
        self.day_tree.tag_configure("senas", background="#fff3cd")  # perkeltas iš ankstesnės dienos
        self.day_tree.tag_configure("velesnis", foreground="#999")  # tęsėsi, perkeltas į vėlesnę dieną
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.day_tree.yview)
        self.day_tree.configure(yscrollcommand=vsb.set)
        self.day_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.day_tree.bind("<<TreeviewSelect>>", self._day_on_select)
        self.day_tree.bind("<Double-1>", lambda e: self._day_toggle())
        self.day_tree.bind("<Delete>", lambda e: self._day_delete())

        right = ttk.Frame(mid)
        mid.add(right, weight=2)
        head = ttk.Frame(right)
        head.pack(fill="x")
        self.day_note_title = ttk.Label(head, text="Darbo eiga (pasirinkite darbą)",
                                        font=("Segoe UI", 10, "bold"))
        self.day_note_title.pack(side="left")
        ttk.Button(head, text="🕒 Įrašas su data",
                   command=self._day_stamp).pack(side="right")
        self.day_notes = tk.Text(right, wrap="word", font=("Segoe UI", 10),
                                 background="#fffbe6", state="disabled")
        self.day_notes.pack(fill="both", expand=True, pady=(4, 4))
        self.day_notes.bind("<KeyRelease>", self._day_notes_changed)
        self.day_notes.bind("<FocusOut>", lambda e: self._day_save_notes())
        ttk.Label(right, text="Eiga išsisaugo pati.", foreground="#777").pack(anchor="w")

        bottom = ttk.Frame(win, padding=(10, 8))
        bottom.pack(fill="x")
        self.day_entry = ttk.Entry(bottom)
        self.day_entry.pack(side="left", fill="x", expand=True)
        self.day_entry.bind("<Return>", lambda e: self._day_add())
        ttk.Button(bottom, text="➕ Pridėti", command=self._day_add).pack(side="left", padx=6)
        ttk.Button(bottom, text="✔ Atlikta / grąžinti",
                   command=self._day_toggle).pack(side="left")
        ttk.Button(bottom, text="🗑", width=3, command=self._day_delete).pack(side="right")

        ttk.Label(win, text="Įrašykite darbą ir spauskite Enter. Dvigubas paspaudimas = atlikta / grąžinti. "
                            "Neatlikti darbai kitą dieną persikelia patys (↻ rodo, kelintą dieną tęsiasi). "
                            "◀ ▶ — kitų dienų peržiūra (→ = darbas tą dieną dar tęsėsi).",
                  foreground="#777", padding=(10, 0, 10, 8)).pack(fill="x")

        self._day_refresh()

    def _day_today(self):
        self.day_var.set(datetime.date.today().isoformat())
        self._day_refresh()

    def _day_step(self, delta: int):
        """◀ ▶ — peržiūrėti kitą dieną (praeitos rodomos su visais darbais)."""
        try:
            d = datetime.date.fromisoformat((self.day_var.get() or "").strip())
        except ValueError:
            d = datetime.date.today()
        self.day_var.set((d + datetime.timedelta(days=delta)).isoformat())
        self._day_refresh()

    def _day_close(self):
        self._day_save_notes()
        self._day_win.destroy()

    def _day_refresh(self):
        self._day_save_notes()  # kad perpiešiant neprapultų rašoma eiga
        data = (self.day_var.get() or datetime.date.today().isoformat()).strip()
        try:
            datetime.date.fromisoformat(data)
        except ValueError:
            data = datetime.date.today().isoformat()
            self.day_var.set(data)
        today_iso = datetime.date.today().isoformat()
        moved = 0
        if data == today_iso:
            moved = storage.carry_over_tasks(data)  # vakarykščiai neatlikti -> šiandien
            tasks = storage.list_day_tasks(data)
        elif data < today_iso:
            # praeita diena: ir atlikti, ir tie, kurie tada dar tęsėsi (vėliau perkelti)
            tasks = storage.list_day_tasks_history(data)
        else:
            tasks = storage.list_day_tasks(data)
        dates = storage.list_task_dates()
        if data not in dates:
            dates.insert(0, data)
        self.day_combo.config(values=dates)
        sel = self.day_tree.selection()
        keep = sel[0] if sel else None
        for item in self.day_tree.get_children():
            self.day_tree.delete(item)
        for t in tasks:
            if t.get("perkeltas_veliau"):
                tags, mark = ("velesnis",), "→"
            elif t["atlikta"]:
                tags, mark = ("done",), "✔"
            else:
                tags = ("senas",) if (t.get("perkelimu") or 0) else ()
                mark = ""
            self.day_tree.insert("", "end", iid=str(t["id"]), tags=tags, values=(
                mark, t["tekstas"],
                f"↻{t['perkelimu']}" if (t.get("perkelimu") or 0) else "",
                "📝" if (t.get("pastabos") or "").strip() else ""))
        if keep and self.day_tree.exists(keep):
            self.day_tree.selection_set(keep)
        liko = sum(1 for t in tasks if not t["atlikta"] and not t.get("perkeltas_veliau"))
        tesiasi = sum(1 for t in tasks if t.get("perkeltas_veliau"))
        info = f"Darbų: {len(tasks)} · neatlikta: {liko}"
        if tesiasi:
            info += f" · tą dieną dar tęsėsi: {tesiasi}"
        if moved:
            info += f" · perkelta iš ankstesnių dienų: {moved}"
        if data != today_iso:
            info = f"Žiūrima {data} · " + info
        self.day_info.config(text=info)

    def _day_on_select(self, _event=None):
        """Persijungiant išsaugo ankstesnio darbo eigą ir parodo pasirinktojo."""
        self._day_save_notes(silent=True)
        sel = self.day_tree.selection()
        self.day_notes.config(state="normal")
        self.day_notes.delete("1.0", "end")
        if not sel:
            self._day_current = None
            self.day_note_title.config(text="Darbo eiga (pasirinkite darbą)")
            self.day_notes.config(state="disabled")
            return
        task = storage.get_task(int(sel[0]))
        if not task:
            self._day_current = None
            return
        self._day_current = task["id"]
        title = task["tekstas"][:60]
        if task.get("perkelimu"):
            title += f"  (tęsiasi {task['perkelimu'] + 1} d., nuo {task.get('pradine_data')})"
        self.day_note_title.config(text=title)
        self.day_notes.insert("1.0", task.get("pastabos") or "")
        self.day_notes.edit_modified(False)

    def _day_notes_changed(self, _event=None):
        """Rašant eigą — išsisaugo pati, nustojus rašyti."""
        if self._day_notes_after:
            self.root.after_cancel(self._day_notes_after)
        self._day_notes_after = self.root.after(1500, self._day_save_notes)

    def _day_save_notes(self, silent: bool = True):
        """Automatinis eigos išsaugojimas: rašant, persijungiant, uždarant."""
        if self._day_notes_after:
            self.root.after_cancel(self._day_notes_after)
            self._day_notes_after = None
        if self._day_current is None:
            return
        try:
            if not self.day_notes.edit_modified():
                return
            text = self.day_notes.get("1.0", "end-1c")
        except tk.TclError:
            return
        storage.set_task_notes(self._day_current, text)
        self.day_notes.edit_modified(False)
        try:
            self.day_tree.set(str(self._day_current), "uzrasai",
                              "📝" if text.strip() else "")
        except tk.TclError:
            pass

    def _day_stamp(self):
        """Prideda naują eilutę su data — patogu rašyti eigą chronologiškai."""
        if self._day_current is None:
            return
        self.day_notes.config(state="normal")
        current = self.day_notes.get("1.0", "end-1c")
        prefix = "\n" if current.strip() else ""
        self.day_notes.insert("end", f"{prefix}{datetime.datetime.now():%m-%d %H:%M} — ")
        self.day_notes.see("end")
        self.day_notes.focus_set()
        self._day_notes_changed()

    def _day_add(self):
        text = self.day_entry.get().strip()
        if not text:
            return
        storage.add_day_task(text, self.day_var.get().strip())
        self.day_entry.delete(0, "end")
        self._day_refresh()

    def _day_toggle(self):
        sel = self.day_tree.selection()
        if not sel:
            return
        task_id = int(sel[0])
        current = next((t for t in storage.list_day_tasks(self.day_var.get().strip())
                        if t["id"] == task_id), None)
        if current:
            storage.set_task_done(task_id, not current["atlikta"])
            self._day_refresh()

    def _day_delete(self):
        sel = self.day_tree.selection()
        if not sel:
            self.day_info.config(text="Pažymėkite darbą, kurį norite pašalinti.")
            return
        if self._day_current == int(sel[0]):
            self._day_current = None  # kad neišsaugotume į ką tik ištrintą įrašą
        storage.delete_day_task(int(sel[0]))
        self._day_refresh()
        self._day_on_select()

    # ---------------- Sutartys ----------------

    def _open_contracts(self):
        """Sutarčių langas: prie kiekvienos sutarties — iš užsakymų atpažinti žmonės.
        Jokių papildomų AI kaštų: sutartį ir klientą AI ištraukia ir taip,
        o susiejimas vyksta vietoje (kodu)."""
        if self._sut_win is not None and self._sut_win.winfo_exists():
            self._sut_win.lift()
            self._sut_refresh()
            return
        win = tk.Toplevel(self.root)
        self._sut_win = win
        win.title("Sutartys")
        win.geometry("1180x620")
        win.minsize(900, 460)

        top = ttk.Frame(win, padding=(10, 8))
        top.pack(fill="x")
        ttk.Button(top, text="➕ Nauja sutartis",
                   command=lambda: self._sut_edit(None)).pack(side="left")
        ttk.Button(top, text="📄 Įkelti sąrašą (.docx)",
                   command=self._sut_import).pack(side="left", padx=6)
        ttk.Button(top, text="🔄 Atnaujinti",
                   command=self._sut_refresh).pack(side="left")
        ttk.Label(top, text="Paieška:").pack(side="left", padx=(12, 2))
        self.sut_search = tk.StringVar()
        se = ttk.Entry(top, textvariable=self.sut_search, width=22)
        se.pack(side="left")
        se.bind("<Return>", lambda e: self._sut_refresh())
        se.bind("<KeyRelease>", lambda e: self._sut_refresh())
        self.sut_info = ttk.Label(top, text="", foreground="#666")
        self.sut_info.pack(side="right")

        body = ttk.Frame(win, padding=(10, 0, 10, 8))
        body.pack(fill="both", expand=True)
        cols = ("padalinys", "adresas", "numeris", "statusas", "asmenys")
        self.sut_tree = ttk.Treeview(body, columns=cols, show="headings",
                                     selectmode="browse")
        self.sut_tree.heading("padalinys", text="Padalinys")
        self.sut_tree.heading("adresas", text="Adresas")
        self.sut_tree.heading("numeris", text="Sutartis")
        self.sut_tree.heading("statusas", text="Statusas")
        self.sut_tree.heading("asmenys", text="Žmonės")
        self.sut_tree.column("padalinys", width=250, minwidth=140)
        self.sut_tree.column("adresas", width=220, minwidth=120)
        self.sut_tree.column("numeris", width=100, minwidth=85, stretch=False)
        self.sut_tree.column("statusas", width=170, minwidth=100)
        self.sut_tree.column("asmenys", width=380, minwidth=160)
        self.sut_tree.tag_configure("neg", foreground="#999")       # nebegalioja
        self.sut_tree.tag_configure("extra", background="#fff3cd")  # nėra Word sąraše
        self.sut_tree.bind("<Double-1>", lambda e: self._sut_edit_selected())
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.sut_tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=self.sut_tree.xview)
        self.sut_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.sut_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        ttk.Label(win, text="Žmonės prie sutarčių prisirašo patys iš gautų užsakymų — "
                            "papildomų AI kaštų nėra. Dvigubas paspaudimas — redaguoti "
                            "(žmones, statusą, adresą), ➕ prideda naują sutartį.",
                  foreground="#777", padding=(10, 0, 10, 8)).pack(fill="x")

        self._sut_refresh()

    def _sut_refresh(self):
        if self._sut_win is None or not self._sut_win.winfo_exists():
            return
        q = (self.sut_search.get() or "").strip().lower()
        suts, be_sutarties = storage.sutartys_su_uzsakymais()
        for item in self.sut_tree.get_children():
            self.sut_tree.delete(item)

        self._sut_rows = {}
        rodoma = 0
        for s in suts:
            # žmonės = sąrašo kontaktai + nauji vardai iš užsakymų (be dublių)
            zmones = [p.strip() for p in (s.get("asmenys") or "").split(",") if p.strip()]
            have = {p.lower() for p in zmones}
            for o in s["uzsakymai"]:
                k = (o.get("klientas") or "").strip()
                if k and k.lower() not in have:
                    zmones.append(k)
                    have.add(k.lower())
            row = (s.get("padalinys") or "", s.get("adresas") or "",
                   s["numeris"], s.get("pastaba") or "", ", ".join(zmones))
            if q and q not in " ".join(row).lower():
                continue
            rodoma += 1
            neg = "NEBEGALIOJA" in (s.get("pastaba") or "").upper()
            tags = ("extra",) if s.get("padalinys") == "(nėra sąraše)" else (
                ("neg",) if neg else ())
            iid = self.sut_tree.insert("", "end", tags=tags, values=row)
            # redagavimui: pseudo eilutei (nėra sąraše) paduodame užpildytą ruošinį
            self._sut_rows[iid] = dict(s, asmenys=", ".join(zmones)) \
                if s.get("padalinys") == "(nėra sąraše)" else s
        if be_sutarties:
            vardai = []
            have = set()
            for o in be_sutarties:
                k = (o.get("klientas") or "").strip()
                if k and k.lower() not in have:
                    vardai.append(k)
                    have.add(k.lower())
            row = ("Be sutarties", "", "—", "", ", ".join(vardai))
            if not q or q in " ".join(row).lower():
                self.sut_tree.insert("", "end", tags=("extra",), values=row)
        self.sut_info.config(text=f"Sutarčių: {rodoma}")

    def _sut_edit_selected(self):
        sel = self.sut_tree.selection()
        if not sel:
            return
        s = self._sut_rows.get(sel[0])
        if s is None:  # „Be sutarties" eilutė — nėra ko redaguoti
            return
        self._sut_edit(s)

    def _sut_edit(self, s=None):
        """Sutarties pridėjimas / redagavimas ranka (žmonės keičiasi, sutartys
        pasirašomos naujos — viską galima pasitaisyti pačiam)."""
        dlg = tk.Toplevel(self._sut_win if self._sut_win
                          and self._sut_win.winfo_exists() else self.root)
        dlg.title("Redaguoti sutartį" if s else "Nauja sutartis")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        s = s or {}
        pad = "" if (s.get("padalinys") == "(nėra sąraše)") else (s.get("padalinys") or "")
        v_pad = tk.StringVar(value=pad)
        v_adr = tk.StringVar(value=s.get("adresas") or "")
        v_num = tk.StringVar(value=s.get("numeris") or "")
        v_pas = tk.StringVar(value=s.get("pastaba") or "")
        for i, (txt, var) in enumerate([("Padalinys:", v_pad), ("Adresas:", v_adr),
                                        ("Sutarties nr.:", v_num), ("Statusas:", v_pas)]):
            ttk.Label(frm, text=txt).grid(row=i, column=0, sticky="w", pady=3)
            ttk.Entry(frm, textvariable=var, width=46).grid(
                row=i, column=1, sticky="w", pady=3)
        ttk.Label(frm, text="(pvz.: NEBEGALIOJA, pasirašytas padidinimas 2,5k — laisvas tekstas)",
                  foreground="#777").grid(row=4, column=1, sticky="w")
        ttk.Label(frm, text="Žmonės:").grid(row=5, column=0, sticky="nw", pady=(8, 3))
        zmones = tk.Text(frm, width=46, height=4, font=("Segoe UI", 10), wrap="word")
        zmones.grid(row=5, column=1, sticky="w", pady=(8, 3))
        zmones.insert("1.0", s.get("asmenys") or "")

        def save():
            numeris = re.sub(r"\s+", "", v_num.get()).upper()
            if not numeris:
                messagebox.showwarning("Sutartys", "Įrašykite sutarties numerį.",
                                       parent=dlg)
                return
            parts = [p.strip() for p in re.split(r"[\n;]+|,", zmones.get("1.0", "end-1c"))
                     if p.strip()]
            storage.save_sutartis({
                "id": s.get("id"), "numeris": numeris,
                "padalinys": v_pad.get().strip(), "adresas": v_adr.get().strip(),
                "pastaba": v_pas.get().strip(), "asmenys": ", ".join(parts)})
            dlg.destroy()
            self._sut_refresh()

        def remove():
            if messagebox.askyesno("Šalinti",
                                   f"Pašalinti sutartį {s.get('numeris')} iš sąrašo?",
                                   parent=dlg):
                storage.delete_sutartis(s["id"])
                dlg.destroy()
                self._sut_refresh()

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=1, sticky="e", pady=(12, 0))
        if s.get("id"):
            ttk.Button(btns, text="🗑 Šalinti", command=remove).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Atšaukti", command=dlg.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="💾 Išsaugoti", command=save).pack(side="left")

    def _sut_import(self):
        path = filedialog.askopenfilename(
            title="Pasirinkite sutarčių sąrašą (Word)",
            filetypes=[("Word dokumentas", "*.docx")])
        if not path:
            return
        try:
            import sutartys_docx
            rows = sutartys_docx.parse_docx(path)
        except Exception as e:
            messagebox.showerror("Klaida", f"Nepavyko perskaityti failo:\n{e}")
            return
        if not rows:
            messagebox.showwarning("Sutartys",
                                   "Dokumente nerasta lentelės su sutarčių numeriais.")
            return
        n = storage.save_sutartys(rows)
        self._sut_refresh()
        self.sut_info.config(text=f"Įkelta sutarčių: {n}")

    # ---------------- Nustatymai ----------------

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Nustatymai")
        win.grab_set()
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Anthropic API raktas:").grid(row=0, column=0, sticky="w")
        key_var = tk.StringVar(value=self.cfg.get_api_key())
        ttk.Entry(frm, textvariable=key_var, width=52, show="•").grid(
            row=0, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Modelis:").grid(row=1, column=0, sticky="w")
        model_var = tk.StringVar(value=self.cfg["model"])
        ttk.Combobox(frm, textvariable=model_var, values=MODELS, width=28,
                     state="readonly").grid(row=1, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Tikrinti kas (min.):").grid(row=2, column=0, sticky="w")
        interval_var = tk.IntVar(value=self.cfg["check_interval_min"])
        ttk.Spinbox(frm, from_=2, to=240, textvariable=interval_var,
                    width=6).grid(row=2, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Pirmo paleidimo aprėptis (dienos):").grid(row=3, column=0, sticky="w")
        lookback_var = tk.IntVar(value=self.cfg["first_run_lookback_days"])
        ttk.Spinbox(frm, from_=1, to=60, textvariable=lookback_var,
                    width=6).grid(row=3, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Tikrinti tik iš šių siuntėjų\n(atskirti kabliataškiais; tuščia = visi):",
                  justify="left").grid(row=4, column=0, sticky="w")
        filters_var = tk.StringVar(value="; ".join(self.cfg["sender_filters"]))
        ttk.Entry(frm, textvariable=filters_var, width=52).grid(
            row=4, column=1, pady=3, sticky="w")

        ttk.Label(frm, foreground="#777", justify="left", wraplength=620, text=(
            "Tuščias laukas — tikrinami visi gauti laiškai (įprasta, kai užsakymai ateina "
            "tiesiai iš klientų). Galima apriboti iki klientų domenų, pvz.: @klientas.lt. "
            "Papildoma galimybė sudėtingesniam srautui: jei užsakymus jums persiunčia kolega — "
            "įrašykite ir jį; tikrąjį klientą AI atpažins iš persiųsto laiško vidaus, o "
            "susirašinėjimo grandines seka automatiškai (tęsiniai neregistruojami kaip nauji).")).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(frm, text="Stebėjimo aplankas vadovui\n(tuščia = išjungta):",
                  justify="left").grid(row=6, column=0, sticky="w")
        share_var = tk.StringVar(value=self.cfg.get("share_dir") or "")
        ttk.Entry(frm, textvariable=share_var, width=52).grid(row=6, column=1, pady=3, sticky="w")

        kwon_var = tk.BooleanVar(value=bool(self.cfg.get("keyword_filter_on")))
        ttk.Checkbutton(frm, text="Raktažodžių filtras (laiškai be šių žodžių ir be\n"
                                  "PDF/nuotraukų priedų AI nesiunčiami):",
                        variable=kwon_var).grid(row=7, column=0, sticky="w")
        kw_var = tk.StringVar(value="; ".join(self.cfg.get("analysis_keywords") or []))
        ttk.Entry(frm, textvariable=kw_var, width=52).grid(row=7, column=1, pady=3, sticky="w")

        ttk.Separator(frm, orient="horizontal").grid(
            row=8, column=0, columnspan=2, sticky="we", pady=(10, 4))
        ttk.Label(frm, text="Pašto prisijungimas (Microsoft Graph)",
                  font=("Segoe UI", 9, "bold")).grid(row=9, column=0, columnspan=2, sticky="w")

        ttk.Label(frm, text="Programėlės ID (Client ID):").grid(row=10, column=0, sticky="w")
        gcid_var = tk.StringVar(value=self.cfg.get("graph_client_id") or "")
        ttk.Entry(frm, textvariable=gcid_var, width=52).grid(row=10, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Katalogo ID (Tenant ID):").grid(row=11, column=0, sticky="w")
        gtid_var = tk.StringVar(value=self.cfg.get("graph_tenant_id") or "")
        ttk.Entry(frm, textvariable=gtid_var, width=52).grid(row=11, column=1, pady=3, sticky="w")

        ttk.Label(frm, text="Jūsų pašto adresas:").grid(row=12, column=0, sticky="w")
        ghint_var = tk.StringVar(value=self.cfg.get("graph_login_hint") or "")
        ttk.Entry(frm, textvariable=ghint_var, width=52).grid(row=12, column=1, pady=3, sticky="w")

        ttk.Label(frm, foreground="#777", justify="left", wraplength=620, text=(
            "Kaip gauti šiuos ID (vienkartinė registracija per portal.azure.com) — žingsnis "
            "po žingsnio aprašyta SKAITYK.md skyriuje „Kaip prijungti savo paštą“. Pakeitus "
            "duomenis, kito tikrinimo metu atsidarys Microsoft prisijungimo langas.")).grid(
            row=13, column=0, columnspan=2, sticky="w", pady=(0, 4))

        ttk.Label(frm, text="API raktas ir pašto žetonas saugomi užšifruoti (Windows DPAPI) tik šiame kompiuteryje.",
                  foreground="#777").grid(row=14, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def save():
            self.cfg.set_api_key(key_var.get().strip())
            self.cfg["model"] = model_var.get()
            self.cfg["check_interval_min"] = max(2, int(interval_var.get()))
            self.cfg["first_run_lookback_days"] = max(1, int(lookback_var.get()))
            self.cfg["sender_filters"] = [s.strip() for s in filters_var.get().split(";") if s.strip()]
            self.cfg["share_dir"] = share_var.get().strip()
            self.cfg["keyword_filter_on"] = bool(kwon_var.get())
            self.cfg["analysis_keywords"] = [s.strip() for s in kw_var.get().split(";") if s.strip()]
            self.cfg["graph_client_id"] = gcid_var.get().strip()
            self.cfg["graph_tenant_id"] = gtid_var.get().strip()
            self.cfg["graph_login_hint"] = ghint_var.get().strip()
            self.cfg.save()
            self.lbl_interval.config(
                text=f"Automatiškai kas {self.cfg['check_interval_min']} min.")
            win.destroy()

        ttk.Button(frm, text="Išsaugoti", command=save).grid(row=15, column=1, sticky="e", pady=(12, 0))

    # ---------------- Fono darbininkas ----------------

    def _worker_loop(self):
        # Pulsas kas 1 min: kas minutę sekamos užklausos (❗ atsakymai),
        # o PILNAS užsakymų tikrinimas su AI — kas check_interval_min (2 min).
        paskutinis_pilnas = datetime.datetime.min
        while True:
            triggered = self.check_now_event.wait(timeout=60)
            if triggered:
                # valome tik pamačius signalą — kitaip tarp wait() ir clear()
                # atėjęs paspaudimas būtų tyliai prarastas
                self.check_now_event.clear()
            if not triggered and not self.auto_var.get():
                continue
            if not self.cfg.get_api_key():
                continue
            interval = max(2, int(self.cfg["check_interval_min"])) * 60
            pilnas = triggered or (
                (datetime.datetime.now() - paskutinis_pilnas).total_seconds()
                >= interval)
            try:
                with self._analysis_lock:  # kad nesusidurtų su vykstančiu importu
                    if pilnas:
                        self._run_check(manual=triggered)
                        paskutinis_pilnas = datetime.datetime.now()
                    else:
                        self._run_track_only()
            except Exception:
                err = traceback.format_exc(limit=3)
                self.ui_queue.put(("status", "Klaida tikrinant paštą — žr. zurnalas.txt"))
                self._log(err)
            finally:
                self.ui_queue.put(("check_state", False))

    def _run_track_only(self):
        """Tarpinis 1 min pulsas: tik užklausų/atsakymų sekimas, be AI analizės."""
        if self.cfg["mail_source"] != "graph":
            return
        last_check = storage.get_meta("last_check")
        if not last_check:
            return
        since = datetime.datetime.fromisoformat(last_check) - datetime.timedelta(minutes=30)
        since = min(since, datetime.datetime.now() - datetime.timedelta(
            days=SAUGOS_LANGO_DIENOS))
        try:
            if self._track_inquiries(since):
                self.ui_queue.put(("refresh", None))
                self._publish_share()
            self._track_klaida = False
        except Exception:
            if not getattr(self, "_track_klaida", False):
                self._track_klaida = True   # kad kas minutę nekartotų žurnale
                self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Užklausų sekimo "
                          f"klaida:\n{traceback.format_exc(limit=2)}")

    def _analyze_emails(self, new_emails, label="Analizuojami laiškai",
                        keyword_filter=True, manual_pick=False):
        """Bendra analizės dalis periodiniam tikrinimui ir istorijos importui.
        Grąžina (found, done, spent_ct, failed_received, skipped_chain)."""
        # Susirašinėjimo grandinės: iš tos pačios Outlook „conversation" analizuojame
        # tik naujausią laišką (jame cituojama visa grandinė) — kiti pažymimi apdorotais.
        groups = {}
        for e in sorted(new_emails, key=lambda x: x["received"], reverse=True):
            key = e.get("conversation_id") or e["entry_id"]
            groups.setdefault(key, []).append(e)
        to_analyze = [g[0] for g in groups.values()]
        siblings = {g[0]["entry_id"]: g[1:] for g in groups.values()}
        skipped_chain = len(new_emails) - len(to_analyze)

        # Raktazodziu filtras (kastu taupymas): be raktazodzio ir be neperskaitomu
        # priedu (PDF/nuotrauka) laiskas AI nesiunciamas. Rankinis pasirinkimas
        # („Rinktis laiskus") filtro nepaiso - ten sprendzia vartotojas.
        skipped_keywords = 0
        if keyword_filter and self.cfg.get("keyword_filter_on"):
            raktai = [_be_diakritiku(k) for k in (self.cfg.get("analysis_keywords") or []) if k]
            if raktai:
                lieka = []
                for e in to_analyze:
                    try:
                        tekstas, neperskaitomi = attachments_mod.searchable_text(
                            e, int(self.cfg["min_image_kb"]))
                    except Exception:
                        lieka.append(e)  # abejone = analizuojam
                        continue
                    tekstas = _be_diakritiku(tekstas)
                    if neperskaitomi or any(k in tekstas for k in raktai):
                        lieka.append(e)
                        continue
                    skipped_keywords += 1
                    storage.mark_processed_simple(e, note="be raktažodžių")
                    for sib in siblings.get(e["entry_id"], []):
                        storage.mark_processed_simple(sib, note="be raktažodžių (grandinė)")
                    self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Praleista be "
                              f"raktažodžių (AI nekviesta): „{e['subject'][:60]}“ — "
                              f"{e['sender_name']}")
                to_analyze = lieka

        client = anthropic.Anthropic(api_key=self.cfg.get_api_key())
        model = self.cfg["model"]
        min_kb = int(self.cfg["min_image_kb"])
        workers = max(1, int(self.cfg.get("parallel_analyses", 4)))
        recent = storage.recent_orders_summary()
        total = len(to_analyze)
        found = 0
        done = 0
        updated = 0  # grandinės papildymai, įrašyti į esamus užsakymus
        spent = 0.0  # sukauptos išlaidos (euro centais)
        failed_received = []
        self.ui_queue.put(("status", f"{label} (0/{total}, po {workers} vienu metu)…"))

        pool = ThreadPoolExecutor(max_workers=workers)
        try:
            def _submit(email):
                # deterministinis grandinės saugiklis: jei ši grandinė jau turi
                # užsakymą DB — AI apie tai informuojamas (0 papildomų tokenų)
                chain = storage.order_by_conversation(email.get("conversation_id") or "")
                return pool.submit(extractor.analyze_email, client, model, email,
                                   min_kb, recent, chain_order=chain,
                                   manual_pick=manual_pick)

            ordered = list(reversed(to_analyze))  # seniausi pirmi
            futures = {}
            if ordered:
                futures[_submit(ordered[0])] = ordered[0]
                if len(ordered) > 1:
                    # pirmasis įrašo sisteminį promptą į talpyklą —
                    # visi likusieji skaito jį ~10x pigiau
                    futures_wait(list(futures))
                for email in ordered[1:]:
                    futures[_submit(email)] = email
            for fut in as_completed(futures):
                email = futures[fut]
                group_all = [email] + siblings.get(email["entry_id"], [])
                done += 1
                try:
                    analysis = fut.result()
                    is_order = bool(analysis.get("ar_uzsakymas"))
                    u = analysis.get("_usage")
                    if u:
                        spent += u["eur_cents"]
                        self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} kaina "
                                  f"{u['eur_cents']:.2f} ct ({u['model']}, "
                                  f"in {u['input']}+{u['cache_read']}c / out {u['output']}) "
                                  f"— {email['subject'][:50]}")
                    relevant = self._relevant_attachments(email["attachments"])
                    storage.record_processed(
                        email, analysis, is_order,
                        priedai="; ".join(n for n, _, _ in relevant))
                    for sib in siblings.get(email["entry_id"], []):
                        storage.mark_processed_simple(sib)
                    if not is_order:
                        upd_kodas = (analysis.get("atnaujina_koda") or "").strip()
                        if upd_kodas and storage.apply_order_update(upd_kodas, analysis, email):
                            updated += 1
                            self.ui_queue.put(("refresh", None))
                except Exception as e:
                    # nepažymime kaip apdorotų (nei grandinės) — bandysime kitą kartą
                    failed_received.append(min(x["received"] for x in group_all))
                    self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} "
                              f"Nepavyko apdoroti „{email['subject']}“: {e}")
                    self.ui_queue.put(("status",
                        f"{label} ({done}/{total}), rasta užsakymų: {found}"))
                    continue
                if is_order:
                    found += 1
                    self.ui_queue.put(("refresh", None))
                self.ui_queue.put(("status",
                    f"{label} ({done}/{total}), rasta užsakymų: {found}"))
        finally:
            # jei ciklas nutrūko dėl klaidos — atšaukiame dar nepradėtas analizes,
            # kad gija nekabėtų laukdama nebereikalingų API užklausų
            pool.shutdown(wait=True, cancel_futures=True)
        self._last_kw_skipped = skipped_keywords
        return found, done, spent, failed_received, skipped_chain, updated

    def _run_check(self, manual: bool = False):
        self.ui_queue.put(("check_state", True))
        self.ui_queue.put(("status", "Skaitomi nauji laiškai iš Outlook…"))

        last_check = storage.get_meta("last_check")
        if last_check:
            since = datetime.datetime.fromisoformat(last_check) - datetime.timedelta(minutes=30)
            # Saugos langas — žr. SAUGOS_LANGO_DIENOS
            since = min(since, datetime.datetime.now() - datetime.timedelta(
                days=SAUGOS_LANGO_DIENOS))
        else:
            since = datetime.datetime.now() - datetime.timedelta(
                days=int(self.cfg["first_run_lookback_days"]))
        check_started = datetime.datetime.now()
        max_count = int(self.cfg["max_emails_per_check"])

        new_emails = self._mail().fetch_emails_since(
            since, max_count=max_count,
            skip=storage.is_processed,
            sender_filters=self.cfg["sender_filters"],
            on_skip=lambda subj, err: self._log(
                f"{datetime.datetime.now():%Y-%m-%d %H:%M} Laiško nepavyko perskaityti "
                f"(bus pakartota kitą patikrą): „{subj}“ — {err}"))

        self._log(f"{check_started:%Y-%m-%d %H:%M} Tikrinimas "
                  f"({'rankinis' if manual else 'automatinis'}): langas nuo "
                  f"{since:%m-%d %H:%M}, naujų laiškų pagal filtrą: {len(new_emails)}")

        try:
            if self._track_inquiries(since):
                self.ui_queue.put(("refresh", None))
                self._publish_share()
        except Exception:
            self._log(f"{datetime.datetime.now():%Y-%m-%d %H:%M} Užklausų sekimo "
                      f"klaida:\n{traceback.format_exc(limit=2)}")

        if not new_emails:
            storage.set_meta("last_check", check_started.isoformat(timespec="seconds"))
            self.ui_queue.put(("status",
                f"Naujų laiškų (pagal filtrą) nėra. Patikrinta {check_started:%H:%M}"))
            return

        found, done, spent, failed_received, skipped_chain, updated = self._analyze_emails(
            new_emails, "Analizuojami laiškai")

        # last_check stumiame tik tiek, kiek tikrai patikrinome: jei sąrašas buvo
        # nukirptas ties max_count arba dalis laiškų nepavyko — kitą kartą grįžtame prie jų
        candidates = [check_started]
        if len(new_emails) >= max_count:
            candidates.append(min(e["received"] for e in new_emails))
        if failed_received:
            candidates.append(min(failed_received) - datetime.timedelta(seconds=1))
        new_last_check = max(min(candidates), since)
        storage.set_meta("last_check", new_last_check.isoformat(timespec="seconds"))
        self._publish_share()

        # Atsiradus naujų užsakymų ar papildymų — atnaujinamas telefono
        # kontaktų failas darbalaukyje (klientai.vcf)
        if found or updated:
            try:
                import eksportuoti_kontaktus
                eksportuoti_kontaktus.eksportuoti()
            except Exception:
                pass

        chain_note = f", grandinių laiškų praleista: {skipped_chain}" if skipped_chain else ""
        kw = getattr(self, "_last_kw_skipped", 0)
        chain_note += f", be raktažodžių praleista: {kw}" if kw else ""
        chain_note += f", 🔔 papildyta užsakymų: {updated}" if updated else ""
        cost_note = f", kaina {spent:.1f} ct" if spent else ""
        if updated:
            self.ui_queue.put(("bell", None))
        if found:
            self.ui_queue.put(("status",
                f"✔ Rasta naujų užsakymų: {found} (patikrinta: {done}{chain_note}{cost_note}, "
                f"{check_started:%H:%M})"))
            self.ui_queue.put(("bell", None))
        else:
            self.ui_queue.put(("status",
                f"Patikrinta laiškų: {done}{chain_note}{cost_note}, "
                f"užsakymų nerasta ({check_started:%H:%M})"))

    def _log(self, text):
        try:
            import os
            with open(os.path.join(cfg_mod.APP_DIR, "zurnalas.txt"), "a", encoding="utf-8") as f:
                f.write(text.rstrip() + "\n")
        except Exception:
            pass


def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    _taikyti_tema(root)
    App(root)
    root.mainloop()


def _taikyti_tema(root):
    """Šviesi moderni tema; be ttkbootstrap — Windows standartinė."""
    global TB
    style = None
    if TB:
        try:
            style = tb.Style(theme=TEMA)
        except Exception:
            TB = False
    if style is None:
        style = ttk.Style(root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
    style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure(".", font=("Segoe UI", 10))


if __name__ == "__main__":
    main()
