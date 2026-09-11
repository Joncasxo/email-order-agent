# -*- coding: utf-8 -*-
"""Programos nustatymai. API raktas saugomas užšifruotas Windows DPAPI (pririštas prie vartotojo paskyros)."""
import base64
import json
import os

import win32crypt

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

DEFAULTS = {
    "api_key_enc": "",
    "model": "claude-opus-5",
    "check_interval_min": 10,
    "first_run_lookback_days": 7,
    "min_image_kb": 15,          # mažesnės nuotraukos (parašų logotipai) praleidžiamos
    "max_emails_per_check": 50,
    "parallel_analyses": 4,      # kiek laiškų analizuojama vienu metu
    "sender_filters": [],        # tuščias sąrašas = tikrinami visi gauti laiškai
    # Bendrinamas aplankas vadovo stebėjimui (ten sinchronizuojama DB kopija
    # ir peržiūros programa). Tuščia = išjungta.
    "share_dir": "",
    "auto_check": True,
    # Microsoft Graph (pašto) prisijungimas — įrašoma per ⚙ Nustatymai.
    # Kaip gauti šiuos ID — žr. SKAITYK.md skyrių „Kaip prijungti savo paštą".
    "graph_client_id": "",
    "graph_tenant_id": "",
    "graph_login_hint": "",      # jūsų pašto adresas
    # Pašto šaltinis: "com" — klasikinis Outlook (COM), "graph" — Microsoft Graph API
    "mail_source": "com",
    # Raktazodziu filtras pries AI analize (kastu taupymas): laiskas siunciamas
    # AI tik jei temoje/tekste/Excel prieduose yra bent vienas raktazodis ARBA
    # turi neperskaitomu priedu (PDF/nuotrauka/doc) - ju turinio pigiai patikrinti
    # negalime, tad saugumo sumetimais tokie visada analizuojami.
    # Lyginama be diakritiku ir mazosiomis. Prie rinkinio verta prideti savo
    # sutarciu numeriu prefiksus (pvz., "sut2") - pritaikykite savo imonei.
    "keyword_filter_on": True,
    "analysis_keywords": ["uzsaky", "vnt", "prasom", "kaina",
                          "reikia", "pirk", "pasiulym", "uzklaus", "reikalin"],
}


def _encrypt(text: str) -> str:
    blob = win32crypt.CryptProtectData(text.encode("utf-8"), None, None, None, None, 0)
    return base64.b64encode(blob).decode("ascii")


def _decrypt(b64: str) -> str:
    raw = base64.b64decode(b64)
    return win32crypt.CryptUnprotectData(raw, None, None, None, 0)[1].decode("utf-8")


class Config:
    def __init__(self):
        self.data = dict(DEFAULTS)
        self.load()

    def load(self):
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    stored = json.load(f)
                self.data.update({k: v for k, v in stored.items() if k in DEFAULTS})
            except Exception:
                pass

    def save(self):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get_api_key(self) -> str:
        enc = self.data.get("api_key_enc") or ""
        if not enc:
            return ""
        try:
            return _decrypt(enc)
        except Exception:
            return ""

    def set_api_key(self, key: str):
        self.data["api_key_enc"] = _encrypt(key.strip()) if key.strip() else ""

    def __getitem__(self, k):
        return self.data[k]

    def get(self, k, default=None):
        return self.data.get(k, default)

    def __setitem__(self, k, v):
        self.data[k] = v
