# -*- coding: utf-8 -*-
"""Sinchronizacija į bendrinamą aplanką vadovo stebėjimui.

Po kiekvieno pasikeitimo į aplanką nukopijuojama duomenų bazės kopija
(saugiu SQLite backup būdu) ir peržiūros programa. Vadovui pabendrinus
aplanką per OneDrive, jis mato beveik realiu laiku atsinaujinantį sąrašą.
"""
import os
import shutil
import sqlite3
import tempfile

import storage

VIEWER_FILES = ["Uzsakymu perziura.exe", "Sukurti_nuoroda_darbalaukyje.bat",
                "Perziura.bat", "perziura.py", "SKAITYK.txt", "INSTRUKCIJA_VADOVUI.txt"]


def publish(share_dir: str):
    """Atnaujina DB kopiją ir peržiūros failus bendrinamame aplanke."""
    if not share_dir:
        return
    os.makedirs(share_dir, exist_ok=True)

    # DB kopija per SQLite backup API — saugu net jei tuo metu rašoma
    fd, tmp = tempfile.mkstemp(dir=share_dir, suffix=".db.tmp")
    os.close(fd)
    try:
        with storage._lock:
            src = sqlite3.connect(storage.DB_PATH)
            dst = sqlite3.connect(tmp)
            with dst:
                src.backup(dst)
            dst.close()
            src.close()
        os.replace(tmp, os.path.join(share_dir, "uzsakymai.db"))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass

    for name in VIEWER_FILES:
        src_path = os.path.join(storage.APP_DIR, name)
        dst_path = os.path.join(share_dir, name)
        try:
            if os.path.exists(src_path) and (
                    not os.path.exists(dst_path)
                    or os.path.getmtime(src_path) > os.path.getmtime(dst_path)):
                shutil.copy2(src_path, dst_path)
        except OSError:
            pass
