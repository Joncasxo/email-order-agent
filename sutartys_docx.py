# -*- coding: utf-8 -*-
"""Sutarčių sąrašo įkėlimas iš Word (.docx) lentelės — be papildomų bibliotekų.

Lentelės eilutėje ieškome langelio su sutarties numeriu (pvz. SUT25-114);
langelis prieš jį — adresas, dar ankstesnis — padalinys, langeliai po jo —
kontaktiniai asmenys. Numerio langelyje gali būti pastaba, pvz.
„SUT25-88(NEBEGALIOJA)" — ji išsaugoma atskirai.
"""
import re
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CONTRACT_RE = re.compile(r"[A-ZĄ-Ž]{2,4}\s?\d{2}\s?-\s?\d{1,4}")


def norm_number(text: str) -> str:
    """„Sut25 - 114" -> „SUT25-114"."""
    return re.sub(r"\s+", "", (text or "").upper())


def find_numbers(text: str):
    """Visi sutarčių numeriai tekste, sunormalizuoti."""
    return [norm_number(m) for m in CONTRACT_RE.findall((text or "").upper())]


def _cell_text(tc) -> str:
    return " ".join("".join(t.text or "" for t in p.iter(f"{W}t")).strip()
                    for p in tc.findall(f"{W}p")).strip()


def parse_docx(path: str):
    """Grąžina sutarčių sąrašą: [{numeris, padalinys, adresas, asmenys, pastaba}]."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    rows = []
    for tbl in root.iter(f"{W}tbl"):
        for tr in tbl.findall(f"{W}tr"):
            cells = [_cell_text(tc) for tc in tr.findall(f"{W}tc")]
            num_idx = next((i for i, c in enumerate(cells)
                            if CONTRACT_RE.search(c.upper())), None)
            if num_idx is None:
                continue
            num_cell = cells[num_idx]
            numeris = find_numbers(num_cell)[0]
            # pastaba — kas liko langelyje be paties numerio (skliaustai ir pan.)
            likutis = CONTRACT_RE.sub("", num_cell, count=1).strip(" ()-–")
            adresas = cells[num_idx - 1].strip() if num_idx >= 1 else ""
            padalinys = cells[num_idx - 2].strip(" ()") if num_idx >= 2 else ""
            asmenys = ", ".join(c for c in cells[num_idx + 1:] if c.strip())
            asmenys = re.sub(r"\s*,[\s,]*", ", ", asmenys).strip(" ,-–")
            rows.append({
                "numeris": numeris,
                "padalinys": padalinys,
                "adresas": adresas,
                "asmenys": asmenys,
                "pastaba": likutis,
            })
    return rows
