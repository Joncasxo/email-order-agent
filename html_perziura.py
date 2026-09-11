# -*- coding: utf-8 -*-
"""Vieno failo interaktyvi užsakymų peržiūra naršyklei.

Sugeneruoja savarankišką HTML: viduje įterpti duomenys ir visa logika
(paieška, filtrai, rūšiavimas, detalės), jokių interneto ar diegimo
reikalavimų — gavėjas tiesiog atidaro failą naršyklėje."""
import datetime
import json
import re

import storage

_TERM_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})(?:[ T]+(\d{1,2})[:.](\d{2}))?")

# tos pačios spalvos kaip programoje
SPALVOS = {
    "p_pasiūlyta": "#cfe2ff", "p_laimėta": "#e6d3f2",
    "p_nelaimėta": "#e8c4c4",
    "p_užsakyta": "#c8f0c8", "n_užsakyta": "#d1ecf1",
}


def _terminas(o):
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
            iki = datetime.datetime(int(y), int(mo), int(d), int(hh), int(mm))
            return f"{y}-{mo}-{d} {int(hh):02d}:{mm}", iki
        iki = datetime.datetime(int(y), int(mo), int(d), 23, 59)
        return f"{y}-{mo}-{d}", iki
    except ValueError:
        return raw[:24], None


def _duomenys():
    dabar = datetime.datetime.now()
    eilutes = []
    for o in storage.list_orders():
        try:
            prekes = json.loads(o.get("prekes_json") or "[]")
        except Exception:
            prekes = []
        term_txt, term_iki = _terminas(o)
        prior = (o.get("kodas") or "").startswith("!")
        eilutes.append({
            "kodas": o.get("kodas") or "",
            "prior": prior,
            "terminas": term_txt,
            "term_praejes": bool(term_iki and term_iki < dabar),
            "gauta": o.get("received_at") or "",
            "klientas": o.get("klientas") or "?",
            "el_pastas": o.get("kliento_el_pastas") or "",
            "adresas": o.get("kliento_adresas") or "",
            "telefonas": o.get("kliento_telefonas") or "",
            "tema": o.get("subject") or "",
            "sutartis": o.get("sutartis") or "",
            "statusas": o.get("statusas") or "",
            "saskaita": int(o.get("saskaita") or 0),
            "priedai": len([s for s in (o.get("priedai") or "").split(";") if s.strip()]),
            "prekes": [{"preke": p.get("preke") or "", "kiekis": p.get("kiekis") or "",
                        "pastabos": p.get("pastabos") or ""} for p in prekes],
            "pastabos": o.get("pastabos") or "",
            "mano_pastabos": o.get("mano_pastabos") or "",
            "santrauka": o.get("santrauka") or "",
            "atnaujinimai": o.get("atnaujinimai") or "",
            "pristatymo_data": o.get("pristatymo_data") or "",
        })
    return eilutes


def generuoti(kelias: str):
    duomenys = _duomenys()
    dabar = datetime.datetime.now()
    djson = json.dumps(duomenys, ensure_ascii=False).replace("</", "<\\/")
    sjson = json.dumps(SPALVOS)
    html = """<!DOCTYPE html>
<html lang="lt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Užsakymų peržiūra</title>
<style>
 :root { font-family: "Segoe UI", Arial, sans-serif; }
 body { margin: 0; background: #f4f5f7; color: #1c1c1c; }
 header { background: #1f3a5f; color: #fff; padding: 12px 18px; display: flex;
          flex-wrap: wrap; gap: 14px; align-items: center; }
 header h1 { font-size: 17px; margin: 0 12px 0 0; font-weight: 600; }
 header .meta { font-size: 12px; opacity: .8; }
 .valdymas { display: flex; flex-wrap: wrap; gap: 8px; padding: 10px 18px;
             background: #fff; border-bottom: 1px solid #ddd; position: sticky;
             top: 0; z-index: 5; align-items: center; }
 .valdymas input[type=text] { padding: 6px 10px; border: 1px solid #bbb;
             border-radius: 6px; width: 260px; font-size: 13px; }
 .valdymas select, .valdymas input[type=date] { padding: 5px 8px; border: 1px solid #bbb;
             border-radius: 6px; font-size: 13px; background: #fff; }
 .valdymas .kiek { margin-left: auto; font-size: 12.5px; color: #555; }
 table { border-collapse: collapse; width: 100%; background: #fff; font-size: 13px; }
 thead th { position: sticky; top: 52px; background: #e9edf2; text-align: left;
            padding: 7px 9px; border-bottom: 2px solid #c9d2dc; cursor: pointer;
            user-select: none; white-space: nowrap; }
 thead th:hover { background: #dde4ec; }
 tbody td { padding: 6px 9px; border-bottom: 1px solid #eee; vertical-align: top; }
 tbody tr.eilute { cursor: pointer; }
 tbody tr.eilute:hover td { filter: brightness(.96); }
 .terminas { color: #c00000; font-weight: 600; white-space: nowrap; }
 .terminas.praejes { color: #8a8a8a; font-weight: 400; }
 .kodas { font-weight: 600; white-space: nowrap; }
 .detales td { background: #fafbfc; padding: 12px 16px 14px; }
 .detales h3 { margin: 8px 0 4px; font-size: 13px; color: #1f3a5f; }
 .detales table.prekes { width: auto; min-width: 60%; font-size: 12.5px;
                         border: 1px solid #e2e6ea; }
 .detales table.prekes td, .detales table.prekes th { padding: 4px 8px;
                         border-bottom: 1px solid #eee; }
 .detales .laukas { margin: 2px 0; font-size: 12.5px; }
 .zyma { display: inline-block; padding: 1px 7px; border-radius: 9px;
         font-size: 11.5px; background: #eee; margin-left: 6px; }
 @media (max-width: 800px) { .slepiama { display: none; } }
</style>
</head>
<body>
<header>
 <h1>📦 Užsakymų peržiūra</h1>
 <span class="meta">Sugeneruota: __DATA__ · tik peržiūrai, duomenys neatsinaujina patys</span>
</header>
<div class="valdymas">
 <input type="text" id="paieska" placeholder="🔎 Paieška: kodas, klientas, prekė…">
 <select id="statusas">
  <option value="">Visi statusai</option>
  <option value="__nauji">Nauji (be statuso)</option>
  <option>užklausta</option><option>pasiūlyta</option>
  <option>laimėta</option><option>nelaimėta</option><option>užsakyta</option>
 </select>
 <select id="tipas">
  <option value="">Visi užsakymai</option>
  <option value="prior">Tik „!" (su terminu)</option>
  <option value="paprasti">Tik paprasti</option>
 </select>
 <label style="font-size:12.5px">nuo <input type="date" id="nuo"></label>
 <label style="font-size:12.5px">iki <input type="date" id="iki"></label>
 <span class="kiek" id="kiek"></span>
</div>
<table>
 <thead><tr>
  <th data-k="kodas">Kodas</th>
  <th data-k="terminas">⏰ Pateikti iki</th>
  <th data-k="gauta">Data</th>
  <th data-k="klientas">Klientas</th>
  <th data-k="prekes" class="slepiama">Prekės</th>
  <th data-k="priedai">📎</th>
  <th data-k="statusas">Statusas</th>
  <th data-k="saskaita">Sąsk.</th>
 </tr></thead>
 <tbody id="lentele"></tbody>
</table>
<script>
const DUOMENYS = __DUOMENYS__;
const SPALVOS = __SPALVOS__;
let rikiuok = "gauta", kryptis = -1, atverta = null;

function saskZenklas(o) {
  if (!o.saskaita) return "";
  if (o.saskaita === 1) return "✅";
  return o.prior ? "🟠" : "◐";
}
function eiluteSpalva(o) {
  const g = o.prior ? "p" : "n";
  return SPALVOS[g + "_" + o.statusas] || "";
}
function prekesTekstas(o) {
  const dalys = o.prekes.slice(0, 4).map(p => p.preke + (p.kiekis ? " (" + p.kiekis + ")" : ""));
  let t = dalys.join("; ");
  if (o.prekes.length > 4) t += " … (+" + (o.prekes.length - 4) + ")";
  return t;
}
function filtruota() {
  const q = document.getElementById("paieska").value.trim().toLowerCase();
  const st = document.getElementById("statusas").value;
  const tp = document.getElementById("tipas").value;
  const nuo = document.getElementById("nuo").value;
  const iki = document.getElementById("iki").value;
  return DUOMENYS.filter(o => {
    if (st === "__nauji" && o.statusas) return false;
    if (st && st !== "__nauji" && o.statusas !== st) return false;
    if (tp === "prior" && !o.prior) return false;
    if (tp === "paprasti" && o.prior) return false;
    const d = o.gauta.slice(0, 10);
    if (nuo && d < nuo) return false;
    if (iki && d > iki) return false;
    if (q) {
      const blob = [o.kodas, o.klientas, o.tema, o.sutartis, o.mano_pastabos,
        o.santrauka, o.prekes.map(p => p.preke + " " + (p.pastabos||"")).join(" ")]
        .join(" ").toLowerCase();
      if (!blob.includes(q)) return false;
    }
    return true;
  }).sort((a, b) => {
    let x = a[rikiuok], y = b[rikiuok];
    if (rikiuok === "prekes") { x = a.prekes.length; y = b.prekes.length; }
    if (x === y) return 0;
    return (x > y ? 1 : -1) * kryptis;
  });
}
function esc(s) { const d = document.createElement("div"); d.textContent = s || ""; return d.innerHTML; }
function piesk() {
  const tb = document.getElementById("lentele");
  const eil = filtruota();
  tb.innerHTML = "";
  eil.forEach((o, i) => {
    const tr = document.createElement("tr");
    tr.className = "eilute";
    tr.style.background = eiluteSpalva(o);
    const termKl = o.term_praejes ? "terminas praejes" : "terminas";
    tr.innerHTML =
      '<td class="kodas">' + esc(o.kodas) + '</td>' +
      '<td class="' + (o.prior ? termKl : "") + '">' + esc(o.terminas) + '</td>' +
      '<td style="white-space:nowrap">' + esc(o.gauta) + '</td>' +
      '<td>' + esc(o.klientas) + '</td>' +
      '<td class="slepiama">' + esc(prekesTekstas(o)) + '</td>' +
      '<td>' + (o.priedai ? "📎" + o.priedai : "") + '</td>' +
      '<td>' + esc(o.statusas || "—") + '</td>' +
      '<td>' + saskZenklas(o) + '</td>';
    tr.onclick = () => atidaryk(tr, o);
    tb.appendChild(tr);
  });
  document.getElementById("kiek").textContent =
    "Rodoma: " + eil.length + " iš " + DUOMENYS.length;
}
function atidaryk(tr, o) {
  if (atverta) { atverta.remove(); if (atverta.pries === tr) { atverta = null; return; } }
  const dtr = document.createElement("tr");
  dtr.className = "detales"; dtr.pries = tr;
  let prekesHtml = "";
  if (o.prekes.length) {
    prekesHtml = '<table class="prekes"><tr><th>Prekė</th><th>Kiekis</th><th>Pastabos</th></tr>' +
      o.prekes.map(p => "<tr><td>" + esc(p.preke) + "</td><td>" + esc(p.kiekis) +
                        "</td><td>" + esc(p.pastabos) + "</td></tr>").join("") + "</table>";
  }
  const laukai = [];
  if (o.tema) laukai.push("<b>Laiško tema:</b> " + esc(o.tema));
  if (o.sutartis) laukai.push("<b>Sutartis:</b> " + esc(o.sutartis));
  if (o.pristatymo_data) laukai.push("<b>Pristatymas:</b> " + esc(o.pristatymo_data));
  if (o.el_pastas) laukai.push("<b>El. paštas:</b> " + esc(o.el_pastas));
  if (o.telefonas) laukai.push("<b>Telefonas:</b> " + esc(o.telefonas));
  if (o.adresas) laukai.push("<b>Adresas:</b> " + esc(o.adresas));
  if (o.santrauka) laukai.push("<b>Santrauka:</b> " + esc(o.santrauka));
  if (o.pastabos) laukai.push("<b>AI pastabos:</b> " + esc(o.pastabos));
  if (o.mano_pastabos) laukai.push("<b>Mano pastabos:</b> " + esc(o.mano_pastabos));
  if (o.atnaujinimai) laukai.push("<b>Atnaujinimai:</b> " + esc(o.atnaujinimai));
  dtr.innerHTML = '<td colspan="8"><h3>Prekės (' + o.prekes.length + ')</h3>' + prekesHtml +
    laukai.map(l => '<div class="laukas">' + l + "</div>").join("") + "</td>";
  tr.after(dtr);
  atverta = dtr;
}
document.querySelectorAll("thead th").forEach(th => th.onclick = () => {
  const k = th.dataset.k;
  if (rikiuok === k) kryptis *= -1; else { rikiuok = k; kryptis = 1; }
  piesk();
});
["paieska", "statusas", "tipas", "nuo", "iki"].forEach(id =>
  document.getElementById(id).addEventListener("input", piesk));
piesk();
</script>
</body>
</html>"""
    html = (html.replace("__DATA__", f"{dabar:%Y-%m-%d %H:%M}")
                .replace("__DUOMENYS__", djson)
                .replace("__SPALVOS__", sjson))
    with open(kelias, "w", encoding="utf-8") as f:
        f.write(html)
    return len(duomenys)
