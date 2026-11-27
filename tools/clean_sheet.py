"""Clean the manually gathered sources sheet.

Reads the original sheet (left untouched) and writes a cleaned copy next to it,
with a `changes` tab that lists every edit and why it was made.

    python tools/clean_sheet.py

Requires: openpyxl
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "data/sources/2026-01-09-finnish-stones-sources.xlsx"
OUT = ROOT / "data/sources/2026-01-09-finnish-stones-sources_cleaned.xlsx"

STONE_COLS = ["stone_location_id", "name", "region", "euref_e", "euref_n", "coord_ref",
              "status", "condition", "location_type", "context"]
SOURCE_COLS = ["stone_location_id", "source_id", "source_url", "source_author", "source_date",
               "cite_date", "cite_ref", "notes_raw_fi", "notes_raw_sv", "notes_en",
               "stone_name_fi", "stone_name_sv"]

changes: list[dict] = []


def log(tab, row, column, old, new, reason):
    changes.append({"tab": tab, "original_row": row, "column": column,
                    "old_value": "" if old is None else str(old),
                    "new_value": "" if new is None else str(new), "reason": reason})


def clean_value(v):
    if isinstance(v, str):
        v = v.strip()
        if v in ("", "NA"):
            return None
    return v


# --- source ids -------------------------------------------------------------
# One id per publication. The sheet sometimes used two ids for the same issue.
ID_RENAMES = {
    "pyrkija15031914": ("pyrkija15051914", "Same issue as pyrkija19140515; the issue is dated 15.05.1914 (cite_ref), not 15.03."),
    "pyrkija19140515": ("pyrkija15051914", "Same issue as pyrkija15031914; one id per publication, ddmmyyyy like most ids."),
    "viikkosanomat19230630": ("viikkosanomat30061923", "Same issue as viikkosanomat30061923 (same binding 1735486)."),
    "urheilulehti01041911": ("suomenurheilulehti01041911", "Same issue as suomenurheilulehti01041911 (same binding 882323)."),
    "itahame23041931": ("itahame23041930", "Id said 1931 but source_date and binding order (1183581 < 1183582 = 30.4.1930) give 23.4.1930."),
    "elamaapyyntipaikalladdmmyyyy": ("elamaapyyntipaikalla", "Placeholder 'ddmmyyyy' removed; publication date unknown."),
}

# --- dates --------------------------------------------------------------------
# (source_id after rename, column, old) -> (new, reason)
DATE_FIXES = {
    ("pyrkija15051914", "source_date", "1914-03-15"): ("1914-05-15", "cite_ref gives 15.05.1914."),
    ("kotiliesi19290815", "source_date", "1929-08-19"): ("1929-08-15", "Id and the uploaded excerpt both give 15.8.1929."),
    ("savokarjala16081895", "source_date", "1895-16-08"): ("1895-08-16", "Month 16 impossible; id gives 16.08.1895."),
    ("budkavlen01011929", "cite_date", "202-10-22"): ("2025-10-22", "Year typo; the other rows of that session are 2025-10-22."),
    ("savonsanomat16011908", "cite_date", "2025-12-105"): ("2025-12-05", "Day 105 impossible; neighbouring rows are 2025-12-05 (guess)."),
    ("elamaapyyntipaikalla", "source_date", "yyyy-mm-dd"): (None, "Placeholder; date unknown."),
}
# Id and date disagree and nothing in hand settles which is right: keep as entered, flag.
DATE_UNRESOLVED = {
    "savo09101924": "Id says 09.10.1924, source_date says 1924-01-09. Check on digi.kansalliskirjasto.fi.",
    "savo09101888": "Id says 09.10.1888, source_date says 1888-01-09. Check on digi.kansalliskirjasto.fi.",
    "kaiku28041928": "Id says 28.04.1928, source_date says 1928-04-08. Check on digi.kansalliskirjasto.fi.",
    "taisteluun10121929": "Id says 10.12.1929, source_date says 1929-12-01. Check on digi.kansalliskirjasto.fi.",
}
DATE_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$")

# --- stone links --------------------------------------------------------------
# (original row) -> (new stone_location_id, reason)
LINK_FIXES = {
    31: ("munsala2", "Row names Stubbåkers grisen (munsala2) but had no stone id."),
    52: ("ahtari1", "Row names Siltalan Kustavin väkikivi (ahtari1) but had no stone id."),
    53: ("ahtari1", "Row names Siltalan Kustavin väkikivi (ahtari1) but had no stone id; text same as Viikkosanomat."),
    54: ("ahtari1", "Row names Siltalan Kustavin väkikivi (ahtari1) but had no stone id; text same as Viikkosanomat."),
    35: ("perho2", "Marjaana Jylhä / Jylhän naisten väkikivi is a different farm from Poranen: moved from perho1 to new perho2 (probable, see source analysis)."),
    125: ("perho2", "Excerpt names 'Jylhän naisten väkikivi'; linked to new perho2."),
}

# --- Stejskal 1954: one row per stone instead of seven ids in one cell ---------
STEJSKAL_STONES = [
    ("alaveteli1", None, "Tåmasasstejne",
     "I Nedervetil kyrkoby finnes i Klockarbacken en sten, som kallas »Tomas' sten» (Tåmasasstejne). Varje år prövar skriftskolepojkarna sina krafter på densamma. Den som orkar lyfta stenen från marken, betraktas av sina kamrater som fullvuxen, medan den som inte mäktar med stenen, anses ha kommit för tidigt till skriftskolan. Stenen väger ungefär 92 kg och är mycket svår att få något tag på. Stenen har enligt hörsägen varit i bruk ett par hundra år. Det berättas, att en person vid namn Tomas, som bodde helt nära kyrkan, burit stenen från närheten av kyrkan ungefär 3/4 km till nuvarande plats, där den sedan dess fått ligga invid vägen. Efter denna person har stenen fått sitt namn (Nv).",
     "In Nedervetil church village, in Klockarbacken, there is a stone called 'Tomas' stone' (Tåmasasstejne). Every year the confirmation-school boys test their strength on it. Whoever manages to lift the stone off the ground is regarded by his peers as grown up, while whoever cannot is considered to have come to confirmation school too early. The stone weighs about 92 kg and is very hard to get a grip on. By hearsay the stone has been in use for a couple of hundred years. It is said that a man named Tomas, who lived very near the church, carried the stone from near the church about 3/4 km to its present place, where it has since lain beside the road. The stone is named after him (Nv)."),
    ("teerijarvi1", None, "Risbackastein",
     "I Terjärv finnes en lyftsten, som kallas »Risbackastein», bild 26. Den väger mellan 150—200 kg och har använts ungefär 200 år. Nu ligger den stödd på tvenne mindre stockar. För att det skall vara lättare att lyfta stenen, begagnar man en stark trädgren, mot vilken stenens vassa undre del vilar och vilken man tar tag i från vardera ändan. I forna tider var det en sed, påstår de gamla, att en man inte fick gifta sig, förrän han orkat lyfta stenen från marken (Tj). [Bild 26. Risbacka lyftsten, D 1. Terjärv.]",
     "In Terjärv there is a lifting stone called 'Risbackastein', picture 26. It weighs between 150 and 200 kg and has been used for about 200 years. It now rests on two smaller logs. To make it easier to lift, a strong branch is used, against which the stone's sharp lower part rests and which is gripped at both ends. In former times, the old people claim, a man was not allowed to marry until he had managed to lift the stone off the ground (Tj). [Picture 26. Risbacka lifting stone, D 1. Terjärv.]"),
    ("munsala1", None, "ungkarlsstenen",
     "På kyrkbacken i Munsala finnes en sten, benämnd »ungkarlsstenen». Gossarna skulle lyfta den innan de ansågs som fullgoda ungkarlar. [Not 2: Uppgiften av Lars Huldén, Munsala.]",
     "On the church hill in Munsala there is a stone called 'ungkarlsstenen' (the bachelor stone). The boys had to lift it before they were regarded as proper bachelors. [Note 2: Information from Lars Huldén, Munsala.]"),
    ("oravainen1", None, None,
     "De stenar, som därvid begagnas, kallas »skriftskolstenar» (M.H. Mb; K.G.A. La). En sådan sten, som skall lyftas upp i famnen, finnes invid kyrkomuren i Oravais.",
     "The stones used for this are called 'skriftskolstenar' (confirmation-school stones) (M.H. Mb; K.G.A. La). One such stone, which is to be lifted up into the arms, is beside the church wall in Oravais."),
    ("pietarsaari1", None, None,
     "I Pedersöre finnes en dylik sten bakom klockstapeln vid kyrkan. Får man bara upp stenen från marken, så är provet klarat.",
     "In Pedersöre there is such a stone behind the bell tower at the church. If you just get the stone up off the ground, the test is passed."),
    ("uusikaarlepyy1", None, "junttansten",
     "Stenen kan också vara försedd med järnband, ring eller krok. En sådan sten fanns före sekelskiftet på rådstugugården i Nykarleby. Den kallades »junttansten», och man använde den till att påla med. Det var särskilt »skriftskoleschacket» som övade sig flitigt i att resa upp den i stående ställning (E.E. Nb).",
     "The stone may also be fitted with an iron band, ring or hook. Such a stone stood in the town-hall yard in Nykarleby before the turn of the century. It was called 'junttansten' and was used for driving piles. The confirmation-school lads especially practised raising it to a standing position (E.E. Nb)."),
    ("munsala2", None, "Stubbåkers grisen",
     "En liknande sten, kallad »Stubbåkers grisen» (139 kg), finnes i Kantlax i Munsala. [Not 3: Se SNF 40—41, Ortnamnen i Munsala.]",
     "A similar stone, called 'Stubbåkers grisen' (139 kg), is in Kantlax in Munsala. [Note 3: See SNF 40–41, Ortnamnen i Munsala.]"),
    ("jomala1", None, "Björsby lyftsten",
     "På Åland i Björsby i Jomala finnes en lyftsten, i vilken årtalet 1783 är inhugget. Till formen är stenen äggrund (Jo), bild 27. ... I vissa fall gäller det att lyfta stenen upp på axeln samt kasta den bakåt (Jo). [Bild 27. Björsby lyftsten (daterad 1783), D 1. Jomala.]",
     "On Åland, in Björsby in Jomala, there is a lifting stone with the year 1783 carved into it. The stone is egg-shaped (Jo), picture 27. ... In some cases the task is to lift the stone onto the shoulder and throw it backwards (Jo). [Picture 27. Björsby lifting stone (dated 1783), D 1. Jomala.]"),
    ("sipoo1", None, "Nilsfafas lyftån",
     "Stundom lägger man ytterligare en eller flere stenar ovanpå, såsom vid »Nilsfafas lyftån» i Gästerby, Si.",
     "Sometimes one or more further stones are laid on top, as at 'Nilsfafas lyftån' in Gästerby, Si."),
]

NEW_STONES = [
    # id, name, region, e, n, coord_ref, status, condition, location_type, context, note
    ["perho2", "Jylhän naisten väkikivi", "Perho", 376080, 7008985, None, "unconfirmed", None, None, "local",
     "Added: split from perho1. SKS archive and Pohjolan helleenit 2007 name a 'Jylhän naisten väkikivi' at the Jylhä farm, a different farm from Poranen. Coordinates from docs/stones.json (map, 2026-05-30)."],
    ["jomala1", "Björsby lyftsten", "Jomala", 110147, 6693064, None, "unknown", None, None, "lyftsten",
     "Added: Stejskal 1954 p. D1 and fig. 27. Year 1783 carved into it, egg-shaped. Coordinates from docs/stones.json (map, 2026-05-30)."],
    ["sipoo1", "Nilsfafas lyftån", "Sipoo", 404498, 6689831, None, "unknown", None, None, "lyftsten",
     "Added: Stejskal 1954 ('i Gästerby, Si.'). Gästerby is a village in Sibbo/Sipoo, and the map has this stone in Sipoo, so 'Si' read as Sibbo, not Sideby (probable). Coordinates from docs/stones.json (map, 2026-05-30)."],
]


def norm_text(s):
    return re.sub(r"\W+", "", (s or "").lower())


def main():
    wb = openpyxl.load_workbook(SRC)

    # ---- stone_location ----
    stones = []
    for i, r in enumerate(wb["stone_location"].iter_rows(min_row=2, values_only=True), start=2):
        if all(v is None for v in r):
            continue
        d = {}
        for c, v in zip(STONE_COLS, r):
            nv = clean_value(v)
            if v == "NA":
                log("stone_location", i, c, v, None, "'NA' and empty both meant missing; now always empty.")
            if isinstance(nv, float) and nv.is_integer():
                nv = int(nv)
            d[c] = nv
        d["original_row"] = i
        d["cleaning_note"] = None
        stones.append(d)
    stones_by_id = {s["stone_location_id"]: s for s in stones}
    stones_by_id["perho1"]["cleaning_note"] = (
        "Perho likely has two stones: this one (Poranen) and Jylhän naisten väkikivi, now perho2.")
    for row in NEW_STONES:
        d = dict(zip(STONE_COLS, row[:10]))
        d["original_row"] = None
        d["cleaning_note"] = row[10]
        stones.append(d)
        log("stone_location", None, "(new row)", None, row[0], row[10])

    # ---- source ----
    raw = []
    for i, r in enumerate(wb["source"].iter_rows(min_row=2, values_only=True), start=2):
        if all(v is None for v in r):
            continue
        if any(v is not None for v in r[12:]):
            raise SystemExit(f"row {i}: unexpected data in headerless columns")
        raw.append((i, r[:12]))
    log("source", None, "(3 headerless columns)", None, None, "Removed: no header and no data.")

    rows = []
    aaltonen_n = 0
    for i, r in raw:
        d = {"original_row": i, "original_source_id": r[1], "cleaning_note": []}
        for c, v in zip(SOURCE_COLS, r):
            nv = clean_value(v)
            if isinstance(v, str) and v != v.strip() and nv:
                log("source", i, c, repr(v), nv, "Removed leading/trailing whitespace.")
            if v == "NA":
                log("source", i, c, v, None, "'NA' and empty both meant missing; now always empty.")
            if isinstance(v, dt.datetime):
                nv = v.date().isoformat()
                log("source", i, c, v, nv, "Cell was an Excel datetime; now a YYYY-MM-DD text like the rest.")
            d[c] = nv
        sid = d["source_id"]
        if sid in ID_RENAMES:
            new, why = ID_RENAMES[sid]
            log("source", i, "source_id", sid, new, why)
            d["source_id"] = new
        elif sid == "aaltonen1950":
            aaltonen_n += 1
            new = f"aaltonen1950{'abcde'[aaltonen_n - 1]}"
            log("source", i, "source_id", sid, new, "Five different Finna photo records shared one id; each record is its own source.")
            d["source_id"] = new
        elif sid is None and d["source_url"] and "kinoregina" in d["source_url"]:
            d["source_id"] = "kivenpyorittajankyla"
            log("source", i, "source_id", None, "kivenpyorittajankyla", "Row had no id; named after the film.")
        for col in ("source_date", "cite_date"):
            key = (d["source_id"], col, d[col])
            if key in DATE_FIXES:
                new, why = DATE_FIXES[key]
                log("source", i, col, d[col], new, why)
                d[col] = new
            if d[col] and not DATE_RE.match(d[col]):
                raise SystemExit(f"row {i}: bad {col} {d[col]!r}")
        if d["source_id"] in DATE_UNRESOLVED:
            d["cleaning_note"].append("DATE CHECK: " + DATE_UNRESOLVED[d["source_id"]])
        if d["source_id"] == "kotiliesi19290815":
            d["cleaning_note"].append("PAGE CHECK: URL says page 14, the uploaded excerpt says page 12.")
        if i in LINK_FIXES:
            new, why = LINK_FIXES[i]
            log("source", i, "stone_location_id", d["stone_location_id"], new, why)
            d["stone_location_id"] = new
        rows.append(d)

    # Same publication cited twice where one excerpt contains the other: keep the longer one.
    by_id = {}
    for d in rows:
        by_id.setdefault(d["source_id"], []).append(d)
    drop = set()
    for sid, group in by_id.items():
        for a in group:
            for b in group:
                if a is b or id(a) in drop or id(b) in drop:
                    continue
                ta = norm_text(a["notes_raw_fi"] or a["notes_raw_sv"])
                tb = norm_text(b["notes_raw_fi"] or b["notes_raw_sv"])
                if ta and tb and ta in tb and len(ta) < len(tb):
                    for c in ("stone_location_id", "stone_name_fi", "stone_name_sv", "source_author", "cite_ref"):
                        if a[c] and not b[c]:
                            log("source", b["original_row"], c, None, a[c], f"Carried over from duplicate row {a['original_row']}.")
                            b[c] = a[c]
                    log("source", a["original_row"], "(row)", sid, None,
                        f"Duplicate: its excerpt is contained in row {b['original_row']} of the same publication. Row removed.")
                    drop.add(id(a))
    rows = [d for d in rows if id(d) not in drop]

    # Explode the Stejskal row.
    out = []
    for d in rows:
        if d["source_id"] != "stejskal1954":
            out.append(d)
            continue
        ids = (d["stone_location_id"] or "").split()
        log("source", d["original_row"], "stone_location_id", d["stone_location_id"].replace("\n", " "), "(one row per stone)",
            "Seven stone ids in one cell. Split into one row per stone, each with the passage about that stone; "
            "added jomala1 and sipoo1, which Stejskal also describes. The full chapter text stays on a row with no stone id.")
        ctx = dict(d, stone_location_id=None, cleaning_note=list(d["cleaning_note"]) + ["Full chapter text; stone passages are on the rows below."])
        out.append(ctx)
        listed = {s[0] for s in STEJSKAL_STONES}
        assert set(ids) <= listed, ids
        for stone_id, name_fi, name_sv, sv, en in STEJSKAL_STONES:
            out.append(dict(d, stone_location_id=stone_id, notes_raw_fi=None, notes_raw_sv=sv, notes_en=en,
                            stone_name_fi=name_fi, stone_name_sv=name_sv,
                            cleaning_note=["Passage extracted from the chapter text; English translation added during cleaning."]))
    rows = out

    # ---- write ----
    wbo = openpyxl.Workbook()
    ws = wbo.active
    ws.title = "stone_location"
    hdr = STONE_COLS + ["cleaning_note", "original_row"]
    ws.append(hdr)
    for s in stones:
        ws.append([s.get(c) for c in hdr])
    ws2 = wbo.create_sheet("source")
    hdr2 = SOURCE_COLS + ["original_source_id", "cleaning_note", "original_row"]
    ws2.append(hdr2)
    for d in rows:
        vals = []
        for c in hdr2:
            v = d.get(c)
            if c == "cleaning_note":
                v = " ".join(v) or None
            vals.append(v)
        ws2.append(vals)
    ws3 = wbo.create_sheet("changes")
    hdr3 = ["tab", "original_row", "column", "old_value", "new_value", "reason"]
    ws3.append(hdr3)
    for ch in changes:
        ws3.append([ch[c] for c in hdr3])
    for w in (ws, ws2, ws3):
        for cell in w[1]:
            cell.font = Font(bold=True)
        w.freeze_panes = "A2"
    wbo.save(OUT)
    print(f"stones: {len(stones)}  source rows: {len(rows)}  changes logged: {len(changes)}")
    print(f"wrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
