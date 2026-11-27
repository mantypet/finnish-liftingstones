"""Build build/stones.sqlite from data/csv/*.csv and check the data.

    python tools/build_db.py

Python standard library only. Fails (exit code 1) if any reference is broken,
a value is outside its allowed list, or a record has no source.
"""
from __future__ import annotations

import csv
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CSV_DIR = ROOT / "data/csv"
SCHEMA = ROOT / "data/schema.sql"
DB = ROOT / "build/stones.sqlite"

# Load order respects foreign keys.
TABLES = ["source", "place", "stone", "attestation", "person", "story", "story_motif", "story_telling",
          "custom", "link", "stone_name", "location", "claim", "media"]

# Allowed link shapes: (from_type, role, to_type)
LINK_ROLES = {
    ("person", "eponym_of", "stone"), ("person", "lifted", "stone"), ("person", "carried", "stone"),
    ("person", "owned", "stone"), ("person", "host_of", "stone"), ("person", "photographed", "stone"),
    ("person", "lived_at", "place"), ("person", "born_at", "place"),
    ("story", "about", "stone"), ("story", "features", "person"), ("story", "set_in", "place"),
    ("story", "variant_of", "story"),
    ("custom", "practised_at", "stone"), ("custom", "practised_in", "place"),
    ("stone", "same_as", "stone"),
}
ID_COL = {"stone": "stone_id", "person": "person_id", "story": "story_id", "custom": "custom_id", "place": "place_id"}


def main() -> int:
    DB.parent.mkdir(exist_ok=True)
    if DB.exists():
        DB.unlink()
    con = sqlite3.connect(DB)
    con.executescript(SCHEMA.read_text(encoding="utf-8"))
    con.execute("PRAGMA foreign_keys = OFF")  # attestation/stone are mutually referenced; check after load
    for t in TABLES:
        with (CSV_DIR / f"{t}.csv").open(encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
        if not rows:
            continue
        cols = list(rows[0].keys())
        con.executemany(
            f"INSERT INTO {t} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
            [[(r[c] if r[c] != "" else None) for c in cols] for r in rows])
    con.commit()

    errors = []
    for row in con.execute("PRAGMA foreign_key_check"):
        errors.append(f"broken reference: table {row[0]} rowid {row[1]} -> {row[2]}")

    for link_id, ft, fid, role, tt, tid in con.execute(
            "SELECT link_id, from_type, from_id, role, to_type, to_id FROM link"):
        if (ft, role, tt) not in LINK_ROLES:
            errors.append(f"link {link_id}: '{ft} {role} {tt}' is not an allowed link shape")
        for typ, i in ((ft, fid), (tt, tid)):
            if not con.execute(f"SELECT 1 FROM {typ} WHERE {ID_COL[typ]} = ?", (i,)).fetchone():
                errors.append(f"link {link_id}: no {typ} '{i}'")

    checks = {
        "claim whose stone differs from its passage's stone":
            "SELECT c.claim_id FROM claim c JOIN attestation a USING (attestation_id) WHERE a.stone_id IS NOT c.stone_id",
        "story with no telling":
            "SELECT story_id FROM story WHERE story_id NOT IN (SELECT story_id FROM story_telling)",
        "custom with no sourced link":
            "SELECT custom_id FROM custom WHERE custom_id NOT IN (SELECT from_id FROM link WHERE from_type='custom')",
        "person with no sourced link":
            "SELECT person_id FROM person WHERE person_id NOT IN "
            "(SELECT from_id FROM link WHERE from_type='person' UNION SELECT to_id FROM link WHERE to_type='person')",
        "stone with more than one current location":
            "SELECT stone_id FROM location WHERE is_current=1 GROUP BY stone_id HAVING count(*) > 1",
        "date not YYYY-MM-DD":
            "SELECT source_id FROM source WHERE (pub_date IS NOT NULL AND pub_date NOT GLOB '[0-9][0-9][0-9][0-9]-[01][0-9]-[0-3][0-9]') "
            "OR (accessed IS NOT NULL AND accessed NOT GLOB '[0-9][0-9][0-9][0-9]-[01][0-9]-[0-3][0-9]')",
    }
    for label, sql in checks.items():
        for (bad,) in con.execute(sql):
            errors.append(f"{label}: {bad}")

    warn_sql = {
        "stones with no source passage yet":
            "SELECT stone_id FROM stone WHERE stone_id NOT IN (SELECT stone_id FROM attestation WHERE stone_id IS NOT NULL)",
        "stones with no coordinates":
            "SELECT stone_id FROM stone WHERE stone_id NOT IN (SELECT stone_id FROM location WHERE easting IS NOT NULL)",
        "sources with an open date question":
            "SELECT source_id FROM source WHERE date_note IS NOT NULL",
    }

    counts = {t: con.execute(f"SELECT count(*) FROM {t}").fetchone()[0] for t in TABLES}
    con.close()

    print(f"built {DB.relative_to(ROOT)}: " + ", ".join(f"{t} {n}" for t, n in counts.items()))
    con = sqlite3.connect(DB)
    for label, sql in warn_sql.items():
        ids = [r[0] for r in con.execute(sql)]
        if ids:
            print(f"note: {label}: {', '.join(ids)}")
    con.close()
    if errors:
        print(f"\n{len(errors)} error(s):", file=sys.stderr)
        for e in errors:
            print("  " + e, file=sys.stderr)
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
