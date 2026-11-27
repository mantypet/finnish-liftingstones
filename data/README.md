# Lifting stone database

Plain CSV files are the database. A small script loads them into SQLite to
check them and to query them.

```
data/
  csv/              the database: one CSV file per table (edit these)
  schema.sql        table definitions, allowed values and views
  sources/
    2026-01-09-finnish-stones-sources.xlsx           original sheet, unchanged
    2026-01-09-finnish-stones-sources_cleaned.xlsx   cleaned copy, with a `changes` tab
tools/
  clean_sheet.py    original sheet -> cleaned sheet (every fix listed in the code and the changes tab)
  import_sheet.py   cleaned sheet + docs/stones.json -> data/csv/ (one-off bootstrap)
  build_db.py       data/csv/ -> build/stones.sqlite, and checks every reference
```

## Why CSV plus SQLite

- **Readable and diffable.** CSV opens in Excel, LibreOffice, R and Python, and
  every change shows up line by line in a pull request.
- **Real checks.** `python tools/build_db.py` loads the CSVs into SQLite with
  foreign keys and allowed-value lists, and fails if a fact has no source or a
  link points at something that does not exist.
- **Portable output.** The built `build/stones.sqlite` is one file that R
  (`RSQLite`), Python, DB Browser for SQLite or a static web page (sql.js) can
  read. It is not committed; rebuild it any time.
- **No server, no dependencies.** `build_db.py` needs only Python 3. The
  clean/import scripts also need `openpyxl`.

PostgreSQL/PostGIS would only pay off once there are many editors or heavy map
queries; the CSVs move there unchanged if that day comes.

## The model

Five kinds of record, and typed links between them:

| Record | What it is | Example |
| --- | --- | --- |
| `stone` | a physical stone, the fact | `alaveteli1` Tomasasstäinä |
| `person` | someone who lifted, carried, owned or named a stone | Tomas of Måsabacka |
| `story` | a story or legend, tagged with motifs in `story_motif` | Tomas carries the stone from the church |
| `custom` | a practice | confirmation test, marriage condition |
| `place` | parish, village, farm, church, museum | Nedervetil > Klockarbacken |

`link` connects any two records with a role, e.g. `person lifted stone`,
`story about stone`, `story features person`, `custom practised_at stone`,
`custom practised_in place`, `person lived_at place`. The allowed shapes are
listed in `tools/build_db.py` (`LINK_ROLES`).

Facts about a stone itself are `claim` rows (weight, dimensions, shape,
marking, dating ...). Where a stone is or was is in `location` (stones move:
Tomas' stone was carried, Mattilan mummu went to a museum). Names in every
spelling are in `stone_name`. Photos and figures are in `media`.

### Provenance

- `source` is a publication, archive record, photo record or dataset.
- `attestation` is one passage in a source, with the original excerpt and an
  English translation. It may be about one stone or only give context.
- Every claim, link and story telling points to an attestation; every location,
  name and stone record points to a source. Nothing is stored without one.
- Reprints point to the original with `source.reprint_of`, so a fact with two
  citations is not mistaken for two independent witnesses.

The sheet and the map are sources too (`sheet20260109`, `map20260530`): the
stone status, condition and coordinates come from the researcher's own work.

### Two labels on every statement

- `status`: what kind of statement it is.
  `measured` (a measurement or a calculation from one), `documented` (stated as
  fact by the source), `local_tradition` (oral history about real people and
  places), `legend` (implausible or supernatural), `fiction`.
- `confidence`: how sure we are of our own reading. `certain`, `probable`,
  `possible`. Anything not `certain` is a guess and says why in `note`.

`claim.method` adds how a number was obtained: `weighed`, `measured`,
`calculated`, `estimated`, `reported` (stated without saying how),
`traditional`.

## Working with it

```sh
python tools/build_db.py           # build and check
```

Example queries (`sqlite3 build/stones.sqlite` or any SQLite tool):

```sql
SELECT * FROM v_stone;                                  -- one row per stone
SELECT * FROM v_claim WHERE field = 'weight';           -- every weight with its source
SELECT * FROM v_link WHERE id = 'perho1';               -- everything linked to a stone
SELECT s.title, m.motif FROM story s JOIN story_motif m USING (story_id);
```

From now on the CSVs are the master copy: add and fix data there and run
`build_db.py` before committing. `import_sheet.py` refuses to overwrite them
unless given `--force`.

### Id conventions

- Stones: `<historic parish or municipality><n>`, e.g. `teerijarvi1`.
- Sources: `<title><ddmmyyyy>` for dated issues, `<author><yyyy>` for books
  (the sheet's convention). Old ids are kept in `source.original_ids`.
- Attestations: `<source_id>:<stone_id>` or `<source_id>:context`, `-2` if repeated.

## Open questions in the data

- Four sources have an id and a date that disagree (`source.date_note`); check them on digi.kansalliskirjasto.fi.
- Kotiliesi 15.8.1929: the URL points to page 14, the excerpt says page 12.
- `sipoo1` Nilsfafas lyftån: Stejskal writes "Gästerby, Si."; read as Sibbo (the map agrees), not Sideby.
- `alavus1` Tohnisen äijän kivi: the 1914 text says Kyrö church, the map says Alavus.
- Seven map stones (`hyrynsalmi1`, `korsnas1`, `mustasaari1`, `perho3`, `teerijarvi2`, and the sheet's `haapavesi2`, `haapavesi3`) have no source passage yet.
- Where the sheet and the map give different coordinates, the map point is current (it is newer); the sheet point is kept.
