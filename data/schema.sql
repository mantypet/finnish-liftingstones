-- Finnish lifting stones database schema (SQLite).
--
-- The CSV files in data/csv/ are the master copy; tools/build_db.py loads them
-- into build/stones.sqlite with this schema and checks every reference.
--
-- Model: five kinds of record and the sourced links between them.
--
--   stone   a physical stone (the fact)
--   person  someone who lifted, carried, owned or named a stone
--   story   a story or legend, tagged with motifs
--   custom  a practice: confirmation test, marriage condition, Sunday sport ...
--   place   parish, village, farm, church or museum
--
--   link    typed connection between any two records, e.g.
--           person --lifter--> stone, story --about--> stone,
--           custom --practised_at--> stone, story --set_in--> place.
--           Every link names the source passage (attestation) it rests on.
--
-- Evidence: source = a publication/record; attestation = one passage in it.
-- Facts about a stone itself (weight, size, marking ...) are claims, each tied
-- to an attestation. Where a stone is or was is a location row tied to a source.
--
-- Two separate judgements on every statement:
--   status     what kind of statement it is:
--              measured | documented | local_tradition | legend | fiction
--   confidence how sure we are of our reading of it:
--              certain | probable | possible   (not 'certain' = a guess; say why in note)

PRAGMA foreign_keys = ON;

-- Evidence ------------------------------------------------------------------

CREATE TABLE source (
    source_id       TEXT PRIMARY KEY,
    source_type     TEXT NOT NULL CHECK (source_type IN (
                        'newspaper','periodical','book','thesis','archive_record',
                        'photo_record','web','film','dataset','analysis')),
    title           TEXT,
    author          TEXT,
    pub_date        TEXT,              -- YYYY-MM-DD; see date_precision
    date_precision  TEXT CHECK (date_precision IN ('day','year')),
    date_note       TEXT,              -- open questions about the date
    url             TEXT,
    accessed        TEXT,              -- YYYY-MM-DD the researcher first read it
    language        TEXT,              -- fi | sv | en
    reprint_of      TEXT REFERENCES source(source_id),
    rating          TEXT,              -- from the source analysis report
    cite_ref        TEXT,
    original_ids    TEXT,              -- other ids used for it in the original sheet
    note            TEXT
);

-- One passage in a source. stone_id is set when the passage is about one stone.
CREATE TABLE attestation (
    attestation_id  TEXT PRIMARY KEY,  -- <source_id>:<stone_id|context>[-n]
    source_id       TEXT NOT NULL REFERENCES source(source_id),
    stone_id        TEXT REFERENCES stone(stone_id),
    mentioned_name  TEXT,              -- stone name as the passage gives it
    page            TEXT,
    url             TEXT,              -- page link, when it differs from the source URL
    excerpt_fi      TEXT,
    excerpt_sv      TEXT,
    excerpt_en      TEXT,              -- translation
    confidence      TEXT NOT NULL DEFAULT 'certain'
                    CHECK (confidence IN ('certain','probable','possible')),
                    -- how sure we are the passage is about this stone
    note            TEXT
);

-- Records -------------------------------------------------------------------

CREATE TABLE place (
    place_id        TEXT PRIMARY KEY,
    name            TEXT NOT NULL,     -- name as used in the sources
    name_alt        TEXT,              -- the other language's name
    kind            TEXT NOT NULL CHECK (kind IN (
                        'parish','village','farm','church','museum','site')),
    parent_id       TEXT REFERENCES place(place_id),
    municipality_fi TEXT,              -- municipality it lies in today
    municipality_sv TEXT,
    municipality_code TEXT,
    note            TEXT
);

CREATE TABLE stone (
    stone_id        TEXT PRIMARY KEY,  -- <historic parish or municipality><n>
    preferred_name  TEXT NOT NULL,
    stone_type      TEXT NOT NULL CHECK (stone_type IN (
                        'väkikivi','lyftsten','skriftskolsten','junttasten','other')),
    status          TEXT NOT NULL CHECK (status IN ('confirmed','unconfirmed','unknown','lost')),
                    -- does it still exist: confirmed = checked by the researcher
    condition       TEXT,
    access          TEXT CHECK (access IN ('public','private','museum')),
    sheet_context   TEXT,              -- raw `context` value from the sheet
    record_confidence TEXT NOT NULL DEFAULT 'certain'
                    CHECK (record_confidence IN ('certain','probable','possible')),
                    -- how sure we are this is a separate, real stone
    source_id       TEXT NOT NULL REFERENCES source(source_id),
                    -- where this record and its editorial fields come from
    note            TEXT
);

CREATE TABLE person (
    person_id       TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    birth_year      INTEGER,
    death_year      INTEGER,
    years_approx    INTEGER NOT NULL DEFAULT 0 CHECK (years_approx IN (0,1)),
    note            TEXT
);

CREATE TABLE story (
    story_id        TEXT PRIMARY KEY,
    title           TEXT NOT NULL,
    summary         TEXT NOT NULL,
    status          TEXT NOT NULL CHECK (status IN (
                        'documented','local_tradition','legend','fiction')),
    note            TEXT
);

CREATE TABLE story_motif (
    story_id        TEXT NOT NULL REFERENCES story(story_id),
    motif           TEXT NOT NULL CHECK (motif IN (
                        'carried far','thrown over a fence','unliftable by ordinary men',
                        'woman lifter','added weight','eponym','stone moved or destroyed',
                        'strongman','giant')),
    PRIMARY KEY (story_id, motif)
);

-- Each place a story is told. A story needs at least one telling.
CREATE TABLE story_telling (
    story_id        TEXT NOT NULL REFERENCES story(story_id),
    attestation_id  TEXT NOT NULL REFERENCES attestation(attestation_id),
    note            TEXT,
    PRIMARY KEY (story_id, attestation_id)
);

CREATE TABLE custom (
    custom_id       TEXT PRIMARY KEY,
    tag             TEXT NOT NULL CHECK (tag IN (
                        'confirmation test','marriage condition','bachelor test','sunday sport',
                        'guest challenge','strength feat','carrying test','women lifters','added weight')),
    lift_style      TEXT CHECK (lift_style IN (
                        'off the ground','to the knees','to the chest','into the arms','upright',
                        'shoulder and throw','over a fence','carry')),
    description     TEXT,
    note            TEXT
);

-- Links ---------------------------------------------------------------------

-- A typed, sourced connection between two records.
-- Allowed (from_type, role, to_type) combinations are checked by build_db.py.
CREATE TABLE link (
    link_id         TEXT PRIMARY KEY,
    from_type       TEXT NOT NULL CHECK (from_type IN ('stone','person','story','custom','place')),
    from_id         TEXT NOT NULL,
    role            TEXT NOT NULL,
    to_type         TEXT NOT NULL CHECK (to_type IN ('stone','person','story','custom','place')),
    to_id           TEXT NOT NULL,
    attestation_id  TEXT NOT NULL REFERENCES attestation(attestation_id),
    status          TEXT NOT NULL CHECK (status IN (
                        'measured','documented','local_tradition','legend','fiction')),
    confidence      TEXT NOT NULL DEFAULT 'certain'
                    CHECK (confidence IN ('certain','probable','possible')),
    note            TEXT
);

-- Facts about a stone ---------------------------------------------------------

CREATE TABLE stone_name (
    stone_id        TEXT NOT NULL REFERENCES stone(stone_id),
    name            TEXT NOT NULL,
    language        TEXT,              -- fi | sv | dialect
    source_id       TEXT NOT NULL REFERENCES source(source_id),
    note            TEXT,
    PRIMARY KEY (stone_id, name, source_id)
);

-- Where a stone is or was. Stones move, so a stone can have several rows.
CREATE TABLE location (
    location_id     TEXT PRIMARY KEY,
    stone_id        TEXT NOT NULL REFERENCES stone(stone_id),
    place_id        TEXT REFERENCES place(place_id),  -- most specific place known
    valid_from      TEXT,
    valid_to        TEXT,
    description     TEXT,
    easting         INTEGER,           -- ETRS-TM35FIN, EPSG:3067
    northing        INTEGER,
    point_marks     TEXT,              -- what the point is: stone, church, museum, house ...
    precision_m     INTEGER,
    is_current      INTEGER NOT NULL DEFAULT 0 CHECK (is_current IN (0,1)),
    source_id       TEXT NOT NULL REFERENCES source(source_id),
    confidence      TEXT NOT NULL DEFAULT 'certain'
                    CHECK (confidence IN ('certain','probable','possible')),
    note            TEXT
);

-- One fact as one source states it. Conflicting values sit side by side.
CREATE TABLE claim (
    claim_id        TEXT PRIMARY KEY,
    attestation_id  TEXT NOT NULL REFERENCES attestation(attestation_id),
    stone_id        TEXT NOT NULL REFERENCES stone(stone_id),
    field           TEXT NOT NULL CHECK (field IN (
                        'weight','dimensions','volume','shape','marking','fittings',
                        'dating','location','carry_distance','handling','use','fate')),
    value           TEXT NOT NULL,     -- as the source says it, in English
    value_num       REAL,
    value_num_max   REAL,              -- upper end of a range such as 150-200 kg
    unit            TEXT,
    method          TEXT CHECK (method IN (
                        'weighed','measured','calculated','estimated','reported','traditional')),
                    -- how a number was obtained; reported = stated without saying how
    status          TEXT NOT NULL CHECK (status IN (
                        'measured','documented','local_tradition','legend','fiction')),
    confidence      TEXT NOT NULL DEFAULT 'certain'
                    CHECK (confidence IN ('certain','probable','possible')),
    note            TEXT
);

CREATE TABLE media (
    media_id        TEXT PRIMARY KEY,
    stone_id        TEXT REFERENCES stone(stone_id),
    source_id       TEXT NOT NULL REFERENCES source(source_id),
    kind            TEXT NOT NULL CHECK (kind IN ('photo','figure','film')),
    archive_id      TEXT,
    url             TEXT,
    year            INTEGER,
    photographer    TEXT,
    licence         TEXT,
    caption         TEXT,
    confidence      TEXT NOT NULL DEFAULT 'certain'
                    CHECK (confidence IN ('certain','probable','possible')),
    note            TEXT
);

-- Views ---------------------------------------------------------------------

-- Every claim with the source it rests on.
CREATE VIEW v_claim AS
SELECT c.stone_id, s.preferred_name, c.field, c.value, c.value_num, c.value_num_max,
       c.unit, c.method, c.status, c.confidence, src.source_id, src.title AS source_title,
       src.pub_date, a.page, COALESCE(a.url, src.url) AS url, c.note
FROM claim c
JOIN attestation a ON a.attestation_id = c.attestation_id
JOIN source src ON src.source_id = a.source_id
JOIN stone s ON s.stone_id = c.stone_id;

-- Every link read in both directions, with its source.
CREATE VIEW v_link AS
SELECT l.from_type AS type, l.from_id AS id, l.role, l.to_type AS other_type, l.to_id AS other_id,
       l.status, l.confidence, a.source_id, l.link_id
FROM link l JOIN attestation a USING (attestation_id)
UNION ALL
SELECT l.to_type, l.to_id, 'inverse of ' || l.role, l.from_type, l.from_id,
       l.status, l.confidence, a.source_id, l.link_id
FROM link l JOIN attestation a USING (attestation_id);

-- One row per stone: current location and what the sources say it weighs.
CREATE VIEW v_stone AS
SELECT s.stone_id, s.preferred_name, s.stone_type, s.status, s.record_confidence,
       COALESCE(par.name, p.name) AS parish, COALESCE(par.municipality_fi, p.municipality_fi) AS municipality,
       l.easting, l.northing, l.point_marks,
       (SELECT group_concat(c.value || ' (' || a.source_id || ')', '; ')
          FROM claim c JOIN attestation a USING (attestation_id)
         WHERE c.stone_id = s.stone_id AND c.field = 'weight') AS weights,
       (SELECT count(DISTINCT a.source_id) FROM attestation a
         WHERE a.stone_id = s.stone_id) AS n_sources,
       (SELECT count(*) FROM link k WHERE (k.to_type = 'stone' AND k.to_id = s.stone_id)
                                      OR (k.from_type = 'stone' AND k.from_id = s.stone_id)) AS n_links
FROM stone s
LEFT JOIN location l ON l.stone_id = s.stone_id AND l.is_current = 1
LEFT JOIN place p ON p.place_id = l.place_id
LEFT JOIN place par ON par.place_id = p.parent_id AND p.kind <> 'parish';
