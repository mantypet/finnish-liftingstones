"""Bootstrap the CSV database from the cleaned sources sheet and the map data.

    python tools/import_sheet.py [--force]

Reads  data/sources/2026-01-09-finnish-stones-sources_cleaned.xlsx
       docs/stones.json (the map's stone list, 2026-05-30)
Writes data/csv/*.csv

This was run once to create data/csv/. From then on the CSV files are the
master copy: edit them directly. Re-running overwrites every CSV, so it refuses
unless --force is given.

The records typed in by hand below (claims, places, people, stories, customs,
links, media) were read
from the excerpts in the sheet and from the source analysis report of
2026-11-27. Each one points to the attestation (source passage) it rests on.

Requires: openpyxl
"""
from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SHEET = ROOT / "data/sources/2026-01-09-finnish-stones-sources_cleaned.xlsx"
MAP = ROOT / "docs/stones.json"
OUT = ROOT / "data/csv"

SHEET_SRC = "sheet20260109"
MAP_SRC = "map20260530"
REPORT_SRC = "report20261127"

# Parishes: place_id, name in the sources, other-language name, municipality fi, sv, code
PARISHES = [
    ("ahtari", "Ähtäri", "Etseri", "Ähtäri", "Etseri", "989"),
    ("haapavesi", "Haapavesi", None, "Haapavesi", "Haapavesi", "071"),
    ("haapajarvi", "Haapajärvi", None, "Haapajärvi", "Haapajärvi", "069"),
    ("kangasniemi", "Kangasniemi", None, "Kangasniemi", "Kangasniemi", "213"),
    ("perho", "Perho", None, "Perho", "Perho", "584"),
    ("kyyjarvi", "Kyyjärvi", None, "Kyyjärvi", "Kyyjärvi", "256"),
    ("sonkajarvi", "Sonkajärvi", None, "Sonkajärvi", "Sonkajärvi", "762"),
    ("sysma", "Sysmä", None, "Sysmä", "Sysmä", "781"),
    ("hartola", "Hartola", "Gustav Adolfs", "Hartola", "Gustav Adolfs", "081"),
    ("nedervetil", "Nedervetil", "Alaveteli", "Kruunupyy", "Kronoby", "288"),
    ("terjarv", "Terjärv", "Teerijärvi", "Kruunupyy", "Kronoby", "288"),
    ("munsala", "Munsala", None, "Uusikaarlepyy", "Nykarleby", "893"),
    ("oravais", "Oravais", "Oravainen", "Vöyri", "Vörå", "946"),
    ("pedersore", "Pedersöre", "Pedersören", "Pietarsaari", "Jakobstad", "598"),
    ("nykarleby", "Nykarleby", "Uusikaarlepyy", "Uusikaarlepyy", "Nykarleby", "893"),
    ("jomala", "Jomala", None, "Jomala", "Jomala", "170"),
    ("sibbo", "Sibbo", "Sipoo", "Sipoo", "Sibbo", "753"),
    ("hyrynsalmi", "Hyrynsalmi", None, "Hyrynsalmi", "Hyrynsalmi", "105"),
    ("korsholm", "Korsholm", "Mustasaari", "Mustasaari", "Korsholm", "499"),
    ("korsnas", "Korsnäs", None, "Korsnäs", "Korsnäs", "280"),
    ("alavus", "Alavus", "Alavo", "Alavus", "Alavo", "010"),
    ("isokyro", "Kyrö (Isokyrö)", "Storkyro", "Isokyrö", "Storkyro", "152"),
]
# Smaller places: place_id, name, kind, parent parish, note
SUBPLACES = [
    ("klockarbacken", "Klockarbacken", "village", "nedervetil", None),
    ("masabacka", "Måsabacka", "farm", "nedervetil", None),
    ("nedervetil_church", "Nedervetil church", "church", "nedervetil", None),
    ("kantlax", "Kantlax", "village", "munsala", None),
    ("munsala_church", "Munsala church", "church", "munsala", None),
    ("oravais_church", "Oravais church", "church", "oravais", None),
    ("pedersore_church", "Pedersöre church", "church", "pedersore", "The old Pedersöre church stands in Jakobstad today."),
    ("nykarleby_townhall", "Nykarleby town-hall yard (rådstugugården)", "site", "nykarleby", None),
    ("bjorsby", "Björsby", "village", "jomala", None),
    ("gasterby", "Gästerby", "village", "sibbo", "Stejskal writes 'Gästerby, Si.'; read as Sibbo (probable)."),
    ("poranen", "Poranen (Mehtiä-Poranen)", "farm", "perho", None),
    ("jylha", "Jylhä", "farm", "perho", None),
    ("nopola", "Nopola", "farm", "kyyjarvi", None),
    ("ritola", "Ritola", "farm", "haapavesi", None),
    ("synsio", "Synsiö", "village", "kangasniemi", None),
    ("vehkomaki", "Vehkomäki", "farm", "kangasniemi", None),
    ("pirttiniemi", "Pirttiniemi (Jalassyrjä)", "farm", "kangasniemi", None),
    ("siltala", "Siltala", "farm", "ahtari", None),
    ("ronkkola", "Rönkkölä", "farm", "sonkajarvi", None),
    ("nostokiven_kallio", "Nostokiven kallio, Kaijanmäki", "site", "sonkajarvi", "Named after the stone."),
    ("istula", "Istula", "farm", "haapajarvi", None),
    ("itahamen_museo", "Itä-Hämeen museo", "museum", "hartola", None),
    ("kyro_church", "Kyrö church", "church", "isokyro", None),
]
# Sheet region / map municipality -> parish place_id
REGION = {
    "Ähtäri": "ahtari", "Haapavesi": "haapavesi", "Kangasniemi": "kangasniemi", "Perho": "perho",
    "Sonkajärvi": "sonkajarvi", "Sysmä": "sysma", "Alaveteli": "nedervetil", "Teerijärvi": "terjarv",
    "Munsala": "munsala", "Oravainen": "oravais", "Pietarsaari": "pedersore", "Uusikaarlepyy": "nykarleby",
    "Jomala": "jomala", "Sipoo": "sibbo", "Hyrynsalmi": "hyrynsalmi", "Mustasaari": "korsholm",
    "Korsnäs": "korsnas", "Alavus": "alavus",
}
# Most specific known place for a stone's current location, where finer than the parish.
STONE_PLACE = {
    "alaveteli1": "klockarbacken", "munsala1": "munsala_church", "munsala2": "kantlax",
    "oravainen1": "oravais_church", "pietarsaari1": "pedersore_church", "uusikaarlepyy1": "nykarleby_townhall",
    "jomala1": "bjorsby", "sipoo1": "gasterby", "perho1": "poranen", "perho2": "jylha",
    "kangasniemi1": "synsio", "kangasniemi2": "pirttiniemi", "sonkajarvi1": "nostokiven_kallio",
}

# Map entry name -> stone id. Stones only on the map get new ids.
MAP_TO_STONE = {
    "Nauris-Oskarin väkikivi": "hyrynsalmi1",
    "Tåmasasstejne": "alaveteli1",
    "Risbackasten": "teerijarvi1",
    "Skrifskolstennar #1": "oravainen1",
    "Skrifskolstennar #2": "pietarsaari1",
    "Skrifskolstennar #3": "teerijarvi2",
    "Junntansten": "uusikaarlepyy1",
    "Ungkarlsstenen": "munsala1",
    "Stubbåkers Grisen": "munsala2",
    "Karperö lyftsten": "mustasaari1",
    "Södra Björköstenet": "korsnas1",
    "Hiiden äijän väkikivi": "perho3",
    "Porasen tuvan väkikivi": "perho1",
    "Jylhän naisten väkikivi": "perho2",
    "Tohnisen äijän kivi": "alavus1",
    "Mattilan Mummu": "sysma1",
    "Nils-Fafas Lyfton": "sipoo1",
    "Björsby Lyftstenen": "jomala1",
}
MAP_ONLY = {
    # id: (name, type, region, historic-parish confidence, note)
    "hyrynsalmi1": ("Nauris-Oskarin väkikivi", "väkikivi", "Hyrynsalmi", "certain", None),
    "teerijarvi2": ("Skriftskolsten 3 (name?)", "skriftskolsten", "Teerijärvi", "probable",
                    "Map calls it 'Skrifskolstennar #3'. The point is about 550 m from Risbacka, so historic parish Terjärv is a guess."),
    "mustasaari1": ("Karperö lyftsten", "lyftsten", "Mustasaari", "certain", None),
    "korsnas1": ("Södra Björköstenet", "lyftsten", "Korsnäs", "certain", None),
    "perho3": ("Hiiden äijän väkikivi", "väkikivi", "Perho", "certain", None),
    "alavus1": ("Tohnisen äijän kivi", "väkikivi", "Alavus", "certain",
                "Suomen Urheilulehti 1914 says Tohni rolled a stone onto the wall of Kyrö church; the map places the stone in Alavus. Check which."),
}

PREFERRED_NAME = {
    "alaveteli1": ("Tomasasstäinä", "Preferred name as printed in Budkavlen 1929; the sheet spelled it 'Tomasassstäina'."),
}

# (source_id, stone_id) -> (confidence, note) for the link between passage and stone
LINK_CONFIDENCE = {
    ("uusisuometar07071876", "kangasniemi1"): ("possible", "Names 'Ison Heikin väkikivi' on the Synsiö shore; may be the same stone as Nuija-Heikin väkikivi."),
    ("paulaharju01011930", "perho1"): ("probable", "Poras-äijä throws 'a great väkikivi' over a fence; that this is the Poranen stone is inferred."),
    ("kansankuvalehti16011931", "perho1"): ("probable", "Reprint of Paulaharju 1930; same inference as there."),
    ("perhojokilaakso13012021", "perho1"): ("probable", "Viitala lived at the Poranen homestead, so the stone in the 1932 photo is probably the Poranen stone."),
    ("kotiliesi19290815", "perho2"): ("possible", "Marjaana Jylhä is called a lifter of väkikivet; no stone is named. Linked to the Jylhä women's stone by place."),
    ("pohjolanhelleenit01012007", "perho2"): ("certain", "Quotes SKS KRA records naming 'Jylhän naisten väkikivi'."),
}
# Original sheet row -> stone link added only in the database (not in the cleaned sheet).
DB_ONLY_LINKS = {
    50: ("alavus1", "possible", "Tohni's stone; the sheet names it 'Tohnisen äijän kivi' but had no stone id. See alavus1 note."),
}
EXTRA_ATTESTATIONS = [
    # source_id, stone_id, page, note
    ("ekrem2020", "teerijarvi1", None,
     "Per the source analysis report: Ekrem 2020 repeats 150-200 kg and the marriage rule and prints a 1932 photo of the stone. Passage not yet copied into the sheet."),
]

SOURCE_META = {
    # source_id: dict(title, author, rating, reprint_of, note, type)
    "budkavlen01011929": dict(title="Budkavlen 1929, no. 3, p. 8: 'En lyftsten'", author="Walter Herrgård", rating="core, primary"),
    "stejskal1954": dict(title="Folklig idrott", author="Maximilian Stejskal", rating="core, mixed",
                         note="Copies Tomas' sten from Budkavlen 1929 and Stubbåkers grisen from Huldén 1952."),
    "hulden1952": dict(title="Ortnamnen i Munsala (SNF 40-41)", author="Lars Huldén", rating="core, primary"),
    "ekrem2020": dict(title="Stora finlandssvenska festboken", author="Bergman & Ekrem", rating="useful, secondary", pub_date="2020-01-01"),
    "kansankuvalehti16011931": dict(title="Kansan Kuvalehti", rating="useful, reprint", reprint_of="paulaharju01011930"),
    "paulaharju01011930": dict(title="Suomenselän vieriltä", author="Samuli Paulaharju"),
    "kotiliesi19290815": dict(title="Kotiliesi", rating="useful for people"),
    "suomenurheilulehti01041911": dict(title="Suomen Urheilulehti", rating="useful, weak"),
    "pyrkija15051914": dict(title="Pyrkijä : nuorison rientojen kannattaja", author="P. P-a", rating="context"),
    "kauppisheikki1912": dict(title="Savolaisia", author="Kauppis-Heikki", rating="context, fiction"),
    "perhojokilaakso13012021": dict(title="Perhonjokilaakso", rating="context as uploaded", author=None,
                                    note="The sheet gives Samuli Paulaharju as author; he took the 1932 photo the article discusses."),
    "pohjolanhelleenit01012007": dict(title="Pohjolan helleenit"),
    "osterbottenstidning09062019": dict(title="Österbottens Tidning"),
    "itahame30041930": dict(title="Itä-Häme"),
    "viikkosanomat30061923": dict(title="Viikkosanomat", reprint_of="vaasa30051923",
                                  note="Same text as Vaasa 30.5.1923, a month earlier; treated as the reprint (probable)."),
    "canadanuutiset19071923": dict(title="Canadan Uutiset", reprint_of="vaasa30051923"),
    "vaasa30051923": dict(title="Vaasa"),
    "turistforeningsarsbok01011898": dict(title="Turistföreningens i Finland årsbok 1898",
                                          note="Probably the Swedish edition of the Suomen Matkailijayhdistyksen vuosikirja 1898."),
    "suomenmatkailijayhdistyksenvuosikirja01011898": dict(title="Suomen Matkailijayhdistyksen vuosikirja 1898"),
    "collan1904": dict(title="Suomen kansan leikkejä", author="Anni Collan"),
    "sarkkinen2010": dict(title="Pyhäjärvi Kaakkomäki itä, tarkastuskertomus (Kyppi)", author="Mika Sarkkinen"),
    "carsten1886": dict(title="Kertomus Haapajärven kihlakunnan muinaisjäännöksistä", author="Carstén",
                        note="Leppänen 2018 cites this as Castrén (1886: 35-36)."),
    "leppanen2018": dict(title="Kivet kipperällä asemalla: Suomen kivipöytien rakenteen ja arkeologisen kontekstin analyysi", author="Jere Leppänen"),
    "kivenpyorittajankyla": dict(title="Kivenpyörittäjän kylä (film)", author="Markus Pölönen"),
}

FIELDS = {
    "source": ["source_id", "source_type", "title", "author", "pub_date", "date_precision", "date_note", "url",
               "accessed", "language", "reprint_of", "rating", "cite_ref", "original_ids", "note"],
    "stone": ["stone_id", "preferred_name", "stone_type", "status", "condition", "access", "sheet_context",
              "record_confidence", "source_id", "note"],
    "attestation": ["attestation_id", "source_id", "stone_id", "mentioned_name", "page", "url", "excerpt_fi",
                    "excerpt_sv", "excerpt_en", "confidence", "note"],
    "stone_name": ["stone_id", "name", "language", "source_id", "note"],
    "location": ["location_id", "stone_id", "place_id", "valid_from", "valid_to", "description", "easting", "northing",
                 "point_marks", "precision_m", "is_current", "source_id", "confidence", "note"],
    "claim": ["claim_id", "attestation_id", "stone_id", "field", "value", "value_num", "value_num_max", "unit",
              "method", "status", "confidence", "note"],
    "place": ["place_id", "name", "name_alt", "kind", "parent_id", "municipality_fi", "municipality_sv",
              "municipality_code", "note"],
    "person": ["person_id", "name", "birth_year", "death_year", "years_approx", "note"],
    "story": ["story_id", "title", "summary", "status", "note"],
    "story_motif": ["story_id", "motif"],
    "story_telling": ["story_id", "attestation_id", "note"],
    "custom": ["custom_id", "tag", "lift_style", "description", "note"],
    "link": ["link_id", "from_type", "from_id", "role", "to_type", "to_id", "attestation_id", "status", "confidence", "note"],
    "media": ["media_id", "stone_id", "source_id", "kind", "archive_id", "url", "year", "photographer", "licence",
              "caption", "confidence", "note"],
}

# ---------------------------------------------------------------------------
# Facts read from the excerpts. (attestation_id, field, value, num, num_max, unit, method, confidence, note)
# Status of a claim follows from how it was obtained, unless overridden here.
CLAIM_STATUS_OVERRIDE = {
    ("turistforeningsarsbok01011898:kangasniemi1", "carry_distance"):
        "legend",  # the same text says the stone is too big to have been carried
}


A_BUD = "budkavlen01011929:alaveteli1"
A_STE = "stejskal1954:{}"
CLAIMS = [
    (A_BUD, "weight", "about 92 kg", 92, None, "kg", "calculated", "certain",
     "Volume 34 dm3 x assumed density 2.7, treating the stone as a perfect prism. Not a weighing."),
    (A_BUD, "dimensions", "triangular prism, sides 56, 46 and 60 cm, height 28 cm", None, None, None, "measured", "certain", None),
    (A_BUD, "volume", "about 34 dm3", 34, None, "dm3", "calculated", "certain", None),
    (A_BUD, "location", "Nedervetil, at the junction where the cart road up to Klockarbacken leaves the highway, on the south side of the highway",
     None, None, None, "reported", "certain", None),
    (A_BUD, "carry_distance", "carried by Tomas from near the church; OCR reads '3 4 km'", None, None, None, "traditional", "possible",
     "The OCR is ambiguous; Stejskal 1954 gives 3/4 km."),
    (A_BUD, "use", "lifted every year by confirmation-school boys; only a couple a year manage it (1929)", None, None, None, "reported", "certain", None),
    (A_STE.format("alaveteli1"), "weight", "about 92 kg", 92, None, "kg", "reported", "certain", "Probably taken from Budkavlen 1929."),
    (A_STE.format("alaveteli1"), "handling", "very hard to get a grip on", None, None, None, "reported", "certain", None),
    (A_STE.format("alaveteli1"), "dating", "in use a couple of hundred years, by hearsay (1954)", None, None, None, "traditional", "certain", None),
    (A_STE.format("alaveteli1"), "carry_distance", "about 3/4 km from near the church to its present place", 0.75, None, "km", "traditional", "certain", None),
    (A_STE.format("alaveteli1"), "location", "in Klockarbacken, Nedervetil church village, beside the road", None, None, None, "reported", "certain", None),
    (A_STE.format("teerijarvi1"), "weight", "between 150 and 200 kg", 150, 200, "kg", "reported", "certain", None),
    (A_STE.format("teerijarvi1"), "dating", "used for about 200 years (1954)", None, None, None, "traditional", "certain", None),
    (A_STE.format("teerijarvi1"), "shape", "sharp lower edge", None, None, None, "reported", "certain", None),
    (A_STE.format("teerijarvi1"), "handling", "rests on two small logs; lifted with a strong branch under its sharp lower edge, gripped at both ends",
     None, None, None, "reported", "certain", None),
    (A_STE.format("munsala1"), "location", "on the church hill (kyrkbacken) in Munsala", None, None, None, "reported", "certain", None),
    ("osterbottenstidning09062019:munsala1", "location", "below the church", None, None, None, "reported", "certain", None),
    ("osterbottenstidning09062019:munsala1", "use", "tradition long dead; confirmands of 1959 and 1970 did not take part", None, None, None, "reported", "certain", None),
    ("hulden1952:munsala2", "weight", "139 kg", 139, None, "kg", "reported", "certain", "Huldén does not say whether it was weighed."),
    ("hulden1952:munsala2", "shape", "elongated and smooth, the size of a pig (hence the name)", None, None, None, "reported", "certain", None),
    ("hulden1952:munsala2", "location", "Stubbåker, a field in Kantlax, Munsala", None, None, None, "reported", "probable",
     "Huldén places the carrying test in Stubbåker; Stejskal says 'in Kantlax in Munsala'."),
    (A_STE.format("munsala2"), "weight", "139 kg", 139, None, "kg", "reported", "certain", "Cites SNF 40-41, i.e. Huldén 1952."),
    (A_STE.format("oravainen1"), "location", "beside the church wall in Oravais", None, None, None, "reported", "certain", None),
    (A_STE.format("pietarsaari1"), "location", "behind the bell tower at Pedersöre church", None, None, None, "reported", "certain", None),
    (A_STE.format("uusikaarlepyy1"), "location", "town-hall yard (rådstugugården), Nykarleby", None, None, None, "reported", "certain", None),
    (A_STE.format("uusikaarlepyy1"), "dating", "there before the turn of the century (before c. 1900)", None, None, None, "reported", "certain", None),
    (A_STE.format("uusikaarlepyy1"), "fittings", "iron band, ring or hook", None, None, None, "reported", "probable",
     "Stejskal says such stones can have fittings and that 'such a stone' was in Nykarleby; read as this stone having them."),
    (A_STE.format("uusikaarlepyy1"), "use", "used as a pile-driving weight", None, None, None, "reported", "certain", None),
    (A_STE.format("jomala1"), "marking", "the year 1783 carved into it", 1783, None, "year", "reported", "certain", None),
    (A_STE.format("jomala1"), "shape", "egg-shaped (äggrund)", None, None, None, "reported", "certain", None),
    (A_STE.format("jomala1"), "location", "Björsby, Jomala, Åland", None, None, None, "reported", "certain", None),
    (A_STE.format("sipoo1"), "location", "Gästerby, 'Si.' (read as Sibbo/Sipoo)", None, None, None, "reported", "probable",
     "'Si' is Stejskal's parish code. Gästerby is a village in Sibbo and the map has the stone in Sipoo; an earlier reading as Sideby is probably wrong."),
    ("turistforeningsarsbok01011898:kangasniemi1", "dimensions", "150 cm high on the west side, 75 cm on the east side, 125 cm long, 60 cm thick",
     None, None, None, "reported", "certain", None),
    ("turistforeningsarsbok01011898:kangasniemi1", "location", "Synsiö village, a short way north-east of Vehkomäki farm, on a rock sloping towards the farm",
     None, None, None, "reported", "certain", None),
    ("turistforeningsarsbok01011898:kangasniemi1", "marking", "the stone's name indistinctly carved on its side", None, None, None, "reported", "certain", None),
    ("turistforeningsarsbok01011898:kangasniemi1", "carry_distance", "carried by Nuija-Heikki from the bog up onto the rock", None, None, None, "traditional", "certain", None),
    ("turistforeningsarsbok01011898:kangasniemi2", "location", "on Jalassyrjä farm land at Pirttiniemi, east shore of Siikajärvi", None, None, None, "reported", "certain", None),
    ("turistforeningsarsbok01011898:kangasniemi2", "marking", "inscription 'Otto Aatunpojan väkikivi'", None, None, None, "reported", "certain", None),
    ("itahame30041930:sysma1", "weight", "180.5 kg", 180.5, None, "kg", "reported", "certain",
     "Stated as 'weighing 180.5 kg'; the precision suggests a weighing, but the text does not say."),
    ("itahame30041930:sysma1", "shape", "smooth, spherical", None, None, None, "reported", "certain", None),
    ("itahame30041930:sysma1", "location", "deposited in Itä-Hämeen museo, Hartola, at a Sisu soirée the Sunday before 30.4.1930",
     None, None, None, "reported", "certain", None),
    ("hakkapeliitta29091931:sonkajarvi1", "carry_distance", "carried by Iso Rönkkö under his arm from Lehtoniemi field to the roadside at Kaijanmäki",
     None, None, None, "traditional", "certain", None),
    ("hakkapeliitta29091931:sonkajarvi1", "location", "Nostokiven kallio (Lifting Stone Rock), by the road at Kaijanmäki", None, None, None, "traditional", "certain", None),
    ("hakkapeliitta29091931:sonkajarvi1", "fate", "Jussi of Lähdeniemi had a horse drag the stone into a mud pit", None, None, None, "traditional", "certain", None),
    ("viikkosanomat30061923:ahtari1", "location", "by a fence in the village", None, None, None, "traditional", "certain", None),
]

# Customs: id, tag, lift_style, description, note
CUSTOMS = [
    ("confirmation_lift_nedervetil", "confirmation test", "off the ground",
     "Confirmation-school boys lift the stone; lifting it means you are grown up, failing means you came to confirmation school too early.", None),
    ("marriage_lift_terjarv", "marriage condition", "off the ground", "A man could not marry until he could lift the stone off the ground.", None),
    ("bachelor_lift_munsala", "bachelor test", None, "Boys lifted the stone to count as proper bachelors.", None),
    ("field_carry_munsala", "carrying test", "carry", "Carrying the stone across one field strip (teg) as a test of strength.", None),
    ("confirmation_lift_oravais", "confirmation test", "into the arms", "Confirmation-school stone lifted into the arms.", None),
    ("confirmation_lift_pedersore", "confirmation test", "off the ground", "Getting the stone off the ground passes the test.", None),
    ("confirmation_upright_nykarleby", "confirmation test", "upright", "Confirmation-school lads raised the stone to standing.", None),
    ("shoulder_throw_jomala", "strength feat", "shoulder and throw", "Lift the stone onto the shoulder and throw it backwards.", None),
    ("added_stones_sibbo", "added weight", None, "Extra stones laid on top of the lifting stone to add weight.", None),
    ("sunday_lifting_perho", "sunday sport", None, "Väkikivi lifting as a regular Sunday pastime; stones in farmyards worn smooth by hands.", None),
    ("women_lifting_perho", "women lifters", None, "Women competed with men in stone lifting; the Jylhä women were famous for it.", None),
    ("guest_challenge_ritola", "guest challenge", None, "Visitors rarely left Ritola before lifting the väkikivi with the master.", None),
    ("holy_evening_lifting_sonkajarvi", "strength feat", "off the ground", "Men competed with the stone on holy evenings (pyhäillat).", None),
]

# People: id, name, birth, death, approx, note
PERSONS = [
    ("tomas_masabacka", "Tomas of Måsabacka", 1800, None, 1, "Born c. 1800 according to one informant (Budkavlen 1929)."),
    ("poras_aija", "Poras-äijä", None, None, 0, "Woodsman who carried two full tar barrels at a time."),
    ("marjaana_jylha", "Marjaana Jylhä", None, None, 0, "Famous woman väkikivi lifter, photographed at 81 (Kotiliesi 1929, Kansan Kuvalehti 1931)."),
    ("matti_ritola", "Matti 'Ukko' Ritola", None, None, 0, "Skier, father of the Haapavesi ski model."),
    ("nuija_heikki", "Nuija-Heikki", None, None, 0, "Legendary."),
    ("otto_aatunpoika", "Otto Aatunpoika", None, None, 0, "Known only from the inscription."),
    ("antti_nousiainen", "Antti Nousiainen", 1800, 1823, 1, "Farmer, died 1823 aged 23."),
    ("siltalan_kustavi", "Kustavi of Siltala", None, None, 0, None),
    ("iso_ronkko", "Iso Rönkkö", None, None, 0, "Blacksmith."),
    ("johannes_viitala", "Johannes Viitala", None, None, 0, "Forest warden living at the Poranen homestead."),
    ("samuli_paulaharju", "Samuli Paulaharju", 1875, 1944, 0, "Folklorist and photographer."),
    ("tohni", "Tohni (Tohnisen äijä)", None, None, 0, "Strongman of Suupohja."),
    ("lemetti_jarvenpaa", "Lemetti Järvenpää", None, None, 0, "Farm master at Alapelto, Haapajärvi, 'two hundred years ago' (1915)."),
    ("perho_hostess", "Strong hostess of Perho (unnamed)", None, None, 0, None),
]

# Stories: id, title, summary, status, motifs, tellings (attestation ids), note
STORIES = [
    ("tomas_carries_stone", "Tomas carries the stone from the church",
     "Tomas of Måsabacka carried the stone from near Nedervetil church to the road junction, and it took his name.",
     "local_tradition", ["carried far", "eponym"], [A_BUD, A_STE.format("alaveteli1")],
     "Distance: OCR '3 4 km' in Budkavlen, 3/4 km in Stejskal."),
    ("nuija_heikki_carries_stone", "Nuija-Heikki carries the stone up the rock",
     "Nuija-Heikki carried the huge stone from the bog up onto the rock; the yearbook notes it is far too big for that.",
     "legend", ["carried far", "eponym", "strongman"], ["turistforeningsarsbok01011898:kangasniemi1"], None),
    ("poras_aija_fence_throw", "Poras-äijä throws the väkikivi over the fence",
     "Finding men straining at a great väkikivi, Poras-äijä threw it over the fence: 'A man ought not to bother lifting a thing like that!'",
     "local_tradition", ["thrown over a fence", "strongman"], ["paulaharju01011930:perho1", "kansankuvalehti16011931:perho1"], None),
    ("siltalan_kustavi_fence_throw", "Siltalan Kustavi throws the village stone",
     "Only two village men could move the väkikivi; young Kustavi of Siltala threw it over the fence, and the men took it badly.",
     "local_tradition", ["thrown over a fence", "unliftable by ordinary men", "strongman"],
     ["vaasa30051923:ahtari1", "viikkosanomat30061923:ahtari1", "canadanuutiset19071923:ahtari1"], None),
    ("antti_nousiainen_fence_throw", "Antti Nousiainen tosses the stone over the manor fence",
     "Antti Nousiainen secretly tossed the former väkikivi over the tall fence of the manor avenue.",
     "local_tradition", ["thrown over a fence", "strongman"], ["urheilija01011931:kangasniemi3"], None),
    ("iso_ronkko_stone", "Iso Rönkkö's stone on Nostokiven kallio",
     "Iso Rönkkö carried a stone under his arm from Lehtoniemi field and dropped it by the road: 'The man who lifts that is neither poor nor weak.' "
     "The rock was named after it; Taavetti of Rönkkölä lifted it, but not with a folding knife on top. Jussi of Lähdeniemi had it dragged into a mud pit.",
     "local_tradition", ["carried far", "strongman", "added weight", "stone moved or destroyed"], ["hakkapeliitta29091931:sonkajarvi1"], None),
    ("perho_hostess_lift", "The pregnant hostess of Perho",
     "A Perho hostess lifted the village's largest väkikivi while pregnant; her daughter later lifted it with a second stone on top.",
     "local_tradition", ["woman lifter", "added weight"], ["pohjolanhelleenit01012007:context"], "From Paulaharju, via Pohjolan helleenit 2007."),
    ("lemetti_vakikivi", "Lemetti's strength lives in his väkikivi",
     "Lemetti Järvenpää of Haapajärvi, dead two hundred years, still lives on in his väkikivi, now the porch step of the old Istula farm.",
     "legend", ["strongman", "eponym", "stone moved or destroyed"], ["liitto14121915:context"], "Stone 'Lemmetin väkikivi' has no record yet."),
    ("tohni_church_stone", "Tohni rolls a stone onto the church wall",
     "Tohni took part in building Kyrö church and rolled onto its wall, alone, a stone ten men could not move.",
     "local_tradition", ["strongman", "unliftable by ordinary men"], ["suomenurheilulehti30041914:alavus1"], None),
    ("peltoniemen_paavo", "Peltoniemen Paavo",
     "A fictional character's youthful hobby of lifting väkikivi.", "fiction", ["strongman"],
     ["kauppisheikki1912:context", "savonsanomat16011908:context", "kaiku15121912:context"],
     "Same story printed three times; not evidence of any stone."),
]

# Links: from_type, from_id, role, to_type, to_id, attestation, status, confidence, note
P_BUD, P_STE = A_BUD, A_STE
LINKS = [
    # people and stones
    ("person", "tomas_masabacka", "eponym_of", "stone", "alaveteli1", A_BUD, "local_tradition", "certain", None),
    ("person", "tomas_masabacka", "carried", "stone", "alaveteli1", A_BUD, "local_tradition", "certain", None),
    ("person", "tomas_masabacka", "lived_at", "place", "masabacka", A_BUD, "local_tradition", "certain", None),
    ("person", "poras_aija", "lifted", "stone", "perho1", "paulaharju01011930:perho1", "local_tradition", "probable", "Which stone is inferred."),
    ("person", "poras_aija", "lived_at", "place", "poranen", "paulaharju01011930:perho1", "local_tradition", "probable", "Inferred from the name."),
    ("person", "marjaana_jylha", "lifted", "stone", "perho2", "kotiliesi19290815:perho2", "documented", "possible", "Called a lifter of väkikivet; no stone named."),
    ("person", "marjaana_jylha", "lived_at", "place", "jylha", "kotiliesi19290815:perho2", "documented", "certain", "Married into Jylhä."),
    ("person", "marjaana_jylha", "born_at", "place", "nopola", "kotiliesi19290815:perho2", "documented", "certain", None),
    ("person", "matti_ritola", "host_of", "stone", "haapavesi1", "suomenurheilulehti01041911:haapavesi1", "documented", "certain", None),
    ("person", "matti_ritola", "lived_at", "place", "ritola", "suomenurheilulehti01041911:haapavesi1", "documented", "certain", None),
    ("person", "nuija_heikki", "eponym_of", "stone", "kangasniemi1", "turistforeningsarsbok01011898:kangasniemi1", "local_tradition", "certain", None),
    ("person", "nuija_heikki", "carried", "stone", "kangasniemi1", "turistforeningsarsbok01011898:kangasniemi1", "legend", "certain", None),
    ("person", "otto_aatunpoika", "eponym_of", "stone", "kangasniemi2", "turistforeningsarsbok01011898:kangasniemi2", "documented", "certain", "Name inscribed on the stone."),
    ("person", "antti_nousiainen", "lifted", "stone", "kangasniemi3", "urheilija01011931:kangasniemi3", "local_tradition", "certain", None),
    ("person", "siltalan_kustavi", "lifted", "stone", "ahtari1", "viikkosanomat30061923:ahtari1", "local_tradition", "certain", None),
    ("person", "siltalan_kustavi", "lived_at", "place", "siltala", "viikkosanomat30061923:ahtari1", "local_tradition", "certain", None),
    ("person", "iso_ronkko", "carried", "stone", "sonkajarvi1", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
    ("person", "iso_ronkko", "lived_at", "place", "ronkkola", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
    ("person", "johannes_viitala", "lifted", "stone", "perho1", "perhojokilaakso13012021:perho1", "documented", "probable", "1932 photo; stone inferred."),
    ("person", "johannes_viitala", "lived_at", "place", "poranen", "perhojokilaakso13012021:perho1", "documented", "certain", None),
    ("person", "samuli_paulaharju", "photographed", "stone", "perho1", "valokuvapaulaharju1928:perho1", "documented", "certain", None),
    ("person", "samuli_paulaharju", "photographed", "stone", "perho1", "perhojokilaakso13012021:perho1", "documented", "probable", None),
    ("person", "tohni", "carried", "stone", "alavus1", "suomenurheilulehti30041914:alavus1", "local_tradition", "possible", "See alavus1 note."),
    ("person", "lemetti_jarvenpaa", "lived_at", "place", "haapajarvi", "liitto14121915:context", "legend", "certain", None),
    ("person", "perho_hostess", "lived_at", "place", "perho", "pohjolanhelleenit01012007:context", "local_tradition", "certain", None),
    # stories
    ("story", "tomas_carries_stone", "about", "stone", "alaveteli1", A_BUD, "local_tradition", "certain", None),
    ("story", "tomas_carries_stone", "features", "person", "tomas_masabacka", A_BUD, "local_tradition", "certain", None),
    ("story", "tomas_carries_stone", "set_in", "place", "nedervetil_church", A_STE.format("alaveteli1"), "local_tradition", "certain", None),
    ("story", "nuija_heikki_carries_stone", "about", "stone", "kangasniemi1", "turistforeningsarsbok01011898:kangasniemi1", "legend", "certain", None),
    ("story", "nuija_heikki_carries_stone", "features", "person", "nuija_heikki", "turistforeningsarsbok01011898:kangasniemi1", "legend", "certain", None),
    ("story", "nuija_heikki_carries_stone", "set_in", "place", "vehkomaki", "turistforeningsarsbok01011898:kangasniemi1", "legend", "certain", None),
    ("story", "poras_aija_fence_throw", "about", "stone", "perho1", "paulaharju01011930:perho1", "local_tradition", "probable", None),
    ("story", "poras_aija_fence_throw", "features", "person", "poras_aija", "paulaharju01011930:perho1", "local_tradition", "certain", None),
    ("story", "poras_aija_fence_throw", "set_in", "place", "perho", "paulaharju01011930:perho1", "local_tradition", "certain", None),
    ("story", "siltalan_kustavi_fence_throw", "about", "stone", "ahtari1", "viikkosanomat30061923:ahtari1", "local_tradition", "certain", None),
    ("story", "siltalan_kustavi_fence_throw", "features", "person", "siltalan_kustavi", "viikkosanomat30061923:ahtari1", "local_tradition", "certain", None),
    ("story", "antti_nousiainen_fence_throw", "about", "stone", "kangasniemi3", "urheilija01011931:kangasniemi3", "local_tradition", "certain", None),
    ("story", "antti_nousiainen_fence_throw", "features", "person", "antti_nousiainen", "urheilija01011931:kangasniemi3", "local_tradition", "certain", None),
    ("story", "iso_ronkko_stone", "about", "stone", "sonkajarvi1", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
    ("story", "iso_ronkko_stone", "features", "person", "iso_ronkko", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
    ("story", "iso_ronkko_stone", "set_in", "place", "nostokiven_kallio", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
    ("story", "perho_hostess_lift", "features", "person", "perho_hostess", "pohjolanhelleenit01012007:context", "local_tradition", "certain", None),
    ("story", "perho_hostess_lift", "set_in", "place", "perho", "pohjolanhelleenit01012007:context", "local_tradition", "certain", None),
    ("story", "lemetti_vakikivi", "features", "person", "lemetti_jarvenpaa", "liitto14121915:context", "legend", "certain", None),
    ("story", "lemetti_vakikivi", "set_in", "place", "istula", "liitto14121915:context", "legend", "certain", None),
    ("story", "tohni_church_stone", "about", "stone", "alavus1", "suomenurheilulehti30041914:alavus1", "local_tradition", "possible", "See alavus1 note."),
    ("story", "tohni_church_stone", "features", "person", "tohni", "suomenurheilulehti30041914:alavus1", "local_tradition", "certain", None),
    ("story", "tohni_church_stone", "set_in", "place", "kyro_church", "suomenurheilulehti30041914:alavus1", "local_tradition", "certain", None),
    # customs
    ("custom", "confirmation_lift_nedervetil", "practised_at", "stone", "alaveteli1", A_BUD, "documented", "certain", "Observed 1929."),
    ("custom", "confirmation_lift_nedervetil", "practised_in", "place", "nedervetil", A_BUD, "documented", "certain", None),
    ("custom", "marriage_lift_terjarv", "practised_at", "stone", "teerijarvi1", P_STE.format("teerijarvi1"), "local_tradition", "certain", "'The old people claim' (1954)."),
    ("custom", "bachelor_lift_munsala", "practised_at", "stone", "munsala1", P_STE.format("munsala1"), "documented", "certain", None),
    ("custom", "field_carry_munsala", "practised_at", "stone", "munsala2", "hulden1952:munsala2", "documented", "certain", None),
    ("custom", "confirmation_lift_oravais", "practised_at", "stone", "oravainen1", P_STE.format("oravainen1"), "documented", "certain", None),
    ("custom", "confirmation_lift_pedersore", "practised_at", "stone", "pietarsaari1", P_STE.format("pietarsaari1"), "documented", "certain", None),
    ("custom", "confirmation_upright_nykarleby", "practised_at", "stone", "uusikaarlepyy1", P_STE.format("uusikaarlepyy1"), "documented", "certain", None),
    ("custom", "shoulder_throw_jomala", "practised_in", "place", "jomala", P_STE.format("jomala1"), "documented", "certain", None),
    ("custom", "shoulder_throw_jomala", "practised_at", "stone", "jomala1", P_STE.format("jomala1"), "documented", "possible",
     "Stejskal gives this for Jomala parish (Jo), not explicitly for the Björsby stone."),
    ("custom", "added_stones_sibbo", "practised_at", "stone", "sipoo1", P_STE.format("sipoo1"), "documented", "certain", None),
    ("custom", "sunday_lifting_perho", "practised_in", "place", "perho", "paulaharju01011930:perho1", "documented", "certain", None),
    ("custom", "sunday_lifting_perho", "practised_at", "stone", "perho1", "paulaharju01011930:perho1", "documented", "possible", None),
    ("custom", "women_lifting_perho", "practised_in", "place", "perho", "paulaharju01011930:perho1", "documented", "certain", None),
    ("custom", "women_lifting_perho", "practised_at", "stone", "perho2", "pohjolanhelleenit01012007:perho2", "local_tradition", "certain", None),
    ("custom", "guest_challenge_ritola", "practised_at", "stone", "haapavesi1", "suomenurheilulehti01041911:haapavesi1", "documented", "certain", None),
    ("custom", "holy_evening_lifting_sonkajarvi", "practised_at", "stone", "sonkajarvi1", "hakkapeliitta29091931:sonkajarvi1", "local_tradition", "certain", None),
]

MEDIA = [
    # id, stone, source, kind, archive_id, url, year, photographer, licence, caption, confidence, note
    ("paulaharju1928_perho1", "perho1", "valokuvapaulaharju1928", "photo", "museovirasto.8319F7526E14A0B65CD5DCF49C0369E9",
     "https://www.finna.fi/Record/museovirasto.8319F7526E14A0B65CD5DCF49C0369E9", 1928, "Samuli Paulaharju", None, None, "certain",
     "Check whether this is the same picture as the 1932 Viitala photo."),
    ("paulaharju1932_viitala", "perho1", "perhojokilaakso13012021", "photo", None, None, 1932, "Samuli Paulaharju", None,
     "The väkikivi rises in the hands of Johannes Viitala; the boy is Aatos Viitala.", "probable",
     "Archive record not found yet; look in the Finnish Heritage Agency picture archive."),
    ("budkavlen1929_fig", "alaveteli1", "budkavlen01011929", "figure", None, "https://digi.kansalliskirjasto.fi/aikakausi/binding/917722?page=8",
     1929, None, None, "Photo printed with the article ('the stone depicted above').", "certain", None),
    ("stejskal1954_fig26", "teerijarvi1", "stejskal1954", "figure", "fig. 26", None, None, None, None, "Risbacka lyftsten, D 1. Terjärv.", "certain", None),
    ("stejskal1954_fig27", "jomala1", "stejskal1954", "figure", "fig. 27", None, None, None, None, "Björsby lyftsten (daterad 1783), D 1. Jomala.", "certain", None),
    ("ekrem2020_risbacka1932", "teerijarvi1", "ekrem2020", "photo", None, None, 1932, None, None, "Risbacka lifting stone, photo from 1932.", "certain",
     "Per the source analysis report."),
    ("virtaranta1960_photo", None, "virtaranta1960", "photo", "museovirasto.858E9FFFDE339CDD18429EC7CABD505A",
     "https://www.finna.fi/Record/museovirasto.858E9FFFDE339CDD18429EC7CABD505A", 1960, "Helmi Virtaranta", None, "Ysö-Pekan väkikivi", "certain",
     "Stone has no record yet."),
] + [
    (f"aaltonen1950{x}_photo", None, f"aaltonen1950{x}", "photo", None, None, 1950, "Esko Aaltonen", None, "Juuka", "certain", None)
    for x in "abcde"
]


def page_of(url):
    if not url:
        return None
    q = parse_qs(urlparse(url).query)
    return q.get("page", [None])[0]


def source_type(sid, url, notes):
    u = url or ""
    if "/sanomalehti/" in u:
        return "newspaper"
    if "/aikakausi/" in u:
        return "periodical"
    if "/teos/" in u or "helka.helsinki" in u or "sls.fi/publications" in u or "doria.fi" in u or "yhdistysavain" in u:
        return "book"
    if "academia.edu" in u:
        return "thesis"
    if "finna.fi" in u:
        return "photo_record" if (notes or "").startswith("Valokuva") else "archive_record"
    if "kyppi.fi" in u or "yksa.disec" in u:
        return "archive_record"
    if "kinoregina" in u:
        return "film"
    if "osterbottenstidning" in u or "perhonjokilaakso" in u:
        return "newspaper"
    if not u and re.search(r"\d{8}$", sid):
        return "newspaper"
    return "web"


def clean_name(n):
    if not n:
        return None
    n = re.sub(r"\(==.*?\)", "", n).replace("?", "").strip()
    return n or None


def write(name, rows):
    path = OUT / f"{name}.csv"
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS[name], lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r.get(k) is None else r.get(k)) for k in FIELDS[name]})
    print(f"  {name:13s} {len(rows):4d} rows")


def main():
    if any(OUT.glob("*.csv")) and "--force" not in sys.argv:
        raise SystemExit("data/csv already exists and is the master copy. Use --force to overwrite it.")
    OUT.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(SHEET)
    sheet_stones = [dict(zip([c.value for c in wb["stone_location"][1]], r))
                    for r in wb["stone_location"].iter_rows(min_row=2, values_only=True)]
    src_rows = [dict(zip([c.value for c in wb["source"][1]], r))
                for r in wb["source"].iter_rows(min_row=2, values_only=True)]
    map_stones = json.loads(MAP.read_text(encoding="utf-8"))["stones"]

    # ---- sources ----
    sources = {}
    for r in src_rows:
        sid = r["source_id"]
        s = sources.get(sid)
        if s is None:
            s = sources[sid] = dict(
                source_id=sid, source_type=source_type(sid, r["source_url"], r["notes_raw_fi"]),
                author=r["source_author"], pub_date=r["source_date"], url=r["source_url"],
                accessed=r["cite_date"], cite_ref=r["cite_ref"], _orig=set(), _notes=[], _lang=set())
        s["author"] = s["author"] or r["source_author"]
        s["pub_date"] = s["pub_date"] or r["source_date"]
        s["cite_ref"] = s["cite_ref"] or r["cite_ref"]
        if r["cite_date"] and (not s["accessed"] or r["cite_date"] < s["accessed"]):
            s["accessed"] = r["cite_date"]
        if r["original_source_id"]:
            s["_orig"].add(r["original_source_id"].strip())
        for part in (r["cleaning_note"] or "").split("  "):
            if part.startswith("DATE CHECK"):
                s["date_note"] = part.replace("DATE CHECK: ", "")
        if r["notes_raw_fi"]:
            s["_lang"].add("fi")
        if r["notes_raw_sv"]:
            s["_lang"].add("sv")
    for s in sources.values():
        s.update({k: v for k, v in SOURCE_META.get(s["source_id"], {}).items()})
        s["original_ids"] = ";".join(sorted(s.pop("_orig") - {s["source_id"]})) or None
        s.pop("_notes")
        lang = s.pop("_lang")
        s["language"] = "sv" if lang == {"sv"} else ("fi" if lang == {"fi"} else None)
        if s.get("pub_date"):
            s["date_precision"] = "year" if s["pub_date"].endswith("-01-01") else "day"
    sources[SHEET_SRC] = dict(source_id=SHEET_SRC, source_type="dataset", title="Finnish stones sources sheet (manually gathered)",
                              pub_date="2026-01-09", date_precision="day",
                              url="data/sources/2026-01-09-finnish-stones-sources_cleaned.xlsx",
                              note="The project's own research log. Basis for stone status, condition and coordinates.")
    sources[MAP_SRC] = dict(source_id=MAP_SRC, source_type="dataset", title="Stone map data (docs/stones.json)",
                            pub_date="2026-05-30", date_precision="day", url="docs/stones.json",
                            note="Coordinates placed by the researcher for the web map; also entry_data/2026-05-30-finstones_mun.geoJSON.")
    sources[REPORT_SRC] = dict(source_id=REPORT_SRC, source_type="analysis", title="Finnish lifting stones: source analysis",
                               pub_date="2026-11-27", date_precision="day", url="https://claude.ai/artifact/7BbyoeSodsiXPLsMRpkVrp",
                               note="Cross-reading of the sheet and ten excerpts. Basis for the reprint links and the Perho split.")

    # ---- stones ----
    stones = []
    for s in sheet_stones:
        sid = s["stone_location_id"]
        ctx = s["context"]
        if sid == "uusikaarlepyy1":
            typ = "junttasten"
        elif ctx == "skriftskolsten":
            typ = "skriftskolsten"
        elif s["region"] in ("Teerijärvi", "Munsala", "Jomala", "Sipoo") or ctx == "lyftsten":
            typ = "lyftsten"
        else:
            typ = "väkikivi"
        note = s["cleaning_note"]
        if ctx and ctx.endswith("?"):
            note = ((note + " ") if note else "") + f"Sheet context '{ctx}' was uncertain; type set to what the sources call it."
        if sid in PREFERRED_NAME:
            name, why = PREFERRED_NAME[sid]
            note = ((note + " ") if note else "") + why
        else:
            name = s["name"]
        stones.append(dict(
            stone_id=sid, preferred_name=name, stone_type=typ, status=s["status"] or "unknown",
            condition=s["condition"], access=s["location_type"] if s["location_type"] in ("public", "private", "museum") else None,
            sheet_context=ctx,
            record_confidence="probable" if sid == "perho2" else "certain",
            source_id={"jomala1": "stejskal1954", "sipoo1": "stejskal1954", "perho2": "pohjolanhelleenit01012007"}.get(sid, SHEET_SRC),
            note=note))
    for sid, (name, typ, region, _, note) in MAP_ONLY.items():
        stones.append(dict(stone_id=sid, preferred_name=name, stone_type=typ, status="unknown", record_confidence="certain",
                           source_id=MAP_SRC, note=("On the map but not yet in the sources sheet. " + (note or "")).strip()))
    stone_ids = {s["stone_id"] for s in stones}

    # ---- attestations ----
    atts, used = [], set()
    for r in src_rows:
        stone = r["stone_location_id"]
        conf, note = "certain", None
        if r["original_row"] in DB_ONLY_LINKS and not stone:
            stone, conf, note = DB_ONLY_LINKS[r["original_row"]]
        if (r["source_id"], stone) in LINK_CONFIDENCE:
            conf, note = LINK_CONFIDENCE[(r["source_id"], stone)]
        base = f"{r['source_id']}:{stone or 'context'}"
        aid, n = base, 1
        while aid in used:
            n += 1
            aid = f"{base}-{n}"
        used.add(aid)
        src_url = sources[r["source_id"]]["url"]
        cn = r["cleaning_note"]
        if cn and not cn.startswith("DATE CHECK"):
            note = ((note + " ") if note else "") + cn
        atts.append(dict(attestation_id=aid, source_id=r["source_id"], stone_id=stone,
                         mentioned_name=clean_name(r["stone_name_fi"] or r["stone_name_sv"]),
                         page=page_of(r["source_url"]), url=r["source_url"] if r["source_url"] != src_url else None,
                         excerpt_fi=r["notes_raw_fi"], excerpt_sv=r["notes_raw_sv"], excerpt_en=r["notes_en"],
                         confidence=conf, note=note))
    for sid, stone, page, note in EXTRA_ATTESTATIONS:
        aid = f"{sid}:{stone}"
        assert aid not in used
        used.add(aid)
        atts.append(dict(attestation_id=aid, source_id=sid, stone_id=stone, page=page, confidence="certain", note=note))
    for a in atts:
        if a["source_id"] == "kotiliesi19290815":
            a["page"] = "12"
            a["note"] = (a["note"] or "") + " Page 12 per the uploaded excerpt; the URL points to page 14."

    # ---- names ----
    names, seen = [], set()

    def add_name(stone, name, lang, src, note=None):
        name = clean_name(name)
        if not stone or not name or stone not in stone_ids:
            return
        key = (stone, name.lower(), src)
        if key in seen:
            return
        seen.add(key)
        names.append(dict(stone_id=stone, name=name, language=lang, source_id=src, note=note))

    for s in sheet_stones:
        add_name(s["stone_location_id"], s["name"], "sv" if s["region"] in ("Alaveteli", "Teerijärvi", "Munsala", "Oravainen",
                 "Pietarsaari", "Uusikaarlepyy", "Jomala", "Sipoo") else "fi", SHEET_SRC)
    for m in map_stones:
        add_name(MAP_TO_STONE[m["name"]], m["name"], None, MAP_SRC)
    for a, r in zip(atts, src_rows):
        if r["stone_name_fi"]:
            add_name(a["stone_id"], r["stone_name_fi"], "fi", r["source_id"],
                     "May be another name for Nuija-Heikin väkikivi." if a["confidence"] == "possible" and "Ison" in r["stone_name_fi"] else None)
        if r["stone_name_sv"]:
            add_name(a["stone_id"], r["stone_name_sv"], "sv", r["source_id"])
    add_name("alaveteli1", "Tomasasstäinä", "dialect", "budkavlen01011929", "Dialect form as printed in Budkavlen.")
    add_name("alaveteli1", "Tomas' sten", "sv", "stejskal1954")
    add_name("teerijarvi1", "Risbacka lyftsten", "sv", "stejskal1954", "Figure caption.")
    add_name("perho1", "Porasen väkikivi", "fi", REPORT_SRC, "Descriptive name used in the analysis.")

    # ---- locations ----
    locs = []
    map_by_stone = {MAP_TO_STONE[m["name"]]: m for m in map_stones}

    def region_of(stone):
        for s in sheet_stones:
            if s["stone_location_id"] == stone:
                return s["region"]
        return MAP_ONLY[stone][2]

    def place_of(sid):
        return STONE_PLACE.get(sid) or REGION[region_of(sid)]

    for s in sheet_stones:
        sid = s["stone_location_id"]
        if sid in ("perho2", "jomala1", "sipoo1"):
            continue  # their sheet coordinates were copied from the map
        if s["euref_e"] is None:
            continue
        locs.append(dict(location_id=f"{sid}-sheet", stone_id=sid, place_id=place_of(sid),
                         easting=int(s["euref_e"]), northing=int(s["euref_n"]),
                         point_marks=s["coord_ref"], is_current=0 if sid in map_by_stone else 1, source_id=SHEET_SRC,
                         confidence="certain"))
    for sid, m in map_by_stone.items():
        conf, note = "certain", None
        if sid == "teerijarvi2":
            conf, note = "probable", "Historic parish is a guess."
        if sid == "sysma1":
            note = "Moved by the researcher on 2026-05-30 (commit 'edit mattilan mummu location')."
        if any(l["stone_id"] == sid for l in locs):
            note = ((note + " ") if note else "") + "Preferred over the sheet point because the map is newer (guess)."
        locs.append(dict(location_id=f"{sid}-map", stone_id=sid, place_id=place_of(sid),
                         easting=round(m["easting"]), northing=round(m["northing"]),
                         point_marks=None, is_current=1, source_id=MAP_SRC, confidence=conf, note=note))
    locs.append(dict(location_id="sysma1-museum1930", stone_id="sysma1", place_id="itahamen_museo", valid_from="1930-04-27",
                     description="Courtyard of Itä-Hämeen museo, Hartola", point_marks="museum", is_current=0,
                     source_id="itahame30041930", confidence="certain",
                     note="Deposited at a soirée 'last Sunday' before 30.4.1930. It later returned to Sysmä; the return is not yet sourced."))
    locs.append(dict(location_id="alaveteli1-original", stone_id="alaveteli1", place_id="nedervetil_church",
                     description="Near the church, Nedervetil", is_current=0, source_id="stejskal1954", confidence="certain",
                     note="Before Tomas of Måsabacka (born c. 1800) carried it about 3/4 km to the road junction."))

    # ---- places ----
    places = [dict(place_id=p, name=n, name_alt=alt, kind="parish", municipality_fi=mfi, municipality_sv=msv,
                   municipality_code=code) for p, n, alt, mfi, msv, code in PARISHES]
    places += [dict(place_id=p, name=n, kind=k, parent_id=par, note=note) for p, n, k, par, note in SUBPLACES]
    place_ids = {p["place_id"] for p in places}
    assert all(l["place_id"] in place_ids for l in locs if l.get("place_id")), [l["place_id"] for l in locs]

    # ---- claims, records and links ----
    att_ids = {a["attestation_id"] for a in atts}

    def claim_status(aid, field, method):
        if (aid, field) in CLAIM_STATUS_OVERRIDE:
            return CLAIM_STATUS_OVERRIDE[(aid, field)]
        return {"weighed": "measured", "measured": "measured", "calculated": "measured",
                "traditional": "local_tradition"}.get(method, "documented")

    claims = []
    for n, (aid, field, value, num, nmax, unit, method, conf, note) in enumerate(CLAIMS, 1):
        assert aid in att_ids, aid
        claims.append(dict(claim_id=f"c{n:04d}", attestation_id=aid, stone_id=aid.split(":")[1].split("-")[0], field=field,
                           value=value, value_num=num, value_num_max=nmax, unit=unit, method=method,
                           status=claim_status(aid, field, method), confidence=conf, note=note))
    customs = [dict(zip(FIELDS["custom"], c)) for c in CUSTOMS]
    persons = [dict(zip(FIELDS["person"], p)) for p in PERSONS]
    stories, motifs, tellings = [], [], []
    for sid, title, summary, status, mots, tells, note in STORIES:
        stories.append(dict(story_id=sid, title=title, summary=summary, status=status, note=note))
        motifs += [dict(story_id=sid, motif=m) for m in mots]
        for t in tells:
            assert t in att_ids, t
            tellings.append(dict(story_id=sid, attestation_id=t))
    links = []
    for n, row in enumerate(LINKS, 1):
        d = dict(zip(FIELDS["link"][1:], row))
        assert d["attestation_id"] in att_ids, row
        links.append(dict(link_id=f"l{n:04d}", **d))
    media = [dict(zip(FIELDS["media"], m)) for m in MEDIA]

    print("wrote data/csv/:")
    for old in OUT.glob("*.csv"):
        old.unlink()
    write("source", sorted(sources.values(), key=lambda s: s["source_id"]))
    write("attestation", atts)
    write("place", places)
    write("stone", sorted(stones, key=lambda s: s["stone_id"]))
    write("person", persons)
    write("story", stories)
    write("story_motif", motifs)
    write("story_telling", tellings)
    write("custom", customs)
    write("link", links)
    write("stone_name", names)
    write("location", locs)
    write("claim", claims)
    write("media", media)


if __name__ == "__main__":
    main()
