library(sf)
library(jsonlite)

geo <- st_read("entry_data/2026-05-30-finstones_mun.GEOjson", quiet = TRUE)
wgs <- st_transform(geo, 4326)
coords_tm  <- st_coordinates(geo)
coords_wgs <- st_coordinates(wgs)

stones <- lapply(seq_len(nrow(geo)), function(i) {
  list(
    id        = i,
    name      = geo$stone_name_fi[i],
    location  = geo$kunta_name_fi[i],
    weight_kg = NULL,
    historical = "",
    lat       = round(coords_wgs[i, 2], 6),
    lon       = round(coords_wgs[i, 1], 6),
    northing  = round(coords_tm[i, 2]),
    easting   = round(coords_tm[i, 1]),
    notes     = ""
  )
})

out <- list(stones = stones)
write(toJSON(out, pretty = TRUE, auto_unbox = TRUE, null = "null"), "docs/stones.json")
cat("Done:", nrow(geo), "stones written\n")
