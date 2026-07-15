"""Shared constants for reading CSV uploads."""

# Encodings tried in order when decoding an uploaded CSV.
# Unleashed / Excel exports on Windows are usually Windows-1252 (cp1252): bytes
# like 0x93/0x94 are "smart quotes" that break pandas' default UTF-8 decode.
# - utf-8-sig  : plain UTF-8, also strips a byte-order mark if present.
# - cp1252     : the actual encoding of most Unleashed/Excel exports.
# - latin-1    : never raises on any byte, so it's the guaranteed last resort.
CSV_ENCODINGS = ("utf-8-sig", "cp1252", "latin-1")
