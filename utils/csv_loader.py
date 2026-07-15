"""Encoding-robust CSV reading for uploaded files."""

import io

import pandas as pd

from constants.csv import CSV_ENCODINGS


def read_csv_bytes(data: bytes, **kwargs) -> pd.DataFrame:
    """Read CSV bytes, trying each encoding in CSV_ENCODINGS until one decodes.

    pandas' read_csv defaults to UTF-8, which raises UnicodeDecodeError on the
    Windows-1252 characters (e.g. curly quotes / byte 0x93) that Unleashed and
    Excel exports contain. Any extra kwargs (skiprows, etc.) are passed through.
    """
    last_err = None
    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(io.BytesIO(data), encoding=encoding, **kwargs)
        except UnicodeDecodeError as err:
            last_err = err
    # latin-1 decodes any byte sequence, so this is effectively unreachable;
    # re-raise the last error rather than silently returning None.
    raise last_err
