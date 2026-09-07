"""Normalise Unleashed boolean-ish export values to the sheet's YES/NO."""

import pandas as pd

from constants.order_sheet import YesNo, UNLEASHED_TRUE_VALUES


def unleashed_flag_to_yesno(series: pd.Series) -> pd.Series:
    """Map an Unleashed boolean column (True/False, Yes/No, 1/0, Y/N, any case)
    to YesNo.YES / YesNo.NO. Blank or unrecognised values become NO."""
    is_on = series.astype(str).str.strip().str.lower().isin(UNLEASHED_TRUE_VALUES)
    return is_on.map({True: YesNo.YES, False: YesNo.NO})
