"""Shared constants for the UNL Stock Ordering Sheet."""


class PlistColumn:
    """Column names in the Unleashed Product List export (input 4)."""
    PRODUCT_CODE = "*Product Code"
    WEIGHT = "Weight"
    IS_PURCHASABLE = "Is Purchasable"


class PoProductColumn:
    """Column names in the Unleashed PO Product Data export (input 1)."""
    OBSOLETE = "Obsolete"


class OutputColumn:
    PRODUCT_CODE = "Product Code"
    UNL_PURCHASABLE = "UNL Purchasable"
    OBSOLETE = "Obsolete"


class YesNo:
    YES = "YES"
    NO = "NO"


# Values Unleashed uses (case-insensitive) for a switched-off boolean flag.
UNLEASHED_FALSE_VALUES = frozenset({"no", "false", "0", "n"})
# Values Unleashed uses (case-insensitive) for a switched-on boolean flag.
UNLEASHED_TRUE_VALUES = frozenset({"yes", "true", "1", "y"})


class HighlightColor:
    # Special-order rows.
    SPECIAL_ORDER_BG = "#FFFF00"
    # Not-purchasable rows (Excel "Bad" style: light red fill, dark red text).
    NOT_PURCHASABLE_BG = "#FFC7CE"
    NOT_PURCHASABLE_FONT = "#9C0006"
