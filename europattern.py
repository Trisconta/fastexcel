""" Euro pattern
"""

import re

EURO_NUMBER_PATTERN = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")

def normalize_number(value: str):
    """ Normalize European-formatted numbers like '36.423,19' into '36423.19'.
    Only applies the transformation if the string matches the expected pattern.
    """
    if not isinstance(value, str):
        return value
    if not EURO_NUMBER_PATTERN.match(value):
        return None
    # Remove thousand separators
    res = value.replace(".", "")
    # Replace decimal comma with decimal point
    return res.replace(",", ".")
