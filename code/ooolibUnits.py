#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import re

###########################
# ooolib-python Units     #
###########################
# Number of each unit in one inch
UNITS_PER_INCH = {
    "in": 1.0,
    "cm": 2.54,
    "mm": 25.4,
    "pt": 72.0,
    "pc": 6.0,
    "px": 96.0,
}

lengthPattern = re.compile(r"^\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)\s*([a-z]*)\s*$")

def parseLength(length):
    """Split a length string such as '2.5in' into (2.5, 'in').

    Returns None if the length can not be parsed."""
    if length is None: return None
    match = lengthPattern.match(str(length))
    if match is None: return None
    return (float(match.group(1)), match.group(2))

def convertLength(length, unit="in", defaultUnit="in"):
    """Convert a length string (e.g. '2.5in', '30mm') to a float in the target unit.

    Numbers without a unit are treated as defaultUnit. Returns None if the
    length can not be parsed or uses an unknown unit."""
    parsed = parseLength(length)
    if parsed is None: return None
    value, fromUnit = parsed
    if fromUnit == "": fromUnit = defaultUnit
    if fromUnit not in UNITS_PER_INCH: return None
    if unit not in UNITS_PER_INCH: return None
    return value / UNITS_PER_INCH[fromUnit] * UNITS_PER_INCH[unit]

def isUnit(unit):
    return unit in UNITS_PER_INCH

def formatNumber(value):
    """Format a number without trailing zeros (2.50 -> '2.5', 3.0 -> '3')."""
    text = "%.4f" % value
    text = text.rstrip("0").rstrip(".")
    if text == "-0": text = "0"
    return text

def formatLength(value, unit="in"):
    """Return an ODF length string.

    Numbers are given the unit (2.5 -> '2.5in'). Strings are returned
    unchanged ('30mm' stays '30mm'). None is returned as None."""
    if value is None: return None
    if isinstance(value, bool):
        raise TypeError("Length must be a number or string, not %r" % value)
    if isinstance(value, (int, float)):
        return formatNumber(value) + unit
    return str(value)


###################
# Execute to Test #
###################
if __name__ == "__main__":
    print(convertLength("1in", "cm"))      # 2.54
    print(convertLength("25.4mm", "in"))   # 1.0
    print(convertLength("72pt", "in"))     # 1.0
    print(convertLength("3", "mm"))        # 76.2
    print(formatLength(2.50, "in"))        # 2.5in
    print(formatLength("30mm", "in"))      # 30mm
