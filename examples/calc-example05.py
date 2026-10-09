#!/usr/bin/python

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Calc()

# Standard Cell Properties
# setCellBold(row, col, value=True)
# setCellItalics(row, col, value=True)
# setCellUnderline(row, col, value=True)
doc.content.activeTable.setCellText(1, 1, "Normal Text")

doc.content.activeTable.setCellText(1, 2, "Bold Text")
doc.content.activeTable.setCellBold(1, 2, True) # True for Bold on, False for off

doc.content.activeTable.setCellText(1, 3, "Italic Text")
doc.content.activeTable.setCellItalics(1, 3) # True is default, therefore optional

doc.content.activeTable.setCellText(1, 4, "Underline Text")
doc.content.activeTable.setCellUnderline(1, 4)


# Cell Colors
# Colors are in the format "#ffffff"
# setCellFontColor(row, col, HexColorString)
# setCellBackgroundColor(row, col, HexColorString)
doc.content.activeTable.setCellFontColor(2, 1, "#0000ff")
doc.content.activeTable.setCellBackgroundColor(2, 1, "#ff0000")
doc.content.activeTable.setCellText(2, 1, "Blue on Red")

doc.content.activeTable.setCellFontColor(2, 2, "#ff0000")
doc.content.activeTable.setCellBackgroundColor(2, 2, "#0000ff")
doc.content.activeTable.setCellText(2, 2, "Red on Blue")

doc.content.activeTable.setCellText(2, 3, "Default Colors")


# Text Font Sizes
# List font sizes as points (ie. "10pt")
# setCellFontSize(row, col, PointsString)
doc.content.activeTable.setCellFontSize(3, 1, "10pt")
doc.content.activeTable.setCellText(3, 1, "Default 10pt")

doc.content.activeTable.setCellFontSize(3, 2, "11pt")
doc.content.activeTable.setCellText(3, 2, "11pt")

doc.content.activeTable.setCellFontSize(3, 3, "12pt")
doc.content.activeTable.setCellText(3, 3, "12pt")

# Write out the document
doc.export("calc-example05.ods")
