#!/usr/bin/python

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Calc()

# Set Column Width
# setColumnWidth(col, width)
doc.content.activeTable.setColumnWidth(1, "0.5in")
doc.content.activeTable.setColumnWidth(2, "1.0in")
doc.content.activeTable.setColumnWidth(3, "1.5in")

# Set Row Height
# setRowHeight(row, height)
doc.content.activeTable.setRowHeight(1, "0.5in")
doc.content.activeTable.setRowHeight(2, "1.0in")
doc.content.activeTable.setRowHeight(3, "1.5in")

# Fill in Cell Data
doc.content.activeTable.setCellText(1, 1, "0.5in x 0.5in")
doc.content.activeTable.setCellText(2, 2, "1.0in x 1.0in")
doc.content.activeTable.setCellText(3, 3, "1.5in x 1.5in")

# Write out the document
doc.export("calc-example04.ods")
