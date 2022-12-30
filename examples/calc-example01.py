#!/usr/bin/python

import sys
sys.path.append('../code')
import ooolib

# Create your document
doc = ooolib.Calc()

# We use the doc object to work with our Calc document.  We need to
# then address different parts of the document.  For spreadsheets,
# we typically want to address the current/active table.  We address
# the activeTable in the content area of the document using:
#
# doc.content.activeTable
#
# Setting float values.  We index into the table using a row and column.
# The row and column values both start at 1.  To set float values we
# use the method setCellFloat:
#
# setCellFloat(row, col, value)
#
for row in range(1, 9):
	for col in range(1, 9):
                value = row*col
                doc.content.activeTable.setCellFloat(row, col, value)

# Export/save the document to the file you want to create
doc.export("calc-example01.ods")
