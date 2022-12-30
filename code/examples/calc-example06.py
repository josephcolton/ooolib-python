#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create your document
doc = ooolib.Calc()

# Spreadsheet Title
doc.content.activeTable.setCellText(1, 1, 'Alignment')
doc.content.activeTable.setCellBold(1, 1)

# Row Heights
doc.content.activeTable.setRowHeight(4, '0.5in')
doc.content.activeTable.setRowHeight(5, '0.5in')
doc.content.activeTable.setRowHeight(6, '0.5in')

# Vertical Alignment Labels
doc.content.activeTable.setCellText(4, 1, 'top')
doc.content.activeTable.setCellText(5, 1, 'middle')
doc.content.activeTable.setCellText(6, 1, 'bottom')

# Horizontal Alignment Labels
doc.content.activeTable.setCellText(3, 2, 'left')
doc.content.activeTable.setCellText(3, 3, 'center')
doc.content.activeTable.setCellText(3, 4, 'right')
doc.content.activeTable.setCellText(3, 5, 'justify')

# Fill in cells with "x" to make it visible
for row in range(4, 6+1):
    for col in range(2, 5+1):
        doc.content.activeTable.setCellText(row, col, 'x')

# Configure Vertical Alignment for cells
# Vertical Alignments are "top", "middle", and "bottom"
# setCellVerticalAlign(row, col, AlignString)
for col in range(2, 5+1):
    doc.content.activeTable.setCellVerticalAlign(4, col, "top")
    doc.content.activeTable.setCellVerticalAlign(5, col, "middle")
    doc.content.activeTable.setCellVerticalAlign(6, col, "bottom")

# Configure Horizontal Alignment for cells
# Horizontal Alignments are "start", "center", "end", and "justify".
# ooolib will convert "left" to "start" and "right" to "end"
# setCellHorizontalAlign(row, col, AlignString)
for row in range(4, 6+1):
    doc.content.activeTable.setCellHorizontalAlign(row, 2, "left")
    doc.content.activeTable.setCellHorizontalAlign(row, 3, "center")
    doc.content.activeTable.setCellHorizontalAlign(row, 4, "right")
    doc.content.activeTable.setCellHorizontalAlign(row, 5, "justify")

# Save the document to the file you want to create
doc.export("calc-example06.ods")
