#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create your document
doc = ooolib.Calc()

# Create numbers for performing formulas on
for row in range(1, 9):
	doc.content.activeTable.setCellFloat(row, 1, row)

# Formula Title Text
doc.content.activeTable.setCellText(1, 2, 'AVERAGE')
doc.content.activeTable.setCellText(2, 2, 'MIN')
doc.content.activeTable.setCellText(3, 2, 'MAX')
doc.content.activeTable.setCellText(4, 2, 'SUM')
doc.content.activeTable.setCellText(5, 2, 'SQRT')
doc.content.activeTable.setCellText(6, 2, 'Custom')

# Actual Formulas
#setCellFormulaAverage(row, col, startCellId, endCellId)
#setCellFormulaMin(row, col, startCellId, endCellId)
#setCellFormulaMax(row, col, startCellId, endCellId)
#setCellFormulaSum(row, col, startCellId, endCellId)
#setCellFormulaSqrt(row, col, targetCellId)
#setCellFormulaCustom(row, col, value)
#convertCellRowCol2Id(row, col)

# Convert row=1, col=1 to "A1" and row=8, col=1 to "A8"
startCellId = doc.content.activeTable.convertCellRowCol2Id(1, 1)
endCellId = doc.content.activeTable.convertCellRowCol2Id(8, 1)

# Create functions
doc.content.activeTable.setCellFormulaAverage(1, 3, startCellId, endCellId)
doc.content.activeTable.setCellFormulaMin(2, 3, startCellId, endCellId)
doc.content.activeTable.setCellFormulaMax(3, 3, startCellId, endCellId)
doc.content.activeTable.setCellFormulaSum(4, 3, startCellId, endCellId)
doc.content.activeTable.setCellFormulaSqrt(5, 3, endCellId)

# If you know the format for a formula, you can manually enter it.
doc.content.activeTable.setCellFormulaCustom(6, 3, "of:=IF(([.A5]>[.A4]);[.A4];[.A1])")

# Save the document to the file you want to create
doc.export("calc-example07.ods")
