#!/usr/bin/python

import sys
sys.path.append('../code')
import ooolib

# Create your document
doc = ooolib.Calc()

# Set Manual Styles
# These styles override any other settings you might have set
# The available styles are as follows:

# Heading Styles
# "Heading"
doc.content.activeTable.setCellText(1, 1, "Heading")
doc.content.activeTable.setCellStyle(1, 1, "Heading")
# "Heading 1"
doc.content.activeTable.setCellText(2, 1, "Heading 1")
doc.content.activeTable.setCellStyle(2, 1, "Heading 1")
# "Heading 2"
doc.content.activeTable.setCellText(3, 1, "Heading 2")
doc.content.activeTable.setCellStyle(3, 1, "Heading 2")

# Note/Link Styles
# "Note"
doc.content.activeTable.setCellText(1, 2, "Note")
doc.content.activeTable.setCellStyle(1, 2, "Note")
# "Footnote"
doc.content.activeTable.setCellText(2, 2, "Footnote")
doc.content.activeTable.setCellStyle(2, 2, "Footnote")
# "Hyperlink" - Note: Hyperlink does not make the link active
doc.content.activeTable.setCellText(3, 2, "Hyperlink")
doc.content.activeTable.setCellStyle(3, 2, "Hyperlink")

# Status Styles
# "Good"
doc.content.activeTable.setCellText(1, 3, "Good")
doc.content.activeTable.setCellStyle(1, 3, "Good")
# "Neutral"
doc.content.activeTable.setCellText(2, 3, "Neutral")
doc.content.activeTable.setCellStyle(2, 3, "Neutral")
# "Bad"
doc.content.activeTable.setCellText(3, 3, "Bad")
doc.content.activeTable.setCellStyle(3, 3, "Bad")
# "Warning"
doc.content.activeTable.setCellText(4, 3, "Warning")
doc.content.activeTable.setCellStyle(4, 3, "Warning")
# "Error"
doc.content.activeTable.setCellText(5, 3, "Error")
doc.content.activeTable.setCellStyle(5, 3, "Error")

# Accent Styles
# "Accent"
doc.content.activeTable.setCellText(1, 4, "Accent")
doc.content.activeTable.setCellStyle(1, 4, "Accent")
# "Accent 1"
doc.content.activeTable.setCellText(2, 4, "Accent 1")
doc.content.activeTable.setCellStyle(2, 4, "Accent 1")
# "Accent 2"
doc.content.activeTable.setCellText(3, 4, "Accent 2")
doc.content.activeTable.setCellStyle(3, 4, "Accent 2")
# "Accent 3"
doc.content.activeTable.setCellText(4, 4, "Accent 3")
doc.content.activeTable.setCellStyle(4, 4, "Accent 3")

# Result Style
# "Result"
doc.content.activeTable.setCellText(1, 5, "Result")
doc.content.activeTable.setCellStyle(1, 5, "Result")

# Save the document to the file you want to create
doc.export("calc-example08.ods")
