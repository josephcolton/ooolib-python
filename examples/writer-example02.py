#!/usr/bin/python3
# Writer Example 02 - Demonstrates lists, paragraph layout, and mixed formatting

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Writer()
doc.meta.setTitle("Writer Example 2")

# Paragraph layout
# addParagraph(text, ..., align=None, indent=None, spaceBefore=None, spaceAfter=None, keepWithNext=False)
doc.content.addParagraph("Centered Title", bold=True, fontsize="16pt", align="center", spaceAfter="0.2in")
doc.content.addParagraph("This paragraph is indented half an inch.", indent="0.5in")
doc.content.addParagraph("This paragraph is justified and has space before and after it. " * 3,
                         align="justify", spaceBefore="0.1in", spaceAfter="0.1in")

# Mixed formatting: a list of runs, each a string or (string, formatting)
doc.content.addParagraph([("Monthly Rent:", {"bold": True}), " $800.00, due on the ",
                          ("first", {"italic": True}), " of each month."])

# Bulleted list with nesting
doc.content.addHeading("Bulleted List", level=2)
doc.content.addListItem("First item")
doc.content.addListItem("Second item")
doc.content.addListItem("Nested item", level=2)
doc.content.addListItem("Deeper item", level=3)
doc.content.addListItem([("Bold label:", {"bold": True}), " third item"])

# Numbered list with bulleted sub-items
doc.content.addHeading("Numbered List", level=2)
doc.content.addListItem("Step one", numbered=True)
doc.content.addListItem("A note about step one", level=2)
doc.content.addListItem("Step two", numbered=True)
doc.content.addListItem("Sub-step a", level=2, numbered=True)
doc.content.addListItem("Sub-step b", level=2, numbered=True)

# endList() starts a fresh list (numbering restarts)
doc.content.endList()
doc.content.addListItem("A new list starts at one", numbered=True)

# Write out the document
doc.export("writer-example02.odt")
print("Created writer-example02.odt")
