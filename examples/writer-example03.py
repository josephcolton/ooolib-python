#!/usr/bin/python3
# Writer Example 03 - Demonstrates tables, font selection, and page breaks

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Writer()
doc.meta.setTitle("Writer Example 3")

# Document-wide fonts
doc.styles.setDefaultFont("Liberation Sans", "11pt")
doc.styles.setHeadingFont("Liberation Serif")

# Font selection on paragraphs and runs
doc.content.addHeading("Fonts", level=2)
doc.content.addHeading("A heading in a monospaced font", level=3, fontname="Liberation Mono")
doc.content.addParagraph("This paragraph uses Liberation Serif.", fontname="Liberation Serif")
doc.content.addParagraph(["Inline code such as ", ("print('hello')", {"fontname": "Liberation Mono"}),
                          " can use a monospaced font."])

# A simple table from a list of rows; the first row is a repeating header row
doc.content.addHeading("Simple Table", level=2)
doc.content.addTable([["Name", "Department", "Extension"],
                      ["Alice", "Accounting", "1234"],
                      ["Bob", "Engineering", "5678"]], headerRows=1)

# A table built row by row with column widths and per-cell formatting
doc.content.addHeading("Invoice", level=2)
table = doc.content.addTable(columnWidths=["3in", "1in", "1.25in"], align="center")
table.addRow(["Item", "Qty", "Price"], header=True, bold=True, background="#dddddd", align="center")
table.addRow(["Widget", {"text": "2", "align": "right"}, {"text": "$10.00", "align": "right"}])
table.addRow(["Gadget", {"text": "1", "align": "right"}, {"text": "$24.50", "align": "right"}])
table.addRow([{"text": "Total", "colspan": 2, "bold": True, "align": "right"},
              {"text": "$44.50", "bold": True, "align": "right", "fontcolor": "#cc0000"}])

# Start the next section on a new page
doc.content.addPageBreak()

# A borderless table, useful for layout
doc.content.addHeading("Borderless Table", level=2)
doc.content.addTable([[("Phone:", {"bold": True}), "555-0100"],
                      [("Email:", {"bold": True}), "info@example.com"]],
                     columnWidths=["1in", "3in"], border=None)

# Write out the document
doc.export("writer-example03.odt")
print("Created writer-example03.odt")
