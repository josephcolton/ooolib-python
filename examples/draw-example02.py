#!/usr/bin/python3
# Draw Example 02 - Page size, numeric coordinates, named shapes, and text in shapes

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Draw()
doc.meta.setTitle("Draw Example 2")
doc.meta.setDescription("A simple process diagram using numeric coordinates.")

# US Letter, portrait (orientation is worked out from the size)
doc.styles.setPageSize("8.5in", "11in")
doc.styles.setPageMargins("0.5in")

# Numbers are measured in the document unit (inches by default)
doc.content.setUnit("in")
page = doc.content.activePage

# Title across the top of the page
page.addTextBox(0.5, 0.5, 7.5, 0.6, "Order Process",
    bold=True, fontsize=24, fontcolor="#1f3864", align="center", name="Title")

# A column of named boxes, placed with a loop
steps = ["Receive Order", "Check Stock", "Pack Items", "Ship Order"]
boxWidth, boxHeight = 3.0, 1.0
x = (8.5 - boxWidth) / 2
y = 1.5
gap = 0.75
for step in steps:
    page.addRectangle(x, y, boxWidth, boxHeight,
        fillcolor="#dae3f3", strokecolor="#4472c4", strokewidth=0.03,
        name=step, text=step, bold=True, fontsize=14, fontcolor="#1f3864")
    # Line down to the next box
    if step != steps[-1]:
        page.addLine(x + boxWidth / 2, y + boxHeight, x + boxWidth / 2, y + boxHeight + gap,
            strokecolor="#4472c4", strokewidth=0.03)
    y += boxHeight + gap

# Ellipse with text, measured in centimeters
doc.content.setUnit("cm")
page.addEllipse(6.5, 22, 8.6, 3, fillcolor="#e2f0d9", strokecolor="#548235",
    name="Done", text="Done!", italic=True, fontsize=18, fontcolor="#375623")

# Shapes are returned so they can be changed after they are added
box = page.getShape("Check Stock")
box.setText("Check Stock\n(warehouse)", bold=True, fontsize=12, fontcolor="#1f3864")

# Write out the document
doc.export("draw-example02.odg")
print("Created draw-example02.odg")
