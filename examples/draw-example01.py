#!/usr/bin/python3
# Draw Example 01 - Demonstrates shapes, lines, and text boxes

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Draw()

# Document metadata
doc.meta.setTitle("Draw Example 1")
doc.meta.setSubject("ooolib-python Draw demonstration")
doc.meta.setDescription("Shows rectangles, ellipses, lines, and text boxes.")
doc.meta.addKeyword("ooolib")
doc.meta.addKeyword("draw")

# --- Page 1: Basic shapes ---
# All coordinates and sizes are strings, e.g. "1in", "2.5in", "50mm"
# addRectangle(x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in")
doc.content.activePage.addRectangle("0.5in", "0.5in", "3in", "2in",
    fillcolor="#4472c4", strokecolor="#1f3864", strokewidth="0.04in")

# addEllipse(x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in")
doc.content.activePage.addEllipse("4in", "0.5in", "2.5in", "2in",
    fillcolor="#ed7d31", strokecolor="#843c00", strokewidth="0.04in")

doc.content.activePage.addRectangle("7in", "0.5in", "2.5in", "2in",
    fillcolor="#70ad47", strokecolor="#375623")

# Shapes with no fill (outline only)
doc.content.activePage.addRectangle("0.5in", "3in", "2in", "1.5in",
    fillcolor=None, strokecolor="#4472c4", strokewidth="0.06in")

doc.content.activePage.addEllipse("3in", "3in", "2in", "1.5in",
    fillcolor=None, strokecolor="#ed7d31", strokewidth="0.06in")

# addLine(x1, y1, x2, y2, strokecolor="#000000", strokewidth="0.02in")
doc.content.activePage.addLine("0.5in", "5in", "9.5in", "5in",
    strokecolor="#595959", strokewidth="0.04in")
doc.content.activePage.addLine("0.5in", "5.2in", "9.5in", "5.2in",
    strokecolor="#595959", strokewidth="0.01in")

# addTextBox(x, y, width, height, text, bold=False, italic=False, underline=False,
#            fontsize=None, fontcolor=None)
doc.content.activePage.addTextBox("0.5in", "5.5in", "4in", "0.8in",
    "Hello from ooolib Draw!", bold=True, fontsize="18pt", fontcolor="#cc0000")

doc.content.activePage.addTextBox("5in", "5.5in", "4in", "0.8in",
    "Italic blue text", italic=True, fontsize="16pt", fontcolor="#0000cc")

# --- Page 2: More examples ---
doc.content.addPage("Page Two")

# Nested feel: large background rect, smaller shapes on top
doc.content.activePage.addRectangle("0.25in", "0.25in", "9.5in", "7in",
    fillcolor="#f2f2f2", strokecolor="#cccccc")

doc.content.activePage.addRectangle("0.75in", "0.75in", "4in", "2.5in",
    fillcolor="#dae3f3", strokecolor="#4472c4")
doc.content.activePage.addTextBox("0.75in", "0.75in", "4in", "0.6in",
    "Rectangle", bold=True, fontcolor="#1f3864")

doc.content.activePage.addEllipse("5.25in", "0.75in", "4in", "2.5in",
    fillcolor="#fce4d6", strokecolor="#ed7d31")
doc.content.activePage.addTextBox("5.25in", "0.75in", "4in", "0.6in",
    "Ellipse", bold=True, fontcolor="#843c00")

doc.content.activePage.addLine("0.75in", "3.75in", "9.25in", "3.75in",
    strokecolor="#4472c4", strokewidth="0.06in")
doc.content.activePage.addTextBox("0.75in", "4in", "8.5in", "0.6in",
    "Line (above)", italic=True, fontcolor="#595959")

doc.content.activePage.addTextBox("2in", "5in", "6in", "1in",
    "Underlined large text", underline=True, fontsize="28pt", fontcolor="#375623")

# Write out the document
doc.export("draw-example01.odg")
print("Created draw-example01.odg")
