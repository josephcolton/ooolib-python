#!/usr/bin/python3
# Read Example 01 - Gathers information from an existing ODF document
#
# Usage: python read-example01.py [document]
# With no argument, reads draw-example01.odg (run draw-example01.py first).

import sys
sys.path.append('../code')
import ooolib

filename = "draw-example01.odg"
if len(sys.argv) > 1: filename = sys.argv[1]

# Open the document (read only)
doc = ooolib.openDocument(filename)

# Document type and metadata
print("File:        %s" % doc.filename)
print("Type:        %s" % doc.type)
print("Title:       %s" % doc.meta["title"])
print("Subject:     %s" % doc.meta["subject"])
print("Keywords:    %s" % ", ".join(doc.meta["keywords"]))
print("Statistics:  %s" % doc.meta["statistics"])
if doc.pageLayout:
    print("Page size:   %s x %s" % (doc.pageLayout["width"], doc.pageLayout["height"]))
print("Images:      %d" % len(doc.images))

# Pages and shapes (drawings and presentations)
for page in doc.pages:
    print("")
    print("Page %d: %s (%d shapes)" % (page.number, page.name, len(page.shapes)))
    for shape in page.getShapes():
        bounds = shape.getBounds("in")
        if bounds:
            position = "at (%.2fin, %.2fin) size %.2fin x %.2fin" % bounds
        else:
            position = ""
        print("  %-8s %s" % (shape.type, position))
        if shape.fillColor:   print("           fill:   %s" % shape.fillColor)
        if shape.strokeColor: print("           stroke: %s %s" % (shape.strokeColor, shape.strokeWidth))
        if shape.text:        print("           text:   %r" % shape.text)
    if page.notes:
        print("  Notes: %r" % page.notes)

# Shape counts by type
print("")
counts = {}
for shape in doc.getShapes():
    counts[shape.type] = counts.get(shape.type, 0) + 1
print("Shape counts: %s" % counts)

# All text in the document
print("")
print("All text:")
print(doc.getText())
