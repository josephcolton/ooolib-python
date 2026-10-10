#!/usr/bin/python3
# Draw Example 03 - Polygons, polylines, arrows, dashed lines, rotation, and groups

import sys, math
sys.path.append('../code')
import ooolib

# Create the document (10in x 7.5in landscape, coordinates in inches)
doc = ooolib.Draw()
doc.meta.setTitle("Draw Example 3")
page = doc.content.activePage

# Polygons: a triangle and a five pointed star built from a list of (x, y) points
page.addPolygon([(1.25, 0.5), (2.25, 2.25), (0.25, 2.25)],
    fillcolor="#fce4d6", strokecolor="#ed7d31", strokewidth=0.03,
    name="Triangle", text="Triangle", fontsize=11)

star = []
for i in range(10):
    radius = 1.0 if i % 2 == 0 else 0.4
    angle = math.radians(90 + i * 36)
    star.append((3.75 + radius * math.cos(angle), 1.4 - radius * math.sin(angle)))
page.addPolygon(star, fillcolor="#ffd966", strokecolor="#bf9000", name="Star")

# Polyline: an open zig-zag line with an arrow on the end
page.addPolyline([(5.5, 2.25), (6.25, 0.5), (7.0, 2.25), (7.75, 0.5), (8.5, 2.25)],
    strokecolor="#4472c4", strokewidth=0.04, endarrow="arrow", name="Zigzag")

# Line styles and line ends
y = 3.0
for linestyle in ["solid", "dash", "dot", "dashdot", "longdash"]:
    page.addTextBox(0.5, y - 0.15, 1.5, 0.3, linestyle, fontsize=11)
    page.addLine(2.0, y, 5.0, y, strokewidth=0.03, linestyle=linestyle)
    y += 0.5
y = 3.0
for marker in ["arrow", "circle", "square", "diamond"]:
    page.addLine(5.75, y, 8.75, y, strokewidth=0.03, startarrow=marker, endarrow=marker)
    page.addTextBox(5.75, y + 0.05, 3.0, 0.3, marker, fontsize=11, align="center")
    y += 0.6

# Rotation (degrees counterclockwise around the shape's center)
for i, angle in enumerate([0, 30, 60, 90]):
    page.addRectangle(0.75 + i * 1.5, 5.75, 1.0, 0.6, fillcolor="#dae3f3", strokecolor="#4472c4",
        text="%d°" % angle, rotation=angle)

# A group: the shapes move and select together in LibreOffice
group = page.addGroup(name="Server")
group.addRectangle(7.0, 5.5, 2.0, 1.2, fillcolor="#e2f0d9", strokecolor="#548235",
    linestyle="dash", name="Server Box", text="Server", bold=True)
group.addEllipse(7.15, 5.65, 0.25, 0.25, fillcolor="#70ad47", name="Status Light")

# Write out the document
doc.export("draw-example03.odg")
print("Created draw-example03.odg")
