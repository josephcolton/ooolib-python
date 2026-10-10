#!/usr/bin/python3
# Impress Example 02 - Widescreen slides with shapes, connectors, groups, and images

import sys
sys.path.append('../code')
import ooolib

# Create the document with 16:9 widescreen slides
doc = ooolib.Impress()
doc.meta.setTitle("Impress Example 2")
doc.styles.setPageSize("13.333in", "7.5in")   # the title and content areas resize to match
doc.content.setUnit("in")

# --- Slide 1: title slide with shapes ---
slide = doc.content.activeSlide
slide.setTitle("Shapes on Slides")
slide.addContent("Slides have the same shape methods as Draw pages")
slide.addRectangle(1.0, 4.5, 3.0, 1.5, fillcolor="#dae3f3", strokecolor="#4472c4",
    text="Rectangle", bold=True, fontsize=20, name="Rectangle")
slide.addEllipse(5.17, 4.5, 3.0, 1.5, fillcolor="#fce4d6", strokecolor="#ed7d31",
    text="Ellipse", fontsize=20, rotation=10, name="Ellipse")
slide.addPolygon([(10.83, 4.5), (12.33, 6.0), (9.33, 6.0)], fillcolor="#e2f0d9",
    strokecolor="#548235", text="Polygon", fontsize=16, name="Polygon")

# --- Slide 2: a diagram built from connectors and a group (title only, no content area) ---
slide = doc.content.addSlide("Diagram", layout="title")
slide.setTitle("How a Request Flows")
boxStyle = {"fillcolor": "#dae3f3", "strokecolor": "#4472c4", "strokewidth": 0.03,
            "bold": True, "fontsize": 18}
slide.addRectangle(0.8, 3.0, 2.5, 1.2, name="Browser", text="Browser", **boxStyle)
slide.addRectangle(5.4, 3.0, 2.5, 1.2, name="Server", text="Server", **boxStyle)
group = slide.addGroup(name="Storage")
group.addRectangle(10.0, 2.2, 2.5, 1.0, name="Cache", text="Cache", **boxStyle)
group.addRectangle(10.0, 4.2, 2.5, 1.0, name="Database", text="Database", **boxStyle)
slide.addConnector("Browser", "Server", startside="right", endside="left",
    endarrow="arrow", strokewidth=0.03, text="request", fontsize=14)
slide.addConnector("Server", "Cache", startside="right", endside="left", endarrow="arrow")
slide.addConnector("Server", "Database", startside="right", endside="left", endarrow="arrow",
    linestyle="dash")

# --- Slide 3: a blank slide (no title or content areas) with an image ---
slide = doc.content.addSlide("Picture", layout="blank")
slide.setBackgroundColor("#f2f2f2")
slide.addTextBox(0.5, 0.5, 12.33, 1.0, "Blank slides hold only what you add",
    fontsize=32, bold=True, fontcolor="#1f3864", align="center")
logo = b"""<svg xmlns="http://www.w3.org/2000/svg" width="3in" height="3in" viewBox="0 0 100 100">
  <circle cx="50" cy="50" r="45" fill="#4472c4" stroke="#1f3864" stroke-width="4"/>
  <text x="50" y="62" font-size="32" text-anchor="middle" fill="white" font-family="sans-serif">ODF</text>
</svg>"""
slide.addImage(logo, 5.17, 2.5, name="Logo")
slide.addLine(1.0, 6.5, 12.33, 6.5, strokecolor="#7f7f7f", linestyle="longdash", startarrow="circle", endarrow="arrow")

# Write out the document
doc.export("impress-example02.odp")
print("Created impress-example02.odp")
