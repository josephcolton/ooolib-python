# ooolib-python

ooolib-python is a pure-Python library for programmatically creating OpenOffice.org and LibreOffice compatible documents in the Open Document Format (ODF). It requires no external dependencies — only the Python standard library.

**Supported document types:**

| Class | Format | Extension |
|---|---|---|
| `ooolib.Calc` | Spreadsheet | `.ods` |
| `ooolib.Writer` | Text document | `.odt` |
| `ooolib.Impress` | Presentation | `.odp` |
| `ooolib.Draw` | Drawing | `.odg` |

---

## Installation

Download or clone the repository. Add the `code/` directory to your Python path before importing:

```python
import sys
sys.path.append('/path/to/ooolib-python/code')
import ooolib
```

No pip install or build step is needed. There are no external dependencies.

---

## Quick Start

```python
import sys
sys.path.append('code')
import ooolib

# Spreadsheet
doc = ooolib.Calc()
doc.content.activeTable.setCellText(1, 1, "Hello, World!")
doc.export("hello.ods")

# Text document
doc = ooolib.Writer()
doc.content.addHeading("Hello, World!", level=1)
doc.content.addParagraph("This is a paragraph.")
doc.export("hello.odt")

# Presentation
doc = ooolib.Impress()
doc.content.activeSlide.setTitle("Hello, World!")
doc.content.activeSlide.addContent("First slide content")
doc.export("hello.odp")

# Drawing
doc = ooolib.Draw()
doc.content.activePage.addRectangle("1in", "1in", "3in", "2in", fillcolor="#4472c4")
doc.export("hello.odg")
```

---

## Common API (all document types)

Every document class has the same `meta` interface and `export` method.

### Document Properties

```python
doc.meta.setTitle("My Document")
doc.meta.setSubject("Subject line")
doc.meta.setDescription("Longer description of the document.")
doc.meta.addKeyword("keyword1")
doc.meta.addKeyword("keyword2")
```

### Exporting

```python
doc.export("output_filename.ods")   # writes the ODF ZIP archive
```

---

## Calc — Spreadsheets

```python
doc = ooolib.Calc()            # creates document with one blank sheet
doc = ooolib.Calc(autoInit=False)  # creates document with no sheets
```

### Sheet Management

```python
doc.content.addTable("SheetName")      # add a new sheet (becomes active)
doc.content.setActiveTable("SheetName") # switch to an existing sheet by name
doc.content.setActiveTable()           # switch to the last sheet
doc.content.activeTable.setTableName("New Name")  # rename current sheet
```

The `addTable()` method returns the sheet object, which can be stored for direct access later:

```python
sheet = doc.content.addTable("Data")
sheet.setCellText(1, 1, "Direct access")
```

### Cell Coordinates

Rows and columns are both 1-indexed. Row 1, Column 1 is the top-left cell (A1).

```python
# Convert between (row, col) and spreadsheet cell IDs like "A1"
cellId = doc.content.activeTable.convertCellRowCol2Id(1, 1)   # → "A1"
row, col = doc.content.activeTable.convertCellId2RowCol("B3") # → (3, 2)
```

### Setting Cell Values

```python
# setCellFloat(row, col, value)
doc.content.activeTable.setCellFloat(1, 1, 42.5)

# setCellText(row, col, text)
doc.content.activeTable.setCellText(1, 2, "Hello")

# setCellDate(row, col, datetime_value)
import datetime
doc.content.activeTable.setCellDate(1, 3, datetime.datetime(2024, 1, 15))
```

### Formulas

Cell IDs for formula ranges can be generated with `convertCellRowCol2Id()`.

```python
start = doc.content.activeTable.convertCellRowCol2Id(1, 1)  # "A1"
end   = doc.content.activeTable.convertCellRowCol2Id(8, 1)  # "A8"

doc.content.activeTable.setCellFormulaAverage(1, 3, start, end)
doc.content.activeTable.setCellFormulaMin(2, 3, start, end)
doc.content.activeTable.setCellFormulaMax(3, 3, start, end)
doc.content.activeTable.setCellFormulaSum(4, 3, start, end)
doc.content.activeTable.setCellFormulaSqrt(5, 3, end)        # single cell target
doc.content.activeTable.setCellFormulaCustom(6, 3, "of:=IF(([.A5]>[.A4]);[.A4];[.A1])")
```

### Cell Formatting

Style methods can be called before or after setting the cell value. All style methods are independent and can be combined freely.

```python
# Text style
doc.content.activeTable.setCellBold(row, col)           # bold=True by default
doc.content.activeTable.setCellBold(row, col, False)    # turn off
doc.content.activeTable.setCellItalics(row, col)
doc.content.activeTable.setCellUnderline(row, col)

# Colors — hex strings in "#rrggbb" format
doc.content.activeTable.setCellFontColor(row, col, "#0000ff")
doc.content.activeTable.setCellBackgroundColor(row, col, "#ffff00")

# Font size — point string
doc.content.activeTable.setCellFontSize(row, col, "14pt")

# Vertical alignment: "top", "middle", "bottom"
doc.content.activeTable.setCellVerticalAlign(row, col, "middle")

# Horizontal alignment: "left", "center", "right", "justify"
# (ooolib converts "left" → "start" and "right" → "end" internally)
doc.content.activeTable.setCellHorizontalAlign(row, col, "center")
```

### Predefined Named Styles

Named styles override any individually set style properties on that cell.

```python
doc.content.activeTable.setCellStyle(row, col, "Style Name")
```

Available named styles:

| Category | Style Names |
|---|---|
| Heading | `"Heading"`, `"Heading 1"`, `"Heading 2"` |
| Note / Link | `"Note"`, `"Footnote"`, `"Hyperlink"` |
| Status | `"Good"`, `"Neutral"`, `"Bad"`, `"Warning"`, `"Error"` |
| Accent | `"Accent"`, `"Accent 1"`, `"Accent 2"`, `"Accent 3"` |
| Other | `"Result"` |

### Column and Row Dimensions

Dimensions are strings with a unit, e.g. `"1in"`, `"2.5in"`, `"25mm"`.

```python
doc.content.activeTable.setColumnWidth(col, "1.5in")
doc.content.activeTable.setRowHeight(row, "0.5in")
```

### Calc Example

```python
import sys
sys.path.append('code')
import ooolib

doc = ooolib.Calc()
doc.meta.setTitle("Sales Report")

# Header row
doc.content.activeTable.setCellText(1, 1, "Product")
doc.content.activeTable.setCellText(1, 2, "Units")
doc.content.activeTable.setCellText(1, 3, "Revenue")
for col in range(1, 4):
    doc.content.activeTable.setCellBold(1, col)
    doc.content.activeTable.setCellBackgroundColor(1, col, "#dae3f3")

# Data rows
doc.content.activeTable.setCellText(2, 1, "Widget A")
doc.content.activeTable.setCellFloat(2, 2, 150)
doc.content.activeTable.setCellFloat(2, 3, 4500.00)

doc.content.activeTable.setCellText(3, 1, "Widget B")
doc.content.activeTable.setCellFloat(3, 2, 80)
doc.content.activeTable.setCellFloat(3, 3, 3200.00)

# Total row using SUM formula
doc.content.activeTable.setCellText(4, 1, "Total")
doc.content.activeTable.setCellBold(4, 1)
start = doc.content.activeTable.convertCellRowCol2Id(2, 3)
end   = doc.content.activeTable.convertCellRowCol2Id(3, 3)
doc.content.activeTable.setCellFormulaSum(4, 3, start, end)

doc.export("sales_report.ods")
```

---

## Writer — Text Documents

```python
doc = ooolib.Writer()
```

### Adding Content

```python
# Headings — level 1 through 4
doc.content.addHeading("Document Title", level=1)
doc.content.addHeading("Chapter One", level=2)
doc.content.addHeading("Section 1.1", level=3)
doc.content.addHeading("Sub-section", level=4)

# Paragraphs with optional formatting
doc.content.addParagraph("Plain paragraph text.")
doc.content.addParagraph("Bold text.", bold=True)
doc.content.addParagraph("Italic text.", italic=True)
doc.content.addParagraph("Underlined text.", underline=True)
doc.content.addParagraph("Combined formatting.", bold=True, italic=True, underline=True)

# Font size — point string
doc.content.addParagraph("Large text.", fontsize="18pt")

# Font color — hex string
doc.content.addParagraph("Red text.", fontcolor="#cc0000")
```

The `addParagraph` signature in full:

```python
doc.content.addParagraph(text,
    bold=False, italic=False, underline=False,
    fontsize=None,    # e.g. "12pt", "18pt"
    fontcolor=None)   # e.g. "#cc0000"
```

Formatting applies to the entire paragraph. The default page layout is US Letter (8.5″ × 11″) with 1″ top/bottom and 1.25″ left/right margins. The default body font is Liberation Serif 12pt.

### Writer Example

```python
import sys
sys.path.append('code')
import ooolib

doc = ooolib.Writer()
doc.meta.setTitle("Project Report")
doc.meta.setSubject("Q3 Status")

doc.content.addHeading("Project Report", level=1)
doc.content.addParagraph("This report summarizes the Q3 project status.")

doc.content.addHeading("Summary", level=2)
doc.content.addParagraph("All milestones were completed on schedule.", bold=True)
doc.content.addParagraph("Budget utilization was within 5% of projections.")

doc.content.addHeading("Action Items", level=2)
doc.content.addParagraph("Schedule Q4 kickoff meeting.", fontcolor="#cc0000")
doc.content.addParagraph("Update stakeholder dashboard.")

doc.export("project_report.odt")
```

---

## Impress — Presentations

```python
doc = ooolib.Impress()   # creates document with one blank slide
```

Default slide size: 10″ × 7.5″ (standard 4:3 landscape).

### Slide Management

```python
doc.content.addSlide("SlideName")        # add a new slide (becomes active)
doc.content.setActiveSlide("SlideName")  # switch to an existing slide by name
doc.content.activeSlide                  # the currently active slide
```

### Adding Content to a Slide

```python
# Title (appears in the large title area at the top)
doc.content.activeSlide.setTitle("Slide Title")

# Content lines (appear in the body area below the title)
doc.content.activeSlide.addContent("Plain content line")
doc.content.activeSlide.addContent("Bold content", bold=True)
doc.content.activeSlide.addContent("Italic content", italic=True)
doc.content.activeSlide.addContent("Underlined content", underline=True)
doc.content.activeSlide.addContent("Larger text", fontsize="36pt")
doc.content.activeSlide.addContent("Colored text", fontcolor="#cc0000")

# Background color for the slide
doc.content.activeSlide.setBackgroundColor("#1f3864")
```

The `addContent` signature in full:

```python
doc.content.activeSlide.addContent(text,
    bold=False, italic=False, underline=False,
    fontsize=None,    # e.g. "28pt" (default content size is 28pt)
    fontcolor=None)   # e.g. "#ffffff"
```

Title text uses 44pt bold Liberation Sans. Content text uses 28pt Liberation Sans. These reflect the Default-title and Default-subtitle presentation styles.

### Impress Example

```python
import sys
sys.path.append('code')
import ooolib

doc = ooolib.Impress()
doc.meta.setTitle("Quarterly Review")

# Slide 1: title slide
doc.content.activeSlide.setTitle("Q3 Quarterly Review")
doc.content.activeSlide.addContent("Presented by the Engineering Team")
doc.content.activeSlide.setBackgroundColor("#1f3864")

# Slide 2: agenda
doc.content.addSlide("Agenda")
doc.content.activeSlide.setTitle("Agenda")
doc.content.activeSlide.addContent("1. Project Status")
doc.content.activeSlide.addContent("2. Budget Overview")
doc.content.activeSlide.addContent("3. Next Steps")

# Slide 3: key results
doc.content.addSlide("Results")
doc.content.activeSlide.setTitle("Key Results")
doc.content.activeSlide.addContent("All milestones completed", bold=True, fontcolor="#006600")
doc.content.activeSlide.addContent("3 items pending review", fontcolor="#cc6600")
doc.content.activeSlide.addContent("Budget: on track", italic=True)

doc.export("quarterly_review.odp")
```

---

## Draw — Drawings

```python
doc = ooolib.Draw()   # creates document with one blank page
```

Default page size: 10″ × 7.5″ (landscape) with 0.5″ margins.

### Page Setup

```python
doc.styles.setPageSize("8.5in", "11in")       # orientation is worked out from the size
doc.styles.setPageSize("11in", "8.5in", orientation="landscape")
doc.styles.setPageMargins("0.5in")            # all four margins
doc.styles.setPageMargins(top="1in", left="0.75in")
doc.styles.getPageSize()                      # ("8.5in", "11in")
```

### Units

Positions and sizes can be strings with a unit (`"1in"`, `"2.5cm"`, `"25mm"`) or plain numbers. Numbers are measured in the document unit, which is inches by default:

```python
doc.content.setUnit("in")    # "in", "cm", "mm", "pt", "pc", or "px"
doc.content.activePage.addRectangle(1, 1, 3, 2)          # same as "1in", "1in", "3in", "2in"
doc.content.setUnit("mm")
doc.content.activePage.addEllipse(20, 20, 50, 30)        # millimeters
```

`setUnit()` only affects shapes added after it is called. Font sizes given as numbers are always points (`fontsize=14` is `"14pt"`).

### Page Management

```python
doc.content.addPage("PageName")        # add a new page (becomes active)
doc.content.setActivePage("PageName")  # switch to an existing page by name
doc.content.activePage                 # the currently active page
```

### Shapes

Every shape method returns the new shape. Shapes can be given a `name`, which LibreOffice shows in the Navigator and which `ooolib.openDocument()` reads back.

```python
# Rectangle
doc.content.activePage.addRectangle(x, y, width, height,
    fillcolor=None,         # e.g. "#4472c4" — None means no fill
    strokecolor=None,       # e.g. "#000000" — None means no border
    strokewidth="0.02in",   # border thickness
    name=None,              # shape name, e.g. "Server A"
    text=None,              # text shown inside the shape ("\n" for a new line)
    bold=False, italic=False, underline=False,
    fontsize=None,          # e.g. "14pt" or 14
    fontcolor=None,         # e.g. "#cc0000"
    align="center",         # "left", "center", "right", or "justify"
    linestyle="solid",      # border style, see Line Styles below
    rotation=None)          # degrees counterclockwise around the center

# Ellipse (also used for circles when width == height) - same options as addRectangle
doc.content.activePage.addEllipse(x, y, width, height, fillcolor=None, strokecolor=None,
    strokewidth="0.02in", name=None, text=None, ...)

# Polygon - a closed shape through a list of (x, y) points, same options as addRectangle
doc.content.activePage.addPolygon([(1, 1), (3, 1), (2, 3)],
    fillcolor="#ffd966", strokecolor="#bf9000", text="Triangle", rotation=None, ...)

# Line
doc.content.activePage.addLine(x1, y1, x2, y2,
    strokecolor="#000000",
    strokewidth="0.02in",
    name=None,
    linestyle="solid",
    startarrow=None,        # line ends: None, "arrow", "circle", "square", or "diamond"
    endarrow=None,
    arrowsize=None)         # width of the line ends; default scales with strokewidth

# Polyline - an open line through a list of (x, y) points
doc.content.activePage.addPolyline([(1, 1), (2, 2), (3, 1)],
    strokecolor="#000000", strokewidth="0.02in", name=None, linestyle="solid",
    startarrow=None, endarrow=None, arrowsize=None, rotation=None)

# Text box
doc.content.activePage.addTextBox(x, y, width, height, text,
    bold=False, italic=False, underline=False,
    fontsize=None,    # e.g. "14pt" or 14
    fontcolor=None,   # e.g. "#cc0000"
    name=None,
    align="left",
    rotation=None)
```

Text in rectangles, ellipses, and polygons is centered vertically. Shapes can be found and changed after they are added:

```python
box = doc.content.activePage.addRectangle(1, 1, 3, 1, name="Step 1")
box.setText("Step 1", bold=True)                       # replaces the text and its formatting
box.setRotation(45)                                    # None removes the rotation
doc.content.activePage.getShape("Step 1").setName("First Step")
doc.content.activePage.getShapes()                     # every shape on the page
```

### Line Styles

`linestyle` works on lines, polylines, and the borders of rectangles, ellipses, and polygons: `"solid"`, `"dash"`, `"dot"`, `"dashdot"`, or `"longdash"`. Dash lengths scale with the line width.

### Rotation

`rotation` turns a shape counterclockwise around its center, in degrees, like the rotation setting in LibreOffice. Lines are not rotated; change their end points instead.

### Groups

Shapes in a group move and select together in LibreOffice. A group has the same add methods as a page, and groups can be nested. `getShape()` searches inside groups.

```python
group = doc.content.activePage.addGroup(name="Server")
group.addRectangle(7, 5.5, 2, 1.2, fillcolor="#e2f0d9", text="Server")
group.addEllipse(7.15, 5.65, 0.25, 0.25, fillcolor="#70ad47", name="Status Light")
doc.content.activePage.getShape("Status Light")
```

### Connectors

A connector joins two shapes and stays attached when the shapes are moved in LibreOffice. Shapes can be given as shape objects or by name, including shapes inside groups. Connectors attach to rectangles, ellipses, polygons, and text boxes.

```python
page = doc.content.activePage
page.addRectangle(1, 1, 2, 1, name="Web Server", text="Web Server")
page.addRectangle(6, 3, 2, 1, name="Database", text="Database")
page.addConnector("Web Server", "Database",
    startside="right",          # "top", "right", "bottom", "left", or None to let LibreOffice choose
    endside="top",
    connectortype="standard",   # "standard" (right angles), "lines", "line" (straight), or "curve"
    strokecolor="#000000", strokewidth="0.02in", linestyle="solid",
    startarrow=None, endarrow="arrow", arrowsize=None,
    name=None,
    text="SQL",                 # optional label, with bold/italic/underline/fontsize/fontcolor
    fontsize=10)
```

LibreOffice works out the connector's route when it opens the file.

### Draw Example

```python
import sys
sys.path.append('code')
import ooolib

doc = ooolib.Draw()
doc.meta.setTitle("Architecture Diagram")

page = doc.content.activePage

# Background panel
page.addRectangle("0.25in", "0.25in", "9.5in", "7in",
    fillcolor="#f5f5f5", strokecolor="#cccccc")

# Three boxes representing services
for i, (label, color) in enumerate([
    ("Web Server",    "#dae3f3"),
    ("App Server",    "#e2efda"),
    ("Database",      "#fce4d6"),
]):
    x = "%.2fin" % (0.75 + i * 3.0)
    page.addRectangle(x, "1in", "2.5in", "1.5in",
        fillcolor=color, strokecolor="#595959")
    page.addTextBox(x, "1.4in", "2.5in", "0.7in",
        label, bold=True, fontsize="14pt")

# Arrows (represented as lines)
page.addLine("3.25in", "1.75in", "3.75in", "1.75in",
    strokecolor="#595959", strokewidth="0.04in")
page.addLine("6.25in", "1.75in", "6.75in", "1.75in",
    strokecolor="#595959", strokewidth="0.04in")

# Caption
page.addTextBox("0.25in", "6in", "9.5in", "0.75in",
    "Three-tier architecture — ooolib Draw example",
    italic=True, fontcolor="#595959")

doc.export("architecture_diagram.odg")
```

---

## Reading Documents

`ooolib.openDocument()` opens an existing ODF file (any type, including files saved by LibreOffice) and gathers information from it. It is read only.

```python
doc = ooolib.openDocument("drawing.odg")

doc.type                  # "graphics", "presentation", "text" or "spreadsheet"
doc.meta["title"]         # also subject, description, keywords, creator,
                          # creationDate, date, generator, statistics, userDefined
doc.pageLayout            # {"width": "10in", "height": "7.5in", "orientation": ..., margins}
doc.images                # [{"name": "Pictures/...", "size": 12345}, ...]
doc.getText()             # all text in the document
doc.toDict()              # everything above as plain dictionaries (JSON friendly)
```

Drawings and presentations also expose their pages (slides) and shapes:

```python
for page in doc.pages:                  # also doc.getPage("Page Two") or doc.getPage(2)
    print(page.number, page.name, page.notes)
    for shape in page.getShapes():      # includes shapes inside groups
        shape.type                      # rect, ellipse, line, textbox, image, custom-shape, group, ...
        shape.name, shape.text
        shape.x, shape.y, shape.width, shape.height   # strings as stored, e.g. "1.5in"
        shape.getBounds("cm")           # (x, y, width, height) as floats in any unit
        shape.fill, shape.fillColor     # resolved through the document's styles
        shape.stroke, shape.strokeColor, shape.strokeWidth
        shape.strokeDash                # dash style name, e.g. "Dash", or None
        shape.markerStart, shape.markerEnd   # line end names, e.g. "Arrow", or None
        shape.rotation                  # degrees counterclockwise, or None
        shape.getPoints("in")           # polygon/polyline points as (x, y) floats
        shape.children                  # shapes inside a group

for startShape, endShape, connector in page.getConnections():
    print(startShape.name, "->", endShape.name, connector.text)
    connector.connectorType             # "standard", "lines", "line", or "curve"
    connector.startSide, connector.endSide   # "top", "right", "bottom", "left", or None

doc.getShapes("textbox")                # every text box in the document
ooolib.convertLength("2.54cm", "in")    # 1.0 (units: in, cm, mm, pt, pc, px)
```

---

## Running the Examples

All examples are in the `examples/` directory and should be run from within that directory:

```
cd examples
python calc-example01.py    # multiplication table
python calc-example02.py    # multiple sheets
python calc-example03.py    # document metadata
python calc-example04.py    # column and row dimensions
python calc-example05.py    # cell formatting (bold, italic, colors, font size)
python calc-example06.py    # cell alignment (vertical and horizontal)
python calc-example07.py    # formulas (AVERAGE, MIN, MAX, SUM, SQRT, custom)
python calc-example08.py    # predefined named styles
python writer-example01.py  # headings, paragraphs, text formatting
python impress-example01.py # slides, titles, content text, background colors
python draw-example01.py    # rectangles, ellipses, lines, text boxes
python draw-example02.py    # page size, numeric coordinates, named shapes, text in shapes
python draw-example03.py    # polygons, polylines, arrows, dashed lines, rotation, groups
python draw-example04.py    # connectors between shapes
python read-example01.py    # read pages, shapes, and metadata from draw-example01.odg
```

---

## Running the Tests

The tests use only the Python standard library. Run them from the repository root:

```
python3 -m unittest discover tests            # every test
python3 -m unittest discover tests -k Draw    # tests whose names match "Draw"
python3 tests/test_draw.py                    # one file
```

The tests create documents, check that they are valid ODF files, and read them back with `ooolib.openDocument()`. They also run every example script. If LibreOffice is installed, `tests/test_libreoffice.py` converts documents with `soffice` to make sure LibreOffice can open them; set `OOOLIB_SKIP_LIBREOFFICE=1` to skip those slower tests.

---

## History

ooolib-python was originally written in 2006 as a port from the ooolib-perl project, created as part of a switch from Perl to Python. It was maintained for several years and then dormant. In 2023 the code was completely rewritten for Python 3 using a more object-oriented design. Since the 2023 rewrite Writer, Impress, and Draw document support has been added alongside the original Calc functionality.

---

## Credits

Written by Joseph Colton &lt;josephcolton@gmail.com&gt;

---

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.
