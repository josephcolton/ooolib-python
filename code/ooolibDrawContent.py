#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import math

# Import ooolib-python modules
import ooolibGlobal
import ooolibDrawStyles
import ooolibUnits
import ooolibXML

##############################
# ooolib-python Draw Content #
##############################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        self.internal_pages = []
        self.activePage = None
        self.global_object.setGlobalInt("drawPageNumber", 1)
        self.global_object.setGlobalString("drawUnit", "in")
        self.global_object.setGlobalInt("drawShapeId", 1)
        # Automatic styles
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles)
        self.global_object.setGlobalObjects("drawAutoStyles", self.automaticStylesInstance)
        # XML structure
        self.prolog = ooolibXML.Prolog("xml")
        self.documentContent = ooolibXML.Element("office:document-content")
        self.documentContent.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentContent.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentContent.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentContent.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.documentContent.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentContent.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentContent.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentContent.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentContent.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentContent.setAttribute("office:version", "1.3")
        # Font face declarations
        self.documentContent.addChild(ooolibXML.Element("office:scripts"))
        officeFontFace = self.documentContent.addChild(ooolibXML.Element("office:font-face-decls"))
        fontFace1 = officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace1.setAttribute("style:name", "Liberation Sans")
        fontFace1.setAttribute("svg:font-family", "'Liberation Sans'")
        fontFace1.setAttribute("style:font-family-generic", "swiss")
        fontFace1.setAttribute("style:font-pitch", "variable")
        # Attach automatic styles
        self.documentContent.addChild(self.automaticStyles)
        # Office body
        officeBody = self.documentContent.addChild(ooolibXML.Element("office:body"))
        self.officeDrawing = officeBody.addChild(ooolibXML.Element("office:drawing"))
        # Add initial page
        self.activePage = self.addPage()

    def addPage(self, name=None):
        page = ContentPage(self.global_object, name)
        self.internal_pages.append(page)
        self.activePage = page
        return page

    def setUnit(self, unit):
        """Set the unit used for numeric positions and sizes ('in', 'cm', 'mm', 'pt', 'pc', 'px').

        Only affects shapes added after the call. Strings such as '25mm' always keep their own unit."""
        if not ooolibUnits.isUnit(unit):
            raise ValueError("Unknown unit %r, expected one of: %s" % (unit, ", ".join(ooolibUnits.UNITS_PER_INCH)))
        self.global_object.setGlobalString("drawUnit", unit)

    def getUnit(self):
        return self.global_object.getGlobalString("drawUnit", "in")

    def setActivePage(self, name):
        for page in self.internal_pages:
            if page.getPageName() == name:
                self.activePage = page
                return self.activePage
        print("setActivePage('%s') - Unable to find page" % name)
        return self.activePage

    def updateStats(self):
        self.global_object.setGlobalInt("drawPageCount", len(self.internal_pages))

    def buildPages(self):
        self.officeDrawing.clearChildren()
        for page in self.internal_pages:
            self.officeDrawing.addChild(page.pageObject())

    def toString(self, indent=False):
        self.buildPages()
        contentString = self.prolog.toString()
        contentString += self.documentContent.toString(indent)
        return contentString


####################
# Automatic Styles #
####################
# Paragraph alignment names mapped to fo:text-align values
PARAGRAPH_ALIGN = {"left": "start", "center": "center", "right": "end", "justify": "justify"}

# Connector sides mapped to the standard glue point numbers every shape has
CONNECTOR_SIDES = {"top": 0, "right": 1, "bottom": 2, "left": 3}

# Connector routes: "standard" uses right angles, "lines" uses straight segments
CONNECTOR_TYPES = ("standard", "lines", "line", "curve")

class AutomaticStyles:
    def __init__(self, global_object, automaticStyles):
        self.global_object = global_object
        self.officeAutomaticStyles = automaticStyles
        self.styleObjects = {}
        self.global_object.setGlobalInt("drawStyleGraphic", 1)
        self.global_object.setGlobalInt("drawStyleParagraph", 1)
        # Create the fixed text-box style once
        self._textBoxStyleName = self._createTextBoxStyle()

    def extractData(self, data, name, default=None):
        if data == None: return default
        if name not in data: return default
        val = data[name]
        if val is None: return default
        return val

    def getGraphicStyle(self, fillcolor=None, strokecolor=None, strokewidth=None, linestyle="solid",
                        startarrow=None, endarrow=None, arrowsize=None, textMiddle=False):
        """Return (creating if needed) a graphic style for shapes and lines."""
        index = ("graphic", fillcolor, strokecolor, strokewidth, linestyle,
                 startarrow, endarrow, arrowsize, textMiddle)
        if index not in self.styleObjects:
            self.styleObjects[index] = self._createGraphicStyle(fillcolor, strokecolor, strokewidth, linestyle,
                                                                startarrow, endarrow, arrowsize, textMiddle)
        return self.styleObjects[index]

    def _createGraphicStyle(self, fillcolor, strokecolor, strokewidth, linestyle,
                            startarrow, endarrow, arrowsize, textMiddle):
        styleNum = self.global_object.incrementGlobalInt("drawStyleGraphic")
        styleName = "gr%d" % styleNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", styleName)
        style.setAttribute("style:family", "graphic")
        graphicProp = style.addChild(ooolibXML.Element("style:graphic-properties"))
        if textMiddle:
            graphicProp.setAttribute("draw:textarea-vertical-align", "middle")
        if fillcolor:
            graphicProp.setAttribute("draw:fill", "solid")
            graphicProp.setAttribute("draw:fill-color", fillcolor)
        else:
            graphicProp.setAttribute("draw:fill", "none")
        if strokecolor:
            if linestyle == "solid":
                graphicProp.setAttribute("draw:stroke", "solid")
            else:
                graphicProp.setAttribute("draw:stroke", "dash")
                graphicProp.setAttribute("draw:stroke-dash", ooolibDrawStyles.LINE_STYLES[linestyle][0])
            graphicProp.setAttribute("svg:stroke-color", strokecolor)
            graphicProp.setAttribute("svg:stroke-width", strokewidth)
        else:
            graphicProp.setAttribute("draw:stroke", "none")
        for end, marker in (("start", startarrow), ("end", endarrow)):
            if marker is None: continue
            drawName, viewBox, path, centered = ooolibDrawStyles.MARKERS[marker]
            graphicProp.setAttribute("draw:marker-%s" % end, drawName)
            graphicProp.setAttribute("draw:marker-%s-width" % end, arrowsize)
            graphicProp.setAttribute("draw:marker-%s-center" % end, "true" if centered else "false")
        return styleName

    def getTextBoxStyle(self):
        """Return the graphic style name for text boxes."""
        return self._textBoxStyleName

    def _createTextBoxStyle(self):
        styleNum = self.global_object.incrementGlobalInt("drawStyleGraphic")
        styleName = "gr%d" % styleNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", styleName)
        style.setAttribute("style:family", "graphic")
        graphicProp = style.addChild(ooolibXML.Element("style:graphic-properties"))
        graphicProp.setAttribute("draw:fill", "none")
        graphicProp.setAttribute("draw:stroke", "none")
        graphicProp.setAttribute("draw:auto-grow-height", "true")
        return styleName

    def getParagraphStyle(self, paraData=None):
        """Return the automatic paragraph style name for the given formatting data."""
        has_formatting = (
            self.extractData(paraData, "bold") or
            self.extractData(paraData, "italic") or
            self.extractData(paraData, "underline") or
            self.extractData(paraData, "fontsize") is not None or
            self.extractData(paraData, "fontcolor") is not None or
            self.extractData(paraData, "align", "left") != "left"
        )
        if not has_formatting:
            return "Standard"
        index = ("para",
                 self.extractData(paraData, "bold"),
                 self.extractData(paraData, "italic"),
                 self.extractData(paraData, "underline"),
                 self.extractData(paraData, "fontsize"),
                 self.extractData(paraData, "fontcolor"),
                 self.extractData(paraData, "align", "left"))
        if index not in self.styleObjects:
            self.styleObjects[index] = self._createParagraphStyle(paraData)
        return self.styleObjects[index]

    def _createParagraphStyle(self, paraData):
        paraNum = self.global_object.incrementGlobalInt("drawStyleParagraph")
        paraName = "P%d" % paraNum
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", paraName)
        styleObject.setAttribute("style:family", "paragraph")
        styleObject.setAttribute("style:parent-style-name", "Standard")
        align = self.extractData(paraData, "align", "left")
        if align != "left":
            paraProp = styleObject.addChild(ooolibXML.Element("style:paragraph-properties"))
            paraProp.setAttribute("fo:text-align", PARAGRAPH_ALIGN[align])
        textProp = None
        if self.extractData(paraData, "bold"):
            if textProp is None: textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            textProp.setAttribute("fo:font-weight", "bold")
            textProp.setAttribute("style:font-weight-asian", "bold")
            textProp.setAttribute("style:font-weight-complex", "bold")
        if self.extractData(paraData, "italic"):
            if textProp is None: textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            textProp.setAttribute("fo:font-style", "italic")
            textProp.setAttribute("style:font-style-asian", "italic")
            textProp.setAttribute("style:font-style-complex", "italic")
        if self.extractData(paraData, "underline"):
            if textProp is None: textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            textProp.setAttribute("style:text-underline-style", "solid")
            textProp.setAttribute("style:text-underline-width", "auto")
            textProp.setAttribute("style:text-underline-color", "font-color")
        fontsize = self.extractData(paraData, "fontsize")
        if fontsize is not None:
            if textProp is None: textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            textProp.setAttribute("fo:font-size", fontsize)
            textProp.setAttribute("style:font-size-asian", fontsize)
            textProp.setAttribute("style:font-size-complex", fontsize)
        fontcolor = self.extractData(paraData, "fontcolor")
        if fontcolor is not None:
            if textProp is None: textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            textProp.setAttribute("fo:color", fontcolor)
        self.styleObjects[paraName] = styleObject
        return paraName


#################
# Content Shape #
#################
class ContentShape:
    """A shape on a Draw page. Returned by the ContentPage add methods."""
    def __init__(self, shapeType, data):
        self.type = shapeType
        self.data = data

    def getType(self):
        return self.type

    def getName(self):
        return self.data.get("name")

    def setName(self, name):
        self.data["name"] = name
        return self

    def getText(self):
        return self.data.get("text")

    def setText(self, text, bold=False, italic=False, underline=False,
                fontsize=None, fontcolor=None, align=None):
        """Set the text shown inside the shape, replacing any existing text."""
        if self.type in ("line", "polyline"):
            raise ValueError("Lines can not contain text")
        if align is None:
            align = self.data.get("defaultAlign")
        if align not in PARAGRAPH_ALIGN:
            raise ValueError("align must be one of: %s" % ", ".join(PARAGRAPH_ALIGN))
        self.data["text"] = text
        self.data["paraData"] = {
            "bold": bold, "italic": italic, "underline": underline,
            "fontsize": formatFontSize(fontsize), "fontcolor": fontcolor,
            "align": align,
        }
        return self

    def getRotation(self):
        return self.data.get("rotation")

    def setRotation(self, rotation):
        """Rotate the shape around its center, in degrees counterclockwise (None for no rotation)."""
        if self.type in ("line", "connector"):
            raise ValueError("Lines can not be rotated, change their end points instead")
        self.data["rotation"] = checkRotation(rotation)
        return self

    def getBounds(self):
        """Return (x, y, width, height) in inches, before rotation."""
        if self.type == "polygon":
            inches = [(toInches(x), toInches(y)) for (x, y) in self.data["points"]]
            xs = [x for (x, y) in inches]
            ys = [y for (x, y) in inches]
            return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
        return (toInches(self.data["x"]), toInches(self.data["y"]),
                toInches(self.data["width"]), toInches(self.data["height"]))

    def getGluePoint(self, side):
        """Return the (x, y) position in inches of the middle of a side, including rotation."""
        x, y, width, height = self.getBounds()
        centerX, centerY = x + width / 2, y + height / 2
        offsets = {"top": (0, -height / 2), "right": (width / 2, 0),
                   "bottom": (0, height / 2), "left": (-width / 2, 0)}
        offsetX, offsetY = offsets[side]
        rotation = self.data.get("rotation")
        if rotation:
            angle = math.radians(rotation)
            offsetX, offsetY = (offsetX * math.cos(angle) + offsetY * math.sin(angle),
                                offsetY * math.cos(angle) - offsetX * math.sin(angle))
        return (centerX + offsetX, centerY + offsetY)

    def __repr__(self):
        return "<ContentShape %s name=%r>" % (self.type, self.getName())

def formatFontSize(fontsize):
    """Font sizes given as numbers are in points (14 -> '14pt')."""
    return ooolibUnits.formatLength(fontsize, "pt")

def checkRotation(rotation):
    if rotation is None: return None
    if isinstance(rotation, bool) or not isinstance(rotation, (int, float)):
        raise TypeError("rotation must be a number of degrees, not %r" % (rotation,))
    rotation = rotation % 360
    if rotation == 0: return None
    return rotation

def checkLineStyle(linestyle):
    if linestyle not in ooolibDrawStyles.LINE_STYLES:
        raise ValueError("linestyle must be one of: %s" % ", ".join(ooolibDrawStyles.LINE_STYLES))
    return linestyle

def checkMarker(marker):
    if marker is None: return None
    if marker not in ooolibDrawStyles.MARKERS:
        raise ValueError("Line ends must be None or one of: %s" % ", ".join(ooolibDrawStyles.MARKERS))
    return marker

def toInches(length):
    """Convert a length string to inches, raising ValueError if it can not be converted."""
    value = ooolibUnits.convertLength(length, "in")
    if value is None:
        raise ValueError("Unable to use length %r for this shape" % (length,))
    return value


###################
# Shape Container #
###################
class ShapeContainer:
    """Shared shape methods for pages and groups."""
    def __init__(self, global_object):
        self.global_object = global_object
        self.shapes = []
        self.automaticStylesInstance = self.global_object.getGlobalObjects("drawAutoStyles")

    def escapeText(self, text):
        return ooolibXML.escapeParagraphText(text)

    def length(self, value):
        """Convert numbers to lengths in the document unit; strings are kept as given."""
        return ooolibUnits.formatLength(value, self.global_object.getGlobalString("drawUnit", "in"))

    ##########
    # Shapes #
    ##########
    def addShape(self, shape):
        self.shapes.append(shape)
        return shape

    def getShapes(self):
        """Return the shapes directly in this page or group (groups are not expanded)."""
        return list(self.shapes)

    def iterShapes(self):
        """Yield every shape, including shapes inside groups."""
        for shape in self.shapes:
            yield shape
            if shape.type == "group":
                for nested in shape.iterShapes():
                    yield nested

    def getShape(self, name):
        """Return the first shape or group with the given name (searching inside groups), or None."""
        for shape in self.iterShapes():
            if shape.getName() == name: return shape
        return None

    def addGroup(self, name=None):
        """Add a group. Shapes added to the group move and select together in LibreOffice."""
        return self.addShape(ContentGroup(self.global_object, name))

    def addRectangle(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in",
                     name=None, text=None, bold=False, italic=False, underline=False,
                     fontsize=None, fontcolor=None, align="center", linestyle="solid", rotation=None):
        """Add a rectangle. Positions and sizes are numbers in the document unit or strings
        such as '1in' or '25mm'. Optional text is centered in the shape by default."""
        return self.addFilledShape("rect", x, y, width, height, fillcolor, strokecolor, strokewidth,
                                   name, text, bold, italic, underline, fontsize, fontcolor, align,
                                   linestyle, rotation)

    def addEllipse(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in",
                   name=None, text=None, bold=False, italic=False, underline=False,
                   fontsize=None, fontcolor=None, align="center", linestyle="solid", rotation=None):
        """Add an ellipse. Positions and sizes are numbers in the document unit or strings
        such as '1in' or '25mm'. Optional text is centered in the shape by default."""
        return self.addFilledShape("ellipse", x, y, width, height, fillcolor, strokecolor, strokewidth,
                                   name, text, bold, italic, underline, fontsize, fontcolor, align,
                                   linestyle, rotation)

    def addFilledShape(self, shapeType, x, y, width, height, fillcolor, strokecolor, strokewidth,
                       name, text, bold, italic, underline, fontsize, fontcolor, align,
                       linestyle, rotation):
        shape = ContentShape(shapeType, {
            "x": self.length(x), "y": self.length(y),
            "width": self.length(width), "height": self.length(height),
            "fillcolor": fillcolor, "strokecolor": strokecolor,
            "strokewidth": self.length(strokewidth), "linestyle": checkLineStyle(linestyle),
            "name": name, "text": None, "paraData": None, "defaultAlign": "center",
            "rotation": None,
        })
        shape.setRotation(rotation)
        if text is not None:
            shape.setText(text, bold, italic, underline, fontsize, fontcolor, align)
        return self.addShape(shape)

    def addPolygon(self, points, fillcolor=None, strokecolor=None, strokewidth="0.02in",
                   name=None, text=None, bold=False, italic=False, underline=False,
                   fontsize=None, fontcolor=None, align="center", linestyle="solid", rotation=None):
        """Add a closed polygon through a list of (x, y) points, e.g. [(1, 1), (3, 1), (2, 3)].
        Points are numbers in the document unit or strings such as '1in'."""
        shape = ContentShape("polygon", {
            "points": self.convertPoints(points, 3),
            "fillcolor": fillcolor, "strokecolor": strokecolor,
            "strokewidth": self.length(strokewidth), "linestyle": checkLineStyle(linestyle),
            "name": name, "text": None, "paraData": None, "defaultAlign": "center",
            "rotation": None,
        })
        shape.setRotation(rotation)
        if text is not None:
            shape.setText(text, bold, italic, underline, fontsize, fontcolor, align)
        return self.addShape(shape)

    def addPolyline(self, points, strokecolor="#000000", strokewidth="0.02in", name=None,
                    linestyle="solid", startarrow=None, endarrow=None, arrowsize=None, rotation=None):
        """Add an open line through a list of (x, y) points, e.g. [(1, 1), (2, 2), (3, 1)].
        Points are numbers in the document unit or strings such as '1in'."""
        shape = ContentShape("polyline", {
            "points": self.convertPoints(points, 2),
            "strokecolor": strokecolor, "strokewidth": self.length(strokewidth),
            "linestyle": checkLineStyle(linestyle),
            "startarrow": checkMarker(startarrow), "endarrow": checkMarker(endarrow),
            "arrowsize": self.length(arrowsize), "name": name, "rotation": None,
        })
        shape.setRotation(rotation)
        return self.addShape(shape)

    def addLine(self, x1, y1, x2, y2, strokecolor="#000000", strokewidth="0.02in", name=None,
                linestyle="solid", startarrow=None, endarrow=None, arrowsize=None):
        """Add a line from (x1,y1) to (x2,y2). Coordinates are numbers in the document unit
        or strings such as '1in'. startarrow/endarrow: 'arrow', 'circle', 'square', or 'diamond'."""
        return self.addShape(ContentShape("line", {
            "x1": self.length(x1), "y1": self.length(y1),
            "x2": self.length(x2), "y2": self.length(y2),
            "strokecolor": strokecolor, "strokewidth": self.length(strokewidth),
            "linestyle": checkLineStyle(linestyle),
            "startarrow": checkMarker(startarrow), "endarrow": checkMarker(endarrow),
            "arrowsize": self.length(arrowsize), "name": name,
        }))

    def addConnector(self, startShape, endShape, startside=None, endside=None, connectortype="standard",
                     strokecolor="#000000", strokewidth="0.02in", name=None, linestyle="solid",
                     startarrow=None, endarrow=None, arrowsize=None, text=None,
                     bold=False, italic=False, underline=False, fontsize=None, fontcolor=None):
        """Add a connector between two shapes. It stays attached when the shapes are moved
        in LibreOffice. Shapes may be given as shape objects or by name.

        startside/endside: 'top', 'right', 'bottom', 'left', or None to let LibreOffice choose.
        connectortype: 'standard' (right angles), 'lines', 'line' (straight), or 'curve'."""
        startShape = self.findConnectable(startShape)
        endShape = self.findConnectable(endShape)
        for side in (startside, endside):
            if side is not None and side not in CONNECTOR_SIDES:
                raise ValueError("Connector sides must be None or one of: %s" % ", ".join(CONNECTOR_SIDES))
        if connectortype not in CONNECTOR_TYPES:
            raise ValueError("connectortype must be one of: %s" % ", ".join(CONNECTOR_TYPES))
        for shape in (startShape, endShape):
            if shape.data.get("id") is None:
                shape.data["id"] = "id%d" % self.global_object.incrementGlobalInt("drawShapeId")
        shape = ContentShape("connector", {
            "startShape": startShape, "endShape": endShape,
            "startside": startside, "endside": endside, "connectortype": connectortype,
            "strokecolor": strokecolor, "strokewidth": self.length(strokewidth),
            "linestyle": checkLineStyle(linestyle),
            "startarrow": checkMarker(startarrow), "endarrow": checkMarker(endarrow),
            "arrowsize": self.length(arrowsize), "name": name,
            "text": None, "paraData": None, "defaultAlign": "center",
        })
        if text is not None:
            shape.setText(text, bold, italic, underline, fontsize, fontcolor)
        return self.addShape(shape)

    def findConnectable(self, shape):
        """Return the ContentShape for a shape object or name that a connector can attach to."""
        if not isinstance(shape, (ContentShape, ContentGroup)):
            name = shape
            shape = self.getShape(name)
            if shape is None:
                raise ValueError("No shape named %r" % (name,))
        if shape.type not in ("rect", "ellipse", "polygon", "textbox"):
            raise ValueError("Connectors can only attach to rectangles, ellipses, polygons, and text boxes")
        return shape

    def addTextBox(self, x, y, width, height, text,
                   bold=False, italic=False, underline=False, fontsize=None, fontcolor=None,
                   name=None, align="left", rotation=None):
        """Add a text box at the given position. Positions and sizes are numbers in the
        document unit or strings such as '1in'."""
        shape = ContentShape("textbox", {
            "x": self.length(x), "y": self.length(y),
            "width": self.length(width), "height": self.length(height),
            "name": name, "defaultAlign": "left", "rotation": None,
        })
        shape.setRotation(rotation)
        shape.setText(text, bold, italic, underline, fontsize, fontcolor, align)
        return self.addShape(shape)

    def convertPoints(self, points, minimum):
        points = [(self.length(x), self.length(y)) for (x, y) in points]
        if len(points) < minimum:
            raise ValueError("At least %d points are needed" % minimum)
        for (x, y) in points:
            toInches(x)
            toInches(y)
        return points

    ##################
    # XML Generation #
    ##################
    def addParagraph(self, parent, shape):
        """Add the shape's text (or an empty paragraph) to parent."""
        p = parent.addChild(ooolibXML.Element("text:p"))
        text = shape.data.get("text")
        if text is None or text == "": return p
        paraStyleName = self.automaticStylesInstance.getParagraphStyle(shape.data["paraData"])
        p.setAttribute("text:style-name", paraStyleName)
        p.setText(self.escapeText(text))
        return p

    def setPosition(self, elem, x, y, width, height, rotation):
        """Set svg:x/y/width/height, or a draw:transform when the shape is rotated."""
        elem.setAttribute("svg:width", width)
        elem.setAttribute("svg:height", height)
        if rotation is None:
            elem.setAttribute("svg:x", x)
            elem.setAttribute("svg:y", y)
            return
        # Rotate around the shape's center: ODF rotates around the shape's
        # top left corner, so translate that corner to keep the center in place.
        x, y = toInches(x), toInches(y)
        halfWidth, halfHeight = toInches(width) / 2, toInches(height) / 2
        angle = math.radians(rotation)
        cornerX = x + halfWidth - (halfWidth * math.cos(angle) + halfHeight * math.sin(angle))
        cornerY = y + halfHeight - (halfHeight * math.cos(angle) - halfWidth * math.sin(angle))
        elem.setAttribute("draw:transform", "rotate (%s) translate (%sin %sin)" % (
            repr(angle), ooolibUnits.formatNumber(cornerX), ooolibUnits.formatNumber(cornerY)))

    def setPoints(self, elem, points, rotation):
        """Set the bounding box, view box, and points of a polygon or polyline."""
        inches = [(toInches(x), toInches(y)) for (x, y) in points]
        minX = min([x for (x, y) in inches])
        minY = min([y for (x, y) in inches])
        width = max([x for (x, y) in inches]) - minX
        height = max([y for (x, y) in inches]) - minY
        # View box units are thousandths of an inch
        viewWidth = max(1, int(round(width * 1000)))
        viewHeight = max(1, int(round(height * 1000)))
        self.setPosition(elem, ooolibUnits.formatLength(minX, "in"), ooolibUnits.formatLength(minY, "in"),
                         ooolibUnits.formatLength(viewWidth / 1000.0, "in"),
                         ooolibUnits.formatLength(viewHeight / 1000.0, "in"), rotation)
        elem.setAttribute("svg:viewBox", "0 0 %d %d" % (viewWidth, viewHeight))
        pointStrings = ["%d,%d" % (int(round((x - minX) * 1000)), int(round((y - minY) * 1000)))
                        for (x, y) in inches]
        elem.setAttribute("draw:points", " ".join(pointStrings))

    def setNameAndId(self, elem, data):
        if data.get("name") is not None: elem.setAttribute("draw:name", data["name"])
        if data.get("id") is not None:
            # draw:id for older readers, xml:id for ODF 1.2 and later
            elem.setAttribute("draw:id", data["id"])
            elem.setAttribute("xml:id", data["id"])

    def connectorEnds(self, data):
        """Work out connector end points in inches. LibreOffice reroutes connectors
        when it opens the file; these points are for other ODF readers."""
        startShape, endShape = data["startShape"], data["endShape"]
        startSide, endSide = data["startside"], data["endside"]
        if startSide is None or endSide is None:
            # Pick the sides that face each other
            startX, startY, startWidth, startHeight = startShape.getBounds()
            endX, endY, endWidth, endHeight = endShape.getBounds()
            deltaX = (endX + endWidth / 2) - (startX + startWidth / 2)
            deltaY = (endY + endHeight / 2) - (startY + startHeight / 2)
            if abs(deltaX) >= abs(deltaY):
                facing = ("right", "left") if deltaX >= 0 else ("left", "right")
            else:
                facing = ("bottom", "top") if deltaY >= 0 else ("top", "bottom")
            if startSide is None: startSide = facing[0]
            if endSide is None: endSide = facing[1]
        return startShape.getGluePoint(startSide) + endShape.getGluePoint(endSide)

    def buildShapes(self, parent):
        """Add the XML elements for every shape to parent."""
        styles = self.automaticStylesInstance
        for shape in self.shapes:
            shapeType = shape.type
            data = shape.data

            if shapeType == "group":
                elem = parent.addChild(ooolibXML.Element("draw:g"))
                self.setNameAndId(elem, data)
                shape.buildShapes(elem)

            elif shapeType in ("rect", "ellipse", "polygon"):
                styleName = styles.getGraphicStyle(data["fillcolor"], data["strokecolor"], data["strokewidth"],
                                                   data["linestyle"], textMiddle=True)
                elem = parent.addChild(ooolibXML.Element("draw:" + shapeType))
                self.setNameAndId(elem, data)
                elem.setAttribute("draw:style-name", styleName)
                if shapeType == "polygon":
                    self.setPoints(elem, data["points"], data["rotation"])
                else:
                    self.setPosition(elem, data["x"], data["y"], data["width"], data["height"], data["rotation"])
                self.addParagraph(elem, shape)

            elif shapeType in ("line", "polyline"):
                styleName = styles.getGraphicStyle(None, data["strokecolor"], data["strokewidth"],
                                                   data["linestyle"], data["startarrow"], data["endarrow"],
                                                   self.arrowSize(data))
                elem = parent.addChild(ooolibXML.Element("draw:" + shapeType))
                self.setNameAndId(elem, data)
                elem.setAttribute("draw:style-name", styleName)
                if shapeType == "polyline":
                    self.setPoints(elem, data["points"], data["rotation"])
                else:
                    elem.setAttribute("svg:x1", data["x1"])
                    elem.setAttribute("svg:y1", data["y1"])
                    elem.setAttribute("svg:x2", data["x2"])
                    elem.setAttribute("svg:y2", data["y2"])

            elif shapeType == "textbox":
                styleName = styles.getTextBoxStyle()
                frame = parent.addChild(ooolibXML.Element("draw:frame"))
                self.setNameAndId(frame, data)
                frame.setAttribute("draw:style-name", styleName)
                self.setPosition(frame, data["x"], data["y"], data["width"], data["height"], data["rotation"])
                textBox = frame.addChild(ooolibXML.Element("draw:text-box"))
                self.addParagraph(textBox, shape)

            elif shapeType == "connector":
                styleName = styles.getGraphicStyle(None, data["strokecolor"], data["strokewidth"],
                                                   data["linestyle"], data["startarrow"], data["endarrow"],
                                                   self.arrowSize(data))
                elem = parent.addChild(ooolibXML.Element("draw:connector"))
                self.setNameAndId(elem, data)
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("draw:type", data["connectortype"])
                x1, y1, x2, y2 = self.connectorEnds(data)
                elem.setAttribute("svg:x1", ooolibUnits.formatLength(x1, "in"))
                elem.setAttribute("svg:y1", ooolibUnits.formatLength(y1, "in"))
                elem.setAttribute("svg:x2", ooolibUnits.formatLength(x2, "in"))
                elem.setAttribute("svg:y2", ooolibUnits.formatLength(y2, "in"))
                elem.setAttribute("draw:start-shape", data["startShape"].data["id"])
                if data["startside"] is not None:
                    elem.setAttribute("draw:start-glue-point", str(CONNECTOR_SIDES[data["startside"]]))
                elem.setAttribute("draw:end-shape", data["endShape"].data["id"])
                if data["endside"] is not None:
                    elem.setAttribute("draw:end-glue-point", str(CONNECTOR_SIDES[data["endside"]]))
                self.addParagraph(elem, shape)

    def arrowSize(self, data):
        """Line end width: arrowsize if given, otherwise scaled from the line width."""
        if data["startarrow"] is None and data["endarrow"] is None: return None
        if data["arrowsize"] is not None: return data["arrowsize"]
        strokeWidth = ooolibUnits.convertLength(data["strokewidth"], "in") or 0.02
        return ooolibUnits.formatLength(max(0.1, strokeWidth * 4), "in")


#################
# Content Group #
#################
class ContentGroup(ShapeContainer):
    """A group of shapes. Has the same add methods as a page."""
    def __init__(self, global_object, name=None):
        ShapeContainer.__init__(self, global_object)
        self.type = "group"
        self.data = {"name": name}

    def getType(self):
        return self.type

    def getName(self):
        return self.data.get("name")

    def setName(self, name):
        self.data["name"] = name
        return self

    def __repr__(self):
        return "<ContentGroup name=%r shapes=%d>" % (self.getName(), len(self.shapes))


#####################
# Content Page      #
#####################
class ContentPage(ShapeContainer):
    def __init__(self, global_object, name=None):
        ShapeContainer.__init__(self, global_object)
        self.name = name
        if self.name is None:
            self.name = "page%d" % self.global_object.incrementGlobalInt("drawPageNumber")

    def getPageName(self):
        return self.name

    def pageObject(self):
        """Build and return the draw:page XML element for this page."""
        page = ooolibXML.Element("draw:page")
        page.setAttribute("draw:name", self.name)
        page.setAttribute("draw:style-name", "dp1")
        page.setAttribute("draw:master-page-name", "Default")
        self.buildShapes(page)
        return page


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    content = Content(g)

    content.activePage.addRectangle("0.5in", "0.5in", "3in", "2in",
        fillcolor="#4472c4", strokecolor="#000000")
    content.activePage.addEllipse("4in", "0.5in", "2.5in", "2in",
        fillcolor="#ed7d31", strokecolor="#000000")
    content.activePage.addLine("0.5in", "3in", "9in", "3in",
        strokecolor="#333333", strokewidth="0.04in")
    content.activePage.addTextBox("0.5in", "3.5in", "4in", "1in",
        "Hello from Draw!", bold=True, fontsize="18pt", fontcolor="#cc0000")

    content.addPage("Page Two")
    content.activePage.addRectangle("1in", "1in", "8in", "5.5in",
        fillcolor="#e8f4e8", strokecolor="#006600")

    string1 = content.toString(indent=True)
    print("Draw Content XML with indentation:")
    print(string1)
