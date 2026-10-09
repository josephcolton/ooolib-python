#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import ooolib-python modules
import ooolibGlobal
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

    def getShapeStyle(self, fillcolor, strokecolor, strokewidth):
        """Return (creating if needed) a graphic style for filled/stroked shapes."""
        index = ("shape", fillcolor, strokecolor, strokewidth)
        if index not in self.styleObjects:
            self.styleObjects[index] = self._createShapeStyle(fillcolor, strokecolor, strokewidth)
        return self.styleObjects[index]

    def _createShapeStyle(self, fillcolor, strokecolor, strokewidth):
        styleNum = self.global_object.incrementGlobalInt("drawStyleGraphic")
        styleName = "gr%d" % styleNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", styleName)
        style.setAttribute("style:family", "graphic")
        graphicProp = style.addChild(ooolibXML.Element("style:graphic-properties"))
        graphicProp.setAttribute("draw:textarea-vertical-align", "middle")
        if fillcolor:
            graphicProp.setAttribute("draw:fill", "solid")
            graphicProp.setAttribute("draw:fill-color", fillcolor)
        else:
            graphicProp.setAttribute("draw:fill", "none")
        if strokecolor:
            graphicProp.setAttribute("draw:stroke", "solid")
            graphicProp.setAttribute("svg:stroke-color", strokecolor)
            graphicProp.setAttribute("svg:stroke-width", strokewidth)
        else:
            graphicProp.setAttribute("draw:stroke", "none")
        self.styleObjects[("shape", fillcolor, strokecolor, strokewidth)] = styleName
        return styleName

    def getLineStyle(self, strokecolor, strokewidth):
        """Return (creating if needed) a graphic style for lines."""
        index = ("line", strokecolor, strokewidth)
        if index not in self.styleObjects:
            self.styleObjects[index] = self._createLineStyle(strokecolor, strokewidth)
        return self.styleObjects[index]

    def _createLineStyle(self, strokecolor, strokewidth):
        styleNum = self.global_object.incrementGlobalInt("drawStyleGraphic")
        styleName = "gr%d" % styleNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", styleName)
        style.setAttribute("style:family", "graphic")
        graphicProp = style.addChild(ooolibXML.Element("style:graphic-properties"))
        graphicProp.setAttribute("draw:fill", "none")
        graphicProp.setAttribute("draw:stroke", "solid")
        graphicProp.setAttribute("svg:stroke-color", strokecolor)
        graphicProp.setAttribute("svg:stroke-width", strokewidth)
        self.styleObjects[("line", strokecolor, strokewidth)] = styleName
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
        if self.type == "line":
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

    def __repr__(self):
        return "<ContentShape %s name=%r>" % (self.type, self.getName())

def formatFontSize(fontsize):
    """Font sizes given as numbers are in points (14 -> '14pt')."""
    return ooolibUnits.formatLength(fontsize, "pt")


#####################
# Content Page      #
#####################
class ContentPage:
    def __init__(self, global_object, name=None):
        self.global_object = global_object
        self.name = name
        if self.name is None:
            self.name = "page%d" % self.global_object.incrementGlobalInt("drawPageNumber")
        self.shapes = []
        self.automaticStylesInstance = self.global_object.getGlobalObjects("drawAutoStyles")

    def getPageName(self):
        return self.name

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
        return list(self.shapes)

    def getShape(self, name):
        """Return the first shape with the given name, or None."""
        for shape in self.shapes:
            if shape.getName() == name: return shape
        return None

    def addRectangle(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in",
                     name=None, text=None, bold=False, italic=False, underline=False,
                     fontsize=None, fontcolor=None, align="center"):
        """Add a rectangle. Positions and sizes are numbers in the document unit or strings
        such as '1in' or '25mm'. Optional text is centered in the shape by default."""
        return self.addFilledShape("rect", x, y, width, height, fillcolor, strokecolor, strokewidth,
                                   name, text, bold, italic, underline, fontsize, fontcolor, align)

    def addEllipse(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in",
                   name=None, text=None, bold=False, italic=False, underline=False,
                   fontsize=None, fontcolor=None, align="center"):
        """Add an ellipse. Positions and sizes are numbers in the document unit or strings
        such as '1in' or '25mm'. Optional text is centered in the shape by default."""
        return self.addFilledShape("ellipse", x, y, width, height, fillcolor, strokecolor, strokewidth,
                                   name, text, bold, italic, underline, fontsize, fontcolor, align)

    def addFilledShape(self, shapeType, x, y, width, height, fillcolor, strokecolor, strokewidth,
                       name, text, bold, italic, underline, fontsize, fontcolor, align):
        shape = ContentShape(shapeType, {
            "x": self.length(x), "y": self.length(y),
            "width": self.length(width), "height": self.length(height),
            "fillcolor": fillcolor, "strokecolor": strokecolor,
            "strokewidth": self.length(strokewidth),
            "name": name, "text": None, "paraData": None, "defaultAlign": "center",
        })
        if text is not None:
            shape.setText(text, bold, italic, underline, fontsize, fontcolor, align)
        return self.addShape(shape)

    def addLine(self, x1, y1, x2, y2, strokecolor="#000000", strokewidth="0.02in", name=None):
        """Add a line from (x1,y1) to (x2,y2). Coordinates are numbers in the document unit
        or strings such as '1in'."""
        return self.addShape(ContentShape("line", {
            "x1": self.length(x1), "y1": self.length(y1),
            "x2": self.length(x2), "y2": self.length(y2),
            "strokecolor": strokecolor, "strokewidth": self.length(strokewidth),
            "name": name,
        }))

    def addTextBox(self, x, y, width, height, text,
                   bold=False, italic=False, underline=False, fontsize=None, fontcolor=None,
                   name=None, align="left"):
        """Add a text box at the given position. Positions and sizes are numbers in the
        document unit or strings such as '1in'."""
        shape = ContentShape("textbox", {
            "x": self.length(x), "y": self.length(y),
            "width": self.length(width), "height": self.length(height),
            "name": name, "defaultAlign": "left",
        })
        shape.setText(text, bold, italic, underline, fontsize, fontcolor, align)
        return self.addShape(shape)

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

    def pageObject(self):
        """Build and return the draw:page XML element for this page."""
        page = ooolibXML.Element("draw:page")
        page.setAttribute("draw:name", self.name)
        page.setAttribute("draw:style-name", "dp1")
        page.setAttribute("draw:master-page-name", "Default")

        for shape in self.shapes:
            shapeType = shape.type
            data = shape.data

            if shapeType in ("rect", "ellipse"):
                styleName = self.automaticStylesInstance.getShapeStyle(
                    data["fillcolor"], data["strokecolor"], data["strokewidth"])
                elem = page.addChild(ooolibXML.Element("draw:" + shapeType))
                if data["name"] is not None: elem.setAttribute("draw:name", data["name"])
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("svg:x", data["x"])
                elem.setAttribute("svg:y", data["y"])
                elem.setAttribute("svg:width", data["width"])
                elem.setAttribute("svg:height", data["height"])
                self.addParagraph(elem, shape)

            elif shapeType == "line":
                styleName = self.automaticStylesInstance.getLineStyle(
                    data["strokecolor"], data["strokewidth"])
                elem = page.addChild(ooolibXML.Element("draw:line"))
                if data["name"] is not None: elem.setAttribute("draw:name", data["name"])
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("svg:x1", data["x1"])
                elem.setAttribute("svg:y1", data["y1"])
                elem.setAttribute("svg:x2", data["x2"])
                elem.setAttribute("svg:y2", data["y2"])

            elif shapeType == "textbox":
                styleName = self.automaticStylesInstance.getTextBoxStyle()
                frame = page.addChild(ooolibXML.Element("draw:frame"))
                if data["name"] is not None: frame.setAttribute("draw:name", data["name"])
                frame.setAttribute("draw:style-name", styleName)
                frame.setAttribute("svg:x", data["x"])
                frame.setAttribute("svg:y", data["y"])
                frame.setAttribute("svg:width", data["width"])
                frame.setAttribute("svg:height", data["height"])
                textBox = frame.addChild(ooolibXML.Element("draw:text-box"))
                self.addParagraph(textBox, shape)

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
