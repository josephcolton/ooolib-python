#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import ooolib-python modules
import ooolibGlobal
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
        fontFace1.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
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
            self.extractData(paraData, "fontcolor") is not None
        )
        if not has_formatting:
            return "Standard"
        index = ("para",
                 self.extractData(paraData, "bold"),
                 self.extractData(paraData, "italic"),
                 self.extractData(paraData, "underline"),
                 self.extractData(paraData, "fontsize"),
                 self.extractData(paraData, "fontcolor"))
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
        text = str(text)
        text = text.replace("&", "&amp;")
        text = text.replace("'", "&apos;")
        text = text.replace('"', "&quot;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace("\t", "<text:tab/>")
        text = text.replace("\n", "<text:line-break/>")
        return text

    def addRectangle(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in"):
        """Add a rectangle. Coordinates and sizes are strings (e.g. '1in', '2.5in')."""
        self.shapes.append({
            "type": "rect",
            "x": x, "y": y, "width": width, "height": height,
            "fillcolor": fillcolor, "strokecolor": strokecolor, "strokewidth": strokewidth,
        })

    def addEllipse(self, x, y, width, height, fillcolor=None, strokecolor=None, strokewidth="0.02in"):
        """Add an ellipse. Coordinates and sizes are strings (e.g. '1in', '2.5in')."""
        self.shapes.append({
            "type": "ellipse",
            "x": x, "y": y, "width": width, "height": height,
            "fillcolor": fillcolor, "strokecolor": strokecolor, "strokewidth": strokewidth,
        })

    def addLine(self, x1, y1, x2, y2, strokecolor="#000000", strokewidth="0.02in"):
        """Add a line from (x1,y1) to (x2,y2). Coordinates are strings (e.g. '1in')."""
        self.shapes.append({
            "type": "line",
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
            "strokecolor": strokecolor, "strokewidth": strokewidth,
        })

    def addTextBox(self, x, y, width, height, text,
                   bold=False, italic=False, underline=False, fontsize=None, fontcolor=None):
        """Add a text box at the given position. Coordinates and sizes are strings."""
        paraData = {
            "bold": bold, "italic": italic, "underline": underline,
            "fontsize": fontsize, "fontcolor": fontcolor,
        }
        self.shapes.append({
            "type": "textbox",
            "x": x, "y": y, "width": width, "height": height,
            "text": text, "paraData": paraData,
        })

    def pageObject(self):
        """Build and return the draw:page XML element for this page."""
        page = ooolibXML.Element("draw:page")
        page.setAttribute("draw:name", self.name)
        page.setAttribute("draw:style-name", "dp1")
        page.setAttribute("draw:master-page-name", "Default")

        for shape in self.shapes:
            shapeType = shape["type"]

            if shapeType == "rect":
                styleName = self.automaticStylesInstance.getShapeStyle(
                    shape["fillcolor"], shape["strokecolor"], shape["strokewidth"])
                elem = page.addChild(ooolibXML.Element("draw:rect"))
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("svg:x", shape["x"])
                elem.setAttribute("svg:y", shape["y"])
                elem.setAttribute("svg:width", shape["width"])
                elem.setAttribute("svg:height", shape["height"])
                elem.addChild(ooolibXML.Element("text:p"))

            elif shapeType == "ellipse":
                styleName = self.automaticStylesInstance.getShapeStyle(
                    shape["fillcolor"], shape["strokecolor"], shape["strokewidth"])
                elem = page.addChild(ooolibXML.Element("draw:ellipse"))
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("svg:x", shape["x"])
                elem.setAttribute("svg:y", shape["y"])
                elem.setAttribute("svg:width", shape["width"])
                elem.setAttribute("svg:height", shape["height"])
                elem.addChild(ooolibXML.Element("text:p"))

            elif shapeType == "line":
                styleName = self.automaticStylesInstance.getLineStyle(
                    shape["strokecolor"], shape["strokewidth"])
                elem = page.addChild(ooolibXML.Element("draw:line"))
                elem.setAttribute("draw:style-name", styleName)
                elem.setAttribute("svg:x1", shape["x1"])
                elem.setAttribute("svg:y1", shape["y1"])
                elem.setAttribute("svg:x2", shape["x2"])
                elem.setAttribute("svg:y2", shape["y2"])

            elif shapeType == "textbox":
                styleName = self.automaticStylesInstance.getTextBoxStyle()
                paraStyleName = self.automaticStylesInstance.getParagraphStyle(shape["paraData"])
                frame = page.addChild(ooolibXML.Element("draw:frame"))
                frame.setAttribute("draw:style-name", styleName)
                frame.setAttribute("svg:x", shape["x"])
                frame.setAttribute("svg:y", shape["y"])
                frame.setAttribute("svg:width", shape["width"])
                frame.setAttribute("svg:height", shape["height"])
                textBox = frame.addChild(ooolibXML.Element("draw:text-box"))
                p = textBox.addChild(ooolibXML.Element("text:p"))
                p.setAttribute("text:style-name", paraStyleName)
                p.setText(self.escapeText(shape["text"]))

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
