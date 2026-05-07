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

##################################
# ooolib-python Impress Content  #
##################################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        self.internal_slides = []
        self.activeSlide = None
        self.global_object.setGlobalInt("impSlideNumber", 1)
        # Automatic styles
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles)
        self.global_object.setGlobalObjects("impAutoStyles", self.automaticStylesInstance)
        # Create the two fixed frame styles used for title and content frames
        self._titleFrameStyle = self.automaticStylesInstance.createFrameStyle()
        self._contentFrameStyle = self.automaticStylesInstance.createFrameStyle()
        self.global_object.setGlobalString("impTitleFrameStyle", self._titleFrameStyle)
        self.global_object.setGlobalString("impContentFrameStyle", self._contentFrameStyle)
        # XML structure
        self.prolog = ooolibXML.Prolog("xml")
        self.documentContent = ooolibXML.Element("office:document-content")
        self.documentContent.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentContent.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentContent.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentContent.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.documentContent.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
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
        self.officePresentation = officeBody.addChild(ooolibXML.Element("office:presentation"))
        # Add initial slide
        self.activeSlide = self.addSlide()

    def addSlide(self, name=None):
        slide = ContentSlide(self.global_object, name)
        self.internal_slides.append(slide)
        self.activeSlide = slide
        return slide

    def setActiveSlide(self, name):
        for slide in self.internal_slides:
            if slide.getSlideName() == name:
                self.activeSlide = slide
                return self.activeSlide
        print("setActiveSlide('%s') - Unable to find slide" % name)
        return self.activeSlide

    def updateStats(self):
        self.global_object.setGlobalInt("slideCount", len(self.internal_slides))

    def buildSlides(self):
        for slide in self.internal_slides:
            self.officePresentation.addChild(slide.slideObject())

    def toString(self, indent=False):
        self.buildSlides()
        contentString = self.prolog.toString()
        contentString += self.documentContent.toString(indent)
        return contentString


######################
# Automatic Styles   #
######################
class AutomaticStyles:
    def __init__(self, global_object, automaticStyles):
        self.global_object = global_object
        self.officeAutomaticStyles = automaticStyles
        self.styleObjects = {}
        self.global_object.setGlobalInt("impStyleFrame", 1)
        self.global_object.setGlobalInt("impStylePage", 1)
        self.global_object.setGlobalInt("impStyleParagraph", 1)

    def extractData(self, data, name, default=None):
        if data == None: return default
        if name not in data: return default
        val = data[name]
        if val is None: return default
        return val

    def createFrameStyle(self):
        """Create a graphic style for presentation frames (no fill, no stroke)."""
        frameNum = self.global_object.incrementGlobalInt("impStyleFrame")
        frameName = "fr%d" % frameNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", frameName)
        style.setAttribute("style:family", "graphic")
        graphicProp = style.addChild(ooolibXML.Element("style:graphic-properties"))
        graphicProp.setAttribute("draw:fill", "none")
        graphicProp.setAttribute("draw:stroke", "none")
        return frameName

    def getDrawingPageStyle(self, bgcolor):
        """Return (creating if needed) an automatic drawing-page style for a background color."""
        index = ("page", bgcolor)
        if index not in self.styleObjects:
            self.styleObjects[index] = self.createDrawingPageStyle(bgcolor)
        return self.styleObjects[index]

    def createDrawingPageStyle(self, bgcolor):
        pageNum = self.global_object.incrementGlobalInt("impStylePage")
        pageName = "dp%d" % pageNum
        style = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        style.setAttribute("style:name", pageName)
        style.setAttribute("style:family", "drawing-page")
        pageProp = style.addChild(ooolibXML.Element("style:drawing-page-properties"))
        pageProp.setAttribute("draw:fill", "solid")
        pageProp.setAttribute("draw:fill-color", bgcolor)
        pageProp.setAttribute("presentation:display-header", "false")
        pageProp.setAttribute("presentation:display-footer", "false")
        pageProp.setAttribute("presentation:display-page-number", "false")
        pageProp.setAttribute("presentation:display-date-time", "false")
        self.styleObjects[("page", bgcolor)] = pageName
        return pageName

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
            self.styleObjects[index] = self.createParagraphStyle(paraData)
        return self.styleObjects[index]

    def createParagraphStyle(self, paraData):
        paraNum = self.global_object.incrementGlobalInt("impStyleParagraph")
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


######################
# Content Slide      #
######################
class ContentSlide:
    # Title area dimensions
    TITLE_X = "0.6in"
    TITLE_Y = "0.3in"
    TITLE_W = "8.8in"
    TITLE_H = "1.2in"
    # Content area dimensions
    CONTENT_X = "0.6in"
    CONTENT_Y = "1.8in"
    CONTENT_W = "8.8in"
    CONTENT_H = "5.3in"

    def __init__(self, global_object, name=None):
        self.global_object = global_object
        self.name = name
        if self.name is None:
            self.name = "Slide%d" % self.global_object.incrementGlobalInt("impSlideNumber")
        self.title_text = ""
        self.content_lines = []  # list of (text, paraData) tuples
        self.backgroundColor = None
        self.automaticStylesInstance = self.global_object.getGlobalObjects("impAutoStyles")

    def getSlideName(self):
        return self.name

    def setTitle(self, text):
        """Set the title text for this slide."""
        self.title_text = text

    def addContent(self, text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None):
        """Add a line of content text with optional formatting."""
        paraData = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "fontsize": fontsize,
            "fontcolor": fontcolor,
        }
        self.content_lines.append((text, paraData))

    def setBackgroundColor(self, color):
        """Set the slide background to a solid color (hex string, e.g. '#336699')."""
        self.backgroundColor = color

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

    def slideObject(self):
        """Build and return the draw:page XML element for this slide."""
        titleFrameStyle = self.global_object.getGlobalString("impTitleFrameStyle")
        contentFrameStyle = self.global_object.getGlobalString("impContentFrameStyle")

        # Resolve drawing-page style
        if self.backgroundColor:
            pageStyleName = self.automaticStylesInstance.getDrawingPageStyle(self.backgroundColor)
        else:
            pageStyleName = "dp1"

        page = ooolibXML.Element("draw:page")
        page.setAttribute("draw:name", self.name)
        page.setAttribute("draw:style-name", pageStyleName)
        page.setAttribute("draw:master-page-name", "Default")

        # Title frame
        titleFrame = page.addChild(ooolibXML.Element("draw:frame"))
        titleFrame.setAttribute("presentation:class", "title")
        titleFrame.setAttribute("draw:layer", "layout")
        titleFrame.setAttribute("draw:style-name", titleFrameStyle)
        titleFrame.setAttribute("presentation:style-name", "Default-title")
        titleFrame.setAttribute("svg:x", self.TITLE_X)
        titleFrame.setAttribute("svg:y", self.TITLE_Y)
        titleFrame.setAttribute("svg:width", self.TITLE_W)
        titleFrame.setAttribute("svg:height", self.TITLE_H)
        titleTextBox = titleFrame.addChild(ooolibXML.Element("draw:text-box"))
        titleP = titleTextBox.addChild(ooolibXML.Element("text:p"))
        if self.title_text:
            titleP.setText(self.escapeText(self.title_text))

        # Content frame
        contentFrame = page.addChild(ooolibXML.Element("draw:frame"))
        contentFrame.setAttribute("presentation:class", "subtitle")
        contentFrame.setAttribute("draw:layer", "layout")
        contentFrame.setAttribute("draw:style-name", contentFrameStyle)
        contentFrame.setAttribute("presentation:style-name", "Default-subtitle")
        contentFrame.setAttribute("svg:x", self.CONTENT_X)
        contentFrame.setAttribute("svg:y", self.CONTENT_Y)
        contentFrame.setAttribute("svg:width", self.CONTENT_W)
        contentFrame.setAttribute("svg:height", self.CONTENT_H)
        contentTextBox = contentFrame.addChild(ooolibXML.Element("draw:text-box"))
        if self.content_lines:
            for (lineText, paraData) in self.content_lines:
                styleName = self.automaticStylesInstance.getParagraphStyle(paraData)
                p = contentTextBox.addChild(ooolibXML.Element("text:p"))
                p.setAttribute("text:style-name", styleName)
                p.setText(self.escapeText(lineText))
        else:
            contentTextBox.addChild(ooolibXML.Element("text:p"))

        return page


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    content = Content(g)

    content.activeSlide.setTitle("Welcome to ooolib Impress")
    content.activeSlide.addContent("Normal content line")
    content.activeSlide.addContent("Bold content line", bold=True)
    content.activeSlide.addContent("Colored content line", fontcolor="#cc0000")
    content.activeSlide.setBackgroundColor("#e8f0fe")

    content.addSlide("Slide Two")
    content.activeSlide.setTitle("Second Slide")
    content.activeSlide.addContent("More content here")

    string1 = content.toString(indent=True)
    print("Impress Content XML with indentation:")
    print(string1)
