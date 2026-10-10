#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import ooolib-python modules
import ooolibGlobal
import ooolibDrawContent
import ooolibXML

# Slide layouts and the placeholder areas each one has
LAYOUTS = {
    "title-content": ("title", "content"),
    "title": ("title",),
    "blank": (),
}

##################################
# ooolib-python Impress Content  #
##################################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        self.internal_slides = []
        self.activeSlide = None
        self.global_object.setGlobalInt("impSlideNumber", 1)
        # Slides use the Draw shape code, which keeps its settings under "draw" names
        self.global_object.setGlobalString("drawUnit", "in")
        self.global_object.setGlobalInt("drawShapeId", 1)
        self.pictureStore = ooolibDrawContent.PictureStore()
        self.global_object.setGlobalObjects("drawPictureStore", self.pictureStore)
        # Automatic styles
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles)
        self.global_object.setGlobalObjects("impAutoStyles", self.automaticStylesInstance)
        self.global_object.setGlobalObjects("drawAutoStyles", self.automaticStylesInstance)
        # Create the two fixed frame styles used for title and content frames
        self._titleFrameStyle = self.automaticStylesInstance.createFrameStyle()
        self._contentFrameStyle = self.automaticStylesInstance.createFrameStyle()
        self.global_object.setGlobalString("impTitleFrameStyle", self._titleFrameStyle)
        self.global_object.setGlobalString("impContentFrameStyle", self._contentFrameStyle)
        # Title and content areas for 10in x 7.5in slides (Styles.setPageSize changes them)
        self.global_object.setGlobalObjects("impPlaceholders", {
            "title": ("0.6in", "0.3in", "8.8in", "1.2in"),
            "content": ("0.6in", "1.8in", "8.8in", "5.3in"),
        })
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
        fontFace1.setAttribute("svg:font-family", "'Liberation Sans'")
        fontFace1.setAttribute("style:font-family-generic", "swiss")
        fontFace1.setAttribute("style:font-pitch", "variable")
        # Attach automatic styles
        self.documentContent.addChild(self.automaticStyles)
        # Office body
        officeBody = self.documentContent.addChild(ooolibXML.Element("office:body"))
        self.officePresentation = officeBody.addChild(ooolibXML.Element("office:presentation"))
        # Add initial slide
        self.activeSlide = self.addSlide()

    def addSlide(self, name=None, layout="title-content"):
        """Add a slide and make it the active slide.

        layout: "title-content" (title and content areas), "title" (title only),
        or "blank" (no title or content areas, for shapes and images)."""
        slide = ContentSlide(self.global_object, name, layout)
        self.internal_slides.append(slide)
        self.activeSlide = slide
        return slide

    # Numeric positions and sizes work the same way as in Draw
    setUnit = ooolibDrawContent.Content.setUnit
    getUnit = ooolibDrawContent.Content.getUnit

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
        self.officePresentation.clearChildren()
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
class AutomaticStyles(ooolibDrawContent.AutomaticStyles):
    """Draw's shape and paragraph styles, plus slide frame and background styles."""
    def __init__(self, global_object, automaticStyles):
        ooolibDrawContent.AutomaticStyles.__init__(self, global_object, automaticStyles)
        self.global_object.setGlobalInt("impStyleFrame", 1)
        # dp1 is the default drawing-page style in styles.xml, so backgrounds start at dp2
        self.global_object.setGlobalInt("impStylePage", 2)

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


######################
# Content Slide      #
######################
class ContentSlide(ooolibDrawContent.ShapeContainer):
    """A slide with a title and content area. Shapes, images, groups, and connectors
    are added with the same methods as a Draw page (addRectangle, addImage, ...)."""
    def __init__(self, global_object, name=None, layout="title-content"):
        ooolibDrawContent.ShapeContainer.__init__(self, global_object)
        self.name = name
        if self.name is None:
            self.name = "Slide%d" % self.global_object.incrementGlobalInt("impSlideNumber")
        self.title_text = ""
        self.content_lines = []  # list of (text, paraData) tuples
        self.backgroundColor = None
        self.layout = None
        self.setLayout(layout)

    def getSlideName(self):
        return self.name

    def getLayout(self):
        return self.layout

    def setLayout(self, layout):
        """Change the slide layout: "title-content", "title", or "blank"."""
        if layout not in LAYOUTS:
            raise ValueError("layout must be one of: %s" % ", ".join(LAYOUTS))
        if self.title_text and "title" not in LAYOUTS[layout]:
            raise ValueError("The %r layout has no title area, but this slide has a title" % layout)
        if self.content_lines and "content" not in LAYOUTS[layout]:
            raise ValueError("The %r layout has no content area, but this slide has content" % layout)
        self.layout = layout

    def setTitle(self, text):
        """Set the title text for this slide."""
        if "title" not in LAYOUTS[self.layout]:
            raise ValueError("The %r layout has no title area" % self.layout)
        self.title_text = text

    def addContent(self, text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None):
        """Add a line of content text with optional formatting."""
        if "content" not in LAYOUTS[self.layout]:
            raise ValueError("The %r layout has no content area" % self.layout)
        paraData = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "fontsize": ooolibDrawContent.formatFontSize(fontsize),
            "fontcolor": fontcolor,
        }
        self.content_lines.append((text, paraData))

    def setBackgroundColor(self, color):
        """Set the slide background to a solid color (hex string, e.g. '#336699')."""
        self.backgroundColor = color

    def slideObject(self):
        """Build and return the draw:page XML element for this slide."""
        titleFrameStyle = self.global_object.getGlobalString("impTitleFrameStyle")
        contentFrameStyle = self.global_object.getGlobalString("impContentFrameStyle")
        placeholders = self.global_object.getGlobalObjects("impPlaceholders")

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
        if "title" in LAYOUTS[self.layout]:
            titleFrame = page.addChild(ooolibXML.Element("draw:frame"))
            titleFrame.setAttribute("presentation:class", "title")
            titleFrame.setAttribute("draw:layer", "layout")
            titleFrame.setAttribute("draw:style-name", titleFrameStyle)
            titleFrame.setAttribute("presentation:style-name", "Default-title")
            self.setPlaceholderPosition(titleFrame, placeholders["title"])
            titleTextBox = titleFrame.addChild(ooolibXML.Element("draw:text-box"))
            titleP = titleTextBox.addChild(ooolibXML.Element("text:p"))
            if self.title_text:
                titleP.setText(self.escapeText(self.title_text))

        # Content frame
        if "content" in LAYOUTS[self.layout]:
            contentFrame = page.addChild(ooolibXML.Element("draw:frame"))
            contentFrame.setAttribute("presentation:class", "subtitle")
            contentFrame.setAttribute("draw:layer", "layout")
            contentFrame.setAttribute("draw:style-name", contentFrameStyle)
            contentFrame.setAttribute("presentation:style-name", "Default-subtitle")
            self.setPlaceholderPosition(contentFrame, placeholders["content"])
            contentTextBox = contentFrame.addChild(ooolibXML.Element("draw:text-box"))
            if self.content_lines:
                for (lineText, paraData) in self.content_lines:
                    styleName = self.automaticStylesInstance.getParagraphStyle(paraData)
                    p = contentTextBox.addChild(ooolibXML.Element("text:p"))
                    p.setAttribute("text:style-name", styleName)
                    p.setText(self.escapeText(lineText))
            else:
                contentTextBox.addChild(ooolibXML.Element("text:p"))

        # Shapes, images, groups, and connectors added to the slide
        self.buildShapes(page)
        return page

    def setPlaceholderPosition(self, frame, geometry):
        for attribute, value in zip(("svg:x", "svg:y", "svg:width", "svg:height"), geometry):
            frame.setAttribute(attribute, value)


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
