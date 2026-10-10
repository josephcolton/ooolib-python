#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import ooolib-python modules
import ooolibGlobal
import ooolibDrawStyles
import ooolibUnits
import ooolibXML

# Title and content placeholders as fractions of the slide: (x, y, width, height)
# On the default 10in x 7.5in slide these are 0.6/0.3/8.8/1.2in and 0.6/1.8/8.8/5.3in.
PLACEHOLDERS = {
    "title": (0.06, 0.04, 0.88, 0.16),
    "content": (0.06, 0.24, 0.88, 5.3 / 7.5),
}

####################################
# ooolib-python Impress Styles object #
####################################
class Styles:
    def __init__(self, global_object):
        self.global_object = global_object
        self.prolog = ooolibXML.Prolog("xml")
        self.documentStyles = ooolibXML.Element("office:document-styles")
        self.documentStyles.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentStyles.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentStyles.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentStyles.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.documentStyles.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
        self.documentStyles.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentStyles.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentStyles.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentStyles.setAttribute("office:version", "1.3")

        #####################
        # Font Face Declarations #
        #####################
        officeFonts = self.documentStyles.addChild(ooolibXML.Element("office:font-face-decls"))
        font1 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font1.setAttribute("style:name", "Liberation Sans")
        font1.setAttribute("svg:font-family", "'Liberation Sans'")
        font1.setAttribute("style:font-family-generic", "swiss")
        font1.setAttribute("style:font-pitch", "variable")
        font2 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font2.setAttribute("style:name", "Liberation Serif")
        font2.setAttribute("svg:font-family", "'Liberation Serif'")
        font2.setAttribute("style:font-family-generic", "roman")
        font2.setAttribute("style:font-pitch", "variable")

        #################
        # Office Styles #
        #################
        officeStyles = self.documentStyles.addChild(ooolibXML.Element("office:styles"))

        # Line end markers and dash styles referenced by shape styles
        ooolibDrawStyles.addLineDefinitions(officeStyles)

        # Default drawing-page style (no background)
        styleDP = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleDP.setAttribute("style:name", "dp1")
        styleDP.setAttribute("style:family", "drawing-page")
        dpProp = styleDP.addChild(ooolibXML.Element("style:drawing-page-properties"))
        dpProp.setAttribute("presentation:display-header", "false")
        dpProp.setAttribute("presentation:display-footer", "false")
        dpProp.setAttribute("presentation:display-page-number", "false")
        dpProp.setAttribute("presentation:display-date-time", "false")

        # Default paragraph style for slide text
        styleStandard = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleStandard.setAttribute("style:name", "Standard")
        styleStandard.setAttribute("style:family", "paragraph")
        styleStandard.setAttribute("style:class", "text")
        standardTextProp = styleStandard.addChild(ooolibXML.Element("style:text-properties"))
        standardTextProp.setAttribute("style:font-name", "Liberation Sans")
        standardTextProp.setAttribute("fo:font-size", "24pt")

        # Presentation style: title frame text
        styleTitle = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleTitle.setAttribute("style:name", "Default-title")
        styleTitle.setAttribute("style:family", "presentation")
        # Without these LibreOffice draws its default blue fill and border on the frame
        titleGraphicProp = styleTitle.addChild(ooolibXML.Element("style:graphic-properties"))
        titleGraphicProp.setAttribute("draw:fill", "none")
        titleGraphicProp.setAttribute("draw:stroke", "none")
        titleParaProp = styleTitle.addChild(ooolibXML.Element("style:paragraph-properties"))
        titleParaProp.setAttribute("fo:text-align", "center")
        titleTextProp = styleTitle.addChild(ooolibXML.Element("style:text-properties"))
        titleTextProp.setAttribute("style:font-name", "Liberation Sans")
        titleTextProp.setAttribute("fo:font-size", "44pt")
        titleTextProp.setAttribute("fo:font-weight", "bold")
        titleTextProp.setAttribute("style:font-weight-asian", "bold")
        titleTextProp.setAttribute("style:font-weight-complex", "bold")

        # Presentation style: content/subtitle frame text
        styleSubtitle = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleSubtitle.setAttribute("style:name", "Default-subtitle")
        styleSubtitle.setAttribute("style:family", "presentation")
        subtitleGraphicProp = styleSubtitle.addChild(ooolibXML.Element("style:graphic-properties"))
        subtitleGraphicProp.setAttribute("draw:fill", "none")
        subtitleGraphicProp.setAttribute("draw:stroke", "none")
        subtitleTextProp = styleSubtitle.addChild(ooolibXML.Element("style:text-properties"))
        subtitleTextProp.setAttribute("style:font-name", "Liberation Sans")
        subtitleTextProp.setAttribute("fo:font-size", "28pt")

        ###########################
        # Office Automatic Styles #
        ###########################
        officeAutomaticStyles = self.documentStyles.addChild(ooolibXML.Element("office:automatic-styles"))
        # Slide page layout: 10" x 7.5" landscape
        pageLayout = officeAutomaticStyles.addChild(ooolibXML.Element("style:page-layout"))
        pageLayout.setAttribute("style:name", "pm1")
        self.pageLayoutProp = pageLayout.addChild(ooolibXML.Element("style:page-layout-properties"))
        pageLayoutProp = self.pageLayoutProp
        pageLayoutProp.setAttribute("fo:margin-top", "0in")
        pageLayoutProp.setAttribute("fo:margin-bottom", "0in")
        pageLayoutProp.setAttribute("fo:margin-left", "0in")
        pageLayoutProp.setAttribute("fo:margin-right", "0in")

        ########################
        # Office Master Styles #
        ########################
        officeMasterStyles = self.documentStyles.addChild(ooolibXML.Element("office:master-styles"))
        # Layer set applies globally to all slides
        layerSet = officeMasterStyles.addChild(ooolibXML.Element("draw:layer-set"))
        for layerName in ["layout", "background", "backgroundobjects", "controls", "measurelines"]:
            layer = layerSet.addChild(ooolibXML.Element("draw:layer"))
            layer.setAttribute("draw:name", layerName)
        # Default master page
        masterPage = officeMasterStyles.addChild(ooolibXML.Element("style:master-page"))
        masterPage.setAttribute("style:name", "Default")
        masterPage.setAttribute("style:page-layout-name", "pm1")
        masterPage.setAttribute("draw:style-name", "dp1")
        # Master page placeholder: title
        titleHolder = masterPage.addChild(ooolibXML.Element("draw:frame"))
        titleHolder.setAttribute("presentation:class", "title")
        titleHolder.setAttribute("presentation:placeholder", "true")
        titleHolder.setAttribute("presentation:style-name", "Default-title")
        titleHolder.setAttribute("draw:layer", "backgroundobjects")
        titleHolder.addChild(ooolibXML.Element("draw:text-box")).addChild(ooolibXML.Element("text:p"))
        # Master page placeholder: body content
        bodyHolder = masterPage.addChild(ooolibXML.Element("draw:frame"))
        # Placeholders are only outlines for editing; without presentation:placeholder
        # LibreOffice draws them as ordinary (blue) boxes behind every slide
        bodyHolder.setAttribute("presentation:class", "outline")
        bodyHolder.setAttribute("presentation:placeholder", "true")
        bodyHolder.setAttribute("presentation:style-name", "Default-subtitle")
        bodyHolder.setAttribute("draw:layer", "backgroundobjects")
        bodyHolder.addChild(ooolibXML.Element("draw:text-box")).addChild(ooolibXML.Element("text:p"))
        self.placeholderFrames = {"title": titleHolder, "content": bodyHolder}
        self.setPageSize("10in", "7.5in")

    ##############
    # Slide Size #
    ##############
    def setPageSize(self, width, height, orientation=None):
        """Set the size of every slide, e.g. setPageSize("13.333in", "7.5in") for 16:9.
        Numbers use the document unit (see Content.setUnit). The title and content
        areas are resized to match. orientation is worked out from the size by default."""
        unit = self.global_object.getGlobalString("drawUnit", "in")
        width = ooolibUnits.formatLength(width, unit)
        height = ooolibUnits.formatLength(height, unit)
        widthIn = ooolibUnits.convertLength(width, "in")
        heightIn = ooolibUnits.convertLength(height, "in")
        if widthIn is None or heightIn is None:
            raise ValueError("Unable to use slide size %r x %r" % (width, height))
        if orientation is None:
            orientation = "portrait" if widthIn < heightIn else "landscape"
        if orientation not in ("landscape", "portrait"):
            raise ValueError("orientation must be 'landscape' or 'portrait', not %r" % orientation)
        self.pageLayoutProp.setAttribute("fo:page-width", width)
        self.pageLayoutProp.setAttribute("fo:page-height", height)
        self.pageLayoutProp.setAttribute("style:print-orientation", orientation)
        # Placeholder areas, shared with the slides
        placeholders = {}
        for name in PLACEHOLDERS:
            fractionX, fractionY, fractionWidth, fractionHeight = PLACEHOLDERS[name]
            geometry = (fractionX * widthIn, fractionY * heightIn, fractionWidth * widthIn, fractionHeight * heightIn)
            geometry = tuple(ooolibUnits.formatLength(round(value, 4), "in") for value in geometry)
            placeholders[name] = geometry
            frame = self.placeholderFrames[name]
            for attribute, value in zip(("svg:x", "svg:y", "svg:width", "svg:height"), geometry):
                frame.setAttribute(attribute, value)
        self.global_object.setGlobalObjects("impPlaceholders", placeholders)

    def getPageSize(self):
        """Return (width, height) as strings, e.g. ('10in', '7.5in')."""
        return (self.pageLayoutProp.getAttribute("fo:page-width"),
                self.pageLayoutProp.getAttribute("fo:page-height"))

    def toString(self, indent=False):
        stylesString = self.prolog.toString()
        stylesString += self.documentStyles.toString(indent)
        return stylesString


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    styles = Styles(g)
    string1 = styles.toString(indent=True)
    print("Impress Styles XML with indentation:")
    print(string1)
