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

#########################
# Line Ends and Dashes  #
#########################
# Line end markers: name -> (draw:name, svg:viewBox, svg:d, centered on the line end)
MARKERS = {
    "arrow": ("Arrow", "0 0 20 30", "M10 0l-10 30h20z", False),
    "circle": ("Circle", "0 0 20 20", "M10 0c5.5 0 10 4.5 10 10s-4.5 10-10 10s-10-4.5-10-10s4.5-10 10-10z", True),
    "square": ("Square", "0 0 10 10", "M0 0h10v10h-10z", True),
    "diamond": ("Diamond", "0 0 10 10", "M5 0l5 5l-5 5l-5-5z", True),
}

# Dashed line styles: name -> (draw:name, dots1, dots1-length, dots2, dots2-length, distance)
# Lengths are percentages of the line width so dashes scale with thicker lines.
LINE_STYLES = {
    "solid": None,
    "dash": ("Dash", 1, "300%", 0, None, "200%"),
    "dot": ("Dot", 1, "100%", 0, None, "100%"),
    "dashdot": ("DashDot", 1, "300%", 1, "100%", "150%"),
    "longdash": ("LongDash", 1, "800%", 0, None, "300%"),
}

def addLineDefinitions(officeStyles):
    """Add the line end markers and dash styles to office:styles (Draw and Impress share them)."""
    for markerName in MARKERS:
        drawName, viewBox, path, centered = MARKERS[markerName]
        marker = officeStyles.addChild(ooolibXML.Element("draw:marker"))
        marker.setAttribute("draw:name", drawName)
        marker.setAttribute("svg:viewBox", viewBox)
        marker.setAttribute("svg:d", path)
    for lineStyle in LINE_STYLES:
        if LINE_STYLES[lineStyle] is None: continue
        drawName, dots1, dots1Length, dots2, dots2Length, distance = LINE_STYLES[lineStyle]
        dash = officeStyles.addChild(ooolibXML.Element("draw:stroke-dash"))
        dash.setAttribute("draw:name", drawName)
        dash.setAttribute("draw:style", "rect")
        dash.setAttribute("draw:dots1", str(dots1))
        dash.setAttribute("draw:dots1-length", dots1Length)
        if dots2:
            dash.setAttribute("draw:dots2", str(dots2))
            dash.setAttribute("draw:dots2-length", dots2Length)
        dash.setAttribute("draw:distance", distance)

##################################
# ooolib-python Draw Styles object #
##################################
class Styles:
    def __init__(self, global_object):
        self.global_object = global_object
        self.prolog = ooolibXML.Prolog("xml")
        self.documentStyles = ooolibXML.Element("office:document-styles")
        self.documentStyles.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentStyles.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentStyles.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentStyles.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
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
        addLineDefinitions(officeStyles)

        # Default drawing-page style (no background)
        styleDP = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleDP.setAttribute("style:name", "dp1")
        styleDP.setAttribute("style:family", "drawing-page")
        styleDP.addChild(ooolibXML.Element("style:drawing-page-properties"))

        # Default paragraph style for text within shapes
        styleStandard = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleStandard.setAttribute("style:name", "Standard")
        styleStandard.setAttribute("style:family", "paragraph")
        styleStandard.setAttribute("style:class", "text")
        standardTextProp = styleStandard.addChild(ooolibXML.Element("style:text-properties"))
        standardTextProp.setAttribute("style:font-name", "Liberation Sans")
        standardTextProp.setAttribute("fo:font-size", "12pt")

        # Default graphic style
        defaultGraphic = officeStyles.addChild(ooolibXML.Element("style:default-style"))
        defaultGraphic.setAttribute("style:family", "graphic")
        defaultGraphicProp = defaultGraphic.addChild(ooolibXML.Element("style:graphic-properties"))
        defaultGraphicProp.setAttribute("draw:fill", "solid")
        defaultGraphicProp.setAttribute("draw:fill-color", "#4472c4")
        defaultGraphicProp.setAttribute("draw:stroke", "solid")
        defaultGraphicProp.setAttribute("svg:stroke-color", "#000000")
        defaultGraphicProp.setAttribute("svg:stroke-width", "0.02in")

        ###########################
        # Office Automatic Styles #
        ###########################
        officeAutomaticStyles = self.documentStyles.addChild(ooolibXML.Element("office:automatic-styles"))
        # Page layout: 10" x 7.5"
        pageLayout = officeAutomaticStyles.addChild(ooolibXML.Element("style:page-layout"))
        pageLayout.setAttribute("style:name", "pm1")
        self.pageLayoutProp = pageLayout.addChild(ooolibXML.Element("style:page-layout-properties"))
        self.setPageSize("10in", "7.5in")
        self.setPageMargins("0.5in")

        ########################
        # Office Master Styles #
        ########################
        officeMasterStyles = self.documentStyles.addChild(ooolibXML.Element("office:master-styles"))
        masterPage = officeMasterStyles.addChild(ooolibXML.Element("style:master-page"))
        masterPage.setAttribute("style:name", "Default")
        masterPage.setAttribute("style:page-layout-name", "pm1")
        masterPage.setAttribute("draw:style-name", "dp1")

    ##############
    # Page Setup #
    ##############
    def setPageSize(self, width, height, orientation=None):
        """Set the size of every page. Numbers use the document unit (see Content.setUnit).

        orientation is 'landscape' or 'portrait'; by default it is worked out from the size."""
        unit = self.global_object.getGlobalString("drawUnit", "in")
        width = ooolibUnits.formatLength(width, unit)
        height = ooolibUnits.formatLength(height, unit)
        if orientation is None:
            widthIn = ooolibUnits.convertLength(width, "in")
            heightIn = ooolibUnits.convertLength(height, "in")
            if widthIn is not None and heightIn is not None and widthIn < heightIn:
                orientation = "portrait"
            else:
                orientation = "landscape"
        if orientation not in ("landscape", "portrait"):
            raise ValueError("orientation must be 'landscape' or 'portrait', not %r" % orientation)
        self.pageLayoutProp.setAttribute("fo:page-width", width)
        self.pageLayoutProp.setAttribute("fo:page-height", height)
        self.pageLayoutProp.setAttribute("style:print-orientation", orientation)

    def setPageMargins(self, margin=None, top=None, bottom=None, left=None, right=None):
        """Set the page margins. margin sets all four; top/bottom/left/right override it.
        Numbers use the document unit (see Content.setUnit)."""
        unit = self.global_object.getGlobalString("drawUnit", "in")
        sides = (("top", top), ("bottom", bottom), ("left", left), ("right", right))
        for side, value in sides:
            if value is None: value = margin
            if value is None: continue
            self.pageLayoutProp.setAttribute("fo:margin-%s" % side, ooolibUnits.formatLength(value, unit))

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
    print("Draw Styles XML with indentation:")
    print(string1)
