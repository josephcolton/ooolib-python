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
# ooolib-python Writer Styles object #
##################################
class Styles:
    def __init__(self, global_object):
        self.global_object = global_object
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Styles
        self.documentStyles = ooolibXML.Element("office:document-styles")
        self.documentStyles.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentStyles.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentStyles.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentStyles.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.documentStyles.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.documentStyles.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentStyles.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentStyles.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentStyles.setAttribute("office:version", "1.3")

        #####################
        # Font Face Declarations #
        #####################
        officeFonts = self.documentStyles.addChild(ooolibXML.Element("office:font-face-decls"))
        font1 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font1.setAttribute("style:name", "Liberation Serif")
        font1.setAttribute("svg:font-family", "&apos;Liberation Serif&apos;")
        font1.setAttribute("style:font-family-generic", "roman")
        font1.setAttribute("style:font-pitch", "variable")
        font2 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font2.setAttribute("style:name", "Liberation Sans")
        font2.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
        font2.setAttribute("style:font-family-generic", "swiss")
        font2.setAttribute("style:font-pitch", "variable")
        font3 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font3.setAttribute("style:name", "Liberation Mono")
        font3.setAttribute("svg:font-family", "&apos;Liberation Mono&apos;")
        font3.setAttribute("style:font-family-generic", "modern")
        font3.setAttribute("style:font-pitch", "fixed")

        #################
        # Office Styles #
        #################
        officeStyles = self.documentStyles.addChild(ooolibXML.Element("office:styles"))

        # Default paragraph style
        defaultStyle = officeStyles.addChild(ooolibXML.Element("style:default-style"))
        defaultStyle.setAttribute("style:family", "paragraph")
        defaultParaProp = defaultStyle.addChild(ooolibXML.Element("style:paragraph-properties"))
        defaultParaProp.setAttribute("style:tab-stop-distance", "0.5in")
        defaultParaProp.setAttribute("style:writing-mode", "page")
        defaultTextProp = defaultStyle.addChild(ooolibXML.Element("style:text-properties"))
        defaultTextProp.setAttribute("style:font-name", "Liberation Serif")
        defaultTextProp.setAttribute("fo:font-size", "12pt")
        defaultTextProp.setAttribute("fo:language", "en")
        defaultTextProp.setAttribute("fo:country", "US")

        # Standard paragraph style (base for body text)
        styleStandard = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleStandard.setAttribute("style:name", "Standard")
        styleStandard.setAttribute("style:family", "paragraph")
        styleStandard.setAttribute("style:class", "text")

        # Heading base style
        styleHeading = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleHeading.setAttribute("style:name", "Heading")
        styleHeading.setAttribute("style:family", "paragraph")
        styleHeading.setAttribute("style:parent-style-name", "Standard")
        styleHeading.setAttribute("style:next-style-name", "Standard")
        styleHeading.setAttribute("style:class", "text")
        headingParaProp = styleHeading.addChild(ooolibXML.Element("style:paragraph-properties"))
        headingParaProp.setAttribute("fo:margin-top", "0.165in")
        headingParaProp.setAttribute("fo:margin-bottom", "0.083in")
        headingParaProp.setAttribute("fo:keep-with-next", "always")
        headingTextProp = styleHeading.addChild(ooolibXML.Element("style:text-properties"))
        headingTextProp.setAttribute("style:font-name", "Liberation Sans")
        headingTextProp.setAttribute("fo:font-weight", "bold")
        headingTextProp.setAttribute("style:font-weight-asian", "bold")
        headingTextProp.setAttribute("style:font-weight-complex", "bold")

        # Heading 1
        styleH1 = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleH1.setAttribute("style:name", "Heading_20_1")
        styleH1.setAttribute("style:display-name", "Heading 1")
        styleH1.setAttribute("style:family", "paragraph")
        styleH1.setAttribute("style:parent-style-name", "Heading")
        styleH1.setAttribute("style:next-style-name", "Standard")
        styleH1.setAttribute("style:default-outline-level", "1")
        h1TextProp = styleH1.addChild(ooolibXML.Element("style:text-properties"))
        h1TextProp.setAttribute("fo:font-size", "24pt")
        h1TextProp.setAttribute("style:font-size-asian", "24pt")
        h1TextProp.setAttribute("style:font-size-complex", "24pt")

        # Heading 2
        styleH2 = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleH2.setAttribute("style:name", "Heading_20_2")
        styleH2.setAttribute("style:display-name", "Heading 2")
        styleH2.setAttribute("style:family", "paragraph")
        styleH2.setAttribute("style:parent-style-name", "Heading")
        styleH2.setAttribute("style:next-style-name", "Standard")
        styleH2.setAttribute("style:default-outline-level", "2")
        h2TextProp = styleH2.addChild(ooolibXML.Element("style:text-properties"))
        h2TextProp.setAttribute("fo:font-size", "18pt")
        h2TextProp.setAttribute("style:font-size-asian", "18pt")
        h2TextProp.setAttribute("style:font-size-complex", "18pt")

        # Heading 3
        styleH3 = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleH3.setAttribute("style:name", "Heading_20_3")
        styleH3.setAttribute("style:display-name", "Heading 3")
        styleH3.setAttribute("style:family", "paragraph")
        styleH3.setAttribute("style:parent-style-name", "Heading")
        styleH3.setAttribute("style:next-style-name", "Standard")
        styleH3.setAttribute("style:default-outline-level", "3")
        h3TextProp = styleH3.addChild(ooolibXML.Element("style:text-properties"))
        h3TextProp.setAttribute("fo:font-size", "14pt")
        h3TextProp.setAttribute("style:font-size-asian", "14pt")
        h3TextProp.setAttribute("style:font-size-complex", "14pt")

        # Heading 4
        styleH4 = officeStyles.addChild(ooolibXML.Element("style:style"))
        styleH4.setAttribute("style:name", "Heading_20_4")
        styleH4.setAttribute("style:display-name", "Heading 4")
        styleH4.setAttribute("style:family", "paragraph")
        styleH4.setAttribute("style:parent-style-name", "Heading")
        styleH4.setAttribute("style:next-style-name", "Standard")
        styleH4.setAttribute("style:default-outline-level", "4")
        h4TextProp = styleH4.addChild(ooolibXML.Element("style:text-properties"))
        h4TextProp.setAttribute("fo:font-size", "12pt")
        h4TextProp.setAttribute("fo:font-style", "italic")
        h4TextProp.setAttribute("style:font-size-asian", "12pt")
        h4TextProp.setAttribute("style:font-size-complex", "12pt")

        ###########################
        # Office Automatic Styles #
        ###########################
        officeAutomaticStyles = self.documentStyles.addChild(ooolibXML.Element("office:automatic-styles"))
        # Page layout: US Letter, 1" top/bottom margins, 1.25" left/right margins
        pageLayout = officeAutomaticStyles.addChild(ooolibXML.Element("style:page-layout"))
        pageLayout.setAttribute("style:name", "pm1")
        pageLayoutProp = pageLayout.addChild(ooolibXML.Element("style:page-layout-properties"))
        pageLayoutProp.setAttribute("fo:page-width", "8.5in")
        pageLayoutProp.setAttribute("fo:page-height", "11in")
        pageLayoutProp.setAttribute("style:print-orientation", "portrait")
        pageLayoutProp.setAttribute("fo:margin-top", "1in")
        pageLayoutProp.setAttribute("fo:margin-bottom", "1in")
        pageLayoutProp.setAttribute("fo:margin-left", "1.25in")
        pageLayoutProp.setAttribute("fo:margin-right", "1.25in")

        ########################
        # Office Master Styles #
        ########################
        officeMasterStyles = self.documentStyles.addChild(ooolibXML.Element("office:master-styles"))
        masterPage = officeMasterStyles.addChild(ooolibXML.Element("style:master-page"))
        masterPage.setAttribute("style:name", "Standard")
        masterPage.setAttribute("style:page-layout-name", "pm1")

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

    print("Writer Styles XML with indentation:")
    print(string1)
