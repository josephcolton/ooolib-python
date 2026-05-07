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

################################
# ooolib-python Writer Content #
################################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        self.paragraphCount = 0
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Content
        self.documentContent = ooolibXML.Element("office:document-content")
        self.documentContent.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentContent.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentContent.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentContent.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentContent.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentContent.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.documentContent.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.documentContent.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentContent.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentContent.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentContent.setAttribute("office:version", "1.3")
        # Office Scripts
        self.documentContent.addChild(ooolibXML.Element("office:scripts"))
        # Office Font Face Decls
        officeFontFace = self.documentContent.addChild(ooolibXML.Element("office:font-face-decls"))
        fontFace1 = officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace1.setAttribute("style:name", "Liberation Serif")
        fontFace1.setAttribute("svg:font-family", "&apos;Liberation Serif&apos;")
        fontFace1.setAttribute("style:font-family-generic", "roman")
        fontFace1.setAttribute("style:font-pitch", "variable")
        fontFace2 = officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace2.setAttribute("style:name", "Liberation Sans")
        fontFace2.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
        fontFace2.setAttribute("style:font-family-generic", "swiss")
        fontFace2.setAttribute("style:font-pitch", "variable")
        # Automatic Styles
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles)
        self.documentContent.addChild(self.automaticStyles)
        # Office Body
        officeBody = self.documentContent.addChild(ooolibXML.Element("office:body"))
        self.officeText = officeBody.addChild(ooolibXML.Element("office:text"))

    ################
    # Text Escaping #
    ################
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

    ####################
    # Content Creation #
    ####################
    def addParagraph(self, text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None):
        """Add a paragraph with optional formatting applied to the entire paragraph text."""
        self.paragraphCount += 1
        paraData = {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "fontsize": fontsize,
            "fontcolor": fontcolor,
        }
        styleName = self.automaticStylesInstance.getParagraphStyle(paraData)
        para = self.officeText.addChild(ooolibXML.Element("text:p"))
        para.setAttribute("text:style-name", styleName)
        para.setText(self.escapeText(text))
        return para

    def addHeading(self, text, level=1):
        """Add a heading at the given level (1-4)."""
        self.paragraphCount += 1
        level = max(1, min(4, int(level)))
        styleName = "Heading_20_%d" % level
        heading = self.officeText.addChild(ooolibXML.Element("text:h"))
        heading.setAttribute("text:style-name", styleName)
        heading.setAttribute("text:outline-level", str(level))
        heading.setText(self.escapeText(text))
        return heading

    ##################
    # XML Generation #
    ##################
    def updateStats(self):
        self.global_object.setGlobalInt("paragraphCount", self.paragraphCount)

    def toString(self, indent=False):
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
        self.global_object.setGlobalInt("automaticStyleWriterParagraph", 1)

    def extractData(self, data, name, default=None):
        if data == None: return default
        if name not in data: return default
        val = data[name]
        if val is None: return default
        return val

    def getParagraphStyle(self, paraData=None):
        # Check whether any non-default formatting is requested
        has_formatting = (
            self.extractData(paraData, "bold") or
            self.extractData(paraData, "italic") or
            self.extractData(paraData, "underline") or
            self.extractData(paraData, "fontsize") is not None or
            self.extractData(paraData, "fontcolor") is not None
        )
        if not has_formatting:
            return "Standard"
        # Build deduplication key
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
        paraNum = self.global_object.incrementGlobalInt("automaticStyleWriterParagraph")
        paraName = "P%d" % paraNum
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", paraName)
        styleObject.setAttribute("style:family", "paragraph")
        styleObject.setAttribute("style:parent-style-name", "Standard")
        # Build text-properties child only when needed
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
        # Cache and return
        self.styleObjects[paraName] = styleObject
        return paraName


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    content = Content(g)

    content.addHeading("Document Title", level=1)
    content.addParagraph("This is a normal paragraph.")
    content.addHeading("Section One", level=2)
    content.addParagraph("This is bold text.", bold=True)
    content.addParagraph("This is italic text.", italic=True)
    content.addParagraph("This is underlined text.", underline=True)
    content.addParagraph("This is large text.", fontsize="18pt")
    content.addParagraph("This is red text.", fontcolor="#cc0000")

    string1 = content.toString(indent=True)

    print("Writer Content XML with indentation:")
    print(string1)
