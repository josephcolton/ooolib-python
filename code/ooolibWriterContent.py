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
        # Open lists, outermost first, and whether each list level is numbered
        self.listStack = []
        self.listLevels = []
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
        fontFace1.setAttribute("svg:font-family", "'Liberation Serif'")
        fontFace1.setAttribute("style:font-family-generic", "roman")
        fontFace1.setAttribute("style:font-pitch", "variable")
        fontFace2 = officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace2.setAttribute("style:name", "Liberation Sans")
        fontFace2.setAttribute("svg:font-family", "'Liberation Sans'")
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
        return ooolibXML.escapeParagraphText(text)

    ####################
    # Content Creation #
    ####################
    def addParagraph(self, text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None,
                     align=None, indent=None, spaceBefore=None, spaceAfter=None, keepWithNext=False):
        """Add a paragraph.

        text is a string, or a list of runs for mixed formatting.  Each run is a string
        or a (string, formatting) tuple, where formatting is a dictionary that may hold
        bold, italic, underline, fontsize, and fontcolor.  The other formatting
        arguments apply to the whole paragraph.  align is "left", "center", "right",
        or "justify"; indent, spaceBefore, and spaceAfter are lengths such as "0.5in"."""
        self.endList()
        paraData = self.paragraphData(bold, italic, underline, fontsize, fontcolor,
                                      align, indent, spaceBefore, spaceAfter, keepWithNext)
        return self.createParagraph(self.officeText, text, paraData)

    def addHeading(self, text, level=1):
        """Add a heading at the given level (1-4).  text may also be a list of runs."""
        self.endList()
        self.paragraphCount += 1
        level = max(1, min(4, int(level)))
        styleName = "Heading_20_%d" % level
        heading = self.officeText.addChild(ooolibXML.Element("text:h"))
        heading.setAttribute("text:style-name", styleName)
        heading.setAttribute("text:outline-level", str(level))
        heading.setText(self.runsText(text))
        return heading

    def addListItem(self, text, level=1, numbered=False, bold=False, italic=False, underline=False,
                    fontsize=None, fontcolor=None, align=None, spaceBefore=None, spaceAfter=None,
                    keepWithNext=False):
        """Add an item to a bulleted (or numbered) list.

        Items added one after another form a single list.  level (1-10) nests the item
        below the previous item one level up.  Adding a paragraph or heading, or
        calling endList(), finishes the list so the next item starts a new one.
        text and the formatting arguments work the same as for addParagraph."""
        level = max(1, min(MAX_LIST_LEVEL, int(level)))
        numbered = bool(numbered)
        # Close lists deeper than this item
        del self.listStack[level:]
        # A different kind of list at this level starts a new list
        if len(self.listStack) == level and self.listStack[-1]["numbered"] != numbered:
            self.listStack.pop()
        # Open lists down to this item's level
        while len(self.listStack) < level:
            if self.listStack:
                parent = self.listStack[-1]
                if parent["item"] is None:
                    parent["item"] = parent["list"].addChild(ooolibXML.Element("text:list-item"))
                container = parent["item"]
            else:
                container = self.officeText
                self.listLevels = [False] * MAX_LIST_LEVEL
            listElement = container.addChild(ooolibXML.Element("text:list"))
            self.listStack.append({"list": listElement, "item": None, "numbered": numbered})
            self.listLevels[len(self.listStack) - 1] = numbered
            # Each list gets a style covering every level as it stands now, so a later
            # sub-list of a different kind does not change this one
            styleName = self.automaticStylesInstance.getListStyle(tuple(self.listLevels))
            listElement.setAttribute("text:style-name", styleName)
        # Add the item
        current = self.listStack[-1]
        current["item"] = current["list"].addChild(ooolibXML.Element("text:list-item"))
        paraData = self.paragraphData(bold, italic, underline, fontsize, fontcolor,
                                      align, None, spaceBefore, spaceAfter, keepWithNext)
        return self.createParagraph(current["item"], text, paraData)

    def endList(self):
        """Finish the current list, if any."""
        self.listStack = []

    ######################
    # Paragraph Building #
    ######################
    def paragraphData(self, bold, italic, underline, fontsize, fontcolor,
                      align, indent, spaceBefore, spaceAfter, keepWithNext):
        return {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "fontsize": fontsize,
            "fontcolor": fontcolor,
            "align": align,
            "indent": indent,
            "spaceBefore": spaceBefore,
            "spaceAfter": spaceAfter,
            "keepWithNext": keepWithNext,
        }

    def createParagraph(self, parent, text, paraData):
        self.paragraphCount += 1
        styleName = self.automaticStylesInstance.getParagraphStyle(paraData)
        para = parent.addChild(ooolibXML.Element("text:p"))
        para.setAttribute("text:style-name", styleName)
        para.setText(self.runsText(text))
        return para

    def runsText(self, text):
        """Return escaped paragraph XML for a string or a list of formatted runs."""
        if not isinstance(text, (list, tuple)):
            return self.escapeText(text)
        parts = []
        for run in text:
            if isinstance(run, (list, tuple)):
                runText, textData = run
            else:
                runText, textData = run, None
            styleName = self.automaticStylesInstance.getTextStyle(textData)
            if styleName is None:
                parts.append(self.escapeText(runText))
            else:
                parts.append('<text:span text:style-name="%s">%s</text:span>' % (styleName, self.escapeText(runText)))
        return "".join(parts)

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
TEXT_KEYS = ("bold", "italic", "underline", "fontsize", "fontcolor")
PARAGRAPH_KEYS = ("align", "indent", "spaceBefore", "spaceAfter", "keepWithNext")
ALIGNMENTS = {"left": "start", "start": "start", "center": "center",
              "right": "end", "end": "end", "justify": "justify"}
MAX_LIST_LEVEL = 10
BULLET_CHARS = ("•", "◦", "▪")
NUMBER_FORMATS = ("1", "a", "i")

class AutomaticStyles:
    def __init__(self, global_object, automaticStyles):
        self.global_object = global_object
        self.officeAutomaticStyles = automaticStyles
        self.styleObjects = {}
        self.global_object.setGlobalInt("automaticStyleWriterParagraph", 1)
        self.global_object.setGlobalInt("automaticStyleWriterText", 1)
        self.global_object.setGlobalInt("automaticStyleWriterList", 1)

    def extractData(self, data, name, default=None):
        if data == None: return default
        if name not in data: return default
        val = data[name]
        if val is None: return default
        return val

    def hasData(self, data, keys):
        for name in keys:
            if self.extractData(data, name) not in (None, False):
                return True
        return False

    def getParagraphStyle(self, paraData=None):
        # Paragraphs without formatting use the Standard style
        if not self.hasData(paraData, TEXT_KEYS + PARAGRAPH_KEYS):
            return "Standard"
        # Build deduplication key
        index = ("para",) + tuple(self.extractData(paraData, name) for name in TEXT_KEYS + PARAGRAPH_KEYS)
        if index not in self.styleObjects:
            self.styleObjects[index] = self.createParagraphStyle(paraData)
        return self.styleObjects[index]

    def getTextStyle(self, textData=None):
        """Return the automatic text style name for a run, or None for plain text."""
        if not self.hasData(textData, TEXT_KEYS):
            return None
        index = ("text",) + tuple(self.extractData(textData, name) for name in TEXT_KEYS)
        if index not in self.styleObjects:
            textNum = self.global_object.incrementGlobalInt("automaticStyleWriterText")
            textName = "T%d" % textNum
            styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
            styleObject.setAttribute("style:name", textName)
            styleObject.setAttribute("style:family", "text")
            self.addTextProperties(styleObject, textData)
            self.styleObjects[index] = textName
        return self.styleObjects[index]

    def getListStyle(self, numberedLevels):
        """Return the list style for a tuple of numbered (True) or bulleted (False) levels."""
        index = ("list", numberedLevels)
        if index not in self.styleObjects:
            listNum = self.global_object.incrementGlobalInt("automaticStyleWriterList")
            listName = "L%d" % listNum
            listStyle = self.officeAutomaticStyles.addChild(ooolibXML.Element("text:list-style"))
            listStyle.setAttribute("style:name", listName)
            for level, numbered in enumerate(numberedLevels, start=1):
                if numbered:
                    levelStyle = listStyle.addChild(ooolibXML.Element("text:list-level-style-number"))
                    levelStyle.setAttribute("text:level", str(level))
                    levelStyle.setAttribute("style:num-suffix", ".")
                    levelStyle.setAttribute("style:num-format", NUMBER_FORMATS[(level - 1) % len(NUMBER_FORMATS)])
                else:
                    levelStyle = listStyle.addChild(ooolibXML.Element("text:list-level-style-bullet"))
                    levelStyle.setAttribute("text:level", str(level))
                    levelStyle.setAttribute("text:bullet-char", BULLET_CHARS[(level - 1) % len(BULLET_CHARS)])
                # Label sits 0.25in left of the text, each level indents another 0.5in
                levelProp = levelStyle.addChild(ooolibXML.Element("style:list-level-properties"))
                levelProp.setAttribute("text:list-level-position-and-space-mode", "label-alignment")
                labelAlign = levelProp.addChild(ooolibXML.Element("style:list-level-label-alignment"))
                labelAlign.setAttribute("text:label-followed-by", "listtab")
                labelAlign.setAttribute("text:list-tab-stop-position", "%gin" % (0.5 * level))
                labelAlign.setAttribute("fo:text-indent", "-0.25in")
                labelAlign.setAttribute("fo:margin-left", "%gin" % (0.5 * level))
            self.styleObjects[index] = listName
        return self.styleObjects[index]

    def createParagraphStyle(self, paraData):
        paraNum = self.global_object.incrementGlobalInt("automaticStyleWriterParagraph")
        paraName = "P%d" % paraNum
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", paraName)
        styleObject.setAttribute("style:family", "paragraph")
        styleObject.setAttribute("style:parent-style-name", "Standard")
        # Paragraph properties come before text properties
        if self.hasData(paraData, PARAGRAPH_KEYS):
            paraProp = styleObject.addChild(ooolibXML.Element("style:paragraph-properties"))
            align = self.extractData(paraData, "align")
            if align is not None:
                paraProp.setAttribute("fo:text-align", ALIGNMENTS.get(str(align).lower(), "start"))
            indent = self.extractData(paraData, "indent")
            if indent is not None:
                paraProp.setAttribute("fo:margin-left", indent)
            spaceBefore = self.extractData(paraData, "spaceBefore")
            if spaceBefore is not None:
                paraProp.setAttribute("fo:margin-top", spaceBefore)
            spaceAfter = self.extractData(paraData, "spaceAfter")
            if spaceAfter is not None:
                paraProp.setAttribute("fo:margin-bottom", spaceAfter)
            if self.extractData(paraData, "keepWithNext"):
                paraProp.setAttribute("fo:keep-with-next", "always")
        self.addTextProperties(styleObject, paraData)
        return paraName

    def addTextProperties(self, styleObject, data):
        """Add a style:text-properties child when any text formatting is requested."""
        if not self.hasData(data, TEXT_KEYS): return
        textProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
        if self.extractData(data, "bold"):
            textProp.setAttribute("fo:font-weight", "bold")
            textProp.setAttribute("style:font-weight-asian", "bold")
            textProp.setAttribute("style:font-weight-complex", "bold")
        if self.extractData(data, "italic"):
            textProp.setAttribute("fo:font-style", "italic")
            textProp.setAttribute("style:font-style-asian", "italic")
            textProp.setAttribute("style:font-style-complex", "italic")
        if self.extractData(data, "underline"):
            textProp.setAttribute("style:text-underline-style", "solid")
            textProp.setAttribute("style:text-underline-width", "auto")
            textProp.setAttribute("style:text-underline-color", "font-color")
        fontsize = self.extractData(data, "fontsize")
        if fontsize is not None:
            textProp.setAttribute("fo:font-size", fontsize)
            textProp.setAttribute("style:font-size-asian", fontsize)
            textProp.setAttribute("style:font-size-complex", fontsize)
        fontcolor = self.extractData(data, "fontcolor")
        if fontcolor is not None:
            textProp.setAttribute("fo:color", fontcolor)


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
    content.addParagraph("Centered text.", align="center")
    content.addParagraph(["Mixed ", ("bold", {"bold": True}), " text."])
    content.addListItem("First bullet")
    content.addListItem("Nested bullet", level=2)
    content.addListItem("Numbered item", numbered=True)

    string1 = content.toString(indent=True)

    print("Writer Content XML with indentation:")
    print(string1)
