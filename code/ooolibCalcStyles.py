#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import datetime

# Import ooolib-python modules
import ooolibGlobal
import ooolibXML

#################################
# ooolib-python Settings object #
#################################
class Styles:
    def __init__(self, global_object):
        self.global_object = global_object
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Styles
        self.documentStyles = ooolibXML.Element("office:document-styles")
        self.documentStyles.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.documentStyles.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentStyles.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentStyles.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentStyles.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.documentStyles.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentStyles.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentStyles.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.documentStyles.setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")
        self.documentStyles.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentStyles.setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
        self.documentStyles.setAttribute("xmlns:rpt", "http://openoffice.org/2005/report")
        self.documentStyles.setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
        self.documentStyles.setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
        self.documentStyles.setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer")
        self.documentStyles.setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc")
        self.documentStyles.setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
        self.documentStyles.setAttribute("xmlns:tableooo", "http://openoffice.org/2009/table")
        self.documentStyles.setAttribute("xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")
        self.documentStyles.setAttribute("xmlns:drawooo", "http://openoffice.org/2010/draw")
        self.documentStyles.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentStyles.setAttribute("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")
        self.documentStyles.setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML")
        self.documentStyles.setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")
        self.documentStyles.setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")
        self.documentStyles.setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events")
        self.documentStyles.setAttribute("xmlns:xhtml", "http://www.w3.org/1999/xhtml")
        self.documentStyles.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.documentStyles.setAttribute("xmlns:css3t", "http://www.w3.org/TR/css3-text/")
        self.documentStyles.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
        self.documentStyles.setAttribute("office:version", "1.3")
        # Office FontFace Decls
        officeFonts = self.documentStyles.addChild(ooolibXML.Element("office:font-face-decls"))
        font1 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font1.setAttribute("style:name", "Liberation Sans")
        font1.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
        font1.setAttribute("style:font-family-generic", "swiss")
        font1.setAttribute("style:font-pitch", "variable")
        font2 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font2.setAttribute("style:name", "Lucida Sans")
        font2.setAttribute("svg:font-family", "&apos;Lucida Sans&apos;")
        font2.setAttribute("style:font-family-generic", "system")
        font2.setAttribute("style:font-pitch", "variable")
        font3 = officeFonts.addChild(ooolibXML.Element("style:font-face"))
        font3.setAttribute("style:name", "Microsoft YaHei")
        font3.setAttribute("svg:font-family", "&apos;Microsoft YaHei&apos;")
        font3.setAttribute("style:font-family-generic", "system")
        font3.setAttribute("style:font-pitch", "variable")

        #################
        # Office Styles #
        #################
        officeStyles = self.documentStyles.addChild(ooolibXML.Element("office:styles"))
        # Table Styles?
        defaultStyle = officeStyles.addChild(ooolibXML.Element("style:default-style"))
        defaultStyle.setAttribute("style:family", "table-cell")
        defaultParagraphProperties = defaultStyle.addChild(ooolibXML.Element("style:paragraph-properties"))
        defaultParagraphProperties.setAttribute("style:tab-stop-distance", "0.5in")
        textProperties = defaultStyle.addChild(ooolibXML.Element("style:text-properties")) # Skipped asian fonts
        textProperties.setAttribute("style:font-name", "Liberation Sans")
        textProperties.setAttribute("fo:font-size", "10pt")
        textProperties.setAttribute("fo:language", "en")
        textProperties.setAttribute("fo:country", "US")
        # Number Style
        numberStyle = officeStyles.addChild(ooolibXML.Element("number:number-style"))
        numberStyle.setAttribute("style:name", "N0")
        numberStyle.addChild(ooolibXML.Element("number:number")).setAttribute("number:min-integer-digits", "1")
        # Style: Default
        officeStyles.addStyle("Default", "table-cell")
        # Style: Heading - Default
        styleHeading = officeStyles.addStyle("Heading", "table-cell", "Default")
        styleHeadingProperties = styleHeading.addChild(ooolibXML.Element("style:text-properties"))
        styleHeadingProperties.setAttribute("fo:color", "#000000")
        styleHeadingProperties.setAttribute("fo:font-size", "24pt")
        styleHeadingProperties.setAttribute("fo:font-style", "normal")
        styleHeadingProperties.setAttribute("fo:font-weight", "bold")
        styleHeadingProperties.setAttribute("style:font-size-asian", "24pt")
        styleHeadingProperties.setAttribute("style:font-style-asian", "normal")
        styleHeadingProperties.setAttribute("style:font-weight-asian", "bold")
        styleHeadingProperties.setAttribute("style:font-size-complex", "24pt")
        styleHeadingProperties.setAttribute("style:font-style-complex", "normal")
        styleHeadingProperties.setAttribute("style:font-weight-complex", "bold")
        # Style: Heading 1 - Heading - Default
        styleHeading1 = officeStyles.addStyle("Heading_20_1", "table-cell", "Heading")
        styleHeading1.setAttribute("style:display-name", "Heading 1")
        styleHeading1Properties = styleHeading1.addChild(ooolibXML.Element("style:text-properties"))
        styleHeading1Properties.setAttribute("fo:font-size", "18pt")
        styleHeading1Properties.setAttribute("style:font-size-asian", "18pt")
        styleHeading1Properties.setAttribute("style:font-size-complex", "18pt")
        # Style: Heading 2 - Heading - Default
        styleHeading2 = officeStyles.addStyle("Heading_20_2", "table-cell", "Heading")
        styleHeading2.setAttribute("style:display-name", "Heading 2")
        styleHeading2Properties = styleHeading1.addChild(ooolibXML.Element("style:text-properties"))
        styleHeading2Properties.setAttribute("fo:font-size", "12pt")
        styleHeading2Properties.setAttribute("style:font-size-asian", "12pt")
        styleHeading2Properties.setAttribute("style:font-size-complex", "12pt")
        # Style: Text - Default
        styleText = officeStyles.addStyle("Text", "table-cell", "Default")
        # Style: Note - Text - Default
        styleNote = officeStyles.addStyle("Note", "table-cell", "Text")
        styleNoteProperties = styleNote.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleNoteProperties.setAttribute("fo:background-color", "#ffffcc")
        styleNoteProperties.setAttribute("style:diagonal-bl-tr", "none")
        styleNoteProperties.setAttribute("style:diagonal-tl-br", "none")
        styleNoteProperties.setAttribute("fo:border", "0.74pt solid #808080")
        styleNoteTextProperties = styleNote.addChild(ooolibXML.Element("style:text-properties"))
        styleNoteTextProperties.setAttribute("fo:color", "#333333")
        # Style: Footnote - Text - Default
        styleFootnote = officeStyles.addStyle("Footnote", "table-cell", "Text")
        styleFootnoteProperties = styleFootnote.addChild(ooolibXML.Element("style:text-properties"))
        styleFootnoteProperties.setAttribute("fo:color", "#808080")
        styleFootnoteProperties.setAttribute("fo:font-style", "italic")
        styleFootnoteProperties.setAttribute("style:font-style-asian", "italic")
        styleFootnoteProperties.setAttribute("style:font-style-complex", "italic")
        # Style: Hyperlink - Text - Default
        styleHyperlink = officeStyles.addStyle("Hyperlink", "table-cell", "Text")
        styleHyperlinkProperties = styleHyperlink.addChild(ooolibXML.Element("style:text-properties"))
        styleHyperlinkProperties.setAttribute("fo:color", "#0000ee")
        styleHyperlinkProperties.setAttribute("style:text-underline-style", "solid")
        styleHyperlinkProperties.setAttribute("style:text-underline-width", "auto")
        styleHyperlinkProperties.setAttribute("style:text-underline-color", "#0000ee")
        # Style: Status - Default
        styleStatus = officeStyles.addStyle("Status", "table-cell", "Default")
        # Style: Good - Status - Default
        styleGood = officeStyles.addStyle("Good", "table-cell", "Status")
        styleGoodProperties = styleGood.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleGoodProperties.setAttribute("fo:background-color", "#ccffcc")
        styleGoodTextProperties = styleGood.addChild(ooolibXML.Element("style:text-properties"))
        styleGoodTextProperties.setAttribute("fo:color", "#006600")
        # Style: Neutral - Status - Default
        styleNeutral = officeStyles.addStyle("Neutral", "table-cell", "Status")
        styleNeutralProperties = styleNeutral.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleNeutralProperties.setAttribute("fo:background-color", "#ffffcc")
        styleNeutralTextProperties = styleNeutral.addChild(ooolibXML.Element("style:text-properties"))
        styleNeutralTextProperties.setAttribute("fo:color", "#996600")
        # Style: Bad - Status - Default
        styleBad = officeStyles.addStyle("Bad", "table-cell", "Status")
        styleBadProperties = styleBad.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleBadProperties.setAttribute("fo:background-color", "#ffcccc")
        styleBadTextProperties = styleBad.addChild(ooolibXML.Element("style:text-properties"))
        styleBadTextProperties.setAttribute("fo:color", "#cc0000")
        # Style: Warning - Status - Default
        styleWarning = officeStyles.addStyle("Warning", "table-cell", "Status")
        styleWarningTextProperties = styleWarning.addChild(ooolibXML.Element("style:text-properties"))
        styleWarningTextProperties.setAttribute("fo:color", "#cc0000")
        # Style: Error - Status - Default
        styleError = officeStyles.addStyle("Error", "table-cell", "Status")
        styleErrorProperties = styleError.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleErrorProperties.setAttribute("fo:background-color", "#cc0000")
        styleErrorTextProperties = styleError.addChild(ooolibXML.Element("style:text-properties"))
        styleErrorTextProperties.setAttribute("fo:color", "#ffffff")
        styleErrorTextProperties.setAttribute("fo:font-weight", "bold")
        styleErrorTextProperties.setAttribute("style:font-weight-asian", "bold")
        styleErrorTextProperties.setAttribute("style:font-weight-complex", "bold")
        # Style: Accent - Default
        styleAccent = officeStyles.addStyle("Accent", "table-cell", "Default")
        styleAccentTextProperties = styleAccent.addChild(ooolibXML.Element("style:text-properties"))
        styleAccentTextProperties.setAttribute("fo:font-weight", "bold")
        styleAccentTextProperties.setAttribute("style:font-weight-asian", "bold")
        styleAccentTextProperties.setAttribute("style:font-weight-complex", "bold")
        # Style: Accent 1 - Accent - Default
        styleAccent1 = officeStyles.addStyle("Accent_20_1", "table-cell", "Accent").setAttribute("style:display-name", "Accent 1")
        styleAccent1Properties = styleAccent1.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleAccent1Properties.setAttribute("fo:background-color", "#000000")
        styleAccent1TextProperties = styleAccent1.addChild(ooolibXML.Element("style:text-properties"))
        styleAccent1TextProperties.setAttribute("fo:color", "#ffffff")
        # Style: Accent 2 - Accent - Default
        styleAccent2 = officeStyles.addStyle("Accent_20_2", "table-cell", "Accent").setAttribute("style:display-name", "Accent 2")
        styleAccent2Properties = styleAccent2.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleAccent2Properties.setAttribute("fo:background-color", "#808080")
        styleAccent2TextProperties = styleAccent2.addChild(ooolibXML.Element("style:text-properties"))
        styleAccent2TextProperties.setAttribute("fo:color", "#ffffff")
        # Style: Accent 3 - Accent - Default
        styleAccent3 = officeStyles.addStyle("Accent_20_3", "table-cell", "Accent").setAttribute("style:display-name", "Accent 3")
        styleAccent3Properties = styleAccent3.addChild(ooolibXML.Element("style:table-cell-properties"))
        styleAccent3Properties.setAttribute("fo:background-color", "#dddddd")
        # Style: Result
        styleResult = officeStyles.addStyle("Result", "table-cell", "Default")
        styleResultTextProperties = styleResult.addChild(ooolibXML.Element("style:text-properties"))
        styleResultTextProperties.setAttribute("fo:font-style", "italic")
        styleResultTextProperties.setAttribute("style:text-underline-style", "solid")
        styleResultTextProperties.setAttribute("style:text-underline-width", "auto")
        styleResultTextProperties.setAttribute("style:text-underline-color", "font-color")
        styleResultTextProperties.setAttribute("fo:font-weight", "bold")
        styleResultTextProperties.setAttribute("style:font-style-asian", "italic")
        styleResultTextProperties.setAttribute("style:font-weight-asian", "bold")
        styleResultTextProperties.setAttribute("style:font-style-complex", "italic")
        styleResultTextProperties.setAttribute("style:font-weight-complex", "bold")

        ###########################
        # Office Automatic Styles #
        ###########################
        officeAutomaticStyles = self.documentStyles.addChild(ooolibXML.Element("office:automatic-styles"))
        # Page Layout: Mpm1
        stylePageLayout1 = officeAutomaticStyles.addChild(ooolibXML.Element("style:page-layout")).setAttribute("style:name", "Mpm1")
        stylePageLayout1.addChild(ooolibXML.Element("style:page-layout-properties")).setAttribute("style:writing-mode", "lr-tb")
        stylePageLayout1Header = stylePageLayout1.addChild(ooolibXML.Element("style:header-style"))
        stylePageLayout1HeaderProperties = stylePageLayout1Header.addChild(ooolibXML.Element("style:header-footer-properties"))
        stylePageLayout1HeaderProperties.setAttribute("fo:min-height", "0.2953in")
        stylePageLayout1HeaderProperties.setAttribute("fo:margin-left", "0in")
        stylePageLayout1HeaderProperties.setAttribute("fo:margin-right", "0in")
        stylePageLayout1HeaderProperties.setAttribute("fo:margin-bottom", "0.0984in")
        stylePageLayout1Footer = stylePageLayout1.addChild(ooolibXML.Element("style:footer-style"))
        stylePageLayout1FooterProperties = stylePageLayout1Footer.addChild(ooolibXML.Element("style:header-footer-properties"))
        stylePageLayout1FooterProperties.setAttribute("fo:min-height", "0.2953in")
        stylePageLayout1FooterProperties.setAttribute("fo:margin-left", "0in")
        stylePageLayout1FooterProperties.setAttribute("fo:margin-right", "0in")
        stylePageLayout1FooterProperties.setAttribute("fo:margin-bottom", "0.0984in")

        ########################
        # Office Master Styles #
        ########################
        officeMasterStyles = self.documentStyles.addChild(ooolibXML.Element("office:master-styles"))
        # Master Page 1: Default (Mpm1)
        styleMasterPage1 = officeMasterStyles.addChild(ooolibXML.Element("style:master-page"))
        styleMasterPage1.setAttribute("style:name", "Default")
        styleMasterPage1.setAttribute("style:page-layout-name", "Mpm1")
        styleMasterPage1Header = styleMasterPage1.addChild(ooolibXML.Element("style:header"))
        styleMasterPage1Header.addChild(ooolibXML.Element("text:p")).addChild(ooolibXML.Element("text:sheet-name", "???"))
        styleMasterPage1.addChild(ooolibXML.Element("style:header-left")).setAttribute("style:display", "false")
        styleMasterPage1.addChild(ooolibXML.Element("style:header-first")).setAttribute("style:display", "false")
        styleMasterPage1Footer = styleMasterPage1.addChild(ooolibXML.Element("style:footer"))
        styleMasterPage1Footer.addChild(ooolibXML.Element("text:p", "Page ")).addChild(ooolibXML.Element("text:page-number", "1"))
        styleMasterPage1.addChild(ooolibXML.Element("style:footer-left")).setAttribute("style:display", "false")
        styleMasterPage1.addChild(ooolibXML.Element("style:footer-first")).setAttribute("style:display", "false")

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

    # Display meta file
    string1 = styles.toString(indent=True)

    print("Styles XML with indentation:")
    print(string1)
