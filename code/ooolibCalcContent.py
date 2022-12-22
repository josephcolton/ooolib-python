#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os

# Import ooolib-python modules
import ooolibGlobal
import ooolibXML

##############################
# ooolib-python Calc Content #
##############################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Content
        self.documentContent = ooolibXML.Element("office:document-content")
        self.documentContent.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.documentContent.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentContent.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.documentContent.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentContent.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentContent.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.documentContent.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.documentContent.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.documentContent.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.documentContent.setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")
        self.documentContent.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.documentContent.setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
        self.documentContent.setAttribute("xmlns:rpt", "http://openoffice.org/2005/report")
        self.documentContent.setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
        self.documentContent.setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
        self.documentContent.setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer")
        self.documentContent.setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc")
        self.documentContent.setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
        self.documentContent.setAttribute("xmlns:tableooo", "http://openoffice.org/2009/table")
        self.documentContent.setAttribute("xmlns:calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0")
        self.documentContent.setAttribute("xmlns:drawooo", "http://openoffice.org/2010/draw")
        self.documentContent.setAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0")
        self.documentContent.setAttribute("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")
        self.documentContent.setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML")
        self.documentContent.setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")
        self.documentContent.setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")
        self.documentContent.setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events")
        self.documentContent.setAttribute("xmlns:xforms", "http://www.w3.org/2002/xforms")
        self.documentContent.setAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        self.documentContent.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        self.documentContent.setAttribute("xmlns:formx", "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0")
        self.documentContent.setAttribute("xmlns:xhtml", "http://www.w3.org/1999/xhtml")
        self.documentContent.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.documentContent.setAttribute("xmlns:css3t", "http://www.w3.org/TR/css3-text/")
        self.documentContent.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
        self.documentContent.setAttribute("office:version", "1.3")
        # Office Scripts
        self.officeScripts = self.documentContent.addChild(ooolibXML.Element("office:scripts"))
        # Office Font Face Decls
        self.officeFontFace = self.documentContent.addChild(ooolibXML.Element("office:font-face-decls"))
        fontFace = self.officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace.setAttribute("style:name", "Liberation Sans")
        fontFace.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
        fontFace.setAttribute("style:font-family-generic", "swiss")
        fontFace.setAttribute("style:font-pitch", "variable")
        # Automatic Styles
        self.officeAutomaticStyles = self.documentContent.addChild(ooolibXML.Element("office:automatic-styles"))
        # Automatic Style Column
        styleCo1 = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleCo1.setAttribute("style:name", "co1")
        styleCo1.setAttribute("style:family", "table-column")
        styleColProp = styleCo1.addChild(ooolibXML.Element("style:table-column-properties"))
        styleColProp.setAttribute("fo:break-before", "auto")
        styleColProp.setAttribute("style:column-width", "0.889in")
        # Automatic Style Row
        styleRo1 = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleRo1.setAttribute("style:name", "ro1")
        styleRo1.setAttribute("style:family", "table-row")
        styleRo1Prop = styleRo1.addChild(ooolibXML.Element("style:table-row-properties"))
        styleRo1Prop.setAttribute("style:row-height", "0.178in")
        styleRo1Prop.setAttribute("fo:break-before", "auto")
        styleRo1Prop.setAttribute("style:use-optimal-row-height", "true")
        # Automatic Style Table
        styleTa1 = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleTa1.setAttribute("style:name", "ta1")
        styleTa1.setAttribute("style:family", "table")
        styleTa1.setAttribute("style:master-page-name", "Default")
        styleTa1Prop = styleTa1.addChild(ooolibXML.Element("style:table-properties"))
        styleTa1Prop.setAttribute("table:display", "true")
        styleTa1Prop.setAttribute("style:writing-mode", "lr-tb")
        # Office Body
        self.officeBody = self.documentContent.addChild(ooolibXML.Element("office:body"))
        self.officeSpreadsheet = self.officeBody.addChild(ooolibXML.Element("office:spreadsheet"))
        calculationSettings = self.officeSpreadsheet.addChild(ooolibXML.Element("table:calculation-settings"))
        calculationSettings.setAttribute("table:automatic-find-labels", "false")
        calculationSettings.setAttribute("table:use-regular-expressions", "false")
        calculationSettings.setAttribute("table:use-wildcards", "true")
        tableSheet1 = self.officeSpreadsheet.addChild(ooolibXML.Element("table:table"))
        tableSheet1.setAttribute("table:name", "Sheet1")
        tableSheet1.setAttribute("table:style-name", "ta1")
        tableColumn = tableSheet1.addChild(ooolibXML.Element("table:table-column"))
        tableColumn.setAttribute("table:style-name", "co1")
        tableColumn.setAttribute("table:default-cell-style-name", "Default")
        tableRow = tableSheet1.addChild(ooolibXML.Element("table:table-row"))
        tableRow.setAttribute("table:style-name", "ro1")
        tableCell = tableRow.addChild(ooolibXML.Element("table:table-cell"))
        tableNamedExpressions = self.officeSpreadsheet.addChild(ooolibXML.Element("table:named-expressions"))
        
    def toString(self, indent=False):
        contentString = self.prolog.toString()
        contentString += self.documentContent.toString(indent)
        return contentString

###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    content = Content(g)

    # Display content file
    string1 = content.toString(indent=True)

    print("Content XML with indentation:")
    print(string1)
