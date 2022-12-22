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
        # Passed in variables
        self.global_object = global_object
        # Internal objects
        self.internal_tables = []
        self.global_object.setGlobalInt("tableSheetNumber", 1) # First sheet number
        self.activeTable = self.addTable()
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

        # Connect sheets here
        """
        tableSheet1 = self.officeSpreadsheet.addChild(ooolibXML.Element("table:table"))
        tableSheet1.setAttribute("table:name", "Sheet1")
        tableSheet1.setAttribute("table:style-name", "ta1")
        tableColumn = tableSheet1.addChild(ooolibXML.Element("table:table-column"))
        tableColumn.setAttribute("table:style-name", "co1")
        tableColumn.setAttribute("table:default-cell-style-name", "Default")
        tableRow = tableSheet1.addChild(ooolibXML.Element("table:table-row"))
        tableRow.setAttribute("table:style-name", "ro1")
        tableCell = tableRow.addChild(ooolibXML.Element("table:table-cell"))
        """
        # Named Expressions
        tableNamedExpressions = self.officeSpreadsheet.addChild(ooolibXML.Element("table:named-expressions"))
        
    def addTable(self, name=None):
        table = ContentTable(self.global_object, name)
        self.internal_tables.append(table)
        return table

    def buildTables(self):
        for table in self.internal_tables:
            tableObject = table.tableObject()
            self.officeSpreadsheet.addChild(tableObject)
        
    def toString(self, indent=False):
        self.buildTables()
        contentString = self.prolog.toString()
        contentString += self.documentContent.toString(indent)
        return contentString

#######################
# Content Table Class #
#######################
class ContentTable:
    def __init__(self, global_object, name=None):
        # Passed in variables
        self.global_object = global_object
        self.name = name
        # Update name if missing
        if (self.name == None):
            self.name = "Sheet%d" % self.global_object.incrementGlobalInt("tableSheetNumber")
        # Build internal variables
        self.columnCount = 1  # First Column = A = 1
        self.rowCount = 1     # First row is 1
        self.tableData = {}
        self.tableStyles = {}
        
    def convertCellRowCol2Id(self, row, col):
        chars = []
        # Convert column into letters
        while(col > 0):
            chValue = col % 26
            col -= chValue
            col = int(col / 26)
            if chValue == 0: chValue = 26
            ch = chr(int(chValue + 64))
            chars.insert(0, ch)
        # Convert to Id
        cellId = "%s%d" % ("".join(chars), row)
        return cellId

    def convertCellId2RowCol(self, cellId):
        row = 0
        col = 0
        chars = list(cellId)
        while(chars):
            ch = chars.pop(0)
            if ch.isalpha():
                ch = ch.upper()
                chValue = ord(ch) - 64
                col *= 26
                col += chValue
            if ch.isdigit():
                chValue = ord(ch) - 48
                row *= 10
                row += chValue
        # Should have row and col calculated
        return (row, col)

    def setCellFloat(self, row, col, value):
        # Set table data
        tableIndex = (row, col)
        cellContents = ("float", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellText(self, row, col, text):
        tableIndex = (row, col)
        cellContents = ("string", text)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)        

    def updateMaximums(self, row, col):
        # Update maximums
        if (row > self.rowCount): self.rowCount = row
        if (col > self.columnCount): self.columnCount = col
        # Update Row Max (highest column number for the row)
        self.getRowMax(row, col)

    def getRowMax(self, row, defaultCol=1):
        rowMaxIndex = ("rowMax", row)
        rowMax = defaultCol
        if rowMaxIndex in self.tableData:
            rowMax = self.tableData[rowMaxIndex]
            if defaultCol > rowMax:
                self.tableData[rowMaxIndex] = defaultCol
                rowMax = defaultCol
        else:
            self.tableData[rowMaxIndex] = defaultCol
        # Return row max
        return rowMax
            
    def tableObject(self):
        # Generate table
        tableSheet = ooolibXML.Element("table:table")
        tableSheet.setAttribute("table:name", self.name)
        tableSheet.setAttribute("table:style-name", "ta1")
        # Create columns
        for col in range(1, self.columnCount + 1):
            tableColumn = tableSheet.addChild(ooolibXML.Element("table:table-column"))
            tableColumn.setAttribute("table:style-name", "co1")
            tableColumn.setAttribute("table:default-cell-style-name", "Default")
        # Create rows
        for row in range(1, self.rowCount + 1):
            tableRow = tableSheet.addChild(ooolibXML.Element("table:table-row"))
            tableRow.setAttribute("table:style-name", "ro1")
            # Create row cells
            rowMax = self.getRowMax(row)
            for col in range(1, rowMax+1):
                tableIndex = (row, col)
                if tableIndex in self.tableData:
                    cellContents = self.tableData[tableIndex]
                    #######################
                    # Table Cell Creation #
                    #######################
                    cellType = cellContents[0]
                    if (cellType == "float"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellFloat(cellValue)
                    if (cellType == "string"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellString(cellValue)
                else:
                    tableRow.addChild(ooolibXML.Element("table:table-cell"))
        # Return completed table sheet
        return tableSheet

###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    content = Content(g)

    # Create content
    for row in range(1, 5):
        for col in range(1, 5):
            value = row * col
            content.activeTable.setCellFloat(row, col, value)
    
    # Display content file
    string1 = content.toString(indent=True)

    print("Content XML with indentation:")
    print(string1)
