#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os
import datetime

# Import ooolib-python modules
import ooolibGlobal
import ooolibXML

##############################
# ooolib-python Calc Content #
##############################
class Content:
    def __init__(self, global_object, autoInit=True):
        # Passed in variables
        self.global_object = global_object
        # Internal objects
        self.internal_tables = []
        self.global_object.setGlobalInt("tableSheetNumber", 1) # First sheet number
        self.activeTable = None
        # Create automatic styles object
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles)
        self.global_object.setGlobalObjects("automaticStyles", self.automaticStylesInstance)
        # Create table if needed
        if autoInit: self.activeTable = self.addTable() # Allow table creation control
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
        fontFace1 = self.officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace1.setAttribute("style:name", "Liberation Sans")
        fontFace1.setAttribute("svg:font-family", "&apos;Liberation Sans&apos;")
        fontFace1.setAttribute("style:font-family-generic", "swiss")
        fontFace1.setAttribute("style:font-pitch", "variable")
        fontFace2 = self.officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace2.setAttribute("style:name", "Lucida Sans")
        fontFace2.setAttribute("svg:font-family", "&apos;Lucida Sans&apos;")
        fontFace2.setAttribute("style:font-family-generic", "system")
        fontFace2.setAttribute("style:font-pitch", "variable")
        fontFace3 = self.officeFontFace.addChild(ooolibXML.Element("style:font-face"))
        fontFace3.setAttribute("style:name", "Microsoft YaHei")
        fontFace3.setAttribute("svg:font-family", "&apos;Microsoft YaHei&apos;")
        fontFace3.setAttribute("style:font-family-generic", "system")
        fontFace3.setAttribute("style:font-pitch", "variable")
        ####################
        # Automatic Styles #
        ####################
        # if autoInit:
        self.global_object.setGlobalString("defaultColumnStyle", self.automaticStylesInstance.getColumnStyle())
        self.global_object.setGlobalString("defaultRowStyle", self.automaticStylesInstance.getRowStyle())
        self.global_object.setGlobalString("defaultTableStyle", self.automaticStylesInstance.getTableStyle())
        self.documentContent.addChild(self.automaticStyles)
        ###############
        # Office Body #
        ###############
        self.officeBody = self.documentContent.addChild(ooolibXML.Element("office:body"))
        self.officeSpreadsheet = self.officeBody.addChild(ooolibXML.Element("office:spreadsheet"))
        calculationSettings = self.officeSpreadsheet.addChild(ooolibXML.Element("table:calculation-settings"))
        calculationSettings.setAttribute("table:automatic-find-labels", "false")
        calculationSettings.setAttribute("table:use-regular-expressions", "false")
        calculationSettings.setAttribute("table:use-wildcards", "true")
        # Connect sheets here
        # Connect named expressions here - After tables or Excel crashes
        #self.officeSpreadsheet.addChild(ooolibXML.Element("table:named-expressions"))

    def addTable(self, name=None):
        table = ContentTable(self.global_object, name)
        self.internal_tables.append(table)
        self.activeTable = table
        return table

    def setActiveTable(self, name=None):
        # Set activeTable to last if none if passed in.  If none exist, create one.
        if (name == None):
            if self.internal_tables:
                self.activeTable = self.internal_tables[-1]
            else:
                self.addTable()
            return self.activeTable
        # Search for the matching table
        for table in self.internal_tables:
            if table.getTableName() == name:
                self.activeTable = table
                return self.activeTable
        # Hmm, didn't find the table
        print("setActiveTable('%s') - Unable to find table sheet" % (name))
        return self.activeTable

    def updateStats(self):
        # Table Count
        self.global_object.setGlobalInt("tableCount", len(self.internal_tables))
        # Count Cells
        cellCount = 0
        for table in self.internal_tables:
            cellCount += table.cellCount
        self.global_object.setGlobalInt("cellCount", cellCount)

    def buildTables(self):
        for table in self.internal_tables:
            tableObject = table.tableObject()
            self.officeSpreadsheet.addChild(tableObject)

    def toString(self, indent=False):
        # Update styles
        # Build tables
        self.buildTables()
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
        # Keep track of styles
        self.styleObjects = {}
        # Initialize Global Ints
        self.global_object.setGlobalInt("automaticStyleColumn", 1)
        self.global_object.setGlobalInt("automaticStyleRow", 1)
        self.global_object.setGlobalInt("automaticStyleTable", 1)
        self.global_object.setGlobalInt("automaticStyleCell", 1)

    def extractData(self, data, name, default=None):
        # Return None if no data or missing key
        if data == None: return default
        if not name in data: return default
        # Return data
        return data[name]

    ###########################
    # Automatic Column Styles #
    ###########################
    def getColumnStyle(self, columnData=None):
        # Build index based on contents of columnData
        index = ("col", self.extractData(columnData, "width"))
        # Create entry if missing
        if not index in self.styleObjects:
            self.styleObjects[index] = self.createColumnStyle(columnData)
        # Return object
        return self.styleObjects[index]

    def createColumnStyle(self, columnData):
        # Get attributes from data
        width = self.extractData(columnData, "width", "0.889in")
        # Get column number
        colNum = self.global_object.incrementGlobalInt("automaticStyleColumn")
        colName = "co%d" % colNum
        # Create object
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", colName)
        styleObject.setAttribute("style:family", "table-column")
        styleProperties = styleObject.addChild(ooolibXML.Element("style:table-column-properties"))
        styleProperties.setAttribute("fo:break-before", "auto")
        styleProperties.setAttribute("style:column-width", width)
        # Save style object
        self.styleObjects[colName] = styleObject
        # Return style name
        return colName

    ########################
    # Automatic Row Styles #
    ########################
    def getRowStyle(self, rowData=None):
        # Build index based on contents of rowData
        index = ("row", self.extractData(rowData, "height"))
        # Create entry if missing
        if not index in self.styleObjects:
            self.styleObjects[index] = self.createRowStyle(rowData)
        # Return object
        return self.styleObjects[index]

    def createRowStyle(self, rowData=None):
        # Get attributes from data
        height = self.extractData(rowData, "height", "0.178in")
        # Get row number
        rowNum = self.global_object.incrementGlobalInt("automaticStyleRow")
        rowName = "ro%d" % rowNum
        # Automatic Style Row
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", rowName)
        styleObject.setAttribute("style:family", "table-row")
        styleObjectProp = styleObject.addChild(ooolibXML.Element("style:table-row-properties"))
        styleObjectProp.setAttribute("style:row-height", height)
        styleObjectProp.setAttribute("fo:break-before", "auto")
        # styleObjectProp.setAttribute("style:use-optimal-row-height", "true")
        # Save style object
        self.styleObjects[rowName] = styleObject
        # Return style name
        return rowName

    ##########################
    # Automatic Table Styles #
    ##########################
    def getTableStyle(self, tableData=None):
        # Build index based on contents of tableData
        index = ("table")
        # Create entry if missing
        if not index in self.styleObjects:
            self.styleObjects[index] = self.createTableStyle(tableData)
        # Return object
        return self.styleObjects[index]

    def createTableStyle(self, tableData=None):
        # Get table number
        tableNum = self.global_object.incrementGlobalInt("automaticStyleTable")
        tableName = "ta%d" % tableNum
        # Automatic Style Table
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", tableName)
        styleObject.setAttribute("style:family", "table")
        styleObject.setAttribute("style:master-page-name", "Default")
        styleObjectProp = styleObject.addChild(ooolibXML.Element("style:table-properties"))
        styleObjectProp.setAttribute("table:display", "true")
        styleObjectProp.setAttribute("style:writing-mode", "lr-tb")
        # Save style object
        self.styleObjects[tableName] = styleObject
        # Return style name
        return tableName

    #########################
    # Automatic Cell Styles #
    #########################
    def getCellStyle(self, cellData=None):
        # Return manual styles
        manualStyle = self.extractData(cellData, "manualStyle", None)
        if manualStyle != None: return manualStyle
        # Build index based on contents of cellData
        index = ("cell",
                 self.extractData(cellData, "bold"),
                 self.extractData(cellData, "italics"),
                 self.extractData(cellData, "underline"),
                 self.extractData(cellData, "fontcolor"),
                 self.extractData(cellData, "bgcolor"),
                 self.extractData(cellData, "fontsize"),
                 self.extractData(cellData, "valign"),
                 self.extractData(cellData, "halign")
                 )
        # Create entry if missing
        if not index in self.styleObjects:
            self.styleObjects[index] = self.createCellStyle(cellData)
        # Return object
        return self.styleObjects[index]

    def createCellStyle(self, cellData=None):
        # Get table number
        cellNum = self.global_object.incrementGlobalInt("automaticStyleCell")
        cellName = "ce%d" % cellNum
        # Automatic Style Table
        styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
        styleObject.setAttribute("style:name", cellName)
        styleObject.setAttribute("style:family", "table-cell")
        styleObject.setAttribute("style:parent-style-name", "Default")
        # Initialize properties variables to None
        styleObjectProp = None     # style:text-properties
        styleObjectCellProp = None # style:table-cell-properties
        styleObjectParaProp = None # style:paragraph-properties
        # Look at contents of cellData
        if (self.extractData(cellData, "bold", False) == True):
            if (styleObjectProp == None): styleObjectProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            styleObjectProp.setAttribute("fo:font-weight", "bold")
            styleObjectProp.setAttribute("style:font-weight-asian", "bold")
            styleObjectProp.setAttribute("style:font-weight-complex", "bold")
        if (self.extractData(cellData, "italics", False) == True):
            if (styleObjectProp == None): styleObjectProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            styleObjectProp.setAttribute("fo:font-style", "italic")
            styleObjectProp.setAttribute("style:font-style-asian", "italic")
            styleObjectProp.setAttribute("style:font-style-complex", "italic")
        if (self.extractData(cellData, "underline", False) == True):
            if (styleObjectProp == None): styleObjectProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            styleObjectProp.setAttribute("style:text-underline-style", "solid")
            styleObjectProp.setAttribute("style:text-underline-width", "auto")
            styleObjectProp.setAttribute("style:text-underline-color", "font-color")
        if (self.extractData(cellData, "fontcolor", None) != None):
            if (styleObjectProp == None): styleObjectProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            styleObjectProp.setAttribute("fo:color", self.extractData(cellData, "fontcolor"))
        if (self.extractData(cellData, "bgcolor", None) != None):
            if (styleObjectCellProp == None): styleObjectCellProp = styleObject.addChild(ooolibXML.Element("style:table-cell-properties"))
            styleObjectCellProp.setAttribute("fo:background-color", self.extractData(cellData, "bgcolor"))
        if (self.extractData(cellData, "fontsize", None) != None):
            if (styleObjectProp == None): styleObjectProp = styleObject.addChild(ooolibXML.Element("style:text-properties"))
            styleObjectProp.setAttribute("fo:font-size", self.extractData(cellData, "fontsize"))
            styleObjectProp.setAttribute("style:font-size-asian", self.extractData(cellData, "fontsize"))
            styleObjectProp.setAttribute("style:font-size-complex", self.extractData(cellData, "fontsize"))
        if (self.extractData(cellData, "valign", None) != None):
            if (styleObjectCellProp == None): styleObjectCellProp = styleObject.addChild(ooolibXML.Element("style:table-cell-properties"))
            styleObjectCellProp.setAttribute("style:text-align-source", "fix")
            styleObjectCellProp.setAttribute("style:repeat-content", "false")
            styleObjectCellProp.setAttribute("style:vertical-align", self.extractData(cellData, "valign"))
        if (self.extractData(cellData, "halign", None) != None):
            if (styleObjectParaProp == None): styleObjectParaProp = styleObject.addChild(ooolibXML.Element("style:paragraph-properties"))
            halign = self.extractData(cellData, "halign")
            if (halign == "left"): halign = "start"
            if (halign == "right"): halign = "end"
            styleObjectParaProp.setAttribute("fo:text-align", halign)
            styleObjectParaProp.setAttribute("fo:margin-left", "0in")
        # Save style object
        self.styleObjects[cellName] = styleObject
        # Return style name
        return cellName


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
        self.cellCount = 0    # Initialized to 0
        self.columnCount = 1  # First Column = A = 1
        self.rowCount = 1     # First row is 1
        self.tableData = {}
        self.tableStyles = {}
        # Automatic Styles
        self.automaticStylesInstance = self.global_object.getGlobalObjects("automaticStyles")

    ######################
    # Conversion Methods #
    ######################
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

    def escapeText(self, text):
        text = str(text)
        text = text.replace("&", "&amp;")
        text = text.replace("'", "&apos;")
        text = text.replace('"', "&quot;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace("\t", "<text:tab-stop/>")
        text = text.replace("\n", "<text:line-break/>")
        return text

    ####################
    # Table Management #
    ####################
    def setTableName(self, name):
        self.name = name
        return self.name

    def getTableName(self):
        return self.name

    def getTableStyle(self):
        index = ("table")
        # Return default if none have been created
        tableStyleName = self.global_object.getGlobalString("defaultTableStyle")
        if not index in self.tableStyles: return tableStyleName
        # Return created style
        return tableStyleName

    def updateMaximums(self, row, col):
        # Update maximums
        if (row > self.rowCount): self.rowCount = row
        if (col > self.columnCount): self.columnCount = col
        # Update Row Max (highest column number for the row)
        self.getRowMax(row, col)

    ###########################
    # Table Column Management #
    ###########################
    def setColumnWidth(self, col, width):
        # Create style
        index = ("column", col)
        if not index in self.tableStyles:
            self.tableStyles[index] = {}
        # Set attribute
        self.tableStyles[index]["width"] = width

    def getColumnStyle(self, col):
        index = ("column", col)
        # Return default if none have been created
        colStyleName = self.global_object.getGlobalString("defaultColumnStyle")
        if not index in self.tableStyles: return colStyleName
        # Pass style dictionary to automatic styles
        colData = self.tableStyles[index]
        colStyleName = self.automaticStylesInstance.getColumnStyle(colData)
        # Return created style
        return colStyleName

    ########################
    # Table Row Management #
    ########################
    def setRowHeight(self, row, height):
        # Create style
        index = ("row", row)
        if not index in self.tableStyles:
            self.tableStyles[index] = {}
        # Set attribute
        self.tableStyles[index]["height"] = height

    def getRowStyle(self, row):
        index = ("row", row)
        # Return default if none have been created
        rowStyleName = self.global_object.getGlobalString("defaultRowStyle")
        if not index in self.tableStyles: return rowStyleName
        # Pass style dictionary to automatic styles
        rowData = self.tableStyles[index]
        rowStyleName = self.automaticStylesInstance.getRowStyle(rowData)
        # Return created style
        return rowStyleName

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

    #########################
    # Table Cell Management #
    #########################
    def __updateCellStyleValue(self, row, col, name, value=None):
        """__updateCellStyleValue - Add/Remove/Update cell style"""
        # Make sure cell style exists
        index = ("cell", row, col)
        if not index in self.tableStyles:
            self.tableStyles[index] = {}
        # Remove unneeded styles
        if (value == None or value == False):
            if name in self.tableStyles[index]:
                self.tableStyles[index].pop(name) # Remove style
            return
        # Must have a value
        self.tableStyles[index][name] = value

    def getCellStyle(self, row, col):
        index = ("cell", row, col)
        # Return default if none have been created
        cellStyleName = self.global_object.getGlobalString("defaultCellStyle")
        if not index in self.tableStyles: return cellStyleName
        # Pass style dictionary to automatic styles
        cellData = self.tableStyles[index]
        cellStyleName = self.automaticStylesInstance.getCellStyle(cellData)
        # Return created style
        return cellStyleName

    # Set Manual Styles
    def setCellStyle(self, row, col, value=None):
        self.__updateCellStyleValue(row, col, "manualStyle", value)

    # Standard cell properties
    def setCellBold(self, row, col, value=True):
        self.__updateCellStyleValue(row, col, "bold", value)

    def setCellItalics(self, row, col, value=True):
        self.__updateCellStyleValue(row, col, "italics", value)

    def setCellUnderline(self, row, col, value=True):
        self.__updateCellStyleValue(row, col, "underline", value)

    # Cell colors
    def setCellFontColor(self, row, col, value):
        self.__updateCellStyleValue(row, col, "fontcolor", value)

    def setCellBackgroundColor(self, row, col, value):
        self.__updateCellStyleValue(row, col, "bgcolor", value)

    # Font sizes
    def setCellFontSize(self, row, col, value=None):
        self.__updateCellStyleValue(row, col, "fontsize", value)

    # Text Alignment
    def setCellVerticalAlign(self, row, col, value=None):
        self.__updateCellStyleValue(row, col, "valign", value)

    def setCellHorizontalAlign(self, row, col, value=None):
        self.__updateCellStyleValue(row, col, "halign", value)

    # Formulas
    def setCellFormulaAverage(self, row, col, startCellId, endCellId):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        value = "of:=AVERAGE([.%s:.%s])" % (startCellId, endCellId)
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellFormulaMin(self, row, col, startCellId, endCellId):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        value = "of:=MIN([.%s:.%s])" % (startCellId, endCellId)
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellFormulaMax(self, row, col, startCellId, endCellId):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        value = "of:=MAX([.%s:.%s])" % (startCellId, endCellId)
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellFormulaSum(self, row, col, startCellId, endCellId):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        value = "of:=SUM([.%s:.%s])" % (startCellId, endCellId)
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellFormulaSqrt(self, row, col, targetCellId):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        value = "of:=SQRT([.%s])" % (targetCellId)
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellFormulaCustom(self, row, col, value):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        # Calculate formula
        cellContents = ("formula", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    # Cell Data
    def createCellCount(self, row, col):
        tableIndex = (row, col)
        if not tableIndex in self.tableData:
            self.cellCount += 1

    def setCellFloat(self, row, col, value):
        self.createCellCount(row, col) # Update cell count
        # Set table data
        tableIndex = (row, col)
        cellContents = ("float", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellText(self, row, col, text):
        self.createCellCount(row, col) # Update cell count
        text = self.escapeText(text)   # Escape text to make it clean
        tableIndex = (row, col)
        cellContents = ("string", text)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    def setCellDate(self, row, col, value):
        self.createCellCount(row, col) # Update cell count
        # Sanitize
        if not isinstance(value, datetime.datetime):
            self.setCellText(row, col, "Date Error: "+str(value))
            return
        # Set the value
        tableIndex = (row, col)
        cellContents = ("date", value)
        self.tableData[tableIndex] = cellContents
        self.updateMaximums(row, col)

    ###########################
    # Table Object Generation #
    ###########################
    def tableObject(self):
        # Generate table
        tableSheet = ooolibXML.Element("table:table")
        tableSheet.setAttribute("table:name", self.name)
        # Figure out table style, default or custom
        tableStyleName = self.getTableStyle()
        # Set table style
        tableSheet.setAttribute("table:style-name", tableStyleName)
        # Create columns
        for col in range(1, self.columnCount + 1):
            tableColumn = tableSheet.addChild(ooolibXML.Element("table:table-column"))
            # Figure out column style, default or custom
            colStyleName = self.getColumnStyle(col)
            # Set column style
            tableColumn.setAttribute("table:style-name", colStyleName)
            tableColumn.setAttribute("table:default-cell-style-name", "Default")
        # Create rows
        for row in range(1, self.rowCount + 1):
            tableRow = tableSheet.addChild(ooolibXML.Element("table:table-row"))
            # Figure out row style, default or custom
            rowStyleName = self.getRowStyle(row)
            # Set row style
            tableRow.setAttribute("table:style-name", rowStyleName)
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
                    cellStyleName = self.getCellStyle(row, col)
                    if (cellType == "float"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellFloat(cellValue, cellStyleName)
                    if (cellType == "string"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellString(cellValue, cellStyleName)
                    if (cellType == "date"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellDate(cellValue, cellStyleName)
                    if (cellType == "formula"):
                        cellValue = cellContents[1]
                        tableRow.addTableCellFormula(cellValue, cellStyleName)
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
