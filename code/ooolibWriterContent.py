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

#############
# Constants #
#############
TEXT_KEYS = ("bold", "italic", "underline", "fontsize", "fontcolor", "fontname")
PARAGRAPH_KEYS = ("align", "indent", "spaceBefore", "spaceAfter", "keepWithNext")
ALIGNMENTS = {"left": "start", "start": "start", "center": "center",
              "right": "end", "end": "end", "justify": "justify"}
MAX_LIST_LEVEL = 10
CELL_KEYS = TEXT_KEYS + ("align", "background", "valign")
VERTICAL_ALIGNMENTS = {"top": "top", "middle": "middle", "center": "middle", "bottom": "bottom"}
TABLE_ALIGNMENTS = ("left", "center", "right", "margins")
DEFAULT_BORDER = "0.5pt solid #000000"
DEFAULT_PADDING = "0.04in"
# Space between the default page margins (8.5in page, 1.25in margins)
DEFAULT_TABLE_WIDTH = 6.0
BULLET_CHARS = ("•", "◦", "▪")
NUMBER_FORMATS = ("1", "a", "i")

################################
# ooolib-python Writer Content #
################################
class Content:
    def __init__(self, global_object):
        self.global_object = global_object
        self.paragraphCount = 0
        self.tableCount = 0
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
        self.documentContent.setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
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
        self.fontFaces = FontFaces(officeFontFace)
        self.fontFaces.addFont("Liberation Serif")
        self.fontFaces.addFont("Liberation Sans")
        # Automatic Styles
        self.automaticStyles = ooolibXML.Element("office:automatic-styles")
        self.automaticStylesInstance = AutomaticStyles(self.global_object, self.automaticStyles, self.fontFaces)
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
                     align=None, indent=None, spaceBefore=None, spaceAfter=None, keepWithNext=False,
                     fontname=None):
        """Add a paragraph.

        text is a string, or a list of runs for mixed formatting.  Each run is a string
        or a (string, formatting) tuple, where formatting is a dictionary that may hold
        bold, italic, underline, fontsize, fontcolor, and fontname.  The other formatting
        arguments apply to the whole paragraph.  align is "left", "center", "right",
        or "justify"; indent, spaceBefore, and spaceAfter are lengths such as "0.5in".
        fontname is a font family name such as "Liberation Sans" or "DejaVu Serif"."""
        self.endList()
        paraData = self.paragraphData(bold=bold, italic=italic, underline=underline, fontsize=fontsize,
                                      fontcolor=fontcolor, fontname=fontname, align=align, indent=indent,
                                      spaceBefore=spaceBefore, spaceAfter=spaceAfter, keepWithNext=keepWithNext)
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
                    keepWithNext=False, fontname=None):
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
        paraData = self.paragraphData(bold=bold, italic=italic, underline=underline, fontsize=fontsize,
                                      fontcolor=fontcolor, fontname=fontname, align=align,
                                      spaceBefore=spaceBefore, spaceAfter=spaceAfter, keepWithNext=keepWithNext)
        return self.createParagraph(current["item"], text, paraData)

    def endList(self):
        """Finish the current list, if any."""
        self.listStack = []

    def addTable(self, data=None, columns=None, columnWidths=None, width=None, align=None,
                 headerRows=0, border=DEFAULT_BORDER, padding=DEFAULT_PADDING, **formatting):
        """Add a table and return it, so more rows can be added with table.addRow().

        data is an optional list of rows to add right away; each row is a list of cells
        (see Table.addRow).  The number of columns comes from columns, columnWidths, or
        the widest row in data.  columnWidths is a list of lengths such as "1.5in"; the
        table is then as wide as their total.  Otherwise width (default: the space
        between the page margins) is shared equally among the columns.  align is
        "left", "center", "right", or "margins" (stretch between the page margins).
        The first headerRows rows of data are header rows, repeated at the top of each
        page.  border is a border such as "0.5pt solid #000000" or None for no borders;
        padding is the space inside each cell.  Other keyword arguments (bold,
        fontname, background, ...) are formatting used by every cell in data."""
        self.endList()
        if columnWidths is not None:
            columnWidths = list(columnWidths)
            if columns is None:
                columns = len(columnWidths)
            elif columns != len(columnWidths):
                raise ValueError("columns is %d but %d column widths were given" % (columns, len(columnWidths)))
        if columns is None and data:
            columns = max(sum(Table.cellSpan(cell) for cell in row) for row in data)
        if not columns or columns < 1:
            raise ValueError("addTable needs columns, columnWidths, or data")
        columns = int(columns)
        # Work out the column widths in inches
        if columnWidths is not None:
            widths = []
            for columnWidth in columnWidths:
                inches = ooolibUnits.convertLength(ooolibUnits.formatLength(columnWidth))
                if inches is None or inches <= 0:
                    raise ValueError("Invalid column width: %r" % (columnWidth,))
                widths.append(inches)
            if align is None: align = "left"
        else:
            if width is None:
                inches = DEFAULT_TABLE_WIDTH
                if align is None: align = "margins"
            else:
                inches = ooolibUnits.convertLength(ooolibUnits.formatLength(width))
                if inches is None or inches <= 0:
                    raise ValueError("Invalid table width: %r" % (width,))
                if align is None: align = "left"
            widths = [inches / columns] * columns
        if str(align).lower() not in TABLE_ALIGNMENTS:
            raise ValueError("Table align must be one of %s" % ", ".join(TABLE_ALIGNMENTS))
        # Build the table element and its columns
        self.tableCount += 1
        tableName = "Table%d" % self.tableCount
        tableElement = self.officeText.addChild(ooolibXML.Element("table:table"))
        tableElement.setAttribute("table:name", tableName)
        tableElement.setAttribute("table:style-name",
                                  self.automaticStylesInstance.getTableStyle(sum(widths), str(align).lower()))
        for inches in widths:
            column = tableElement.addChild(ooolibXML.Element("table:table-column"))
            column.setAttribute("table:style-name", self.automaticStylesInstance.getColumnStyle(inches))
        table = Table(self, tableElement, columns, border, padding)
        for rowNumber, row in enumerate(data or []):
            table.addRow(row, header=rowNumber < headerRows, **formatting)
        return table

    ######################
    # Paragraph Building #
    ######################
    def paragraphData(self, **formatting):
        """Return a paragraph formatting dictionary holding every text and paragraph key."""
        return dict((name, formatting.get(name)) for name in TEXT_KEYS + PARAGRAPH_KEYS)

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
        # A single (string, formatting) run
        if isinstance(text, tuple) and len(text) == 2 and isinstance(text[1], (dict, type(None))):
            text = [text]
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
        self.global_object.setGlobalInt("tableCount", self.tableCount)

    def toString(self, indent=False):
        contentString = self.prolog.toString()
        contentString += self.documentContent.toString(indent)
        return contentString


#########
# Table #
#########
class Table:
    """A table in a Writer document, returned by Content.addTable()."""
    def __init__(self, content, element, columns, border, padding):
        self.content = content
        self.element = element
        self.columns = columns
        self.border = border
        self.padding = padding
        self.headerElement = None
        self.rowCount = 0
        self.bodyRowCount = 0

    @staticmethod
    def cellSpan(cell):
        if isinstance(cell, dict):
            return max(1, int(cell.get("colspan", 1)))
        return 1

    def addRow(self, cells, header=False, **formatting):
        """Add a row of cells and return the table:table-row element.

        Each cell is a string, a list of runs (as for addParagraph), or a dictionary
        holding "text" plus formatting for that cell only.  Formatting is bold, italic,
        underline, fontsize, fontcolor, fontname, align, background (a color such as
        "#dddddd"), and valign ("top", "middle", or "bottom"); a cell dictionary may
        also hold colspan.  Keyword arguments give formatting for every cell in the
        row.  Rows with fewer cells than the table has columns are filled with empty
        cells.  header=True makes the row a header row, repeated at the top of each
        page; header rows must come before the other rows."""
        unknown = set(formatting) - set(CELL_KEYS)
        if unknown:
            raise TypeError("Unknown table formatting: %s" % ", ".join(sorted(unknown)))
        if header:
            if self.bodyRowCount:
                raise ValueError("Header rows must be added before the other rows")
            if self.headerElement is None:
                self.headerElement = self.element.addChild(ooolibXML.Element("table:table-header-rows"))
            parent = self.headerElement
        else:
            parent = self.element
            self.bodyRowCount += 1
        span = sum(self.cellSpan(cell) for cell in cells)
        if span > self.columns:
            raise ValueError("Row has %d cells but the table has %d columns" % (span, self.columns))
        self.rowCount += 1
        row = parent.addChild(ooolibXML.Element("table:table-row"))
        styles = self.content.automaticStylesInstance
        column = 0
        for cell in list(cells) + [""] * (self.columns - span):
            cellData = dict(formatting)
            if isinstance(cell, dict):
                cellData.update(cell)
                text = cellData.pop("text", "")
                colspan = max(1, int(cellData.pop("colspan", 1)))
                unknown = set(cellData) - set(CELL_KEYS)
                if unknown:
                    raise TypeError("Unknown table cell formatting: %s" % ", ".join(sorted(unknown)))
            else:
                text, colspan = cell, 1
            if text is None: text = ""
            cellElement = row.addChild(ooolibXML.Element("table:table-cell"))
            # Cells draw their right and bottom borders, plus the left border in the first
            # column and the top border in the first row, so shared borders are not doubled
            cellElement.setAttribute("table:style-name", styles.getCellStyle(
                self.border, self.padding, cellData.get("background"), cellData.get("valign"),
                left=column == 0, top=self.rowCount == 1))
            column += colspan
            cellElement.setAttribute("office:value-type", "string")
            if colspan > 1:
                cellElement.setAttribute("table:number-columns-spanned", str(colspan))
            self.content.createParagraph(cellElement, text, self.content.paragraphData(**cellData))
            for _ in range(colspan - 1):
                row.addChild(ooolibXML.Element("table:covered-table-cell"))
        return row


##############
# Font Faces #
##############
# Generic families for common fonts, so a missing font falls back sensibly
FONT_GENERIC = {
    "liberation serif": ("roman", "variable"),
    "liberation sans": ("swiss", "variable"),
    "liberation mono": ("modern", "fixed"),
    "dejavu serif": ("roman", "variable"),
    "dejavu sans": ("swiss", "variable"),
    "dejavu sans mono": ("modern", "fixed"),
    "times new roman": ("roman", "variable"),
    "georgia": ("roman", "variable"),
    "arial": ("swiss", "variable"),
    "helvetica": ("swiss", "variable"),
    "verdana": ("swiss", "variable"),
    "calibri": ("swiss", "variable"),
    "courier new": ("modern", "fixed"),
    "courier": ("modern", "fixed"),
}

class FontFaces:
    """Declares each font used in a document once, in an office:font-face-decls element."""
    def __init__(self, element):
        self.element = element
        self.names = set()

    def addFont(self, name):
        name = str(name)
        if name not in self.names:
            self.names.add(name)
            fontFace = self.element.addChild(ooolibXML.Element("style:font-face"))
            fontFace.setAttribute("style:name", name)
            fontFace.setAttribute("svg:font-family", "'%s'" % name if " " in name else name)
            generic = FONT_GENERIC.get(name.lower())
            if generic is not None:
                fontFace.setAttribute("style:font-family-generic", generic[0])
                fontFace.setAttribute("style:font-pitch", generic[1])
        return name


####################
# Automatic Styles #
####################
class AutomaticStyles:
    def __init__(self, global_object, automaticStyles, fontFaces=None):
        self.global_object = global_object
        self.officeAutomaticStyles = automaticStyles
        self.fontFaces = fontFaces
        self.styleObjects = {}
        self.global_object.setGlobalInt("automaticStyleWriterParagraph", 1)
        self.global_object.setGlobalInt("automaticStyleWriterText", 1)
        self.global_object.setGlobalInt("automaticStyleWriterList", 1)
        self.global_object.setGlobalInt("automaticStyleWriterTable", 1)
        self.global_object.setGlobalInt("automaticStyleWriterColumn", 1)
        self.global_object.setGlobalInt("automaticStyleWriterCell", 1)

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

    def getTableStyle(self, widthInches, align):
        index = ("table", round(widthInches, 4), align)
        if index not in self.styleObjects:
            tableNum = self.global_object.incrementGlobalInt("automaticStyleWriterTable")
            tableName = "Tb%d" % tableNum
            styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
            styleObject.setAttribute("style:name", tableName)
            styleObject.setAttribute("style:family", "table")
            tableProp = styleObject.addChild(ooolibXML.Element("style:table-properties"))
            tableProp.setAttribute("style:width", ooolibUnits.formatLength(widthInches))
            tableProp.setAttribute("table:align", align)
            self.styleObjects[index] = tableName
        return self.styleObjects[index]

    def getColumnStyle(self, widthInches):
        index = ("column", round(widthInches, 4))
        if index not in self.styleObjects:
            columnNum = self.global_object.incrementGlobalInt("automaticStyleWriterColumn")
            columnName = "Co%d" % columnNum
            styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
            styleObject.setAttribute("style:name", columnName)
            styleObject.setAttribute("style:family", "table-column")
            columnProp = styleObject.addChild(ooolibXML.Element("style:table-column-properties"))
            columnProp.setAttribute("style:column-width", ooolibUnits.formatLength(widthInches))
            # Relative widths in thousandths of an inch keep the proportions exact
            columnProp.setAttribute("style:rel-column-width", "%d*" % max(1, round(widthInches * 1000)))
            self.styleObjects[index] = columnName
        return self.styleObjects[index]

    def getCellStyle(self, border, padding, background=None, valign=None, left=True, top=True):
        if valign is not None:
            if str(valign).lower() not in VERTICAL_ALIGNMENTS:
                raise ValueError("valign must be top, middle, or bottom")
            valign = VERTICAL_ALIGNMENTS[str(valign).lower()]
        if not border:
            border = None
            left = top = True
        index = ("cell", border, padding, background, valign, bool(left), bool(top))
        if index not in self.styleObjects:
            cellNum = self.global_object.incrementGlobalInt("automaticStyleWriterCell")
            cellName = "Ce%d" % cellNum
            styleObject = self.officeAutomaticStyles.addChild(ooolibXML.Element("style:style"))
            styleObject.setAttribute("style:name", cellName)
            styleObject.setAttribute("style:family", "table-cell")
            cellProp = styleObject.addChild(ooolibXML.Element("style:table-cell-properties"))
            if border is None:
                cellProp.setAttribute("fo:border", "none")
            elif left and top:
                cellProp.setAttribute("fo:border", border)
            else:
                cellProp.setAttribute("fo:border-left", border if left else "none")
                cellProp.setAttribute("fo:border-right", border)
                cellProp.setAttribute("fo:border-top", border if top else "none")
                cellProp.setAttribute("fo:border-bottom", border)
            if padding is not None:
                cellProp.setAttribute("fo:padding", ooolibUnits.formatLength(padding))
            if background is not None:
                cellProp.setAttribute("fo:background-color", background)
            if valign is not None:
                cellProp.setAttribute("style:vertical-align", valign)
            self.styleObjects[index] = cellName
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
        fontname = self.extractData(data, "fontname")
        if fontname is not None:
            if self.fontFaces is not None:
                self.fontFaces.addFont(fontname)
            textProp.setAttribute("style:font-name", str(fontname))


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
    content.addParagraph("Sans serif text.", fontname="Liberation Sans")
    table = content.addTable([["Name", "Qty"], ["Apples", "3"]], headerRows=1)
    table.addRow([{"text": "Total: 3", "colspan": 2, "align": "right"}])

    string1 = content.toString(indent=True)

    print("Writer Content XML with indentation:")
    print(string1)
