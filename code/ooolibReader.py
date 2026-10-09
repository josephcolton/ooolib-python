#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import re
import sys
import xml.etree.ElementTree as ElementTree

# Import ooolib-python modules
import ooolibFile
import ooolibUnits

##############
# Namespaces #
##############
NAMESPACES = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "presentation": "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}

def q(name):
    """Convert a prefixed name such as 'draw:rect' to ElementTree '{uri}rect' form."""
    prefix, local = name.split(":", 1)
    return "{%s}%s" % (NAMESPACES[prefix], local)

def localName(tag):
    """Convert an ElementTree tag such as '{uri}rect' to 'prefix:rect'."""
    if not tag.startswith("{"): return tag
    uri, local = tag[1:].split("}", 1)
    for prefix in NAMESPACES:
        if NAMESPACES[prefix] == uri:
            return "%s:%s" % (prefix, local)
    return local

def attr(element, name, default=None):
    return element.get(q(name), default)

# Shape element names mapped to the shape type reported by the reader
SHAPE_TYPES = {
    "draw:rect": "rect",
    "draw:ellipse": "ellipse",
    "draw:circle": "circle",
    "draw:line": "line",
    "draw:polygon": "polygon",
    "draw:polyline": "polyline",
    "draw:path": "path",
    "draw:regular-polygon": "regular-polygon",
    "draw:connector": "connector",
    "draw:measure": "measure",
    "draw:caption": "caption",
    "draw:custom-shape": "custom-shape",
    "draw:frame": "frame",
    "draw:g": "group",
}

# Paragraph-level text elements
PARAGRAPH_TAGS = (q("text:p"), q("text:h"))

################
# Text Helpers #
################
# Placeholders for whitespace that ODF does not collapse
KEPT_SPACE, KEPT_TAB, KEPT_BREAK = "\x00", "\x01", "\x02"

def elementText(element):
    """Return the plain text of a text:p / text:h / text:span element.

    Whitespace is collapsed the same way LibreOffice displays it."""
    text = rawElementText(element)
    text = re.sub(r" {2,}", " ", text).strip(" ")
    text = text.replace(KEPT_SPACE, " ").replace(KEPT_TAB, "\t").replace(KEPT_BREAK, "\n")
    return text

def collapseWhitespace(text):
    if not text: return ""
    return re.sub(r"[ \t\r\n]+", " ", text)

def rawElementText(element):
    parts = [collapseWhitespace(element.text)]
    for child in element:
        tag = child.tag
        if tag == q("text:s"):
            parts.append(KEPT_SPACE * int(child.get(q("text:c"), "1")))
        elif tag == q("text:tab"):
            parts.append(KEPT_TAB)
        elif tag == q("text:line-break"):
            parts.append(KEPT_BREAK)
        elif tag in (q("office:annotation"), q("text:note")):
            pass # Skip comments and footnote bodies
        else:
            parts.append(rawElementText(child))
        parts.append(collapseWhitespace(child.tail))
    return "".join(parts)

def paragraphTexts(element):
    """Return the text of every paragraph below element, in document order.

    Paragraphs inside nested shapes are not included (they belong to those shapes)."""
    paragraphs = []
    for child in element:
        if child.tag in PARAGRAPH_TAGS:
            paragraphs.append(elementText(child))
        elif localName(child.tag) in SHAPE_TYPES:
            continue
        else:
            paragraphs.extend(paragraphTexts(child))
    return paragraphs


##########
# Styles #
##########
class StyleSheet:
    """Collects the styles from styles.xml and content.xml so shape
    properties can be resolved through parent styles and default styles."""
    def __init__(self):
        self.styles = {}        # (family, name) -> style element
        self.defaultStyles = {} # family -> style element

    def addStyles(self, root):
        if root is None: return
        for container in ("office:styles", "office:automatic-styles"):
            for parent in root.iter(q(container)):
                for style in parent.findall(q("style:style")):
                    key = (attr(style, "style:family"), attr(style, "style:name"))
                    self.styles[key] = style
                for style in parent.findall(q("style:default-style")):
                    self.defaultStyles[attr(style, "style:family")] = style

    def getProperties(self, family, name, propertiesName):
        """Return a dictionary of properties ('draw:fill', 'svg:stroke-color', ...)
        for the named style, including inherited and default values."""
        chain = []
        seen = set()
        while name is not None and (family, name) not in seen:
            seen.add((family, name))
            style = self.styles.get((family, name))
            if style is None: break
            chain.append(style)
            name = attr(style, "style:parent-style-name")
        if family in self.defaultStyles:
            chain.append(self.defaultStyles[family])
        # Apply from most general to most specific
        properties = {}
        for style in reversed(chain):
            props = style.find(q("style:%s" % propertiesName))
            if props is None: continue
            for key in props.attrib:
                properties[localName(key)] = props.attrib[key]
        return properties


#########
# Shape #
#########
class Shape:
    """A shape read from a drawing or presentation page."""
    def __init__(self, element, styleSheet):
        self.element = element
        self.tag = localName(element.tag)
        self.type = SHAPE_TYPES.get(self.tag, self.tag)
        self.name = attr(element, "draw:name")
        self.id = attr(element, "draw:id") or element.get("{http://www.w3.org/XML/1998/namespace}id")
        self.layer = attr(element, "draw:layer")
        self.presentationClass = attr(element, "presentation:class")
        self.transform = attr(element, "draw:transform")
        # Position and size (strings as stored, e.g. '1.5in')
        self.x = attr(element, "svg:x")
        self.y = attr(element, "svg:y")
        self.width = attr(element, "svg:width")
        self.height = attr(element, "svg:height")
        # Line and connector end points
        self.x1 = attr(element, "svg:x1")
        self.y1 = attr(element, "svg:y1")
        self.x2 = attr(element, "svg:x2")
        self.y2 = attr(element, "svg:y2")
        # Geometry details
        self.points = attr(element, "draw:points")
        self.path = attr(element, "svg:d")
        self.geometry = None
        if self.type == "custom-shape":
            geometry = element.find(q("draw:enhanced-geometry"))
            if geometry is not None: self.geometry = attr(geometry, "draw:type")
        # Frame contents decide the frame's type
        self.href = None
        if self.type == "frame":
            if element.find(q("draw:text-box")) is not None:
                self.type = "textbox"
            elif element.find(q("draw:image")) is not None:
                self.type = "image"
                self.href = attr(element.find(q("draw:image")), "xlink:href")
            elif element.find(q("draw:object")) is not None:
                self.type = "object"
                self.href = attr(element.find(q("draw:object")), "xlink:href")
            elif element.find(q("table:table")) is not None:
                self.type = "table"
        # Styles
        self.styleName = attr(element, "draw:style-name")
        self.presentationStyleName = attr(element, "presentation:style-name")
        if self.styleName is not None:
            self.graphicProperties = styleSheet.getProperties("graphic", self.styleName, "graphic-properties")
        elif self.presentationStyleName is not None:
            self.graphicProperties = styleSheet.getProperties("presentation", self.presentationStyleName, "graphic-properties")
        else:
            self.graphicProperties = styleSheet.getProperties("graphic", None, "graphic-properties")
        self.fill = self.graphicProperties.get("draw:fill")
        self.fillColor = self.graphicProperties.get("draw:fill-color")
        self.stroke = self.graphicProperties.get("draw:stroke")
        self.strokeColor = self.graphicProperties.get("svg:stroke-color")
        self.strokeWidth = self.graphicProperties.get("svg:stroke-width")
        # Colors inherited from default styles are meaningless when fill/stroke is off
        if self.fill == "none": self.fillColor = None
        if self.stroke == "none":
            self.strokeColor = None
            self.strokeWidth = None
        # Text
        self.paragraphs = paragraphTexts(element)
        self.text = "\n".join(self.paragraphs)
        # Child shapes (groups)
        self.children = []
        if self.type == "group":
            self.children = readShapes(element, styleSheet)

    def getLength(self, name, unit="in"):
        """Return a position or size attribute ('x', 'width', 'x1', ...) as a float in unit."""
        return ooolibUnits.convertLength(getattr(self, name), unit)

    def getBounds(self, unit="in"):
        """Return (x, y, width, height) as floats in unit, or None if unknown."""
        if self.type in ("line", "connector", "measure"):
            x1, y1 = self.getLength("x1", unit), self.getLength("y1", unit)
            x2, y2 = self.getLength("x2", unit), self.getLength("y2", unit)
            if None in (x1, y1, x2, y2): return None
            return (min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
        bounds = (self.getLength("x", unit), self.getLength("y", unit),
                  self.getLength("width", unit), self.getLength("height", unit))
        if None in bounds: return None
        return bounds

    def iterShapes(self):
        """Yield this shape and every shape nested inside it."""
        yield self
        for child in self.children:
            for shape in child.iterShapes():
                yield shape

    def getText(self):
        """Return the text of this shape and any shapes nested inside it."""
        texts = [shape.text for shape in self.iterShapes() if shape.text]
        return "\n".join(texts)

    def toDict(self):
        data = {
            "type": self.type, "name": self.name, "text": self.text,
            "x": self.x, "y": self.y, "width": self.width, "height": self.height,
            "fill": self.fill, "fillColor": self.fillColor,
            "stroke": self.stroke, "strokeColor": self.strokeColor, "strokeWidth": self.strokeWidth,
        }
        for key in ("x1", "y1", "x2", "y2", "points", "path", "geometry", "href",
                    "layer", "presentationClass", "transform"):
            if getattr(self, key) is not None: data[key] = getattr(self, key)
        if self.children: data["children"] = [child.toDict() for child in self.children]
        return data

    def __repr__(self):
        description = "<Shape %s" % self.type
        if self.name: description += " name=%r" % self.name
        if self.text: description += " text=%r" % (self.text if len(self.text) <= 30 else self.text[:27] + "...")
        return description + ">"

def readShapes(parent, styleSheet):
    """Return Shape objects for the direct shape children of parent."""
    shapes = []
    for child in parent:
        if localName(child.tag) in SHAPE_TYPES:
            shapes.append(Shape(child, styleSheet))
    return shapes


########
# Page #
########
class Page:
    """A page from a drawing (.odg) or slide from a presentation (.odp)."""
    def __init__(self, element, styleSheet, number):
        self.element = element
        self.number = number
        self.name = attr(element, "draw:name")
        self.masterPageName = attr(element, "draw:master-page-name")
        self.shapes = readShapes(element, styleSheet)
        # Presentation speaker notes
        self.notes = None
        notes = element.find(q("presentation:notes"))
        if notes is not None:
            noteTexts = [shape.getText() for shape in readShapes(notes, styleSheet)
                         if shape.presentationClass == "notes"]
            self.notes = "\n".join([text for text in noteTexts if text]) or None

    def getPageName(self):
        return self.name

    def iterShapes(self):
        """Yield every shape on the page, including shapes inside groups."""
        for shape in self.shapes:
            for nested in shape.iterShapes():
                yield nested

    def getShapes(self, shapeType=None):
        """Return every shape on the page (including grouped shapes), optionally filtered by type."""
        return [shape for shape in self.iterShapes() if shapeType is None or shape.type == shapeType]

    def getShape(self, name):
        """Return the first shape with the given draw:name, or None."""
        for shape in self.iterShapes():
            if shape.name == name: return shape
        return None

    def getText(self):
        texts = [shape.text for shape in self.iterShapes() if shape.text]
        return "\n".join(texts)

    def toDict(self):
        data = {"number": self.number, "name": self.name,
                "shapes": [shape.toDict() for shape in self.shapes]}
        if self.notes: data["notes"] = self.notes
        return data

    def __repr__(self):
        return "<Page %d name=%r shapes=%d>" % (self.number, self.name, len(self.shapes))


############
# Document #
############
class Document:
    """Document - Read information from an existing ODF document

    Works with any ODF file. Pages and shapes are available for drawings
    (.odg) and presentations (.odp)."""
    def __init__(self, filename):
        self.filename = filename
        readFile = ooolibFile.ReadFile(filename)
        try:
            self.fileList = readFile.getFileList()
            self.mimetype = (readFile.readFileString("mimetype") or "").strip()
            self.type = self.mimetype.rsplit(".", 1)[-1] if self.mimetype else None
            contentRoot = self.parse(readFile, "content.xml")
            stylesRoot = self.parse(readFile, "styles.xml")
            metaRoot = self.parse(readFile, "meta.xml")
            self.images = []
            for name in self.fileList:
                if name.startswith("Pictures/") and not name.endswith("/"):
                    self.images.append({"name": name, "size": readFile.getFileSize(name)})
        finally:
            readFile.close()
        # Styles
        self.styleSheet = StyleSheet()
        self.styleSheet.addStyles(stylesRoot)
        self.styleSheet.addStyles(contentRoot)
        # Information
        self.meta = self.readMeta(metaRoot)
        self.contentRoot = contentRoot
        self.pages = []
        if contentRoot is not None:
            for number, page in enumerate(contentRoot.iter(q("draw:page")), 1):
                self.pages.append(Page(page, self.styleSheet, number))
        masterPageName = self.pages[0].masterPageName if self.pages else None
        self.pageLayout = self.readPageLayout(stylesRoot, masterPageName)

    def parse(self, readFile, name):
        data = readFile.readFileBytes(name)
        if data is None: return None
        return ElementTree.fromstring(data)

    def readMeta(self, metaRoot):
        meta = {
            "title": None, "subject": None, "description": None, "keywords": [],
            "creator": None, "initialCreator": None, "creationDate": None,
            "date": None, "generator": None, "language": None, "statistics": {},
            "userDefined": {},
        }
        if metaRoot is None: return meta
        officeMeta = metaRoot.find(q("office:meta"))
        if officeMeta is None: return meta
        simple = {
            "dc:title": "title", "dc:subject": "subject", "dc:description": "description",
            "dc:creator": "creator", "meta:initial-creator": "initialCreator",
            "meta:creation-date": "creationDate", "dc:date": "date",
            "meta:generator": "generator", "dc:language": "language",
        }
        for child in officeMeta:
            name = localName(child.tag)
            if name in simple:
                meta[simple[name]] = child.text
            elif name == "meta:keyword":
                if child.text: meta["keywords"].append(child.text)
            elif name == "meta:document-statistic":
                for key in child.attrib:
                    value = child.attrib[key]
                    meta["statistics"][localName(key).split(":", 1)[-1]] = int(value) if value.isdigit() else value
            elif name == "meta:user-defined":
                meta["userDefined"][attr(child, "meta:name")] = child.text
        return meta

    def readPageLayout(self, stylesRoot, masterPageName=None):
        """Return the page size and margins used by the named master page
        (or the first master page), or None."""
        if stylesRoot is None: return None
        masterPages = list(stylesRoot.iter(q("style:master-page")))
        layoutName = None
        for masterPage in masterPages:
            if attr(masterPage, "style:name") == masterPageName:
                layoutName = attr(masterPage, "style:page-layout-name")
        if layoutName is None and masterPages:
            layoutName = attr(masterPages[0], "style:page-layout-name")
        layouts = list(stylesRoot.iter(q("style:page-layout")))
        if not layouts: return None
        pageLayout = layouts[0]
        for candidate in layouts:
            if attr(candidate, "style:name") == layoutName: pageLayout = candidate
        layout = pageLayout.find(q("style:page-layout-properties"))
        if layout is None: return None
        return {
            "width": attr(layout, "fo:page-width"),
            "height": attr(layout, "fo:page-height"),
            "orientation": attr(layout, "style:print-orientation"),
            "marginTop": attr(layout, "fo:margin-top"),
            "marginBottom": attr(layout, "fo:margin-bottom"),
            "marginLeft": attr(layout, "fo:margin-left"),
            "marginRight": attr(layout, "fo:margin-right"),
        }

    ###########
    # Queries #
    ###########
    def getPageNames(self):
        return [page.name for page in self.pages]

    def getPage(self, name):
        """Return a page by name or by 1-based page number, or None."""
        for page in self.pages:
            if page.name == name or page.number == name: return page
        return None

    def getShapes(self, shapeType=None):
        """Return every shape in the document, optionally filtered by type."""
        shapes = []
        for page in self.pages:
            shapes.extend(page.getShapes(shapeType))
        return shapes

    def getText(self):
        """Return all text in the document.

        For drawings and presentations this is the text of each page's shapes;
        for other documents it is every paragraph in the document body."""
        if self.pages:
            return "\n".join([page.getText() for page in self.pages if page.getText()])
        if self.contentRoot is None: return ""
        body = self.contentRoot.find(q("office:body"))
        if body is None: return ""
        texts = [elementText(p) for p in body.iter() if p.tag in PARAGRAPH_TAGS]
        return "\n".join(texts)

    def toDict(self):
        return {
            "filename": self.filename, "type": self.type, "meta": self.meta,
            "pageLayout": self.pageLayout, "images": self.images,
            "pages": [page.toDict() for page in self.pages],
        }

    def __repr__(self):
        return "<Document %r type=%s pages=%d>" % (self.filename, self.type, len(self.pages))


###################
# Execute to Test #
###################
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: %s document.odg" % sys.argv[0])
        sys.exit(1)
    doc = Document(sys.argv[1])
    print(doc)
    print("Title: %s" % doc.meta["title"])
    print("Keywords: %s" % ", ".join(doc.meta["keywords"]))
    print("Statistics: %s" % doc.meta["statistics"])
    print("Page layout: %s" % doc.pageLayout)
    for page in doc.pages:
        print(page)
        for shape in page.shapes:
            print("    %r bounds=%s fill=%s stroke=%s" % (shape, shape.getBounds(), shape.fillColor, shape.strokeColor))
