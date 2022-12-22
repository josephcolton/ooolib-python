#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

############################
# ooolib-python XML Prolog #
############################
class Prolog:
    """Prolog - XML prolog element

    Create a special XML element prolog
    """
    def __init__(self, name, default=True):
        # Parameters
        self.name = name
        # Internal variables
        self.attributes = None
        # Default prolog
        self.initDefault()

    def initDefault(self):
        self.setAttribute("version", "1.0")
        self.setAttribute("encoding", "UTF-8")

    ###########
    # Element #
    ###########
    def getElementName(self):
        return self.name

    def setElementName(self, name):
        self.name = name

    ##############
    # Attributes #
    ##############
    def setAttribute(self, name, value):
        # Mark element as having attributes
        if (self.attributes == None):
            self.attributes = {}
        # Add attribute to attributes dictionary
        self.attributes[name] = value
        # Return element for chaining
        return self

    def getAttribute(self, name, default=None):
        # Make sure we have attributes
        if (self.attributes == None): return default
        # Make sure we have this attribute
        if name in self.attributes:
            return self.attributes[name]
        # Return default value
        return default

    def removeAttribute(self, name, default=None):
        # Make sure we have attributes
        if (self.attributes == None): return default
        # Make sure we have this attribute
        if name in self.attributes:
            return self.attributes.pop(name)
        # Return default value
        return default

    def __attributeString(self):
        attributeStr = ""
        # No attrinbutes
        if (self.attributes == None):
            return attributeStr
        # Has attributes
        for name in self.attributes:
            value = self.attributes[name]
            attributeStr += " %s=\"%s\"" % (name, value)
        # Return attribute string
        return attributeStr

    #####################
    # String Conversion #
    #####################
    def toString(self, indent=False):
        # Get attribute string
        attributeStr = self.__attributeString()

        # Create XML string
        xmlString = "<?%s%s?>\n" % (self.name, attributeStr)
        # Return string
        return xmlString

#############################
# ooolib-python XML Element #
#############################
class Element:
    """Element - XML element creation

    Used to create XML elements for use in ooolib-python document files.
    """
    def __init__(self, name, text=None):
        # Default values
        self.indentChar = " "
        self.indentCharNum = 1
        self.indentEndline = "\n"
        # Parameters
        self.name = name
        self.text = text
        # Internal variables
        self.attributes = None
        self.children = None

    def config(self, name, value):
        # XML Display Indentation
        if (name == "indentChar"): self.indentChar = str(value)
        if (name == "indentCharNum"): self.indentCharNum = int(value)
        if (name == "indentEndline"): self.indentEndline = str(value)

    ###########
    # Element #
    ###########
    def getElementName(self):
        return self.name

    def setElementName(self, name):
        self.name = name

    ########
    # Text #
    ########
    def setText(self, text):
        self.text = text

    def getText(self):
        return self.text

    def removeText(self):
        text = self.text
        self.text = None
        return text

    ##############
    # Attributes #
    ##############
    def setAttribute(self, name, value):
        # Mark element as having attributes
        if (self.attributes == None):
            self.attributes = {}
        # Add attribute to attributes dictionary
        self.attributes[name] = value
        # Return element for chaining
        return self

    def getAttribute(self, name, default=None):
        # Make sure we have attributes
        if (self.attributes == None): return default
        # Make sure we have this attribute
        if name in self.attributes:
            return self.attributes[name]
        # Return default value
        return default

    def removeAttribute(self, name, default=None):
        # Make sure we have attributes
        if (self.attributes == None): return default
        # Make sure we have this attribute
        if name in self.attributes:
            return self.attributes.pop(name)
        # Return default value
        return default

    def __attributeString(self):
        attributeStr = ""
        # No attrinbutes
        if (self.attributes == None):
            return attributeStr
        # Has attributes
        for name in self.attributes:
            value = self.attributes[name]
            attributeStr += " %s=\"%s\"" % (name, value)
        # Return attribute string
        return attributeStr

    ############
    # Children #
    ############
    def addChild(self, child):
        # Mark element as having children
        if (self.children == None):
            self.children = []
        # Add child to element
        self.children.append(child)
        # Return child
        return child

    def removeChild(self, child):
        # Only remove children if they exist
        if (self.children == None): return
        # Try to remove child
        self.children.remove(child)
        # Return child
        return child

    def removeChildIndex(self, index):
        # Only remove children if they exist
        if (self.children == None): return
        # Make sure index is an integer
        index = int(index)
        # Make index is valid
        count = len(self.children)
        if (index < -count): return # From back
        if (index >= count): return # From front
        # Remove the index
        child = self.children.pop(index)
        # Return child
        return child

    ##########################
    # Special Style Children #
    ##########################
    def addStyle(self, styleName, styleFamily=None, parentStyleName=None):
        child = self.addChild(Element("style:style"))
        child.setAttribute("style:name", styleName)
        if (styleFamily != None): child.setAttribute("style:family", styleFamily)
        if (parentStyleName != None): child.setAttribute("style:parent-style-name", parentStyleName)
        return child

    ##########################
    # Special Table Children #
    ##########################
    def addTableCellFloat(self, value):
        child = self.addChild(Element("table:table-cell"))
        child.setAttribute("office:value-type", "float")
        child.setAttribute("office:value", value)
        child.setAttribute("calcext:value-type", "float")
        # Display text
        child.addChild(Element("text:p", value))
        return child

    def addTableCellString(self, value):
        child = self.addChild(Element("table:table-cell"))
        child.setAttribute("office:value-type", "string")
        child.setAttribute("calcext:value-type", "string")
        # Display text
        child.addChild(Element("text:p", value))
        return child

    #####################
    # String Conversion #
    #####################
    def toString(self, indent=False, level=0):
        # Indentation formatting
        if (indent):
            indentPre = self.indentChar * self.indentCharNum * level
            indentPost = self.indentEndline
        else:
            indentPre = ""
            indentPost = ""

        # Get attribute string
        attributeStr = self.__attributeString()

        # Element Text
        if (self.text == None): text = ""
        else: text = self.text

        # Create XML string
        xmlString = ""
        if (self.children == None):
            if (self.text == None):
                # No children or text
                xmlString += "%s<%s%s/>%s" % (indentPre, self.name, attributeStr, indentPost)
            else:
                # No children, but text
                xmlString = "%s<%s%s>%s" % (indentPre, self.name, attributeStr, text)
                xmlString += "</%s>%s" % (self.name, indentPost)
        else:
            if (self.text == None):
                # Children, but no text
                xmlString = "%s<%s%s>%s" % (indentPre, self.name, attributeStr, indentPost)
                for child in self.children:
                    xmlString += child.toString(indent, level+1)
                xmlString += "%s</%s>%s" % (indentPre, self.name, indentPost)
            else:
                # Children and text (this is a mess)
                xmlString = "%s<%s%s>%s" % (indentPre, self.name, attributeStr, text)
                for child in self.children:
                    xmlString += child.toString(indent=False)
                xmlString += "</%s>%s" % (self.name, indentPost)
        # Return string
        return xmlString


###################
# Execute to Test #
###################
if __name__ == "__main__":
    # XML Prolog
    p = Prolog("xml")

    # Main Element
    x = Element("element")
    x.setAttribute("attrib1", 1)
    x.setAttribute("attrib2", "xyz")
    # Child 1
    c1 = Element("child1", "Some text")
    c1.setAttribute("attrib3", "123")
    x.addChild(c1)
    # Child 2
    c2 = Element("child2")
    c2.setAttribute("attrib4", "abc")
    x.addChild(c2)
    # Child 3
    c3 = Element("child3")
    x.addChild(c3)
    # Get the resulting string
    string0 = p.toString()
    string1 = x.toString(indent=True)
    string2 = x.toString()

    print("XML prolog:")
    print(string0)

    print("XML document with indentation:")
    print(string1)

    print("XML document without indentation:")
    print(string2)
