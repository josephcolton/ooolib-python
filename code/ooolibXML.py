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

    def setAttribute(self, name, value):
        # Mark element as having attributes
        if (self.attributes == None):
            self.attributes = {}
        # Add attribute to attributes dictionary
        self.attributes[name] = value

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
        self.indentCharNum = 2
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

    def setAttribute(self, name, value):
        # Mark element as having attributes
        if (self.attributes == None):
            self.attributes = {}
        # Add attribute to attributes dictionary
        self.attributes[name] = value

    def addChild(self, child):
        # Mark element as having children
        if (self.children == None):
            self.children = []
        # Add child to element
        self.children.append(child)

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
                xmlString += "%s<%s%s />%s" % (indentPre, self.name, attributeStr, indentPost)   
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
                    xmlString += child.toString(indent, level+1)
                xmlString += "%s</%s>%s" % (indentPre, self.name, indentPost)
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
