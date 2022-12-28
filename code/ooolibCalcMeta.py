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

#############################
# ooolib-python Meta object #
#############################
class Meta:
    def __init__(self, global_object):
        # Get passed in objects
        self.global_object = global_object
        # Document variables
        self.metaData = {}
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Meta
        self.documentMeta = ooolibXML.Element("office:document-meta")
        self.documentMeta.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.documentMeta.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.documentMeta.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentMeta.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentMeta.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentMeta.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.documentMeta.setAttribute("office:version", "1.3")
        # Office Meta
        self.officeMeta = self.documentMeta.addChild(ooolibXML.Element("office:meta"))
        # Creation Date
        dateStr = datetime.datetime.now().isoformat()
        self.creationDate = self.officeMeta.addChild(ooolibXML.Element("meta:creation-date", dateStr))
        # Document Statistics - Fill in later
        self.documentStatistic = self.officeMeta.addChild(ooolibXML.Element("meta:document-statistic"))
        # Generator (this python module)
        self.officeMeta.addChild(ooolibXML.Element("meta:generator", self.global_object.getVersion()))

        # Initialize Blank Objects for Meta Objects
        self.metaData["title"] = None
        self.metaData["subject"] = None
        self.metaData["description"] = None
        self.metaData["keywords"] = []

    ##########################
    # Set Document Meta Data #
    ##########################
    def setTitle(self, title):
        # Create child data
        if self.metaData["title"] == None:
            self.metaData["title"] = self.officeMeta.addChild(ooolibXML.Element("dc:title"))
        # Set/reset the title
        self.metaData["title"].setText(title)

    def setSubject(self, subject):
        # Create child data
        if self.metaData["subject"] == None:
            self.metaData["subject"] = self.officeMeta.addChild(ooolibXML.Element("dc:subject"))
        # Set/reset the subject
        self.metaData["subject"].setText(subject)

    def setDescription(self, description):
        # Create child data
        if self.metaData["description"] == None:
            self.metaData["description"] = self.officeMeta.addChild(ooolibXML.Element("dc:description"))
        # Set/reset the description
        self.metaData["description"].setText(description)

    def addKeyword(self, keyword):
        # See if the keyword already exists
        for keywordObject in self.metaData["keywords"]:
            if (keywordObject.getText() == keyword): return
        # Add the keyword object
        keywordObject = self.officeMeta.addChild(ooolibXML.Element("meta:keyword"))
        keywordObject.setText(keyword)
        self.metaData["keywords"].append(keywordObject)

    ##################
    # XML Generation #
    ##################
    def updateObjects(self):
        pass

    def updateStatistics(self):
        # Get statistics from global object
        tableCount = self.global_object.getGlobalInt("tableCount")
        cellCount = self.global_object.getGlobalInt("cellCount")
        objectCount = self.global_object.getGlobalInt("objectCount")
        # Update statistics
        self.documentStatistic.setAttribute("meta:table-count", tableCount)
        self.documentStatistic.setAttribute("meta:cell-count", cellCount)
        self.documentStatistic.setAttribute("meta:object-count", objectCount)

    def toString(self, indent=False):
        self.updateObjects()
        self.updateStatistics()
        metaString = self.prolog.toString()
        metaString += self.documentMeta.toString(indent)
        return metaString


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    meta = Meta(g)

    # Display meta file
    string1 = meta.toString(indent=True)

    print("Meta XML with indentation:")
    print(string1)
