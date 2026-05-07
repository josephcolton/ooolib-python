#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import datetime

# Import ooolib-python modules
import ooolibGlobal
import ooolibXML

#################################
# ooolib-python Writer Meta object #
#################################
class Meta:
    def __init__(self, global_object):
        self.global_object = global_object
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
        # Document Statistics - filled in on export
        self.documentStatistic = self.officeMeta.addChild(ooolibXML.Element("meta:document-statistic"))
        # Generator
        self.officeMeta.addChild(ooolibXML.Element("meta:generator", self.global_object.getVersion()))

        self.metaData["title"] = None
        self.metaData["subject"] = None
        self.metaData["description"] = None
        self.metaData["keywords"] = []

    ##########################
    # Set Document Meta Data #
    ##########################
    def setTitle(self, title):
        if self.metaData["title"] == None:
            self.metaData["title"] = self.officeMeta.addChild(ooolibXML.Element("dc:title"))
        self.metaData["title"].setText(title)

    def setSubject(self, subject):
        if self.metaData["subject"] == None:
            self.metaData["subject"] = self.officeMeta.addChild(ooolibXML.Element("dc:subject"))
        self.metaData["subject"].setText(subject)

    def setDescription(self, description):
        if self.metaData["description"] == None:
            self.metaData["description"] = self.officeMeta.addChild(ooolibXML.Element("dc:description"))
        self.metaData["description"].setText(description)

    def addKeyword(self, keyword):
        for keywordObject in self.metaData["keywords"]:
            if (keywordObject.getText() == keyword): return
        keywordObject = self.officeMeta.addChild(ooolibXML.Element("meta:keyword"))
        keywordObject.setText(keyword)
        self.metaData["keywords"].append(keywordObject)

    ##################
    # XML Generation #
    ##################
    def updateStatistics(self):
        paragraphCount = self.global_object.getGlobalInt("paragraphCount")
        self.documentStatistic.setAttribute("meta:paragraph-count", paragraphCount)

    def toString(self, indent=False):
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

    string1 = meta.toString(indent=True)

    print("Writer Meta XML with indentation:")
    print(string1)
