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
        self.global_object = global_object
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
        self.officeMeta = ooolibXML.Element("office:meta")
        # Creation Date
        dateStr = datetime.datetime.now().isoformat()
        self.creationDate = self.officeMeta.addChild(ooolibXML.Element("meta:creation-date", dateStr))
        # Document Statistics - Fill in later
        self.documentStatistic = self.officeMeta.addChild(ooolibXML.Element("meta:document-statistic"))
        # Generator (this python module)
        self.officeMeta.addChild(ooolibXML.Element("meta:generator", self.global_object.getVersion()))
        # Connect components
        self.documentMeta.addChild(self.officeMeta)

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
