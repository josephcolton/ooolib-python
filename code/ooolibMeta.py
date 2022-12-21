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
        self.creationDate = ooolibXML.Element("meta:creation-date", dateStr)
        # Document Statistics
        self.documentStatistic = ooolibXML.Element("meta:document-statistic")
        self.documentStatistic.setAttribute("meta:table-count", "1")
        self.documentStatistic.setAttribute("meta:cell-count", "0")
        self.documentStatistic.setAttribute("meta:object-count", "0")
        # Generator (this python module)
        self.metaGenerator = ooolibXML.Element("meta:generator", self.global_object.getVersion())
        # Connect components
        self.officeMeta.addChild(self.creationDate)
        self.officeMeta.addChild(self.documentStatistic)
        self.officeMeta.addChild(self.metaGenerator)
        self.documentMeta.addChild(self.officeMeta)

    def updateStatistics(self):
        pass

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
