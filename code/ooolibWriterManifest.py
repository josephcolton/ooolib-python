#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import ooolib-python modules
import ooolibGlobal
import ooolibXML

########################################
# ooolib-python Writer Manifest object #
########################################
class Manifest:
    def __init__(self, global_object):
        self.global_object = global_object
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Manifest Object
        self.manifest = ooolibXML.Element("manifest:manifest")
        self.manifest.setAttribute("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
        self.manifest.setAttribute("manifest:version", "1.3")
        # Included Files
        fileEntry1 = self.manifest.addChild(ooolibXML.Element("manifest:file-entry"))
        fileEntry1.setAttribute("manifest:full-path", "/")
        fileEntry1.setAttribute("manifest:version", "1.3")
        fileEntry1.setAttribute("manifest:media-type", "application/vnd.oasis.opendocument.text")
        fileEntry2 = self.manifest.addChild(ooolibXML.Element("manifest:file-entry"))
        fileEntry2.setAttribute("manifest:full-path", "styles.xml")
        fileEntry2.setAttribute("manifest:media-type", "text/xml")
        fileEntry3 = self.manifest.addChild(ooolibXML.Element("manifest:file-entry"))
        fileEntry3.setAttribute("manifest:full-path", "meta.xml")
        fileEntry3.setAttribute("manifest:media-type", "text/xml")
        fileEntry4 = self.manifest.addChild(ooolibXML.Element("manifest:file-entry"))
        fileEntry4.setAttribute("manifest:full-path", "content.xml")
        fileEntry4.setAttribute("manifest:media-type", "text/xml")
        fileEntry5 = self.manifest.addChild(ooolibXML.Element("manifest:file-entry"))
        fileEntry5.setAttribute("manifest:full-path", "settings.xml")
        fileEntry5.setAttribute("manifest:media-type", "text/xml")

    def toString(self, indent=False):
        manifestString = self.prolog.toString()
        manifestString += self.manifest.toString(indent)
        return manifestString


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    manifest = Manifest(g)

    string1 = manifest.toString(indent=True)

    print("Writer Manifest XML with indentation:")
    print(string1)
