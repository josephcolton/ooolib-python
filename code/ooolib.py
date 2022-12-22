#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os

# Import ooolib-python modules
import ooolibGlobal
import ooolibFile
import ooolibMeta
import ooolibCalcContent
import ooolibCalcManifest
import ooolibCalcSettings
import ooolibCalcStyles

###################################
# ooolib-python Calc Spreadsheets #
###################################
class Calc:
    """Calc - Open Document Format (ODF) spreadsheets

    Base object for creating ODF spreadsheets
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.manifest = ooolibCalcManifest.Manifest(self.global_object)
        self.meta = ooolibMeta.Meta(self.global_object)
        self.content = ooolibCalcContent.Content(self.global_object)
        self.settings = ooolibCalcSettings.Settings(self.global_object)
        self.styles = ooolibCalcStyles.Styles(self.global_object)

    def config(self, name, value):
        pass

    def export(self, filename):
        f = ooolibFile.SaveFile(filename)
        f.insertFileString("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        f.insertFileString("META-INF/manifest.xml", self.manifest.toString(indent=True))
        f.insertFileString("meta.xml", self.meta.toString())
        f.insertFileString("content.xml", self.content.toString())
        f.insertFileString("settings.xml", self.settings.toString())
        f.insertFileString("styles.xml", self.styles.toString())
        f.close()

class Writer:
    """Writer - Open Document Format (ODF) text documents

    Base object for creating ODF text documents
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.meta = ooolibMeta.Meta(self.global_object)

    def config(self, name, value):
        pass

###################
# Execute to Test #
###################
if __name__ == "__main__":
    # Make sure we can make simple documents
    os.makedirs("example", exist_ok=True)
    # Generate an empty calc
    calc = Calc()
    calc.export("example/simple-calc.ods")
