#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os, sys

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
    #################################
    # Example 1 - Blank Spreadsheet #
    #################################
    calc = Calc()
    calc.export("example/example1-blank.ods")
    ########################################
    # Example 2 - Spreadsheet with numbers #
    ########################################
    calc = Calc()
    for row in range(1, 5+1):
        for col in range(1, 5+1):
            value = row * col
            calc.content.activeTable.setCellFloat(row, col, value)
    calc.export("example/example2-numbers.ods")
    #################################################
    # Example 3 - Spreadsheet with text and numbers #
    #################################################
    calc = Calc()
    # Row 1
    calc.content.activeTable.setCellText(1, 1, "Year")
    calc.content.activeTable.setCellText(1, 2, "Event")
    # Row 2
    calc.content.activeTable.setCellFloat(2, 1, 2000)
    calc.content.activeTable.setCellText(2, 2, "Y2K Issues")
    # Row 3
    calc.content.activeTable.setCellFloat(3, 1, 2038)
    calc.content.activeTable.setCellText(3, 2, "Epochalypse")
    calc.export("example/example3-number_text.ods")

