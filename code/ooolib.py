#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os, sys

# Import ooolib-python modules
import ooolibGlobal
import ooolibFile

# Calc Modules
import ooolibCalcMeta
import ooolibCalcContent
import ooolibCalcManifest
import ooolibCalcSettings
import ooolibCalcStyles

# Writer Modules
import ooolibWriterMeta
import ooolibWriterContent
import ooolibWriterManifest
import ooolibWriterSettings
import ooolibWriterStyles

# Impress Modules
import ooolibImpressMeta
import ooolibImpressContent
import ooolibImpressManifest
import ooolibImpressSettings
import ooolibImpressStyles

# Draw Modules
import ooolibDrawMeta
import ooolibDrawContent
import ooolibDrawManifest
import ooolibDrawSettings
import ooolibDrawStyles

###################################
# ooolib-python Calc Spreadsheets #
###################################
class Calc:
    """Calc - Open Document Format (ODF) spreadsheets

    Base object for creating ODF spreadsheets
    """
    def __init__(self, autoInit=True):
        self.global_object = ooolibGlobal.Global()
        self.manifest = ooolibCalcManifest.Manifest(self.global_object)
        self.meta = ooolibCalcMeta.Meta(self.global_object)
        self.content = ooolibCalcContent.Content(self.global_object, autoInit)
        self.settings = ooolibCalcSettings.Settings(self.global_object)
        self.styles = ooolibCalcStyles.Styles(self.global_object)

    def export(self, filename):
        # Update statistics first
        self.content.updateStats()
        # Create file
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
        self.manifest = ooolibWriterManifest.Manifest(self.global_object)
        self.meta = ooolibWriterMeta.Meta(self.global_object)
        self.content = ooolibWriterContent.Content(self.global_object)
        self.settings = ooolibWriterSettings.Settings(self.global_object)
        self.styles = ooolibWriterStyles.Styles(self.global_object)

    def export(self, filename):
        self.content.updateStats()
        f = ooolibFile.SaveFile(filename)
        f.insertFileString("mimetype", "application/vnd.oasis.opendocument.text")
        f.insertFileString("META-INF/manifest.xml", self.manifest.toString(indent=True))
        f.insertFileString("meta.xml", self.meta.toString())
        f.insertFileString("content.xml", self.content.toString())
        f.insertFileString("settings.xml", self.settings.toString())
        f.insertFileString("styles.xml", self.styles.toString())
        f.close()

####################################
# ooolib-python Impress Presentations #
####################################
class Impress:
    """Impress - Open Document Format (ODF) presentation documents

    Base object for creating ODF presentation documents
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.manifest = ooolibImpressManifest.Manifest(self.global_object)
        self.meta = ooolibImpressMeta.Meta(self.global_object)
        self.content = ooolibImpressContent.Content(self.global_object)
        self.settings = ooolibImpressSettings.Settings(self.global_object)
        self.styles = ooolibImpressStyles.Styles(self.global_object)

    def export(self, filename):
        self.content.updateStats()
        f = ooolibFile.SaveFile(filename)
        f.insertFileString("mimetype", "application/vnd.oasis.opendocument.presentation")
        f.insertFileString("META-INF/manifest.xml", self.manifest.toString(indent=True))
        f.insertFileString("meta.xml", self.meta.toString())
        f.insertFileString("content.xml", self.content.toString())
        f.insertFileString("settings.xml", self.settings.toString())
        f.insertFileString("styles.xml", self.styles.toString())
        f.close()


##############################
# ooolib-python Draw Documents #
##############################
class Draw:
    """Draw - Open Document Format (ODF) drawing documents

    Base object for creating ODF drawing documents
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.manifest = ooolibDrawManifest.Manifest(self.global_object)
        self.meta = ooolibDrawMeta.Meta(self.global_object)
        self.content = ooolibDrawContent.Content(self.global_object)
        self.settings = ooolibDrawSettings.Settings(self.global_object)
        self.styles = ooolibDrawStyles.Styles(self.global_object)

    def export(self, filename):
        self.content.updateStats()
        f = ooolibFile.SaveFile(filename)
        f.insertFileString("mimetype", "application/vnd.oasis.opendocument.graphics")
        f.insertFileString("META-INF/manifest.xml", self.manifest.toString(indent=True))
        f.insertFileString("meta.xml", self.meta.toString())
        f.insertFileString("content.xml", self.content.toString())
        f.insertFileString("settings.xml", self.settings.toString())
        f.insertFileString("styles.xml", self.styles.toString())
        f.close()


###################
# Execute to Test #
###################
if __name__ == "__main__":
    # Make sure we can make a simple document
    calc = Calc()
    calc.export("blank.ods")
