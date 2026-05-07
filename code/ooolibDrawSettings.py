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

####################################
# ooolib-python Draw Settings object #
####################################
class Settings:
    def __init__(self, global_object):
        self.global_object = global_object
        self.prolog = ooolibXML.Prolog("xml")
        self.documentSettings = ooolibXML.Element("office:document-settings")
        self.documentSettings.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentSettings.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentSettings.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentSettings.setAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0")
        self.documentSettings.setAttribute("office:version", "1.3")
        self.settings = self.documentSettings.addChild(ooolibXML.Element("office:settings"))
        configSet = self.settings.addChild(ooolibXML.Element("config:config-item-set"))
        configSet.setAttribute("config:name", "ooo:configuration-settings")
        self.addConfigItem(configSet, "AllowPrintJobCancel", "boolean", "true")
        self.addConfigItem(configSet, "ApplyUserData", "boolean", "true")
        self.addConfigItem(configSet, "CharacterCompressionType", "short", "0")
        self.addConfigItem(configSet, "EmbedAsianScriptFonts", "boolean", "true")
        self.addConfigItem(configSet, "EmbedComplexScriptFonts", "boolean", "true")
        self.addConfigItem(configSet, "EmbedFonts", "boolean", "false")
        self.addConfigItem(configSet, "EmbedLatinScriptFonts", "boolean", "true")
        self.addConfigItem(configSet, "EmbedOnlyUsedFonts", "boolean", "false")
        self.addConfigItem(configSet, "IsSnapToRaster", "boolean", "false")
        self.addConfigItem(configSet, "LoadReadonly", "boolean", "false")
        self.addConfigItem(configSet, "PrinterName", "string", None)
        self.addConfigItem(configSet, "PrinterPaperFromSetup", "boolean", "false")
        self.addConfigItem(configSet, "PrinterSetup", "base64Binary", None)
        self.addConfigItem(configSet, "RasterIsVisible", "boolean", "false")
        self.addConfigItem(configSet, "SaveVersionOnClose", "boolean", "false")
        self.addConfigItem(configSet, "UpdateFromTemplate", "boolean", "true")

    def addConfigItem(self, element, configName, configType, value=None):
        child = element.addChild(ooolibXML.Element("config:config-item", value))
        child.setAttribute("config:name", configName)
        child.setAttribute("config:type", configType)
        return child

    def toString(self, indent=False):
        settingsString = self.prolog.toString()
        settingsString += self.documentSettings.toString(indent)
        return settingsString


###################
# Execute to Test #
###################
if __name__ == "__main__":
    g = ooolibGlobal.Global()
    settings = Settings(g)
    string1 = settings.toString(indent=True)
    print("Draw Settings XML with indentation:")
    print(string1)
