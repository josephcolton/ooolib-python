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

#################################
# ooolib-python Settings object #
#################################
class Settings:
    def __init__(self, global_object):
        self.global_object = global_object
        # Create XML components
        self.prolog = ooolibXML.Prolog("xml")
        # Office Document Settings
        self.documentSettings = ooolibXML.Element("office:document-settings")
        self.documentSettings.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.documentSettings.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.documentSettings.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.documentSettings.setAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0")
        self.documentSettings.setAttribute("office:version", "1.3")
        # Office Settings
        self.settings = self.documentSettings.addChild(ooolibXML.Element("office:settings"))
        # View Settings - Item Set 1
        configSet1 = self.settings.addChild(ooolibXML.Element("config:config-item-set"))
        configSet1.setAttribute("config:name", "ooo:view-settings")
        self.addConfigItem(configSet1, "VisibleAreaTop", "int", "0")
        self.addConfigItem(configSet1, "VisibleAreaLeft", "int", "0")
        self.addConfigItem(configSet1, "VisibleAreaWidth", "int", "2258")
        self.addConfigItem(configSet1, "VisibleAreaHeight", "int", "452")
        configItemMapIndexed = configSet1.addChild(ooolibXML.Element("config:config-item-map-indexed"))
        configItemMapIndexed.setAttribute("config:name", "Views")
        configItemMapEntry1 = configItemMapIndexed.addChild(ooolibXML.Element("config:config-item-map-entry"))
        self.addConfigItem(configItemMapEntry1, "ViewId", "string", "view1")
        # View Settings - Item Set 1 - Tables
        configItemMapNamed1 = configItemMapEntry1.addChild(ooolibXML.Element("config:config-item-map-named"))
        configItemMapNamed1.setAttribute("config:name", "Tables")
        configItemMapEntry2 = configItemMapNamed1.addChild(ooolibXML.Element("config:config-item-map-entry"))
        configItemMapEntry2.setAttribute("config:name", "Sheet1")
        self.addConfigItem(configItemMapEntry2, "CursorPositionX", "int", "0")
        self.addConfigItem(configItemMapEntry2, "CursorPositionY", "int", "0")
        self.addConfigItem(configItemMapEntry2, "ActiveSplitRange", "short", "2")
        self.addConfigItem(configItemMapEntry2, "PositionLeft", "int", "0")
        self.addConfigItem(configItemMapEntry2, "PositionRight", "int", "0")
        self.addConfigItem(configItemMapEntry2, "PositionTop", "int", "0")
        self.addConfigItem(configItemMapEntry2, "PositionBottom", "int", "0")
        self.addConfigItem(configItemMapEntry2, "ZoomType", "short", "0")
        self.addConfigItem(configItemMapEntry2, "ZoomValue", "int", "100")
        self.addConfigItem(configItemMapEntry2, "PageViewZoomValue", "int", "60")
        self.addConfigItem(configItemMapEntry2, "ShowGrid", "boolean", "true")
        self.addConfigItem(configItemMapEntry2, "AnchoredTextOverflowLegacy", "boolean", "false")
        # View Settings - Item Set 1 - Continued
        self.addConfigItem(configItemMapEntry1, "ActiveTable", "string", "Sheet1")
        self.addConfigItem(configItemMapEntry1, "HorizontalScrollbarWidth", "int", "1839")
        self.addConfigItem(configItemMapEntry1, "ZoomType", "short", "0")
        self.addConfigItem(configItemMapEntry1, "ZoomValue", "int", "100")
        self.addConfigItem(configItemMapEntry1, "PageViewZoomValue", "int", "60")
        self.addConfigItem(configItemMapEntry1, "ShowPageBreakPreview", "boolean", "false")
        self.addConfigItem(configItemMapEntry1, "ShowZeroValues", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "ShowNotes", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "ShowGrid", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "GridColor", "int", "12632256") # C0C0C0
        self.addConfigItem(configItemMapEntry1, "ShowPageBreaks", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "HasColumnRowHeaders", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "HasSheetTabs", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "IsOutlineSymbolsSet", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "IsValueHighlightingEnabled", "boolean", "false")
        self.addConfigItem(configItemMapEntry1, "IsSnapToRaster", "boolean", "false")
        self.addConfigItem(configItemMapEntry1, "RasterIsVisible", "boolean", "false")
        self.addConfigItem(configItemMapEntry1, "RasterResolutionX", "int", "1270")
        self.addConfigItem(configItemMapEntry1, "RasterResolutionY", "int", "1270")
        self.addConfigItem(configItemMapEntry1, "RasterSubdivisionX", "int", "1")
        self.addConfigItem(configItemMapEntry1, "RasterSubdivisionY", "int", "1")
        self.addConfigItem(configItemMapEntry1, "IsRasterAxisSynchronized", "boolean", "true")
        self.addConfigItem(configItemMapEntry1, "AnchoredTextOverflowLegacy", "boolean", "false")        
        # Configuration Settings - Item Set 2
        configSet2 = self.settings.addChild(ooolibXML.Element("config:config-item-set"))
        self.addConfigItem(configSet2, "AllowPrintJobCancel", "boolean", "true")
        self.addConfigItem(configSet2, "ApplyUserData", "boolean", "true")
        self.addConfigItem(configSet2, "AutoCalculate", "boolean", "true")
        self.addConfigItem(configSet2, "CharacterCompressionType", "short", "0")
        self.addConfigItem(configSet2, "EmbedAsianScriptFonts", "boolean", "true")
        self.addConfigItem(configSet2, "EmbedComplexScriptFonts", "boolean", "true")
        self.addConfigItem(configSet2, "EmbedFonts", "boolean", "false")
        self.addConfigItem(configSet2, "EmbedLatinScriptFonts", "boolean", "true")
        self.addConfigItem(configSet2, "EmbedOnlyUsedFonts", "boolean", "false")
        self.addConfigItem(configSet2, "GridColor", "int", "12632256")
        self.addConfigItem(configSet2, "HasColumnRowHeaders", "boolean", "true")
        self.addConfigItem(configSet2, "HasSheetTabs", "boolean", "true")
        self.addConfigItem(configSet2, "IsDocumentShared", "boolean", "false")
        self.addConfigItem(configSet2, "IsKernAsianPunctuation", "boolean", "false")
        self.addConfigItem(configSet2, "IsOutlineSymbolsSet", "boolean", "true")
        self.addConfigItem(configSet2, "IsRasterAxisSynchronized", "boolean", "true")
        self.addConfigItem(configSet2, "IsSnapToRaster", "boolean", "false")
        self.addConfigItem(configSet2, "LinkUpdateMode", "short", "3")
        self.addConfigItem(configSet2, "LoadReadonly", "boolean", "false")
        self.addConfigItem(configSet2, "PrinterName", "string", None)
        self.addConfigItem(configSet2, "PrinterPaperFromSetup", "boolean", "false")
        self.addConfigItem(configSet2, "PrinterSetup", "base64Binary", None)
        self.addConfigItem(configSet2, "RasterIsVisible", "boolean", "false")
        self.addConfigItem(configSet2, "RasterResolutionX", "int", "1270")
        self.addConfigItem(configSet2, "RasterResolutionY", "int", "1270")
        self.addConfigItem(configSet2, "RasterSubdivisionX", "int", "1")
        self.addConfigItem(configSet2, "RasterSubdivisionY", "int", "1")
        self.addConfigItem(configSet2, "SaveThumbnail", "boolean", "true")
        self.addConfigItem(configSet2, "SaveVersionOnClose", "boolean", "false")
        self.addConfigItem(configSet2, "ShowGrid", "boolean", "true")
        self.addConfigItem(configSet2, "ShowNotes", "boolean", "true")
        self.addConfigItem(configSet2, "ShowPageBreaks", "boolean", "true")
        self.addConfigItem(configSet2, "ShowZeroValues", "boolean", "true")
        self.addConfigItem(configSet2, "UpdateFromTemplate", "boolean", "true")
        # Configuration Settings - Item Set 2 - Sheet1
        configItemMapNamed2 = configSet2.addChild(ooolibXML.Element("config:config-item-map-named"))
        configItemMapNamed2.setAttribute("config:name", "ScriptConfiguration")
        configItemMapEntry3 = configItemMapNamed2.addChild(ooolibXML.Element("config:config-item-map-entry"))
        configItemMapEntry3.setAttribute("config:name", "Sheet1")
        self.addConfigItem(configItemMapEntry3, "CodeName", "string", "Sheet1")
        
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

    # Display meta file
    string1 = settings.toString(indent=True)

    print("Settings XML with indentation:")
    print(string1)
