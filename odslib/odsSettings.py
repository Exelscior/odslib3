from odsXML import *

class odsSettings:
    def __init__(self):
        self.docSettings = Element('office:document-settings')
        self.initialize()

    def toString(self):
        sstring = '<?xml version="1.0" encoding="UTF-8"?>\n'
        sstring += self.docSettings.toString()
        return sstring

    def makeConfigItem(self, configName, configType, data=None):
        item = Element("config:config-item", data)
        item.setAttribute("config:name", configName)
        item.setAttribute("config:type", configType)
        return item
		
    def initialize(self):
        # Create document
        self.docSettings.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.docSettings.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.docSettings.setAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0")
        self.docSettings.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.docSettings.setAttribute("office:version", "1.2")

        settings = self.docSettings.addChild(Element("office:settings"))
        viewSettings = settings.addChild(Element("config:config-item-set"))
        confSettings = settings.addChild(Element("config:config-item-set"))

        # View Settings
        viewSettings.addChild(self.makeConfigItem("VisibleAreaTop", "int", "0"))
        viewSettings.addChild(self.makeConfigItem("VisibleAreaLeft", "int", "0"))
        viewSettings.addChild(self.makeConfigItem("VisibleAreaWidth", "int", "2257"))
        viewSettings.addChild(self.makeConfigItem("VisibleAreaHeight", "int", "478"))

        # Views Item Map
        viewIndex = viewSettings.addChild(Element("config:config-item-map-indexed"))
        viewIndex.setAttribute("config:name", "Views")
        viewIndexItemMap = viewIndex.addChild(Element("config:config-item-map-entry"))
        viewIndexItemMap.addChild(self.makeConfigItem("ViewID", "string", "view1"))		

        # Table Sheet1 Configurations
        viewTables = viewIndexItemMap.addChild(Element("config:config-item-map-named"))
        viewTables.setAttribute("config:name", "Tables")
        tableSheet1 = viewTables.addChild(Element("config:config-item-map-entry"))
        tableSheet1.setAttribute("config:name", "Sheet1")
        tableSheet1.addChild(self.makeConfigItem("CursorPositionX", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("CursorPositionY", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("HorizontalSplitMode", "short", "0"))
        tableSheet1.addChild(self.makeConfigItem("VerticalSplitMode", "short", "0"))
        tableSheet1.addChild(self.makeConfigItem("HorizontalSplitPosition", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("VerticalSplitPosition", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("ActiveSplitRange", "short", "2"))
        tableSheet1.addChild(self.makeConfigItem("PositionLeft", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("PositionRight", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("PositionTop", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("PositionBottom", "int", "0"))
        tableSheet1.addChild(self.makeConfigItem("ZoomType", "short", "0"))
        tableSheet1.addChild(self.makeConfigItem("ZoomValue", "int", "100"))
        tableSheet1.addChild(self.makeConfigItem("PageViewZoomValue", "int", "60"))
        tableSheet1.addChild(self.makeConfigItem("ShowGrid", "boolean", "true"))

        viewIndexItemMap.addChild(self.makeConfigItem("ActiveTable", "string", "Sheet1"))		
        viewIndexItemMap.addChild(self.makeConfigItem("HorizontalScrollbarWidth", "int", "270"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ZoomType", "short", "0"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ZoomValue", "int", "100"))		
        viewIndexItemMap.addChild(self.makeConfigItem("PageViewZoomValue", "int", "60"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ShowPageBreakPreview", "boolean", "false"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ShowZeroValues", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ShowNotes", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ShowGrid", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("GridColor", "long", "12632256"))		
        viewIndexItemMap.addChild(self.makeConfigItem("ShowPageBreaks", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("HasColumnRowHeaders", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("HasSheetTabs", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("IsOutlineSymbolsSet", "boolean", "true"))		
        viewIndexItemMap.addChild(self.makeConfigItem("IsSnapToRaster", "boolean", "false"))		
        viewIndexItemMap.addChild(self.makeConfigItem("RasterIsVisible", "boolean", "false"))		
        viewIndexItemMap.addChild(self.makeConfigItem("RasterResolutionX", "int", "1270"))		
        viewIndexItemMap.addChild(self.makeConfigItem("RasterResolutionY", "int", "1270"))		
        viewIndexItemMap.addChild(self.makeConfigItem("RasterSubdivisionX", "int", "1"))		
        viewIndexItemMap.addChild(self.makeConfigItem("RasterSubdivisionY", "int", "1"))		
        viewIndexItemMap.addChild(self.makeConfigItem("IsRasterAxisSynchronized", "boolean", "true"))		

        # Configuration Settings
        confSettings.addChild(self.makeConfigItem("LoadReadonly", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("UpdateFromTemplate", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("PrinterSetup", "base64Binary", None))
        confSettings.addChild(self.makeConfigItem("AutoCalculate", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("IsDocumentShared", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("ShowNotes", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("HasSheetTabs", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("SaveVersionOnClose", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("RasterIsVisible", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("PrinterName", "string", None))
        confSettings.addChild(self.makeConfigItem("LinkUpdateMode", "short", "3"))
        confSettings.addChild(self.makeConfigItem("ApplyUserData", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("IsKernAsianPunctuation", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("RasterResolutionX", "int", "1270"))
        confSettings.addChild(self.makeConfigItem("IsRasterAxisSynchronized", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("RasterResolutionY", "int", "1270"))
        confSettings.addChild(self.makeConfigItem("IsOutlineSymbolsSet", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("ShowPageBreaks", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("ShowGrid", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("CharacterCompressionType", "short", "0"))
        confSettings.addChild(self.makeConfigItem("GridColor", "long", "12632256"))
        confSettings.addChild(self.makeConfigItem("ShowZeroValues", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("AllowPrintJobCancel", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("HasColumnRowHeaders", "boolean", "true"))
        confSettings.addChild(self.makeConfigItem("IsSnapToRaster", "boolean", "false"))
        confSettings.addChild(self.makeConfigItem("RasterSubdivisionX", "int", "1"))
        confSettings.addChild(self.makeConfigItem("RasterSubdivisionY", "int", "1"))
