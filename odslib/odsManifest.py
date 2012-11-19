from odsXML import *

class odsManifest:
    def __init__(self):
        self.docManifest = Element('manifest:manifest')
        self.initialize()

    def toString(self):
        mstring = '<?xml version="1.0" encoding="UTF-8"?>\n'
        mstring += self.docManifest.toString()
        return mstring

    def initialize(self):
        # Create document
        self.docManifest.setAttribute("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")
        self.docManifest.setAttribute("manifest:version", "1.2")

        # Manifest contents
        version = self.docManifest.addChild(Element("manifest:file-entry"))
        meta = self.docManifest.addChild(Element("manifest:file-entry"))
        settings = self.docManifest.addChild(Element("manifest:file-entry"))
        content = self.docManifest.addChild(Element("manifest:file-entry"))
        current = self.docManifest.addChild(Element("manifest:file-entry"))
        config2 = self.docManifest.addChild(Element("manifest:file-entry"))
        styles = self.docManifest.addChild(Element("manifest:file-entry"))
		
        # Version line
        version.setAttribute("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet")
        version.setAttribute("manifest:version", "1.2")
        version.setAttribute("manifest:full-path", "/")

        # Meta line
        meta.setAttribute("manifest:media-type", "text/xml")
        meta.setAttribute("manifest:full-path", "meta.xml")

        # Settings line
        settings.setAttribute("manifest:media-type", "text/xml")
        settings.setAttribute("manifest:full-path", "settings.xml")

        # Content line
        content.setAttribute("manifest:media-type", "text/xml")
        content.setAttribute("manifest:full-path", "content.xml")

        # Current line
        current.setAttribute("manifest:media-type", "")
        current.setAttribute("manifest:full-path", "Configurations2/accelerator/current.xml")

        # Config2 line
        config2.setAttribute("manifest:media-type", "application/vnd.sun.xml.ui.configuration")
        config2.setAttribute("manifest:full-path", "Configurations2/")

        # Styles line
        styles.setAttribute("manifest:media-type", "text/xml")
        styles.setAttribute("manifest:full-path", "styles.xml")
