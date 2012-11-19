import time
from odsXML import *

class odsMeta:
    def __init__(self):
        self.docMeta = Element('office:document-meta')
        self.initialize()

    def toString(self):
        mstring = '<?xml version="1.0" encoding="UTF-8"?>\n'
        mstring += self.docMeta.toString()
        return mstring

    def initialize(self):
        # Create document
        self.docMeta.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.docMeta.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.docMeta.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.docMeta.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.docMeta.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.docMeta.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.docMeta.setAttribute("office:version", "1.2")
        self.docMeta.setAttribute("grddl:transformation", "http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl")

        # Create meta element
        meta = self.docMeta.addChild(Element('office:meta'))

        # Create meta values
        self.generator    = meta.addChild(Element("meta:generator", "odslib-0"))
        self.description  = meta.addChild(Element("dc:description", ""))
        self.subject      = meta.addChild(Element("dc:subject", ""))
        self.title        = meta.addChild(Element("dc:title", ""))

        # Initial Creator
        self.initCreator  = meta.addChild(Element("meta:initial-creator", ""))
        self.creationDate = meta.addChild(Element("meta:creation-date", self.getISO8601()))

        # Modifier
        self.creator      = meta.addChild(Element("dc:creator", ""))
        self.date         = meta.addChild(Element("dc:date", self.getISO8601()))

    def setTitle(self, title):
        self.title.setData(title)
        return self

    def setDescription(self, description):
        self.description.setData(description)
        return self

    def setSubject(self, subject):
        self.subject.setData(subject)
        return self

    def setCreator(self, name):
        self.initCreator.setData(name)
        return self

    def setModifier(self, name):
        self.creator.setData(name)
        return self

    def getISO8601(self):
        "Calculate and return localtime in ISO 8601 format"
        t = time.localtime()
        stamp = "%04d-%02d-%02dT%02d:%02d:%02d" % (t[0], # Year
                                                   t[1], # Month
                                                   t[2], # MDay
                                                   t[3], # Hour
                                                   t[4], # Minute
                                                   t[5]) # Second
        return stamp
