from odsXML import *

class odsStyles:
    def __init__(self):
        self.docStyles = Element('office:document-styles')
        self.initialize()

    def toString(self):
        sstring = '<?xml version="1.0" encoding="UTF-8"?>\n'
        sstring += self.docStyles.toString()
        return sstring

    def initialize(self):
        # Create document
        self.docStyles.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.docStyles.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.docStyles.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.docStyles.setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
        self.docStyles.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.docStyles.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.docStyles.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.docStyles.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.docStyles.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.docStyles.setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
        self.docStyles.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
        self.docStyles.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.docStyles.setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
        self.docStyles.setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")
        self.docStyles.setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML")
        self.docStyles.setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")
        self.docStyles.setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")
        self.docStyles.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.docStyles.setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer")
        self.docStyles.setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc")
        self.docStyles.setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events")
        self.docStyles.setAttribute("xmlns:rpt", "http://openoffice.org/2005/report")
        self.docStyles.setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
        self.docStyles.setAttribute("xmlns:xhtml", "http://www.w3.org/1999/xhtml")
        self.docStyles.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.docStyles.setAttribute("xmlns:tableooo", "http://openoffice.org/2009/table")
        self.docStyles.setAttribute("xmlns:css3t", "http://www.w3.org/TR/css3-text/")
        self.docStyles.setAttribute("office:version", "1.2")
