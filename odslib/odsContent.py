from odsXML import *

class sheetCell:
    def __init__(self):
        self.cell = Element("table:table-cell")
        self.styleBold = False
        self.styleItalic = False
        self.styleUnderline = False
        self.styleFontColor = ""
        self.styleCellColor = ""
        self.styleFontSize = ""
        self.styleAlignVertical = ""
        self.styleAlignHorizontal = ""
        self.styleRotation = 0
        # Cell Borders
        self.styleBorder = ""
        self.styleBorderWidth = "0.0008in"
        self.styleBorderStyle = "solid"
        self.styleBorderColor = "#000000"

    # Cell Styles
    def setStyle(self, style):
        self.cell.setAttribute("table:style-name", style)
        return self

    def setBold(self, value):
        self.styleBold = value
        return self

    def getBold(self):
        return self.styleBold

    def setItalic(self, value):
        self.styleItalic = value
        return self

    def getItalic(self):
        return self.styleItalic

    def setUnderline(self, value):
        self.styleUnderline = value
        return self

    def getUnderline(self):
        return self.styleUnderline

    # Cell Colors
    def setFontColor(self, color):
        self.styleFontColor = color
        return self

    def getFontColor(self):
        return self.styleFontColor

    def setCellColor(self, color):
        self.styleCellColor = color
        return self

    def getCellColor(self):
        return self.styleCellColor

    # Cell Fonts
    def setFontSize(self, size):
        self.styleFontSize = size
        return self

    def getFontSize(self):
        return self.styleFontSize

    # Cell Alignment and Orientation
    def setAlignVertical(self, align):
        if (align == "center"): align = "middle"
        # Make sure we have a valid alignment
        if (align not in ["top", "middle", "bottom", "justify"]): return self
        self.styleAlignVertical = align
        return self

    def setAlignHorizontal(self, align):
        if (align == "right"): align = "end"
        if (align == "left"): align = "start"
        # Make sure we have a valid alignment
        if (align not in ["start", "center", "end", "justify"]): return self
        self.styleAlignHorizontal = align
        return self

    def getAlignVertical(self):
        return self.styleAlignVertical

    def getAlignHorizontal(self):
        return self.styleAlignHorizontal

    def setRotation(self, value):
        self.styleRotation = value
        return self

    def getRotation(self):
        return self.styleRotation

    # Cell Borders
    def setBorderWidth(self, value):
        self.styleBorderWidth = value
        self.styleBorder = "%s %s %s" % (self.styleBorderWidth, self.styleBorderStyle, self.styleBorderColor)
        return self

    def setBorderStyle(self, value):
        self.styleBorderStyle = value
        self.styleBorder = "%s %s %s" % (self.styleBorderWidth, self.styleBorderStyle, self.styleBorderColor)
        return self

    def setBorderColor(self, value):
        self.styleBorderColor = value
        self.styleBorder = "%s %s %s" % (self.styleBorderWidth, self.styleBorderStyle, self.styleBorderColor)
        return self

    def setBorder(self, value=True):
        if value:
            self.styleBorder = "%s %s %s" % (self.styleBorderWidth, self.styleBorderStyle, self.styleBorderColor)
        else:
            self.styleBorder = ""
        return self
    
    def getBorder(self):
        return self.styleBorder

    # Cell Values
    def floatValue(self, value):
        value = unicode(value)
        self.cell.setAttribute("office:value-type", "float")
        self.cell.setAttribute("office:value", "%s" % value)
        self.cell.addChild(Element("text:p", value))
        return self

    def stringValue(self, value):
        value = unicode(value)
        self.cell.setAttribute("office:value-type", "string")
        self.cell.addChild(Element("text:p", value))
        return self

    def floatFormula(self, value, formula):
        value = unicode(value)
        self.cell.setAttribute("office:value-type", "float")
        self.cell.setAttribute("office:value", "%s" % value)
        self.cell.addChild(Element("text:p", value))
        self.cell.setAttribute("table:formula", "of:%s" % formula)
        return self

    # Cell hiding
    def setCover(self, covered):
        if covered:
            self.cell.setElement("table:covered-table-cell")
        else:
            self.cell.setElement("table:table-cell")
        return self

    # Cell XML
    def toString(self):
        return self.cell.toString()	

class sheetColumn:
    def __init__(self):
        self.column = Element("table:table-column")
        self.setStyle("co1")
        self.column.setAttribute("table:default-cell-style-name", "Default")
        # Column values
        self.width = "0.8925in"

    def setStyle(self, style):
        self.column.setAttribute("table:style-name", style)

    def setWidth(self, width):
        self.width = width
        return self

    def getWidth(self):
        return self.width

    def setHidden(self, hide):
        if hide:
            self.column.setAttribute("table:visibility", "collapse")
        else:
            self.column.removeAttribute("table:visibility")
        return self

    def toString(self):
        return self.column.toString()	

class sheetRow:
    def __init__(self):
        self.row = Element("table:table-row")
        self.setStyle("ro1")
        # Column values
        self.height = "0.1783in"

    def setStyle(self, style):
        self.row.setAttribute("table:style-name", style)

    def setHeight(self, height):
        self.height = height
        return self

    def getHeight(self):
        return self.height

    def removeChildren(self):
        self.row.removeChildren()

    def addChild(self, child):
        self.row.addChild(child)

    def setHidden(self, hide):
        if hide:
            self.row.setAttribute("table:visibility", "collapse")
        else:
            self.row.removeAttribute("table:visibility")
        return self

    def toString(self):
        return self.row.toString()	

class sheetTable:
    def __init__(self, sheetName):
        self.sheet = Element("table:table")
        self.sheet.setAttribute("table:name", sheetName)
        self.sheet.setAttribute("table:style-name", "ta1")
        self.sheet.setAttribute("table:print", "false")

        # Sheet variables
        self.maxColumn = 0
        self.maxRow = 0
        self.sheetRows = []
        self.sheetColumns = []
        self.cellObjects = {}

        # Initial Cell
        self.makeCell(0, 0)        

    def setStyle(self, style):
        self.sheet.setAttribute("table:style-name", style)

    def updateColumns(self, columnNum):
        while (columnNum >= len(self.sheetColumns)):
            column = self.sheet.addChild(sheetColumn())
            self.sheetColumns.append(column)
        if columnNum > self.maxColumn: self.maxColumn = columnNum
        return self.sheetColumns[columnNum]

    def updateRows(self, rowNum):
        # Make sure all rows exist
        while (rowNum >= len(self.sheetRows)):
            row = self.sheet.addChild(sheetRow())
            self.sheetRows.append(row)
        # Update the maxRow counter
        if rowNum > self.maxRow: self.maxRow = rowNum
        # Rebuild list of children
        row = self.sheetRows[rowNum]
        row.removeChildren()
        for columnNum in range(0, self.maxColumn + 1):
            row.addChild(self.getCell(columnNum, rowNum))
        return row
		
    def setSheetName(self, sheetName):
        self.sheet.setAttribute("table:name", sheetName)

    def getSheetName(self):
        return self.sheet.getAttribute("table:name")

    def getColumn(self, columnNum):
        return self.updateColumns(columnNum)

    def getColumnIndexList(self):
        return range(0, self.maxColumn + 1)

    def getRow(self, rowNum):
        return self.updateRows(rowNum)

    def getRowIndexList(self):
        return range(0, self.maxRow + 1)

    def setSheetStyle(self, styleID):
        self.sheet.setAttribute("table:style-name", styleID)

    def getSheetStyle(self):
        return self.sheet.getAttribute("table:style-name")

    def makeCell(self, columnNum, rowNum):
        self.cellObjects[(columnNum, rowNum)] = sheetCell()
        self.updateColumns(columnNum)
        self.updateRows(rowNum)
        return self.cellObjects[(columnNum, rowNum)]

    def getCell(self, columnNum, rowNum):
        # See if the cell exists
        cellIndex = (columnNum, rowNum)
        if cellIndex in self.cellObjects:
            cell = self.cellObjects[cellIndex]
        else:
            cell = self.makeCell(columnNum, rowNum)
        return cell
		
    def updateObjects(self):
        self.sheet.removeChildren()
        # Add Columns
        for column in self.sheetColumns:
            self.sheet.addChild(column)
        # Add Rows
        for rowNum in range(0, self.maxRow + 1):
            row = self.updateRows(rowNum)
            self.sheet.addChild(row)

    def toString(self):
        self.updateObjects()
        return self.sheet.toString()

class odsContentStyles:
    def __init__(self, content):
        self.content = content
        self.autostyles = Element("office:automatic-styles")
        self.resetData()

    def resetData(self):
        # Remove existing data
        self.autostyles.removeChildren()
        # Style counters
        self.columnIndex = 1
        self.rowIndex = 1
        self.cellIndex = 1
        self.tableIndex = 1

    def buildColumnStyles(self):
        # Save values
        currentSheet = self.content.getSheetIndex()
        # Build a list of attributes
        columnAttributes = {}
        for sheetIndex in self.content.getSheetIndexList():
            sheet = self.content.getSheet(sheetIndex)
            for columnNum in sheet.getColumnIndexList():
                column = self.content.getColumn(columnNum)
                width = column.getWidth()
                attribTuple = (width, )
                # Insert tuple into list
                if not attribTuple in columnAttributes:
                    columnAttributes[attribTuple] = []
                columnAttributes[attribTuple].append(column)
        # Build needed automatic styles
        for attribTuple in columnAttributes:
            # Get attributes
            width = attribTuple[0]
            # Create the style
            styleID = "co%d" % self.columnIndex
            self.columnIndex += 1
            # Create the elements
            style = self.autostyles.addChild(Element("style:style"))
            style.setAttribute("style:name", styleID)
            style.setAttribute("style:family", "table-column")
            styleProp = style.addChild(Element("style:table-column-properties"))
            styleProp.setAttribute("fo:break-before", "auto")
            styleProp.setAttribute("style:column-width", width)
            # Set all columns to use the new style
            for column in columnAttributes[attribTuple]:
                column.setStyle(styleID)
        # Restore values
        self.content.getSheet(currentSheet)

    def buildRowStyles(self):
        # Save values
        currentSheet = self.content.getSheetIndex()
        # Build a list of attributes
        rowAttributes = {}
        for sheetIndex in self.content.getSheetIndexList():
            sheet = self.content.getSheet(sheetIndex)
            for rowNum in sheet.getRowIndexList():
                row = self.content.getRow(rowNum)
                height = row.getHeight()
                attribTuple = (height, )
                # Insert tuple into list
                if not attribTuple in rowAttributes:
                    rowAttributes[attribTuple] = []
                rowAttributes[attribTuple].append(row)
        # Build needed automatic styles
        for attribTuple in rowAttributes:
            # Get attributes
            height = attribTuple[0]
            # Create the style
            styleID = "ro%d" % self.rowIndex
            self.rowIndex += 1
            # Create the elements
            style = self.autostyles.addChild(Element("style:style"))
            style.setAttribute("style:name", styleID)
            style.setAttribute("style:family", "table-row")
            styleProp = style.addChild(Element("style:table-row-properties"))
            styleProp.setAttribute("fo:break-before", "auto")
            styleProp.setAttribute("style:row-height", height)
			# If the default height is used, we want to use the following line
            # styleProp.setAttribute("style:use-optimal-row-height", "true")
            # Set all rows to use the new style
            for row in rowAttributes[attribTuple]:
                row.setStyle(styleID)
        # Restore values
        self.content.getSheet(currentSheet)

    def buildCellStyles(self):
        # Save values
        currentSheet = self.content.getSheetIndex()
        # Build a list of attributes
        cellAttributes = {}
        for sheetIndex in self.content.getSheetIndexList():
            sheet = self.content.getSheet(sheetIndex)
            for c in sheet.getColumnIndexList():
                for r in sheet.getRowIndexList():
                    cell = self.content.getCell(c, r)
                    # Cell attributes
                    bold = cell.getBold()
                    italic = cell.getItalic()
                    underline = cell.getUnderline()
                    fontColor = cell.getFontColor()
                    cellColor = cell.getCellColor()
                    fontSize = cell.getFontSize()
                    VAlign = cell.getAlignVertical()
                    HAlign = cell.getAlignHorizontal()
                    rotation = cell.getRotation()
                    border = cell.getBorder()
                    attribTuple = (bold, italic, underline, fontColor, cellColor, fontSize,
                                   VAlign, HAlign, rotation, border)
                    # Insert tuple into list
                    if not attribTuple in cellAttributes:
                        cellAttributes[attribTuple] = []
                    cellAttributes[attribTuple].append(cell)
        # Build needed automatic styles
        for attribTuple in cellAttributes:
            # Get attributes
            bold = attribTuple[0]
            italic = attribTuple[1]
            underline = attribTuple[2]
            fontColor = attribTuple[3]
            cellColor = attribTuple[4]
            fontSize = attribTuple[5]
            VAlign = attribTuple[6]
            HAlign = attribTuple[7]
            rotation = attribTuple[8]
            border = attribTuple[9]
            # Create the style
            styleID = "ce%d" % self.cellIndex
            self.cellIndex += 1
            # Create the elements
            style = self.autostyles.addChild(Element("style:style"))
            style.setAttribute("style:name", styleID)
            style.setAttribute("style:family", "table-cell")
            style.setAttribute("style:parent-style-name", "Default")
            # Cell Properties
            styleCellProp = style.addChild(Element("style:table-cell-properties"))
            if cellColor:
                styleCellProp.setAttribute("fo:background-color", cellColor)
            if VAlign:
                styleCellProp.setAttribute("style:vertical-align", VAlign)
                styleCellProp.setAttribute("style:repeat-content", "false")
                styleCellProp.setAttribute("style:text-align-source", "fix")
            if rotation:
                styleCellProp.setAttribute("style:rotation-angle", rotation)
            if border:
                styleCellProp.setAttribute("fo:border", border)
            # Cell Paragraph Properties
            styleParaProp = style.addChild(Element("style:paragraph-properties"))
            if HAlign:
                styleParaProp.setAttribute("fo:margin-left", "0in")
                styleParaProp.setAttribute("fo:text-align", HAlign)
            # Cell Text Properties
            styleTextProp = style.addChild(Element("style:text-properties"))
            if bold:
                styleTextProp.setAttribute("fo:font-weight", "bold")
                styleTextProp.setAttribute("style:font-weight-asian", "bold")
                styleTextProp.setAttribute("style:font-weight-complex", "bold")
            if italic:
                styleTextProp.setAttribute("fo:font-style", "italic")
                styleTextProp.setAttribute("style:font-style-asian", "italic")
                styleTextProp.setAttribute("style:font-style-complex", "italic")
            if underline:
                styleTextProp.setAttribute("style:text-underline-style", "solid")
                styleTextProp.setAttribute("style:text-underline-width", "auto")
                styleTextProp.setAttribute("style:text-underline-color", "font-color")
            if fontColor:
                styleTextProp.setAttribute("fo:color", fontColor)
            if fontSize:
                styleTextProp.setAttribute("fo:font-size", fontSize)
                styleTextProp.setAttribute("style:font-size-asian", fontSize)
                styleTextProp.setAttribute("style:font-size-complex", fontSize)
            # Set all cells to use the new style
            for cell in cellAttributes[attribTuple]:
                cell.setStyle(styleID)
        # Restore values
        self.content.getSheet(currentSheet)

    def buildTableStyles(self):
        style = self.autostyles.addChild(Element("style:style"))
        style.setAttribute("style:name", "ta1")
        style.setAttribute("style:family", "table")
        style.setAttribute("style:master-page-name", "Default")
        styleProp = style.addChild(Element("style:table-properties"))
        styleProp.setAttribute("table:display", "true")
        styleProp.setAttribute("style:writing-mode", "lr-tb")
        
    def toString(self):
        self.resetData()
        self.buildColumnStyles()
        self.buildRowStyles()
        self.buildCellStyles()
        self.buildTableStyles()
        return self.autostyles.toString()
	
class odsContent:
    def __init__(self):
        self.docContent = Element('office:document-content')
        self.initialize()
        self.currentSheet = 0
		
    def toString(self):
        cstring = '<?xml version="1.0" encoding="UTF-8"?>\n'
        cstring += self.docContent.toString()
        return cstring

    def getSheet(self, sheetIndex):
        self.currentSheet = sheetIndex
        return self.documentSheets.getChild(sheetIndex)

    def getSheetIndex(self):
        return self.currentSheet

    def getSheetIndexList(self):
        return range(0, self.documentSheets.childCount())

    def makeSheet(self, sheetName):
        self.currentSheet = self.documentSheets.childCount()
        return self.documentSheets.addChild(sheetTable(sheetName))

    def getColumn(self, columnNum):
        sheet = self.getSheet(self.currentSheet)
        if not (sheet == None):
            return sheet.getColumn(columnNum)
        return None

    def getRow(self, rowNum):
        sheet = self.getSheet(self.currentSheet)
        if not (sheet == None):
            return sheet.getRow(rowNum)
        return None

    def getCell(self, columnNum, rowNum):
        sheet = self.getSheet(self.currentSheet)
        if not (sheet == None):
            return sheet.getCell(columnNum, rowNum)
        return None

    def mergeCells(self, columnNum, rowNum, colSpan=1, rowSpan=1):
        "This merge hides later cells, use with caution."
        cell = self.getCell(columnNum, rowNum)
        # Configure the start cell
        cell.cell.setAttribute("table:number-columns-spanned", colSpan)
        cell.cell.setAttribute("table:number-rows-spanned", rowSpan)
        # Hide the covered cells
        for c in range(int(columnNum), int(columnNum) + colSpan + 1):
            for r in range(int(rowNum), int(rowNum) + rowSpan + 1):
                if (c == columnNum and r == rowNum):
                    self.getCell(c, r).setCover(False)
                else:
                    self.getCell(c, r).setCover(True)
        return cell

    def initialize(self):
        # Create document
        self.docContent.setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
        self.docContent.setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
        self.docContent.setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
        self.docContent.setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
        self.docContent.setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
        self.docContent.setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
        self.docContent.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink")
        self.docContent.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
        self.docContent.setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
        self.docContent.setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
        self.docContent.setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
        self.docContent.setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
        self.docContent.setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
        self.docContent.setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")
        self.docContent.setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML")
        self.docContent.setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0")
        self.docContent.setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0")
        self.docContent.setAttribute("xmlns:ooo", "http://openoffice.org/2004/office")
        self.docContent.setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer")
        self.docContent.setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc")
        self.docContent.setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events")
        self.docContent.setAttribute("xmlns:xforms", "http://www.w3.org/2002/xforms")
        self.docContent.setAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        self.docContent.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        self.docContent.setAttribute("xmlns:rpt", "http://openoffice.org/2005/report")
        self.docContent.setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
        self.docContent.setAttribute("xmlns:xhtml", "http://www.w3.org/1999/xhtml")
        self.docContent.setAttribute("xmlns:grddl", "http://www.w3.org/2003/g/data-view#")
        self.docContent.setAttribute("xmlns:tableooo", "http://openoffice.org/2009/table")
        self.docContent.setAttribute("xmlns:field", "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")
        self.docContent.setAttribute("office:version", "1.2")
        self.docContent.setAttribute("grddl:transformation", "http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl")

        # Create document subsections
        contentScripts = self.docContent.addChild(Element("office:scripts"))
        contentFonts   = self.docContent.addChild(Element("office:font-face-decls"))
        contentStyles  = self.docContent.addChild(odsContentStyles(self))
        contentBody    = self.docContent.addChild(Element("office:body"))
        
        # Content Fonts
        font1 = contentFonts.addChild(Element("style:font-face"))
        font1.setAttribute("style:name", "Arial")
        font1.setAttribute("svg:font-family", "Arial")
        font1.setAttribute("style:font-family-generic", "swiss")
        font1.setAttribute("style:font-pitch", "variable")
        
        font2 = contentFonts.addChild(Element("style:font-face"))
        font2.setAttribute("style:name", "Arial Unicode MS")
        font2.setAttribute("svg:font-family", "&apos;Arial Unicode MS&apos;")
        font2.setAttribute("style:font-family-generic", "system")
        font2.setAttribute("style:font-pitch", "variable")
        
        font3 = contentFonts.addChild(Element("style:font-face"))
        font3.setAttribute("style:name", "Mangal")
        font3.setAttribute("svg:font-family", "Mangal")
        font3.setAttribute("style:font-family-generic", "system")
        font3.setAttribute("style:font-pitch", "variable")
        
        font4 = contentFonts.addChild(Element("style:font-face"))
        font4.setAttribute("style:name", "Microsoft YaHei")
        font4.setAttribute("svg:font-family", "&apos;Microsoft YaHei&apos;")
        font4.setAttribute("style:font-family-generic", "system")
        font4.setAttribute("style:font-pitch", "variable")
        
        font5 = contentFonts.addChild(Element("style:font-face"))
        font5.setAttribute("style:name", "Tahoma")
        font5.setAttribute("svg:font-family", "Tahoma")
        font5.setAttribute("style:font-family-generic", "system")
        font5.setAttribute("style:font-pitch", "variable")
        
        # Document Body Spreadsheets
        self.documentSheets = contentBody.addChild(Element("office:spreadsheet"))
        
        # First Spreadsheet
        self.documentSheets.addChild(sheetTable("Sheet1"))
