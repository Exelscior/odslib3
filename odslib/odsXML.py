class Data:
    def __init__(self, data):
        self.type = 'Data'
        self.data = self.cleanString(data)

    def setData(self, data):
        self.data = self.cleanString(data)

    def getData(self):
        return self.data

    def cleanString(self, string):
        string = string.replace("&", "&amp;")
        string = string.replace('"', "&quot;")
        string = string.replace("'", "&apos;")
        string = string.replace("<", "&lt;")
        string = string.replace(">", "&gt;")
        return string

    def toString(self):
        return self.data


class Element:
    def __init__(self, element, data=None):
        self.type = 'Element'
        self.element = ''
        self.dataChild = None
        self.setElement(element)
        self.attributeValues = {}
        self.childrenList = []
        if not (data == None):
            self.dataChild = Data(data)
            self.addChild(self.dataChild)

    def cleanString(self, string):
        string = str(string)
        string = string.replace("&", "&amp;")
        string = string.replace('"', "&quot;")
        string = string.replace("'", "&apos;")
        string = string.replace("<", "&lt;")
        string = string.replace(">", "&gt;")
        return string

    def setElement(self, element):
        self.element = self.cleanString(element)

    def getElement(self):
        return self.element

    def setData(self, data):
        if not (self.dataChild == None):
            self.dataChild.setData(data)
            return self.dataChild
        return None

    def getData(self):
        if not (self.dataChild == None):
            return self.dataChild.getData()
        return None

    def setAttribute(self, name, value):
        self.attributeValues[name] = self.cleanString(value)

    def getAttribute(self, name):
        if name in self.attributeValues:
            return self.attributeValues[name]
        return None

    def removeAttribute(self, name):
        if name in self.attributeValues:
            self.attributeValues.pop(name)

    def getAttributeString(self):
        rlist = []
        rline = ''
        for key in self.attributeValues:
            value = self.attributeValues[key]
            rlist.append("%s='%s'" % (key, value))
        rline = ' '.join(rlist)
        return rline

    def addChild(self, child):
        self.childrenList.append(child)
        return child

    def removeChild(self, index):
        if index < len(self.childrenList):
            self.childrenList.pop(index)

    def removeChildren(self):
        self.childrenList = []

    def getChild(self, index):
        if index < len(self.childrenList):
            return self.childrenList[index]
        return None

    def childCount(self):
        return len(self.childrenList)

    def toString(self):
        # Open the element
        if self.attributeValues:
            # Display attributes
            rstring = '<%s %s' % (self.element, self.getAttributeString())
        else:
            rstring = '<%s' % self.element
        # Display children if they exist
        if self.childrenList:
            rstring += '>'
            for child in self.childrenList:
                rstring += child.toString()
            # close with children
            rstring += '</%s>' % self.element
        else:
            # close without children
            rstring += '/>'
        return rstring
    
