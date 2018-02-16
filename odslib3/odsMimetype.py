class odsMimetype:
    def __init__(self):
        self.mimetype = "application/vnd.oasis.opendocument.spreadsheet"

    def setMimetype(self, mimetype):
        self.mimetype = mimetype
        return self.mimetype

    def getMimetype(self):
        return self.mimetype

    def toString(self):
        return self.mimetype
