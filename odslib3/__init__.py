import zipfile
import time
import os, re
from io import open

from . import odsContent
from . import odsMeta
from . import odsSettings
from . import odsStyles
from . import odsMimetype
from . import odsManifest

HERE = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(HERE, 'VERSION.txt'), encoding='utf-8') as f:
        VERSION = f.read()

def get_version():
    "Returns a PEP 386-compliant version number from VERSION."
    return VERSION

class ODS:
    def __init__(self):
        # Create instances
        self.content = odsContent.odsContent()
        self.meta = odsMeta.odsMeta()
        self.mimetype = odsMimetype.odsMimetype()
        self.settings = odsSettings.odsSettings()
        self.styles = odsStyles.odsStyles()
        self.manifest = odsManifest.odsManifest()

    def save(self, filename):
        self.savefile = zipfile.ZipFile(filename, "w")
        self._zip_insert(self.savefile, "meta.xml", self.meta.toString())
        self._zip_insert(self.savefile, "mimetype", self.mimetype.toString())
        self._zip_insert(self.savefile, "Configurations2/accelerator/current.xml", "")
        self._zip_insert(self.savefile, "content.xml", self.content.toString())
        self._zip_insert(self.savefile, "settings.xml", self.settings.toString())
        self._zip_insert(self.savefile, "styles.xml", self.styles.toString())
        self._zip_insert(self.savefile, "META-INF/manifest.xml", self.manifest.toString())
        self.savefile.close()

    def _zip_insert(self, file, filename, data):
        "Insert a file into the zip archive"

        # zip seems to struggle with non-ascii characters
        data = data.encode('utf-8')

        now = time.localtime(time.time())[:6]
        info = zipfile.ZipInfo(filename)
        info.date_time = now
        info.compress_type = zipfile.ZIP_DEFLATED
        file.writestr(info, data)
