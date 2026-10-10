#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os
import zipfile

################################
# ooolib-python File Interface #
################################
class SaveFile:
    """SaveFile - Open Document Format (ODF) File Interface

    For writing ODF zip files.
    """
    def __init__(self, filename):
        self.filename = filename
        self.zipArchive = zipfile.ZipFile(self.filename, "w")
        
    def insertFileString(self, filename, string):
        string = string.encode("UTF-8")
        # ODF requires the mimetype file to be stored uncompressed
        if filename == "mimetype":
            self.zipArchive.writestr(filename, string, zipfile.ZIP_STORED)
        else:
            self.zipArchive.writestr(filename, string, zipfile.ZIP_DEFLATED)

    def insertFileBytes(self, filename, data, compress=True):
        # Already compressed files (PNG, JPEG, GIF) are stored as they are
        if compress:
            self.zipArchive.writestr(filename, data, zipfile.ZIP_DEFLATED)
        else:
            self.zipArchive.writestr(filename, data, zipfile.ZIP_STORED)

    def close(self):
        self.zipArchive.close()
        
class ReadFile:
    """ReadFile - Open Document Format (ODF) File Interface

    For reading ODF zip files.
    """
    def __init__(self, filename):
        self.filename = filename
        self.zipArchive = zipfile.ZipFile(self.filename, "r")
        self.contents = self.zipArchive.namelist()

    def hasFile(self, filename):
        return filename in self.contents

    def getFileList(self):
        return list(self.contents)

    def readFileBytes(self, filename, default=None):
        if not self.hasFile(filename): return default
        return self.zipArchive.read(filename)

    def readFileString(self, filename, default=None):
        data = self.readFileBytes(filename)
        if data is None: return default
        return data.decode("UTF-8")

    def getFileSize(self, filename, default=None):
        if not self.hasFile(filename): return default
        return self.zipArchive.getinfo(filename).file_size

    def close(self):
        self.zipArchive.close()

###################
# Execute to Test #
###################
if __name__ == "__main__":
    f = SaveFile("test.zip")
    f.insertFileString("hello.txt", "Hello World!")
    f.close()

