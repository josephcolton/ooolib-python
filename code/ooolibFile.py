#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
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
        self.zipArchive.writestr(filename, string, zipfile.ZIP_DEFLATED)

    def close(self):
        self.zipArchive.close()
        
class ReadFile:
    def __init__(self, filename):
        self.filename = filename
        self.zipArchive = zipfile.ZipFile(self.filename, "r")        
        self.contents = self.zipArchive.namelist()
        
    
###################
# Execute to Test #
###################
if __name__ == "__main__":
    f = SaveFile("test.zip")
    f.insertFileString("hello.txt", "Hello World!")
    f.close()

