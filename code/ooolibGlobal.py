#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

class Global:
    """Global - ooolib global objects and variables

    Used to keep track of global variables and objects
    """
    def __init__(self):
        self.project = "ooolib-python"
        self.version = "1.2"

    def getVersion(self):
        versionStr = "%s %s" % (self.project, self.version)
        return versionStr
