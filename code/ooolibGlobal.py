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
        self.version = "0.2.0"
        self.globalInts = {}
        self.globalStrings = {}
        self.globalObjects = {}

    def getVersion(self):
        versionStr = "%s %s" % (self.project, self.version)
        return versionStr

    ###################
    # Global Integers #
    ###################
    def getGlobalInt(self, name, setDefault=0):
        #print("getGlobalInt(%s)" % (name))
        # Sanitize
        setDefault = int(setDefault)
        # Simple get if existing
        if name in self.globalInts:
            return self.globalInts[name]
        # Missing case
        else:
            self.globalInts[name] = setDefault
            return setDefault

    def setGlobalInt(self, name, value):
        #print("setGlobalInt(%s, %s)" % (name, value))
        value = int(value)             # Sanitize value
        self.globalInts[name] = value # Set value
        return value                   # Return set value

    def incrementGlobalInt(self, name, incValue=1):
        incValue = int(incValue)                  # Sanitize incValue
        value = self.getGlobalInt(name, incValue) # Get current value
        self.setGlobalInt(name, value + incValue) # Increment saved value
        return value                              # Return old value

    ##################
    # Global Strings #
    ##################
    def getGlobalString(self, name, default=""):
        if name in self.globalStrings:
            return self.globalStrings[name]
        # No value
        return default

    def setGlobalString(self, name, value):
        self.globalStrings[name] = value

    ##################
    # Global Objects #
    ##################
    def getGlobalObjects(self, name):
        # Return object if found
        if name in self.globalObjects:
            return self.globalObjects[name]
        # No value
        print("Missing Object: %s" % name)
        return None

    def setGlobalObjects(self, name, value):
        self.globalObjects[name] = value
