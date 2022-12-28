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
        self.global_ints = {}

    def getVersion(self):
        versionStr = "%s %s" % (self.project, self.version)
        return versionStr

    def getGlobalInt(self, name, setDefault=0):
        #print("getGlobalInt(%s)" % (name))
        # Sanitize
        setDefault = int(setDefault)
        # Simple get if existing
        if name in self.global_ints:
            return self.global_ints[name]
        # Missing case
        else:
            self.global_ints[name] = setDefault
            return setDefault

    def setGlobalInt(self, name, value):
        #print("setGlobalInt(%s, %s)" % (name, value))
        value = int(value)             # Sanitize value
        self.global_ints[name] = value # Set value
        return value                   # Return set value

    def incrementGlobalInt(self, name, incValue=1):
        incValue = int(incValue)                  # Sanitize incValue
        value = self.getGlobalInt(name, incValue) # Get current value
        self.setGlobalInt(name, value + incValue) # Increment saved value
        return value                              # Return old value
