#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2023  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import os

# Import ooolib-python modules
import ooolibGlobal
import ooolibMeta

###################################
# ooolib-python Calc Spreadsheets #
###################################
class Calc:
    """Calc - Open Document Format (ODF) spreadsheets

    Base object for creating ODF spreadsheets
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.meta = ooolibMeta.Meta(self.global_object)

    def config(self, name, value):
        pass

class Writer:
    """Writer - Open Document Format (ODF) text documents

    Base object for creating ODF text documents
    """
    def __init__(self):
        self.global_object = ooolibGlobal.Global()
        self.meta = ooolibMeta.Meta(self.global_object)

    def config(self, name, value):
        pass

###################
# Execute to Test #
###################
if __name__ == "__main__":
    calc = Calc()
