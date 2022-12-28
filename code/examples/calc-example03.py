#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create the document
doc = ooolib.Calc()

# Document Properties
doc.meta.setTitle("The Search")
doc.meta.setSubject("Searching for the Grail")
doc.meta.setDescription("This document is all about finding the grail.")

# Keywords
doc.meta.addKeyword("grail")
doc.meta.addKeyword("coconut")
doc.meta.addKeyword("king")
doc.meta.addKeyword("search")

# Single Cell Instructions
doc.content.activeTable.setCellText(1, 1, "To see the changes select: File -> Properties...")

# Write out the document
doc.export("calc-example03.ods")
