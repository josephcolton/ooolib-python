#!/usr/bin/python

import sys
sys.path.append('../code')
import ooolib


# Create the document
# When you create the document, it is assumed that you want to start
# with a blank table.  This is automatically down with the autoInit.
# you can choose to instead make a blank document without any sheets
# using:
#
# autoInit=False
#
doc = ooolib.Calc(autoInit=False)

# Since we do not have any sheets, we need to create them before we
# have an activeTable.
# We can create a new table (sheet) with the addTable method.

doc.content.addTable("First") # First sheet
doc.content.activeTable.setCellText(1, 1, "First Sheet")

# We can create a second sheet using the addTable method again.  The
# activeTable is automatically updated to the newest sheet/table that
# you create.

doc.content.addTable("Second") # Second sheet
doc.content.activeTable.setCellText(1, 1, "Second Sheet")

# Let's go ahead and create another one.  This time we will remember
# where it is using a table object.

thirdSheet = doc.content.addTable("Third") # Third sheet
thirdSheet.setCellText(1, 1, "Third Sheet")

# The activeTable is once again updated.  If we want to move back,
# we can use the setActiveTable method to find it.

# Move back to the first sheet
doc.content.setActiveTable("Second")
doc.content.activeTable.setCellText(2, 1, "More text on the second sheet")

# You can move the the last sheet by not using a name in the lookup.
doc.content.setActiveTable() # Should take us to the last sheet
doc.content.activeTable.setCellText(2, 1, "Text on the last sheet")

# You can also rename a sheet.  You first need to find the sheet,
# then you can rename it.

doc.content.setActiveTable("First")
doc.content.activeTable.setTableName("Number 1")

# Write out the document
doc.export("calc-example02.ods")
