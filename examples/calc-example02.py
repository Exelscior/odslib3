#!/usr/bin/python

import sys
sys.path.append('..')
import odslib

# Create the document
doc = odslib.ODS()
doc.content.getSheet(0).setSheetName("First")

# By default you start on Sheet1.  This has an index of 0 in ooolib.
# We will create a merged cell with a title.  3 columns, 2 rows.
doc.content.getCell(0, 0).stringValue("Welcome to \"First\"")
doc.content.mergeCells(0, 0, 3, 2)

# Create a new sheet by passing the title.  You will automatically
# move to that sheet.
doc.content.makeSheet("Second")
doc.content.getCell(0, 0).stringValue("I'm on Sheet2")

# Create another one
doc.content.makeSheet("Sheet3")
doc.content.getCell(0, 0).stringValue("This is Sheet3")

# Move back to the first sheet
doc.content.getSheet(0)
doc.content.getCell(0, 2).stringValue("Sheet1's index is 0")

# Move back to the second sheet
doc.content.getSheet(1)
doc.content.getCell(0, 1).stringValue("Sheet2's index is 1")

# Write out the document
doc.save("calc-example02.ods")
