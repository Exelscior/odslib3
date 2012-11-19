#!/usr/bin/python

import sys
sys.path.append('..')
import odslib

# Create the document
doc = odslib.ODS()

# Set Column Width
doc.content.getColumn(0).setWidth('0.5in')
doc.content.getColumn(1).setWidth('1.0in')
doc.content.getColumn(2).setWidth('1.5in')

# Set Row Height
doc.content.getRow(0).setHeight('0.5in')
doc.content.getRow(1).setHeight('1.0in')
doc.content.getRow(2).setHeight('1.5in')

# Fill in Cell Data
doc.content.getCell(0, 0).stringValue("0.5in x 0.5in")
doc.content.getCell(1, 1).stringValue("1.0in x 1.0in")
doc.content.getCell(2, 2).stringValue("1.5in x 1.5in")
doc.content.getCell(3, 3).stringValue("Where am I?")

# Hide a column and a row
doc.content.getColumn(3).setHidden(True)
doc.content.getRow(3).setHidden(True)

# Write out the document
doc.save("calc-example04.ods")
