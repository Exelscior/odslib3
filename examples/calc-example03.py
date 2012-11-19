#!/usr/bin/python

import sys
sys.path.append('..')
import odslib

# Create the document
doc = odslib.ODS()

# Document Properties
doc.meta.setTitle('The Search')
doc.meta.setSubject('Searching for the Holy Grail')
doc.meta.setDescription('This document is all about finding the grail, vorpal bunnies, and stuff like that.')

# Set meta data for the document
doc.meta.setCreator('King Arther')
doc.meta.setModifier('Sir Robin')

# Single Cell
doc.content.getCell(0, 0).stringValue("To see the changes select: File -> Properties...")

# Write out the document
doc.save("calc-example03.ods")