#!/usr/bin/python

import sys
sys.path.append('..')
import odslib

# Create the document
doc = odslib.ODS()

# Standard Cell Properties
# The standard bold, italics, and underline properties can be
# set using a boolean method.  You set them to True to add
# the property and False to turn it off.  Everything is
# False by default, so you only need to set True.

# Once again, you can combine methods into a single line to
# save space and clarify things.
doc.content.getCell(0, 0).stringValue("Normal Text")
doc.content.getCell(0, 1).stringValue("Bold Text").setBold(True)
doc.content.getCell(0, 2).stringValue("Italic Text").setItalic(True)
doc.content.getCell(0, 3).stringValue("Underline Text").setUnderline(True)

# Colors
# Colors in the ODS documents are written in standard HTML
# style.  Here we set both the background color and the
# foreground/text color.

# You do not need to combine methods, but I find it easier to
# read than uncombined.
doc.content.getCell(1, 0).stringValue("Blue on Red")
doc.content.getCell(1, 0).setFontColor("#0000FF")
doc.content.getCell(1, 0).setCellColor("#FF0000")
doc.content.getCell(1, 1).stringValue("Red on Blue")
doc.content.getCell(1, 1).setFontColor("#FF0000")
doc.content.getCell(1, 1).setCellColor("#0000FF")
doc.content.getCell(1, 2).stringValue("Default Colors")

# Text Font Sizes
# After not combining methods for the colors, I decided to
# combine them this time.
doc.content.getCell(2, 0).stringValue("Default 10pt")
doc.content.getCell(2, 1).stringValue("11pt").setFontSize("11pt")
doc.content.getCell(2, 2).stringValue("12pt").setFontSize("12pt")

# Write out the document
doc.save("calc-example05.ods")
