#!/usr/bin/python

import sys
sys.path.append('..')
import odslib

# Create your document
doc = odslib.ODS()

# Example Title
titleCell = doc.content.getCell(0, 0)
titleCell.stringValue('Alignment & Rotation')
titleCell.setFontSize("14pt")
titleCell.setBold(True)
titleCell.setAlignVertical('center')
titleCell.setAlignHorizontal('center')
doc.content.mergeCells(0, 0, 5, 2)

# Set the row height
doc.content.getRow(2).setHeight('0.5in')
doc.content.getRow(3).setHeight('0.5in')
doc.content.getRow(4).setHeight('0.5in')
doc.content.getRow(5).setHeight('0.5in')

# Vertical Alignment Labels
doc.content.getCell(0, 3).stringValue('top').setBold(True)
doc.content.getCell(0, 4).stringValue('middle').setBold(True)
doc.content.getCell(0, 5).stringValue('bottom').setBold(True)

# Horizontal Alignment Labels
# Rotation is calculated in degrees (0-359) rotating counter clockwise.
doc.content.getCell(1, 2).stringValue('left').setBold(True).setRotation(0)
doc.content.getCell(2, 2).stringValue('center').setBold(True).setRotation(90)
doc.content.getCell(3, 2).stringValue('right').setBold(True).setRotation(180)
doc.content.getCell(4, 2).stringValue('justify').setBold(True).setRotation(270)

# Fill in aligned properties
# Vertical Align top
doc.content.getCell(1, 3).stringValue('x').setAlignVertical('top').setAlignHorizontal('left')
doc.content.getCell(2, 3).stringValue('x').setAlignVertical('top').setAlignHorizontal('center')
doc.content.getCell(3, 3).stringValue('x').setAlignVertical('top').setAlignHorizontal('right')
doc.content.getCell(4, 3).stringValue('x').setAlignVertical('top').setAlignHorizontal('justify')

# Vertical Align middle
doc.content.getCell(1, 4).stringValue('x').setAlignVertical('middle').setAlignHorizontal('left')
doc.content.getCell(2, 4).stringValue('x').setAlignVertical('middle').setAlignHorizontal('center')
doc.content.getCell(3, 4).stringValue('x').setAlignVertical('middle').setAlignHorizontal('right')
doc.content.getCell(4, 4).stringValue('x').setAlignVertical('middle').setAlignHorizontal('justify')

# Vertical Align bottom
doc.content.getCell(1, 5).stringValue('x').setAlignVertical('bottom').setAlignHorizontal('left')
doc.content.getCell(2, 5).stringValue('x').setAlignVertical('bottom').setAlignHorizontal('center')
doc.content.getCell(3, 5).stringValue('x').setAlignVertical('bottom').setAlignHorizontal('right')
doc.content.getCell(4, 5).stringValue('x').setAlignVertical('bottom').setAlignHorizontal('justify')

# Save the document to the file you want to create
doc.save("calc-example06.ods")
