#!/usr/bin/python

import sys
sys.path.insert(0, '..')
import odslib

# Create your document.  In all examples I will call the object doc.
doc = odslib.ODS()

# Before we can put data into the cell, we need to get the cell object.
# The getCell(col, row) method can be used for this.  The columns and
# rows both start from 0.  Column A is represented as "0" and Row 1 is
# represented as "0".

# Now we are going to insert values into cells.  There are two methods
# that are used for normal data.  floatValue() is used for number data
# and stringValue() is used for strings.

for row in range(1, 9):
    for col in range(1, 9):
        # Because the getCell(col, row) method returns an object, we
        # can immediately tack the floatValue() method on the end with
        # the value we want put into the cell.
        doc.content.getCell(col, row).floatValue(col * row).setBorder()
        # The floatValue() method also returns a copy of the cell object
        # so we could comtinue to tack cell modifiers on the end.  Some
        # of the later examples will do this, so be prepared.

# Save the document to the file you want to create
doc.save("calc-example01.ods")
