.. ODSlib3 documentation master file, created by
   sphinx-quickstart on Mon Nov 19 09:41:11 2012.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to ODSlib3's documentation!
==================================

Contents:

About
=====

This project (https://github.com/Exelscior/odslib3) is fork of odslib (https://github.com/TauPan/odslib), which is a fork of odslib-python (http://code.google.com/p/odslib-python/), which is fork of ooolib-python. (phew)

Installation
============
::

  pip install odslib3

or to install from the source

::

   git clone https://github.com/Exelscior/odslib3.git && cd odslib3
   python setup.py {build, install}


API
===

.. toctree::
   :maxdepth: 2

   content
   meta
   mimetype

Example (django)
================
::

  from odslib3 import ODS

  def report(request):
      ods = ODS()
      
      # sheet title
      sheet = ods.content.getSheet(0)
      sheet.setSheetName('Totals')
      
      # title
      sheet.getCell(0, 0).stringValue("Nice cool report").setFontSize('14pt')
      sheet.getRow(0).setHeight('18pt')
      sheet.getColumn(0).setWidth('10cm')
      
      # Cell1
      sheet.getCell(0, 1).stringValue("Foo")
      sheet.getCell(1, 1).floatValue(2)

      # Cell2
      sheet.getCell(0, 2).stringValue("Bar")
      sheet.getCell(1, 2).floatValue(3)

      # Cell3 with formula
      sheet.getCell(0, 3).stringValue("Total").setBold(True)
      sheet.getCell(1, 3).floatFormula(0, '=SUM(B2:B3').setBold(True)
      
      # generating response
      response = HttpResponse(mimetype=ods.mimetype.toString())
      response['Content-Disposition'] = 'attachment; filename="report.ods"'
      ods.save(response)
      
      return response
 

See more examples in examples directory.


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

