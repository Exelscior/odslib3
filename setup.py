#!/usr/bin/env python

from setuptools import setup, find_packages
version = '1.1.1'

if __name__ == '__main__':
    setup(name='odslib3',
          version=version,
          description='An easy to use module that creates ODS documents. Fork of odslib (https://github.com/TauPan/odslib) as looking unmaintained. Original author: Joseph Colton. Adapted for Python3.x support by Joshua Logan',
          author='Joshua Logan',
          author_email='joshua.logan@eveco.re',
          url='https://github.com/Exelscior/odslib3',
          packages=find_packages(),
          license='GPL',
          classifiers=[
              "Development Status :: 5 - Production/Stable", 
              "Intended Audience :: Developers",
              "License :: OSI Approved :: GNU General Public License (GPL)",
              "Natural Language :: English",
              "Programming Language :: Python",
              "Topic :: Software Development :: Libraries :: Python Modules",
              ],
          )
