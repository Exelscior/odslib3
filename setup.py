#!/usr/bin/env python

from setuptools import setup, find_packages
version = '1.0'

if __name__ == '__main__':
    setup(name='odslib',
          version=version,
          description='An easy to use module that creates ODS documents. Fork of original odslib-python (http://code.google.com/p/odslib-python/). Original author: Joseph Colton',
          author='Alexey Loshkarev',
          author_email='elf2001@gmail.com',
          url='https://github.com/angry-elf/odslib/',
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
