#!/bin/bash
echo "Generating docs"
python setup.py build_sphinx
echo "Creating sdist"
python setup.py sdist --formats=gztar
echo "Creating wheel package"
python setup.py bdist_wheel
echo "Uploading via twine"
twine upload dist/*
