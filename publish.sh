#!/bin/bash
echo "Generating docs"
./setup.py build_sphinx
echo "Creating sdist"
./setup.py sdist --formats=bztar
echo "Creating wheel package"
./setup.py bdist_wheel
echo "Uploading via twine"
twine upload dist/*
