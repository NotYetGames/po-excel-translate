#!/usr/bin/env bash

# https://packaging.python.org/tutorials/packaging-projects/

# Build
python3 -m pip install --user --upgrade setuptools wheel twine
python3 setup.py sdist bdist_wheel

# Upload
python3 -m pip install --user --upgrade twine
# python3 -m twine upload --repository-url https://test.pypi.org/legacy/ dist/*
python3 -m twine upload dist/*
