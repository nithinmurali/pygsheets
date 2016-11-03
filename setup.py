#!/usr/bin/env python

import os.path
import re
import sys

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

if sys.argv[-1] == 'publish':
    os.system('python setup.py sdist upload')
    sys.exit()

def read(filename):
    return open(os.path.join(os.path.dirname(__file__), filename)).read()

description = 'Google Spreadsheets Python API'

long_description = """
{index}

License
-------
MIT

Download
========
"""

long_description = long_description.lstrip("\n").format(index=read('docs/index.txt'))

version = re.search(r'^__version__\s*=\s*[\'"]([^\'"]*)[\'"]',
                    read('pygsheets/__init__.py'), re.MULTILINE).group(1)

setup(
    name='pygsheets',
    packages=['pygsheets'],
    description=description,
    long_description=long_description,
    version=version,
    author='Nithin Murali',
    author_email='imnmfotmal@gmail.com',
    url='https://github.com/nithinmurali/pygsheets',
    keywords=['spreadsheets', 'google-spreadsheets', 'pygsheets'],
    install_requires=['google-api-python-client', 'enum'],
    download_url='https://github.com/nithinmurali/pygsheets/tarball/'+version,
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 2",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Science/Research",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules"
        ],
    license='MIT'
    )
