#!/usr/bin/env python

import os.path
import re
import sys

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup
import sys

if sys.argv[-1] == 'publish':
    os.system('python setup.py sdist upload -r pypi')
    sys.exit()

def read(filename):
    return open(os.path.join(os.path.dirname(__file__), filename)).read()

description = 'Google Spreadsheets Python API v4'

long_description = """
{index}

License
-------
MIT

Download
========
"""

version = re.search(r'^__version__\s*=\s*[\'"]([^\'"]*)[\'"]',
                    read('pygsheets/__init__.py'), re.MULTILINE).group(1)

setup(
    name='pygsheets',
    packages=['pygsheets'],
    description=description,
    version=version,
    author='Nithin Murali',
    author_email='imnmfotmal@gmail.com',
    url='https://github.com/nithinmurali/pygsheets',
    keywords=['spreadsheets', 'google-spreadsheets', 'pygsheets'],
    install_requires=['google-api-python-client>=1.5.5'],
    extras_require={
        ':python_version<="2.7"': ["enum"],
    },
    download_url='https://github.com/nithinmurali/pygsheets/tarball/'+version,
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
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
