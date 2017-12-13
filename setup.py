#!/usr/bin/env python

from __future__ import absolute_import, division, print_function

import os

import setuptools


def readme():
    filename = 'README.md'
    filepath = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
    with open(filepath) as f:
        return f.read()


setuptools.setup(
    name='OfficePie',
    version='1.0.0.dev',
    description='',
    long_description=readme(),
    url='https://github.com/auneri/officepie',
    author='Ali Uneri',
    license='MIT',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 3'],
    packages=setuptools.find_packages(),
    install_requires=[
        'pywin32',
        'qtpy',
        'six'])
