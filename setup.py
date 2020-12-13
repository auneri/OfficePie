from __future__ import absolute_import, division, print_function

import os
import re

import setuptools


def readme():
    filepath = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'README.md')
    with open(filepath) as f:
        return f.read()


def version():
    filepath = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'office', '__init__.py')
    with open(filepath) as f:
        version_match = re.search(r"^__version__ = [']([^']*)[']", f.read(), re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError('Failed to find version string')


setuptools.setup(
    name='OfficePie',
    version=version(),
    description='Microsoft Office automation using Python',
    long_description=readme(),
    long_description_content_type='text/markdown',
    url='https://auneri.github.io/OfficePie',
    author='Ali Uneri',
    license='MIT',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3'],
    packages=setuptools.find_packages(),
    install_requires=[
        'pywin32',
        'qtpy'],
    python_requires='>=3.6')
