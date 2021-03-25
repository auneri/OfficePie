import inspect
import pathlib

import setuptools


def read(filename):
    filepath = pathlib.Path(inspect.getfile(inspect.currentframe())).resolve().parent / filename
    with filepath.open() as f:
        return f.read()


setuptools.setup(
    name='office',
    version='1.0.0.dev',
    description=read('README.md').splitlines()[2],
    long_description=read('README.md'),
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
        'pyqt5',
        'pywin32'],
    python_requires='>=3.6')
