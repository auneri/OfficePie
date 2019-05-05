# [OfficePie](https://www.urbandictionary.com/define.php?term=Office%20Pie)

Microsoft Office [automation](https://msdn.microsoft.com/en-us/VBA/office-shared-vba/articles/getting-started-with-vba-in-office) using Python.

[![license](https://img.shields.io/github/license/auneri/OfficePie.svg)](https://github.com/auneri/OfficePie/blob/master/LICENSE.md)
[![build](https://img.shields.io/appveyor/ci/auneri/OfficePie.svg)](https://ci.appveyor.com/project/auneri/OfficePie)

## Getting started

```batch
pip install git+https://github.com/auneri/OfficePie.git
```

## Creating Portable Applications

```batch
set app=WordRevisions
cd OfficePie\office\%app%
pyinstaller %app%.spec
dist\%app%.exe
```
