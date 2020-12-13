# [OfficePie](https://www.urbandictionary.com/define.php?term=Office%20Pie)

Microsoft Office [automation](https://msdn.microsoft.com/en-us/VBA/office-shared-vba/articles/getting-started-with-vba-in-office) using Python.

[![license](https://img.shields.io/github/license/auneri/OfficePie.svg)](https://github.com/auneri/OfficePie/blob/master/LICENSE.md)
![build](https://github.com/auneri/OfficePie/workflows/office/badge.svg)

## Getting started

```batch
pip install git+https://github.com/auneri/OfficePie.git
```

## Creating portable applications

```batch
set app=WordRevisions
cd OfficePie\office\%app%
pyinstaller %app%.spec
dist\%app%.exe
```
