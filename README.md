# OfficePie

Microsoft Office [automation](https://msdn.microsoft.com/en-us/VBA/office-shared-vba/articles/getting-started-with-vba-in-office) using Python.

[![license](https://img.shields.io/github/license/auneri/OfficePie.svg)](https://github.com/auneri/OfficePie/blob/master/LICENSE.md)

## Creating Portable Applications

```batch
set app=WordRevisions
cd OfficePie\office\%app%
pyinstaller %app%.spec
dist\%app%.exe
```
