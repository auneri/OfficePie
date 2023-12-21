# [OfficePie](https://www.urbandictionary.com/define.php?term=Office%20Pie)

Microsoft Office [automation](https://msdn.microsoft.com/en-us/VBA/office-shared-vba/articles/getting-started-with-vba-in-office) using Python.

[![license](https://img.shields.io/github/license/auneri/OfficePie.svg)](https://github.com/auneri/OfficePie/blob/main/LICENSE.md)
[![build](https://img.shields.io/github/actions/workflow/status/auneri/OfficePie/main.yml)](https://github.com/auneri/OfficePie/actions)

## Getting started

```batchfile
pip install git+https://github.com/auneri/OfficePie.git
```

## Creating portable applications

```batchfile
set app=WordRevisions
cd OfficePie\office\%app%
pyinstaller %app%.spec
dist\%app%.exe
```
