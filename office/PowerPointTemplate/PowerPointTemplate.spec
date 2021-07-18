#!/usr/bin/env python
import inspect
import pathlib

DEBUG = False

application_name = specnm  # noqa: F821
application_path = pathlib.Path(inspect.getfile(inspect.currentframe())).resolve().parent  # noqa: F821
script_path = application_path / f'{application_name}.py'
icon_path = application_path / f'{application_name}.ico'

analysis = Analysis(  # noqa: F821
    [str(script_path)])

pyz = PYZ(  # noqa: F821
    analysis.pure,
    analysis.zipped_data)

exe = EXE(  # noqa: F821
    pyz,
    analysis.scripts,
    *(() if DEBUG else (analysis.binaries, analysis.zipfiles, analysis.datas)),
    console=DEBUG,
    debug=DEBUG,
    name=application_name,
    exclude_binaries=DEBUG,
    icon=str(icon_path),
    upx=False)

if DEBUG:
    collect = COLLECT(  # noqa: F821
        exe,
        analysis.binaries,
        analysis.zipfiles,
        analysis.datas,
        name=application_name)
