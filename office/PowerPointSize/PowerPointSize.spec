#!/usr/bin/env python
import inspect
import pathlib

DEBUG = False

application_name = specnm  # type: ignore # noqa: F821
application_path = pathlib.Path(inspect.getfile(inspect.currentframe())).resolve().parent
script_path = application_path / f'{application_name}.py'
icon_path = application_path / f'{application_name}.ico'

analysis = Analysis(  # type: ignore # noqa: F821
    [str(script_path)])

pyz = PYZ(  # type: ignore # noqa: F821
    analysis.pure)

exe = EXE(  # type: ignore # noqa: F821
    pyz,
    analysis.scripts,
    *(() if DEBUG else (analysis.binaries, analysis.datas)),
    console=DEBUG,
    debug=DEBUG,
    name=application_name,
    exclude_binaries=DEBUG,
    icon=str(icon_path),
    upx=False)

if DEBUG:
    collect = COLLECT(  # type: ignore # noqa: F821
        exe,
        analysis.binaries,
        analysis.datas,
        name=application_name)
