#!/usr/bin/env python
from __future__ import absolute_import, division, print_function

import inspect
import os

DEBUG = False

application_name = specnm  # noqa: F821
application_path = os.path.abspath(os.path.join(inspect.getfile(inspect.currentframe()), '..', SPEC, '..'))  # noqa: F821
package_path = os.path.abspath(os.path.join(application_path, '..', '..', application_name))
script_path = os.path.join(application_path, '{}.py'.format(application_name))
icon_path = os.path.join(application_path, '{}.ico'.format(application_name))

a = Analysis(  # noqa: F821
    [script_path],
    pathex=[package_path],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None)

pyz = PYZ(  # noqa: F821
    a.pure,
    a.zipped_data,
    cipher=None)

if DEBUG:
    exe = EXE(  # noqa: F821
        pyz,
        a.scripts,
        exclude_binaries=True,
        name=application_name,
        debug=True,
        strip=False,
        upx=False,
        console=True)

    coll = COLLECT(  # noqa: F821
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=None,
        upx=False,
        name=application_name)
else:
    exe = EXE(  # noqa: F821
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        name=application_name,
        debug=False,
        strip=False,
        upx=True,
        runtime_tmpdir=None,
        console=False,
        icon=icon_path)
