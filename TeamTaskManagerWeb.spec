# -*- mode: python ; coding: utf-8 -*-
import os

script_path = os.path.join(SPECPATH, 'app.py') if os.path.exists(os.path.join(SPECPATH, 'app.py')) else 'team_task_web/app.py'
web_source = os.path.join(SPECPATH, 'web') if os.path.exists(os.path.join(SPECPATH, 'web')) else 'team_task_web/web'
icon_source = [os.path.join(SPECPATH, 'app_icon.ico')] if os.path.exists(os.path.join(SPECPATH, 'app_icon.ico')) else ['team_task_web/app_icon.ico']

crt_binaries = []
sys32 = r"C:\Windows\System32"
for dll in [
    "msvcp140.dll",
    "msvcp140_1.dll",
    "msvcp140_2.dll",
    "msvcp140_atomic_wait.dll",
    "msvcp140_codecvt_ids.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
]:
    dll_path = os.path.join(sys32, dll)
    if os.path.exists(dll_path):
        crt_binaries.append((dll_path, '.'))

a = Analysis(
    [script_path],
    pathex=[],
    binaries=crt_binaries,
    datas=[(web_source, 'web')],
    hiddenimports=['openpyxl', 'et_xmlfile', 'cloud_excel', 'win32com.client', 'pythoncom', 'pywintypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Excel rows use Python scalars; the native picker uses Windows Forms.
    # Avoid importing NumPy or initializing/bundling the unused Tcl/Tk runtime.
    excludes=['numpy', 'tkinter'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TeamTaskManagerWeb',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_source,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='TeamTaskManagerWeb',
)
