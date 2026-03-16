# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

project_dir = Path(SPECPATH).resolve()

# Collect CustomTkinter assets/themes and watchdog internals.
datas = []
datas += collect_data_files("customtkinter")

hiddenimports = []
hiddenimports += collect_submodules("customtkinter")
hiddenimports += collect_submodules("watchdog")

# Bundle default docs/templates so first run has references.
for name in [
    "README.md",
    "MULTIPLE_KEYWORDS.md",
    "settings/.gitkeep",
    "state/.gitkeep",
]:
    src = project_dir / name
    if src.exists():
        datas.append((str(src), "."))

for name in [
    "logo_quantum_telecom.png",
    "ico_quantum_telecom.ico",
]:
    src = project_dir / "icon" / name
    if src.exists():
        datas.append((str(src), "icon"))

# Bundle optional startup script for Windows users.
startup_bat = project_dir / "start_gui.bat"
if startup_bat.exists():
    datas.append((str(startup_bat), "."))

a = Analysis(
    ["gui_app.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="QuantumBot",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=(
        str(project_dir / "icon/ico_quantum_telecom.ico")
        if (project_dir / "icon/ico_quantum_telecom.ico").exists()
        else None
    ),
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="QuantumBot",
)
