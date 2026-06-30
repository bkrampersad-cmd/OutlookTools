# outlook_tools.spec  —  v1.5
# Build with:  pyinstaller outlook_tools.spec --clean
#
# This builds in ONEDIR mode (matches the Calculator Suite layout):
#   dist/BeeransOutlookTools/BeeransOutlookTools.exe   <- run this
#   dist/BeeransOutlookTools/_internal/                <- everything else
#
# IMPORTANT: Run  python generate_sounds.py  ONCE before building
#            to produce the sounds/ folder that gets bundled here.
# IMPORTANT: Place icon.ico and logo.png next to this spec file / monitor.py
#            before building. icon.ico is used for the .exe icon AND bundled
#            so the app can use it for the window/tray icon at runtime;
#            logo.png is bundled for the sidebar logo.

from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
import os

base_dir = os.path.dirname(os.path.abspath(SPEC))

# Bundle the entire sounds/ subfolder into the exe
sounds_dir = os.path.join(base_dir, "sounds")
sound_datas = [
    (os.path.join(sounds_dir, f), "sounds")
    for f in os.listdir(sounds_dir)
    if f.endswith(".wav")
] if os.path.isdir(sounds_dir) else []

# Bundle icon.ico and logo.png so the running app can load them
# (window icon, tray icon, sidebar logo)
icon_path = os.path.join(base_dir, "icon.ico")
logo_path = os.path.join(base_dir, "logo.png")
icon_datas = [(icon_path, ".")] if os.path.isfile(icon_path) else []
icon_datas += [(logo_path, ".")] if os.path.isfile(logo_path) else []

a = Analysis(
    ["monitor.py"],
    pathex=[],
    binaries=[],
    datas=sound_datas + icon_datas,
    hiddenimports=[
        "customtkinter",
        "plyer.platforms.win.notification",
        "win32com",
        "win32com.client",
        "pythoncom",
        "pytz",
        "pystray",
        "pystray._win32",
        "PIL",
        "PIL._tkinter_finder",
        "PIL.Image",
        "PIL.ImageDraw",
        "pkg_resources.py2_compat",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="BeeransOutlookTools",
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon="icon.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="BeeransOutlookTools",
)
