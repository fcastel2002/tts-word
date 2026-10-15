# PyInstaller: empaqueta helper/service.py con la voz de Piper y el panel compilado (addin/dist).
# Lo corre build.ps1; deja el resultado en installer/build/dist/tts-helper/.
import os

from PyInstaller.utils.hooks import collect_all, collect_submodules

ROOT = os.path.dirname(SPECPATH)
HELPER = os.path.join(ROOT, "helper")

piper_datas, piper_binaries, piper_hidden = collect_all("piper")  # incluye espeak-ng-data

a = Analysis(
    [os.path.join(HELPER, "service.py")],
    pathex=[HELPER],
    binaries=piper_binaries,
    datas=piper_datas
    + [
        (os.path.join(HELPER, "voices"), "voices"),
        (os.path.join(ROOT, "addin", "dist"), "web"),
    ],
    # uvicorn importa sus protocolos y loops por nombre, PyInstaller no los ve.
    hiddenimports=piper_hidden + collect_submodules("uvicorn") + ["app"],
    excludes=["tkinter"],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="tts-helper",
    console=False,  # sin ventana: corre en segundo plano
    icon=os.path.join(SPECPATH, "tts-word.ico"),
)
coll = COLLECT(exe, a.binaries, a.datas, name="tts-helper")
