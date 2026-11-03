# -*- mode: python ; coding: utf-8 -*-
"""Especificação PyInstaller do Conversor CRM para Windows."""

from pathlib import Path


diretorio_projeto = Path(SPECPATH)
caminho_icone = diretorio_projeto / "assets" / "conversor_crm.ico"
dados = [
    (
        str(diretorio_projeto / "assets" / "marca-SESI-SENAI_reduzida_cor.png"),
        "assets",
    ),
    (
        str(caminho_icone),
        "assets",
    ),
]

a = Analysis(
    [str(diretorio_projeto / "app.py")],
    pathex=[str(diretorio_projeto)],
    binaries=[],
    datas=dados,
    hiddenimports=[
        "customtkinter",
        "openpyxl",
        "pandas",
        "PIL.Image",
        "processamento",
        "adapters",
        "pipeline_v2",
        "blip",
        "cursos",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Conversor_CRM",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=str(caminho_icone),
)
