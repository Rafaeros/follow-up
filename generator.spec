# -*- mode: python ; coding: utf-8 -*-
"""Arquivo de configuração (spec) do PyInstaller para compilar o app Follow-Up.

Uso:
  pyinstaller generator.spec

Notas:
- O PyInstaller não faz compilação cruzada (cross-compile). Para gerar um .exe, rode no Windows.
- Arquivos dinâmicos (como configs.json e emails_cc.json) foram removidos do 'datas'
  pois são gerados e alterados em tempo de execução pela aplicação.
"""

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# A variável SPECPATH é injetada globalmente pelo PyInstaller com o caminho absoluto deste arquivo
project_root = Path(SPECPATH)

# Garante que o PyInstaller encontre os pacotes do código-fonte
pathex = [
    str(project_root),
    str(project_root / "src"),
]

# Inclui todos os submódulos da pasta src (garante que imports dinâmicos sejam coletados)
hiddenimports = collect_submodules("src")

# Inclui APENAS arquivos estáticos de interface e recursos inalteráveis.
# IMPORTANTE: Arquivos JSON criados pelo programa NÃO entram aqui.
datas = [
    # Inclui o arquivo de tema visual usado pelo frontend
    (str(project_root / "src" / "frontend" / "theme.qss"), "src/frontend"),
]

# Coleta quaisquer outros arquivos de dados estáticos que existam dentro da pasta src
datas += collect_data_files("src")

a = Analysis(
    [str(project_root / "main.py")],
    pathex=pathex,
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="follow-up",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True, # Dica: Mude para False se quiser esconder a janela preta do CMD no futuro
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="follow-up",
)