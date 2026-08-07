# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # 결과 엑셀 양식 파일을 번들에 포함 (template_writer.resource_path 에서 참조)
    datas=[('resources/outbound_template.xlsx', 'resources')],
    hiddenimports=['xlrd'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 실제로 설치되어 있으면서 실행에 불필요한 모듈만 제외.
    # numpy/pandas는 더 이상 쓰지 않는다. openpyxl이 numpy를 try/except로 임포트해
    # 셀에 쓸 수 있는 숫자 타입을 넓히지만, 값은 문자열로만 쓰므로 없어도 무방하다.
    excludes=[
        'unittest',
        'test',
        'numpy',
        'pandas',
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Outbound_filter_tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
