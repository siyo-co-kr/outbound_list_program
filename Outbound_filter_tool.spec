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

# onedir(폴더) 형태로 빌드한다. onefile은 실행할 때마다 번들 전체를 임시 폴더에
# 풀어내며, 실측 기준 그 압축 해제에만 2.17초가 들어 기동 시간의 71%를 차지했다.
# onedir은 그 단계가 없어 평상시 기동이 3.05초에서 0.51초로 줄었다.
# UPX도 끈다. 압축된 DLL은 로드할 때마다 메모리에서 다시 풀어야 한다.
exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,          # 바이너리는 COLLECT가 폴더에 배치
    name='Outbound_filter_tool',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Outbound_filter_tool',
)
