# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # 결과 엑셀 양식 파일을 번들에 포함 (template_writer.resource_path 에서 참조)
    # app.ico는 exe 아이콘과 별개로 창 아이콘(root.iconbitmap)에도 쓰이므로
    # 런타임에 읽을 수 있도록 번들에 함께 넣는다.
    datas=[
        ('resources/outbound_template.xlsx', 'resources'),
        ('resources/app.ico', 'resources'),
    ],
    hiddenimports=['xlrd'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 실제로 설치되어 있으면서 실행에 불필요한 모듈만 제외.
    # openpyxl은 numpy(셀에 쓸 수 있는 숫자 타입 확장)와 PIL(시트에 이미지 삽입)을
    # try/except ImportError로 임포트한다. 값은 문자열로만 쓰고 이미지는 넣지 않으므로
    # 둘 다 없어도 동작한다. PIL은 아이콘 변환용으로만 설치돼 있는데, 놔두면
    # 번들에 11MB가 딸려 들어간다.
    excludes=[
        'unittest',
        'test',
        'numpy',
        'pandas',
        'PIL',
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
    icon='resources/app.ico',
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
