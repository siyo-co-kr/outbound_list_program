import sys
import os
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QSplashScreen
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QColor, QFont, QPainter


def _create_desktop_shortcut_once():
    """최초 실행 시 바탕화면 바로가기 자동 생성 (PyInstaller 빌드 환경 전용)"""
    if not getattr(sys, 'frozen', False):
        return  # 소스 코드 직접 실행 시 건너뜀

    import ctypes
    import ctypes.wintypes
    import subprocess

    # 실제 바탕화면 경로 조회 (OneDrive 리다이렉트 환경 포함 대응)
    CSIDL_DESKTOPDIRECTORY = 0x0010
    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
    desktop = buf.value

    shortcut_path = os.path.join(desktop, "아웃바운드 도구.lnk")
    if os.path.exists(shortcut_path):
        return  # 이미 바로가기가 있으면 건너뜀

    exe_path = sys.executable
    exe_dir = os.path.dirname(exe_path)

    ps_script = (
        f'$ws = New-Object -ComObject WScript.Shell; '
        f'$sc = $ws.CreateShortcut("{shortcut_path}"); '
        f'$sc.TargetPath = "{exe_path}"; '
        f'$sc.WorkingDirectory = "{exe_dir}"; '
        f'$sc.Description = "아웃바운드 도구 v2.0.0"; '
        f'$sc.Save()'
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True
    )


def _make_splash_pixmap() -> QPixmap:
    """로딩 중 표시할 스플래시 화면 픽스맵 생성"""
    pixmap = QPixmap(420, 160)
    pixmap.fill(QColor("#2E7D32"))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    title_font = QFont()
    title_font.setPointSize(18)
    title_font.setBold(True)
    painter.setFont(title_font)
    painter.setPen(QColor("white"))
    painter.drawText(0, 0, 420, 100, Qt.AlignmentFlag.AlignCenter, "아웃바운드 도구 v2.0.0")

    sub_font = QFont()
    sub_font.setPointSize(11)
    painter.setFont(sub_font)
    painter.setPen(QColor("#A5D6A7"))
    painter.drawText(0, 90, 420, 60, Qt.AlignmentFlag.AlignCenter, "로딩 중...")

    painter.end()
    return pixmap


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("아웃바운드 도구 v2.0.0")
        self.setMinimumWidth(600)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # pandas·msoffcrypto 등 무거운 모듈은 스플래시 화면이 표시된 후 임포트
        from app_modules.ui_components.outbound_list_filter_ui import OutboundApp
        from app_modules.ui_components.outbound_limit_ui import OutboundLimitApp

        self.outbound_tab = OutboundApp()
        self.tabs.addTab(self.outbound_tab, "📄아웃바운드 리스트 필터")

        self.limit_tab = OutboundLimitApp()
        self.tabs.addTab(self.limit_tab, "🚫아웃바운드 제한 리스트")


def main():
    # 최초 실행 시 바탕화면 바로가기 자동 생성
    _create_desktop_shortcut_once()

    app = QApplication(sys.argv)

    # 스플래시 화면을 즉시 표시 (무거운 모듈 임포트 전)
    splash = QSplashScreen(_make_splash_pixmap())
    splash.show()
    app.processEvents()

    # MainWindow 생성 시 app_modules 및 pandas·numpy 임포트 발생
    window = MainWindow()
    window.show()

    # 메인 윈도우가 준비되면 스플래시 닫기
    splash.finish(window)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
