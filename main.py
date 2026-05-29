import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QSplashScreen
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap, QColor, QFont, QPainter


def _make_splash_pixmap() -> QPixmap:
    """로딩 중 표시할 스플래시 화면 픽스맵 생성"""
    pixmap = QPixmap(420, 160)
    pixmap.fill(QColor("#2E7D32"))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    # 제목
    title_font = QFont()
    title_font.setPointSize(18)
    title_font.setBold(True)
    painter.setFont(title_font)
    painter.setPen(QColor("white"))
    painter.drawText(0, 0, 420, 100, Qt.AlignmentFlag.AlignCenter, "아웃바운드 도구 v2.0.0")

    # 부제목
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
