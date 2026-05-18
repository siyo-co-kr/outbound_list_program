import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget
from app_modules.ui_components.outbound_list_filter_ui import OutboundApp

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("아웃바운드 리스트 필터링 도구 v2.0.0")
        self.setMinimumWidth(600)

        # 메인 탭 위젯 생성
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # 1. 아웃바운드 필터 탭 추가
        self.outbound_tab = OutboundApp()
        self.tabs.addTab(self.outbound_tab, "데이터 필터링")

        # 2. (추후 예시) 다른 기능 탭 추가 시
        # self.other_tab = OtherFeatureApp()
        # self.tabs.addTab(self.other_tab, "추가 기능")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()