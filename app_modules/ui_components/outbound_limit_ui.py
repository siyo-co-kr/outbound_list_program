import os
import pandas as pd
from PySide6.QtWidgets import (QWidget, QPushButton, QVBoxLayout,
                               QFileDialog, QLineEdit, QLabel, QMessageBox,
                               QHBoxLayout)
from PySide6.QtCore import Qt
from ..processors.outbound_limit_processor import outbound_limit

class OutboundLimitApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 전체 레이아웃
        layout = QVBoxLayout()

        # 1. 파일 선택부
        file_layout = QVBoxLayout()
        self.file_label = QLabel("1. 제한 대상 환자 정보 엑셀 파일을 선택하세요.")
        self.file_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        file_layout.addWidget(self.file_label)

        self.btn_browse = QPushButton("제한 명단 파일 찾기")
        self.btn_browse.setStyleSheet("background-color: #E53935; color: white; font-weight: bold; height: 40px;")
        self.btn_browse.setFixedHeight(40)
        self.btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(self.btn_browse)

        self.file_path_display = QLabel("선택된 파일 없음")
        self.file_path_display.setStyleSheet("color: gray; margin-top: 5px;")
        file_layout.addWidget(self.file_path_display)
        layout.addLayout(file_layout)

        # 2. 암호 입력부 (비밀번호 입력 UI - 두 줄 구성)
        pass_layout = QVBoxLayout()
        self.pass_label = QLabel("2. 파일 비밀번호 입력 (없을 시 미입력)")
        self.pass_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        pass_layout.addWidget(self.pass_label)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("비밀번호가 걸려있는 경우 입력하세요.")
        pass_layout.addWidget(self.password_edit)
        layout.addLayout(pass_layout)

        # 3. 설명 라벨 (사용자 가이드)
        info_layout = QVBoxLayout()
        self.info_label = QLabel(
            "💡 안내:\n"
            "· 필수 항목: 차트번호, 이름, 연락처, 마지막 내원일자\n"
            "· '아웃바운드 제한 설정' 컬럼이 없으면 자동으로 생성되어 'O'가 입력됩니다."
        )
        self.info_label.setStyleSheet(
            "color: #555555; background-color: #F5F5F5; "
            "border: 1px solid #E0E0E0; border-radius: 4px; "
            "padding: 10px; margin-top: 15px; line-height: 140%;"
        )
        info_layout.addWidget(self.info_label)
        layout.addLayout(info_layout)

        # 4. 실행 버튼
        self.btn_run = QPushButton("제한 명단 추출 및 저장")
        self.btn_run.setStyleSheet("background-color: #E53935; color: white; font-weight: bold; height: 40px; margin-top: 20px;")
        self.btn_run.clicked.connect(self.run_filter)
        layout.addWidget(self.btn_run)

        # 최종 여백 추가 및 레이아웃 적용
        layout.addStretch()
        self.setLayout(layout)

    def browse_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "제한 명단 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if fname:
            self.file_path_display.setText(fname)

    def run_filter(self):
        file_path = self.file_path_display.text()
        password = self.password_edit.text()

        # 파일 선택 여부 검증 (첫 번째 탭의 버그 수정본 반영)
        if file_path == "선택된 파일 없음":
            QMessageBox.warning(self, "경고", "파일을 먼저 선택해주세요.")
            return

        try:
            # 아웃바운드 제한 로직 프로세서 호출
            df_result = outbound_limit(file_path=file_path, password=password)

            # 결과 저장 다이얼로그
            save_path, _ = QFileDialog.getSaveFileName(
                self, "제한 명단 결과 저장", "outbound_limit_list.xlsx", "Excel Files (*.xlsx)"
            )

            if save_path:
                df_result.to_excel(save_path, index=False)
                QMessageBox.information(
                    self, "완료", f"아웃바운드 제한 명단 생성 완료!\n총 {len(df_result)}건이 저장되었습니다."
                )

        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")