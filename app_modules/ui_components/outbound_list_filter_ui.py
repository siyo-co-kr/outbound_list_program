from PySide6.QtWidgets import (QWidget, QPushButton, QVBoxLayout,
                               QFileDialog, QLineEdit, QLabel, QMessageBox,
                               QRadioButton, QButtonGroup, QHBoxLayout, QDateEdit,
                               QCheckBox)
from PySide6.QtCore import Qt, QDate
from ..processors.outbound_list_filter_processor import outbound_list_filter

class OutboundApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 전체 레이아웃
        layout = QVBoxLayout()

        # 1. 파일 선택부
        file_layout = QVBoxLayout()
        self.file_label = QLabel("1. 환자 정보 엑셀 파일을 선택하세요.")
        self.file_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        file_layout.addWidget(self.file_label)

        self.btn_browse = QPushButton("파일 찾기")
        self.btn_browse.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px;")
        self.btn_browse.setFixedHeight(40)
        self.btn_browse.clicked.connect(self.browse_file)
        file_layout.addWidget(self.btn_browse)

        self.file_path_display = QLabel("선택된 파일 없음")
        self.file_path_display.setStyleSheet("color: gray; margin-top: 5px;")
        file_layout.addWidget(self.file_path_display)
        layout.addLayout(file_layout)

        # 2. 암호 입력부
        pass_layout = QVBoxLayout()
        self.pass_label = QLabel("2. 파일 비밀번호 입력 (없을 시 미입력)")
        self.pass_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        pass_layout.addWidget(self.pass_label)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        pass_layout.addWidget(self.password_edit)
        layout.addLayout(pass_layout)

        # 3. 내원 기간 설정
        period_layout = QVBoxLayout()

        self.period_label = QLabel("3. 내원일 기준 필터")
        self.period_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        period_layout.addWidget(self.period_label)

        period_content_layout = QHBoxLayout()

        self.period_value = QLineEdit("6")
        self.period_value.setFixedWidth(50)
        self.period_value.setEnabled(False)
        period_content_layout.addWidget(self.period_value)

        self.period_group = QButtonGroup(self)
        self.radio_all = QRadioButton("전체")
        self.radio_month = QRadioButton("개월")
        self.radio_year = QRadioButton("년")

        self.radio_all.setChecked(True)

        self.period_group.addButton(self.radio_all)
        self.period_group.addButton(self.radio_month)
        self.period_group.addButton(self.radio_year)

        self.radio_all.toggled.connect(lambda: self.period_value.setEnabled(False))
        self.radio_month.toggled.connect(lambda: self.period_value.setEnabled(True))
        self.radio_year.toggled.connect(lambda: self.period_value.setEnabled(True))

        period_content_layout.addWidget(self.radio_all)
        period_content_layout.addWidget(self.radio_month)
        period_content_layout.addWidget(self.radio_year)
        period_content_layout.addStretch()

        period_layout.addLayout(period_content_layout)
        layout.addLayout(period_layout)

        # 4. 생년월일 필터링 (선택 사항)
        birth_layout = QVBoxLayout()

        self.birth_label = QLabel("4. 생년월일 기준 필터")
        self.birth_label.setStyleSheet("font-weight: bold; margin-bottom: 5px; margin-top: 10px;")
        birth_layout.addWidget(self.birth_label)

        birth_content_layout = QHBoxLayout()

        self.check_birth = QCheckBox("생년월일 필터 사용")
        self.date_start = QDateEdit(QDate.currentDate().addYears(-30))
        self.date_end = QDateEdit(QDate.currentDate())
        self.date_start.setCalendarPopup(True)
        self.date_end.setCalendarPopup(True)

        birth_content_layout.addWidget(self.check_birth)
        birth_content_layout.addWidget(self.date_start)
        birth_content_layout.addWidget(QLabel("~"))
        birth_content_layout.addWidget(self.date_end)
        birth_content_layout.addStretch()

        birth_layout.addLayout(birth_content_layout)
        layout.addLayout(birth_layout)

        # 5. 실행 버튼
        self.btn_run = QPushButton("필터링 및 저장")
        self.btn_run.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px; margin-top: 10px;")
        self.btn_run.clicked.connect(self.run_filter)
        layout.addWidget(self.btn_run)

        # 최종 여백 추가 및 레이아웃 적용
        layout.addStretch()
        self.setLayout(layout)

    def browse_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if fname:
            self.file_path_display.setText(fname)

    def run_filter(self):
        file_path = self.file_path_display.text()
        password = self.password_edit.text()

        if file_path == "선택된 파일 없음":
            QMessageBox.warning(self, "경고", "파일을 먼저 선택해주세요.")
            return

        try:
            # 1. 프로세서 호출에 필요한 인자 정리
            period_type = "전체"
            if self.radio_month.isChecked():
                period_type = "개월"
            elif self.radio_year.isChecked():
                period_type = "년"

            # 2. 데이터 처리 실행
            df_result = outbound_list_filter(
                file_path=file_path,
                password=password,
                period_type=period_type,
                period_value=self.period_value.text(),
                use_birth=self.check_birth.isChecked(),
                start_date=self.date_start.date().toPython(),
                end_date=self.date_end.date().toPython()
            )

            # 3. 결과 저장
            save_path, _ = QFileDialog.getSaveFileName(self, "결과 저장", "filtered_outbound.xlsx", "Excel Files (*.xlsx)")
            if save_path:
                df_result.to_excel(save_path, index=False)
                QMessageBox.information(self, "완료", f"필터링 완료!\n총 {len(df_result)}건이 저장되었습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")