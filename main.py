import sys
import pandas as pd
import msoffcrypto
from io import BytesIO
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                               QWidget, QFileDialog, QLineEdit, QLabel, QMessageBox,
                               QRadioButton, QButtonGroup, QHBoxLayout, QDateEdit,
                               QCheckBox, QTabWidget)
from PySide6.QtCore import Qt, QDate

class HospitalDataApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("아웃바운드 리스트 통합 관리 시스템")
        self.setMinimumWidth(750)
        self.initUI()

    def initUI(self):
        # 메인 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 탭 위젯 생성
        self.tabs = QTabWidget()

        # 각 탭 구성
        self.outbound_tab = QWidget()
        self.restriction_tab = QWidget()

        self.tabs.addTab(self.outbound_tab, "🔍 아웃바운드 명단 추출")
        self.tabs.addTab(self.restriction_tab, "🚫 제한 리스트 생성")

        self.setup_outbound_tab()
        self.setup_restriction_tab()

        main_layout.addWidget(self.tabs)

    # ---------------------------------------------------------
    # [탭 1] 아웃바운드 명단 추출 (연락처 중복 제거 포함)
    # ---------------------------------------------------------
    def setup_outbound_tab(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("### 1. 환자 정보 엑셀 파일을 선택하세요."))
        self.file_btn = QPushButton("파일 찾기")
        self.file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.file_btn)

        self.file_path_label = QLabel("선택된 파일 없음")
        self.file_path_label.setStyleSheet("color: gray;")
        layout.addWidget(self.file_path_label)

        layout.addSpacing(10)
        layout.addWidget(QLabel("### 2. 엑셀 비밀번호 (있는 경우만 입력)"))
        self.pw_input = QLineEdit()
        self.pw_input.setPlaceholderText("비밀번호 입력...")
        self.pw_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.pw_input)

        layout.addSpacing(10)
        layout.addWidget(QLabel("### 3. 마지막 내원일 기준 필터링"))
        self.group = QButtonGroup(self)
        radio_layout = QHBoxLayout()
        self.periods = ["전체", "1개월 이내", "3개월 이내", "6개월 이내", "1년 이내", "3년 이내", "5년 이내"]
        for i, text in enumerate(self.periods):
            rb = QRadioButton(text)
            if i == 0: rb.setChecked(True)
            self.group.addButton(rb, i)
            radio_layout.addWidget(rb)
        layout.addLayout(radio_layout)

        layout.addSpacing(10)
        layout.addWidget(QLabel("### 4. 생년월일 범위 필터링 (선택사항)"))
        birth_filter_layout = QHBoxLayout()
        self.use_birth_filter = QCheckBox("필터 사용")
        birth_filter_layout.addWidget(self.use_birth_filter)

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate(1950, 1, 1))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        birth_filter_layout.addWidget(QLabel("시작:"))
        birth_filter_layout.addWidget(self.start_date)
        birth_filter_layout.addWidget(QLabel("~ 종료:"))
        birth_filter_layout.addWidget(self.end_date)
        layout.addLayout(birth_filter_layout)

        layout.addSpacing(20)
        self.run_btn = QPushButton("데이터 정제 및 저장")
        self.run_btn.setFixedHeight(45)
        self.run_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.run_btn.clicked.connect(self.process_outbound_data)
        layout.addWidget(self.run_btn)

        self.outbound_tab.setLayout(layout)
        self.selected_file = None

    def select_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', '', 'Excel Files (*.xlsx)')
        if fname:
            self.selected_file = fname
            self.file_path_label.setText(fname.split('/')[-1])

    def process_outbound_data(self):
        if not self.selected_file:
            QMessageBox.warning(self, "경고", "파일을 먼저 선택해주세요.")
            return

        password = self.pw_input.text()
        period_option = self.group.checkedButton().text()

        try:
            with open(self.selected_file, "rb") as f:
                data_to_read = BytesIO(f.read())

            office_file = msoffcrypto.OfficeFile(data_to_read)
            if office_file.is_encrypted():
                if not password:
                    QMessageBox.warning(self, "알림", "비밀번호가 필요합니다.")
                    return
                office_file.load_key(password=password)
                decrypted_data = BytesIO()
                office_file.decrypt(decrypted_data)
                data_to_read = decrypted_data

            df = pd.read_excel(data_to_read, sheet_name=1, dtype={'연락처': str})
            df.columns = df.columns.str.strip()

            required = ['차트번호', '이름', '연락처', '마지막 내원일자']
            if self.use_birth_filter.isChecked():
                required.append('생년월일')

            if not all(col in df.columns for col in required):
                raise ValueError(f"필수 항목({', '.join(required)})이 부족합니다.")

            # 내원일 필터
            df['마지막 내원일자'] = pd.to_datetime(df['마지막 내원일자'], errors='coerce')
            if period_option != "전체":
                today = datetime.now()
                delta_map = {"1개월 이내": 30, "3개월 이내": 90, "6개월 이내": 180, "1년 이내": 365, "3년 이내": 1095, "5년 이내": 1825}
                threshold_date = today - timedelta(days=delta_map[period_option])
                df = df[df['마지막 내원일자'] >= threshold_date]

            # 생년월일 필터
            if self.use_birth_filter.isChecked():
                df['생년월일'] = pd.to_datetime(df['생년월일'], errors='coerce')
                s_date = datetime.combine(self.start_date.date().toPython(), datetime.min.time())
                e_date = datetime.combine(self.end_date.date().toPython(), datetime.max.time())
                df = df[(df['생년월일'] >= s_date) & (df['생년월일'] <= e_date)]

            # 연락처 정제 및 중복 제거
            df['연락처'] = df['연락처'].astype(str).str.replace(r'[^0-9]', '', regex=True)
            mask_fix = (df['연락처'].str.len() == 10) & (df['연락처'].str.startswith('10'))
            df.loc[mask_fix, '연락처'] = '010' + df['연락처'].str[2:]

            df_final = df.drop_duplicates(subset=['연락처']) # 탭 1은 중복 제거 수행
            df_final = df_final[df_final['연락처'].str.match(r'^010\d{8}$')].copy()

            df_final['차트번호'] = ""
            df_out = df_final[['차트번호', '이름', '연락처']]

            save_path, _ = QFileDialog.getSaveFileName(self, '결과 저장', 'filtered_outbound.xlsx', 'Excel Files (*.xlsx)')
            if save_path:
                df_out.to_excel(save_path, index=False)
                QMessageBox.information(self, "완료", f"총 {len(df_out)}건의 데이터가 저장되었습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류 발생: {str(e)}")

    # ---------------------------------------------------------
    # [탭 2] 아웃바운드 제한 리스트 (중복 유지 + 특수문자 제거)
    # ---------------------------------------------------------
    def setup_restriction_tab(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("### 1. 원본 환자 리스트 파일을 선택하세요."))
        self.res_file_btn = QPushButton("파일 찾기")
        self.res_file_btn.clicked.connect(self.select_res_file)
        layout.addWidget(self.res_file_btn)

        self.res_file_label = QLabel("선택된 파일 없음")
        self.res_file_label.setStyleSheet("color: gray;")
        layout.addWidget(self.res_file_label)

        layout.addSpacing(15)
        layout.addWidget(QLabel("### 2. 엑셀 비밀번호 (있는 경우만 입력)"))
        self.res_pw_input = QLineEdit()
        self.res_pw_input.setPlaceholderText("비밀번호 입력...")
        self.res_pw_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.res_pw_input)

        layout.addSpacing(30)
        self.res_run_btn = QPushButton("제한 리스트 생성 및 저장")
        self.res_run_btn.setFixedHeight(50)
        self.res_run_btn.setStyleSheet("background-color: #E91E63; color: white; font-weight: bold;")
        self.res_run_btn.clicked.connect(self.process_restriction_data)
        layout.addWidget(self.res_run_btn)

        layout.addStretch()
        self.restriction_tab.setLayout(layout)
        self.res_selected_file = None

    def select_res_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', '', 'Excel Files (*.xlsx)')
        if fname:
            self.res_selected_file = fname
            self.res_file_label.setText(fname.split('/')[-1])

    def process_restriction_data(self):
        if not self.res_selected_file:
            QMessageBox.warning(self, "경고", "파일을 선택해주세요.")
            return

        try:
            with open(self.res_selected_file, "rb") as f:
                data_to_read = BytesIO(f.read())

            office_file = msoffcrypto.OfficeFile(data_to_read)
            if office_file.is_encrypted():
                pw = self.res_pw_input.text()
                if not pw:
                    QMessageBox.warning(self, "알림", "비밀번호가 필요합니다.")
                    return
                office_file.load_key(password=pw)
                decrypted_data = BytesIO()
                office_file.decrypt(decrypted_data)
                data_to_read = decrypted_data

            # 데이터 로드
            df = pd.read_excel(data_to_read, sheet_name=1)
            df.columns = df.columns.str.strip()

            # 필수 컬럼 정의
            required = ['차트번호', '이름', '생년월일', '휴대폰번호', '마지막 내원 일자']
            for col in required:
                if col not in df.columns:
                    raise ValueError(f"필수 컬럼 '{col}'이 엑셀에 존재하지 않습니다.")

            # 1. 숫자 정제 (특수문자 제거)
            clean_cols = ['생년월일', '휴대폰번호', '마지막 내원 일자']
            for col in clean_cols:
                df[col] = df[col].astype(str).str.replace(r'[^0-9]', '', regex=True)

            # 2. 아웃바운드 제한 설정 'O' 자동 입력
            df['아웃바운드 제한 설정'] = 'O'

            # 3. 필수값 검증 (빈 칸이 하나라도 있는 행 제거)
            final_cols = required + ['아웃바운드 제한 설정']
            df_final = df[final_cols].copy()
            df_final = df_final.dropna(subset=final_cols)
            for col in final_cols:
                df_final = df_final[df_final[col].astype(str).str.strip() != '']

            # *중복 제거 로직은 요청대로 넣지 않았습니다 (원본 행 보존)*

            save_path, _ = QFileDialog.getSaveFileName(self, '제한 리스트 저장', 'restriction_list.xlsx', 'Excel Files (*.xlsx)')
            if save_path:
                df_final.to_excel(save_path, index=False)
                QMessageBox.information(self, "완료", f"총 {len(df_final)}건의 제한 리스트가 생성되었습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HospitalDataApp()
    window.show()
    sys.exit(app.exec())