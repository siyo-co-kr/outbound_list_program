import sys
import pandas as pd
import msoffcrypto
from io import BytesIO
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                               QWidget, QFileDialog, QLineEdit, QLabel, QMessageBox,
                               QRadioButton, QButtonGroup, QHBoxLayout)
from PySide6.QtCore import Qt

class OutboundApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 아웃바운드 양식 자동 필터링 도구")
        self.setMinimumWidth(500)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # 1. 파일 선택 섹션
        self.label = QLabel("1. 환자 정보 엑셀 파일을 선택하세요.")
        layout.addWidget(self.label)

        self.file_btn = QPushButton("파일 찾기")
        self.file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.file_btn)

        self.file_path_label = QLabel("선택된 파일 없음")
        self.file_path_label.setStyleSheet("color: gray;")
        layout.addWidget(self.file_path_label)

        layout.addSpacing(20)

        # 2. 비밀번호 입력 섹션
        layout.addWidget(QLabel("2. 엑셀 비밀번호 (있는 경우만 입력)"))
        self.pw_input = QLineEdit()
        self.pw_input.setPlaceholderText("비밀번호 입력...")
        self.pw_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.pw_input)

        layout.addSpacing(20)

        # 3. 기간 선택 섹션 (라디오 버튼)
        layout.addWidget(QLabel("3. 내원일 기준 필터링 기간 선택"))
        self.group = QButtonGroup(self)
        radio_layout = QHBoxLayout()

        self.periods = ["전체", "1개월 이내", "3개월 이내", "6개월 이내", "1년 이내", "3년 이내", "5년 이내"]
        for i, text in enumerate(self.periods):
            rb = QRadioButton(text)
            if i == 0: rb.setChecked(True)
            self.group.addButton(rb, i)
            radio_layout.addWidget(rb)
        layout.addLayout(radio_layout)

        layout.addSpacing(30)

        # 4. 실행 버튼
        self.run_btn = QPushButton("데이터 정제 및 저장")
        self.run_btn.setFixedHeight(50)
        self.run_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.run_btn.clicked.connect(self.process_data)
        layout.addWidget(self.run_btn)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.selected_file = None

    def select_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', '', 'Excel Files (*.xlsx)')
        if fname:
            self.selected_file = fname
            self.file_path_label.setText(fname.split('/')[-1])

    def process_data(self):
        if not self.selected_file:
            QMessageBox.warning(self, "경고", "파일을 먼저 선택해주세요.")
            return

        password = self.pw_input.text()
        period_option = self.group.checkedButton().text()

        try:
            # 파일 읽기 로직
            with open(self.selected_file, "rb") as f:
                data_to_read = BytesIO(f.read())

            # 1. 파일이 암호화되어 있는지 확인
            office_file = msoffcrypto.OfficeFile(data_to_read)
            is_encrypted = office_file.is_encrypted()

            if is_encrypted:
                if not password:
                    # 비밀번호가 있는데 입력 안 한 경우
                    QMessageBox.warning(self, "알림", "이 파일은 암호화되어 있습니다. 비밀번호가 필요합니다.")
                    return

                try:
                    decrypted_data = BytesIO()
                    office_file.load_key(password=password)
                    office_file.decrypt(decrypted_data)
                    data_to_read = decrypted_data
                except Exception as e:
                    # 발생한 에러 메시지 안에 'password'라는 단어가 있는지 체크
                    if "password" in str(e).lower():
                        QMessageBox.critical(self, "오류", "비밀번호가 틀렸습니다. 다시 확인해 주세요.")
                    else:
                        QMessageBox.critical(self, "오류", f"처리 중 오류 발생: {e}")
                    return
            else:
                if password:
                    # 비밀번호가 없는데 입력한 경우 (경고 후 진행하거나 차단)
                    QMessageBox.information(self, "알림", "비밀번호가 설정되지 않은 파일입니다. 입력한 비밀번호를 무시하고 진행합니다.")
                    # 무시하고 진행하려면 password = "" 처리 혹은 그냥 통과

            if password:
                decrypted_data = BytesIO()
                office_file = msoffcrypto.OfficeFile(data_to_read)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted_data)
                data_to_read = decrypted_data

            # 데이터 가공 (기존 로직)
            df = pd.read_excel(data_to_read, sheet_name=1, dtype={'연락처': str})
            df.columns = df.columns.str.strip()

            target_cols = ['차트번호', '이름', '연락처', '마지막 내원일자']
            if not all(col in df.columns for col in target_cols):
                raise ValueError(f"필수 항목({', '.join(target_cols)})이 부족합니다.")

            df['마지막 내원일자'] = pd.to_datetime(df['마지막 내원일자'], errors='coerce')

            # 필터링
            if period_option != "전체":
                today = datetime.now()
                delta_map = {"1개월 이내": 30, "3개월 이내": 90, "6개월 이내": 180, "1년 이내": 365, "3년 이내": 1095, "5년 이내": 1825}
                threshold_date = today - timedelta(days=delta_map[period_option])
                df = df[df['마지막 내원일자'] >= threshold_date]

            # 정제
            df['연락처'] = df['연락처'].str.replace(r'[^0-9]', '', regex=True)
            df.loc[df['연락처'].str.len() == 10, '연락처'] = '010' + df['연락처'].str[2:]

            df_deduped = df.drop_duplicates(subset=['연락처'])
            df_final = df_deduped[df_deduped['연락처'].str.match(r'^010\d{8}$')].copy()

            df_final['차트번호'] = ""
            df_final = df_final[['차트번호', '이름', '연락처']]

            # 저장 경로 선택
            save_path, _ = QFileDialog.getSaveFileName(self, '결과 저장', 'filtered_data.xlsx', 'Excel Files (*.xlsx)')
            if save_path:
                df_final.to_excel(save_path, index=False)
                QMessageBox.information(self, "완료", f"총 {len(df_final)}건의 데이터가 저장되었습니다.")

        except Exception as e:
            QMessageBox.critical(self, "오류", f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutboundApp()
    window.show()
    sys.exit(app.exec())