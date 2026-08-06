import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ..processors.outbound_limit_processor import outbound_limit
from ..processors.template_writer import save_limit_list
from .widgets import (EXCEL_FILETYPES, EXCEL_SAVE_FILETYPES, FONT_BASE, FONT_SMALL,
                      GRAY_TEXT, NO_FILE, RED, RED_ACTIVE, accent_button, section_label)

INFO_TEXT = (
    "💡 안내:\n"
    "· 필수 항목: 차트번호, 이름, 연락처, 마지막 내원일자\n"
    "· '아웃바운드 제한 설정' 컬럼이 없으면 자동으로 생성되어 'O'가 입력됩니다."
)


class OutboundLimitApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(16, 10))
        self.file_path = None
        self.init_ui()

    def init_ui(self):
        self.columnconfigure(0, weight=1)
        row = 0

        # 1. 파일 선택부
        section_label(self, "1. 제한 대상 환자 정보 엑셀 파일을 선택하세요.").grid(
            row=row, column=0, sticky="w", pady=(6, 4))
        row += 1

        self.btn_browse = accent_button(self, "제한 명단 파일 찾기", self.browse_file, RED, RED_ACTIVE)
        self.btn_browse.grid(row=row, column=0, sticky="ew")
        row += 1

        self.file_path_display = ttk.Label(self, text=NO_FILE, foreground=GRAY_TEXT,
                                           font=FONT_SMALL, wraplength=560, justify="left")
        self.file_path_display.grid(row=row, column=0, sticky="w", pady=(5, 0))
        row += 1

        # 2. 암호 입력부
        section_label(self, "2. 파일 비밀번호 입력 (없을 시 미입력)").grid(
            row=row, column=0, sticky="w", pady=(14, 4))
        row += 1

        self.password_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.password_var, show="●", font=FONT_BASE).grid(
            row=row, column=0, sticky="ew")
        row += 1

        # 3. 설명 라벨 (사용자 가이드)
        info = tk.Label(self, text=INFO_TEXT, justify="left", anchor="w", font=FONT_SMALL,
                        fg="#555555", bg="#F5F5F5", padx=10, pady=10,
                        relief="solid", borderwidth=1)
        info.grid(row=row, column=0, sticky="ew", pady=(16, 0))
        row += 1

        # 4. 실행 버튼
        self.btn_run = accent_button(self, "제한 명단 추출 및 저장", self.run_filter, RED, RED_ACTIVE)
        self.btn_run.grid(row=row, column=0, sticky="ew", pady=(20, 6))
        row += 1

        self.rowconfigure(row, weight=1)

    def browse_file(self):
        fname = filedialog.askopenfilename(parent=self, title="제한 명단 파일 선택",
                                           filetypes=EXCEL_FILETYPES)
        if fname:
            self.file_path = fname
            self.file_path_display.config(text=fname)

    def run_filter(self):
        if not self.file_path:
            messagebox.showwarning("경고", "파일을 먼저 선택해주세요.", parent=self)
            return

        self.btn_run.config(state="disabled", text="처리 중...")
        self.update_idletasks()
        try:
            # 아웃바운드 제한 로직 프로세서 호출
            df_result = outbound_limit(file_path=self.file_path, password=self.password_var.get())

            # 결과 저장 (템플릿과 동일한 서식 적용)
            save_path = filedialog.asksaveasfilename(
                parent=self, title="제한 명단 결과 저장", defaultextension=".xlsx",
                initialfile="outbound_limit_list.xlsx", filetypes=EXCEL_SAVE_FILETYPES,
            )
            if save_path:
                count = save_limit_list(df_result, save_path)
                messagebox.showinfo(
                    "완료", f"아웃바운드 제한 명단 생성 완료!\n총 {count}건이 저장되었습니다.", parent=self)

        except Exception as e:
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{e}", parent=self)
        finally:
            self.btn_run.config(state="normal", text="제한 명단 추출 및 저장")
