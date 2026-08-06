import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk

from ..processors.outbound_list_filter_processor import outbound_list_filter
from ..processors.template_writer import save_outbound_list
from .widgets import (EXCEL_FILETYPES, EXCEL_SAVE_FILETYPES, FONT_BASE, FONT_SMALL,
                      GRAY_TEXT, GREEN, GREEN_ACTIVE, NO_FILE, DateEntry, PeriodFilter,
                      accent_button, section_label, years_ago)


class OutboundApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(16, 10))
        self.file_path = None
        self.init_ui()

    def init_ui(self):
        self.columnconfigure(0, weight=1)
        row = 0

        # 1. 파일 선택부
        section_label(self, "1. 환자 정보 엑셀 파일을 선택하세요.").grid(
            row=row, column=0, sticky="w", pady=(6, 4))
        row += 1

        self.btn_browse = accent_button(self, "파일 찾기", self.browse_file, GREEN, GREEN_ACTIVE)
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

        # 3. 내원일 기준 필터 (설정 기준 이내에 내원한 사람만)
        section_label(self, "3. 내원일 기준 필터").grid(
            row=row, column=0, sticky="w", pady=(14, 4))
        row += 1

        self.period = PeriodFilter(self)
        self.period.grid(row=row, column=0, sticky="w")
        row += 1

        # 4. 장기 미내원 기준 필터 (설정 기준보다 오래된 사람만)
        section_label(self, "4. 내원일 기준 필터(설정 기준보다 오래된 사람만)").grid(
            row=row, column=0, sticky="w", pady=(14, 4))
        row += 1

        self.period_old = PeriodFilter(self)
        self.period_old.grid(row=row, column=0, sticky="w")
        row += 1

        # 5. 생년월일 필터링 (선택 사항)
        section_label(self, "5. 생년월일 기준 필터").grid(
            row=row, column=0, sticky="w", pady=(14, 4))
        row += 1

        birth_frame = ttk.Frame(self)
        birth_frame.grid(row=row, column=0, sticky="w")
        row += 1

        self.use_birth_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(birth_frame, text="생년월일 필터 사용", variable=self.use_birth_var,
                        command=self._sync_birth_enabled).pack(side="left", padx=(0, 12))

        self.date_start = DateEntry(birth_frame, years_ago(30))
        self.date_start.pack(side="left")
        ttk.Label(birth_frame, text="~", font=FONT_BASE).pack(side="left", padx=6)
        self.date_end = DateEntry(birth_frame, date.today())
        self.date_end.pack(side="left")
        self._sync_birth_enabled()

        # 6. 실행 버튼
        self.btn_run = accent_button(self, "필터링 및 저장", self.run_filter, GREEN, GREEN_ACTIVE)
        self.btn_run.grid(row=row, column=0, sticky="ew", pady=(18, 6))
        row += 1

        self.rowconfigure(row, weight=1)

    def _sync_birth_enabled(self):
        enabled = self.use_birth_var.get()
        self.date_start.set_enabled(enabled)
        self.date_end.set_enabled(enabled)

    def browse_file(self):
        fname = filedialog.askopenfilename(parent=self, title="파일 선택", filetypes=EXCEL_FILETYPES)
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
            # 1. 입력값 검증
            self.period.validate("3. 내원일 기준 필터")
            self.period_old.validate("4. 내원일 기준 필터(오래된 사람만)")
            start_date = self.date_start.get_date() if self.use_birth_var.get() else None
            end_date = self.date_end.get_date() if self.use_birth_var.get() else None
            if start_date and start_date > end_date:
                raise ValueError("생년월일 필터의 시작일이 종료일보다 늦습니다.")

            # 2. 데이터 처리 실행
            df_result = outbound_list_filter(
                file_path=self.file_path,
                password=self.password_var.get(),
                period_type=self.period.period_type,
                period_type_old=self.period_old.period_type,
                period_value=self.period.period_value,
                period_value_old=self.period_old.period_value,
                use_birth=self.use_birth_var.get(),
                start_date=start_date,
                end_date=end_date,
            )

            # 3. 결과 저장 (resources/outbound_template.xlsx 양식 그대로)
            save_path = filedialog.asksaveasfilename(
                parent=self, title="결과 저장", defaultextension=".xlsx",
                initialfile="filtered_outbound.xlsx", filetypes=EXCEL_SAVE_FILETYPES,
            )
            if save_path:
                count = save_outbound_list(df_result, save_path)
                messagebox.showinfo("완료", f"필터링 완료!\n총 {count}건이 저장되었습니다.", parent=self)

        except Exception as e:
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{e}", parent=self)
        finally:
            self.btn_run.config(state="normal", text="필터링 및 저장")
