"""tkinter UI에서 공통으로 쓰는 스타일 상수와 위젯"""

import tkinter as tk
from datetime import date, datetime
from tkinter import ttk

FONT_FAMILY = "맑은 고딕"
FONT_BASE = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_SMALL = (FONT_FAMILY, 9)

GREEN = "#4CAF50"
GREEN_ACTIVE = "#3E8E41"
RED = "#E53935"
RED_ACTIVE = "#C62828"
GRAY_TEXT = "#777777"

NO_FILE = "선택된 파일 없음"

EXCEL_FILETYPES = [("Excel Files", "*.xlsx *.xls"), ("모든 파일", "*.*")]
EXCEL_SAVE_FILETYPES = [("Excel Files", "*.xlsx")]


def section_label(parent, text):
    """번호가 붙은 굵은 소제목 라벨"""
    return ttk.Label(parent, text=text, font=FONT_BOLD)


def accent_button(parent, text, command, color, active_color):
    """색상이 채워진 실행 버튼 (ttk는 Windows에서 배경색이 적용되지 않아 tk.Button 사용)"""
    return tk.Button(
        parent, text=text, command=command,
        bg=color, fg="white", activebackground=active_color, activeforeground="white",
        font=FONT_BOLD, relief="flat", borderwidth=0, cursor="hand2", height=2,
        disabledforeground="#E0E0E0",
    )


class PeriodFilter(ttk.Frame):
    """'숫자 + 전체/개월/년' 라디오 조합의 기간 필터 입력"""

    OPTIONS = ("전체", "개월", "년")

    def __init__(self, parent, default_value="6"):
        super().__init__(parent)
        self.type_var = tk.StringVar(value=self.OPTIONS[0])
        self.value_var = tk.StringVar(value=default_value)

        self.value_entry = ttk.Entry(self, textvariable=self.value_var, width=6,
                                     justify="center", font=FONT_BASE)
        self.value_entry.pack(side="left", padx=(0, 12))

        for option in self.OPTIONS:
            ttk.Radiobutton(self, text=option, value=option, variable=self.type_var,
                            command=self._sync_enabled).pack(side="left", padx=(0, 10))

        self._sync_enabled()

    def _sync_enabled(self):
        self.value_entry.state(["!disabled"] if self.type_var.get() != "전체" else ["disabled"])

    @property
    def period_type(self):
        return self.type_var.get()

    @property
    def period_value(self):
        return self.value_var.get().strip()

    def validate(self, label):
        """'전체'가 아닐 때 숫자가 입력됐는지 확인"""
        if self.period_type == "전체":
            return
        if not self.period_value.isdigit() or int(self.period_value) <= 0:
            raise ValueError(f"{label}: 기간에는 1 이상의 숫자를 입력하세요.")


class DateEntry(ttk.Frame):
    """YYYY-MM-DD 형식 날짜 입력 (표준 라이브러리만 사용)"""

    def __init__(self, parent, initial):
        super().__init__(parent)
        self.var = tk.StringVar(value=initial.strftime("%Y-%m-%d"))
        self.entry = ttk.Entry(self, textvariable=self.var, width=12,
                               justify="center", font=FONT_BASE)
        self.entry.pack()

    def set_enabled(self, enabled):
        self.entry.state(["!disabled"] if enabled else ["disabled"])

    def get_date(self):
        text = self.var.get().strip().replace("/", "-").replace(".", "-")
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError(f"날짜 형식이 올바르지 않습니다: '{self.var.get()}' (YYYY-MM-DD)")


def years_ago(years):
    """오늘로부터 n년 전 날짜 (2월 29일 보정 포함)"""
    today = date.today()
    try:
        return today.replace(year=today.year - years)
    except ValueError:
        return today.replace(year=today.year - years, day=28)
