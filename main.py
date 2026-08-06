import os
import sys
import tkinter as tk
from tkinter import ttk

APP_VERSION = "2.2.0"
APP_TITLE = f"아웃바운드 도구 v{APP_VERSION}"

SPLASH_BG = "#2E7D32"
SPLASH_SUB = "#A5D6A7"
FONT_FAMILY = "맑은 고딕"


def _enable_dpi_awareness():
    """고해상도 모니터에서 창이 흐리게 보이는 문제 방지 (실패해도 무방)"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def _create_desktop_shortcut_once():
    """최초 실행 시 바탕화면 바로가기 자동 생성 (PyInstaller 빌드 환경 전용)

    COM(IShellLink)을 직접 호출하므로 powershell.exe 등 외부 프로세스에 의존하지 않는다.
    바로가기 생성은 부가 기능이므로, 어떤 이유로 실패하더라도 프로그램 실행을 막지 않는다.
    """
    if not getattr(sys, "frozen", False):
        return  # 소스 코드 직접 실행 시 건너뜀

    try:
        from app_modules.win_shortcut import create_shortcut, get_desktop_dir

        shortcut_path = os.path.join(get_desktop_dir(), "아웃바운드 도구.lnk")
        if os.path.exists(shortcut_path):
            return  # 이미 바로가기가 있으면 건너뜀

        exe_path = sys.executable
        create_shortcut(
            shortcut_path=shortcut_path,
            target_path=exe_path,
            working_dir=os.path.dirname(exe_path),
            description=APP_TITLE,
        )
    except Exception:
        pass


def _show_splash(root):
    """무거운 모듈을 임포트하는 동안 표시할 로딩 화면"""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg=SPLASH_BG)
    splash.attributes("-topmost", True)

    width, height = 420, 160
    x = (splash.winfo_screenwidth() - width) // 2
    y = (splash.winfo_screenheight() - height) // 2
    splash.geometry(f"{width}x{height}+{x}+{y}")

    tk.Label(splash, text=APP_TITLE, bg=SPLASH_BG, fg="white",
             font=(FONT_FAMILY, 18, "bold")).pack(expand=True, pady=(34, 0))
    tk.Label(splash, text="로딩 중...", bg=SPLASH_BG, fg=SPLASH_SUB,
             font=(FONT_FAMILY, 11)).pack(expand=True, pady=(0, 28))

    splash.update()
    return splash


def _apply_styles(root):
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure(".", font=(FONT_FAMILY, 10))
    style.configure("TNotebook.Tab", font=(FONT_FAMILY, 10, "bold"), padding=(14, 7))


def main():
    # 최초 실행 시 바탕화면 바로가기 자동 생성
    _create_desktop_shortcut_once()
    _enable_dpi_awareness()

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("640x660")
    root.minsize(600, 560)
    root.withdraw()

    # 로딩 화면을 먼저 띄운 뒤 무거운 모듈을 임포트
    splash = _show_splash(root)
    _apply_styles(root)

    # 아래 임포트 시점에 app_modules 및 pandas·openpyxl 로딩이 발생
    from app_modules.ui_components.outbound_limit_ui import OutboundLimitApp
    from app_modules.ui_components.outbound_list_filter_ui import OutboundApp

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=8, pady=8)
    notebook.add(OutboundApp(notebook), text="아웃바운드 리스트 필터")
    notebook.add(OutboundLimitApp(notebook), text="아웃바운드 제한 리스트")
    notebook.enable_traversal()   # Ctrl+Tab 으로 탭 이동

    # 메인 윈도우가 준비되면 로딩 화면 닫기
    splash.destroy()
    root.deiconify()
    root.lift()

    root.mainloop()


if __name__ == "__main__":
    main()
