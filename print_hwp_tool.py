import os
import sys
import time
import threading
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32print
import win32com.client as win32

# ==========================================
# 1. Core Printing & HWP Automation Engine
# ==========================================

class HwpPrinterEngine:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback
        self.hwp = None
        self.printer_restore_info = None  # Stores (printer_name, original_duplex)
        
    def log(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def get_available_printers(self):
        """List all printer names available on the system."""
        try:
            printers = [p[2] for p in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )]
            return printers
        except Exception as e:
            self.log(f"[오류] 프린터 목록 조회 실패: {e}")
            return []

    def configure_duplex(self, printer_name, mode_str):
        """
        Configure printer duplex mode.
        mode_str options: 'default', 'simplex', 'long', 'short'
        Returns True if successful, False otherwise.
        """
        if mode_str == 'default':
            self.log(f"프린터 '{printer_name}'의 기본 설정을 유지합니다.")
            return True

        # Duplex mapping: 1 = Simplex, 2 = Duplex (Long edge), 3 = Duplex (Short edge)
        duplex_val = 1
        if mode_str == 'simplex':
            duplex_val = 1
            mode_korean = "단면 인쇄"
        elif mode_str == 'long':
            duplex_val = 2
            mode_korean = "양면 인쇄 (긴 쪽 넘김)"
        elif mode_str == 'short':
            duplex_val = 3
            mode_korean = "양면 인쇄 (짧은 쪽 넘김)"
        else:
            return False

        try:
            # DesiredAccess PRINTER_ALL_ACCESS is needed to change settings.
            # If the user lacks permission, it will throw Access Denied.
            print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            handle = win32print.OpenPrinter(printer_name, print_defaults)
        except Exception as e:
            self.log(f"[경고] 프린터 '{printer_name}' 설정을 변경할 권한이 없습니다. (기본 설정 사용)")
            self.log(f"       (오류 원인: {e})")
            return False

        try:
            level = 2
            info = win32print.GetPrinter(handle, level)
            devmode = info['pDevMode']
            
            if devmode is None:
                self.log(f"[경고] 프린터 '{printer_name}'의 DEVMODE 구조체를 가져올 수 없습니다. 설정을 변경하지 않습니다.")
                return False

            # Keep original duplex setting to restore later
            orig_duplex = devmode.Duplex
            self.printer_restore_info = (printer_name, orig_duplex)
            
            # Apply new duplex setting
            devmode.Duplex = duplex_val
            devmode.Fields |= win32print.DM_DUPLEX
            
            win32print.SetPrinter(handle, level, info, 0)
            self.log(f"[설정] 프린터 '{printer_name}'을(를) '{mode_korean}' 모드로 임시 설정했습니다.")
            return True
        except Exception as e:
            self.log(f"[경고] 프린터 설정을 적용하는 중 오류가 발생했습니다: {e}")
            self.printer_restore_info = None
            return False
        finally:
            win32print.ClosePrinter(handle)

    def restore_printer_settings(self):
        """Restore the printer to its original duplex setting."""
        if not self.printer_restore_info:
            return

        printer_name, orig_duplex = self.printer_restore_info
        try:
            print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            handle = win32print.OpenPrinter(printer_name, print_defaults)
            level = 2
            info = win32print.GetPrinter(handle, level)
            devmode = info['pDevMode']
            if devmode is not None:
                devmode.Duplex = orig_duplex
                win32print.SetPrinter(handle, level, info, 0)
                self.log(f"[설정] 프린터 '{printer_name}' 설정을 원래 상태(Duplex={orig_duplex})로 복구했습니다.")
        except Exception as e:
            self.log(f"[경고] 프린터 설정을 원래대로 복구하는 데 실패했습니다: {e}")
        finally:
            self.printer_restore_info = None

    def initialize_hwp(self):
        """Initialize HwpObject."""
        if self.hwp is not None:
            return True
        
        try:
            self.log("한글 프로그램(HwpObject)을 초기화하는 중...")
            # Use gencache to support autocomplete and correct object binding
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            
            # Register Security Module to bypass API warning popup
            # It's optional but highly recommended if the user has the module.
            # Usually, you can register a custom security module.
            # We register a common name; if it doesn't exist, HWP will just ignore it.
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            # Run in background (Visible=False)
            try:
                self.hwp.XHwpWindows.Item(0).Visible = False
            except Exception as ev:
                self.log(f"[알림] 한글 창 숨기기 설정 실패 (무시하고 진행): {ev}")
            self.log("한글 프로그램이 백그라운드에서 정상적으로 로드되었습니다.")
            return True
        except Exception as e:
            self.log(f"[오류] 한글 프로그램 초기화 실패: {e}")
            self.log("이유: 한컴오피스 한글이 설치되어 있지 않거나 pywin32 바인딩 오류입니다.")
            self.hwp = None
            return False

    def print_file(self, file_path, printer_name, copies=1):
        """Open, print, and close a single HWP/HWPX file."""
        if not self.hwp:
            if not self.initialize_hwp():
                return False

        abs_path = os.path.abspath(file_path)
        self.log(f"[진행] 파일 여는 중: {os.path.basename(abs_path)}")

        try:
            # Open the file
            # Open parameters: (path, format, arg)
            opened = self.hwp.Open(abs_path)
            if not opened:
                self.log(f"[오류] 파일을 열 수 없습니다: {os.path.basename(abs_path)} (보안 경고 팝업이 활성화되었거나 파일이 손상되었을 수 있습니다.)")
                return False

            # Create Print Action
            act = self.hwp.CreateAction("Print")
            pset = act.CreateSet()
            act.GetDefault(pset)

            # Set print configuration parameters
            pset.SetItem("PrinterName", printer_name)
            pset.SetItem("Collate", 1)  # Collate copies
            pset.SetItem("NumCopy", copies)  # Number of copies

            self.log(f"[진행] 인쇄 명령 전송 중... (프린터: {printer_name}, 부수: {copies})")
            
            # Execute print
            success = act.Execute(pset)
            if success:
                self.log(f"[완료] 인쇄 작업이 전송되었습니다: {os.path.basename(abs_path)}")
                # Give a small delay to let the print command safely spool
                time.sleep(1.0)
                return True
            else:
                self.log(f"[오류] 인쇄 실행 실패: {os.path.basename(abs_path)}")
                return False

        except Exception as e:
            self.log(f"[오류] 파일 처리 중 예외 발생 ({os.path.basename(abs_path)}): {e}")
            return False
        finally:
            try:
                # Clear document without saving changes (1: discard changes)
                self.hwp.Clear(1)
            except Exception:
                pass

    def save_as_pdf(self, file_path, output_dir):
        """Open a single HWP/HWPX file, save it as PDF to output_dir, and close it."""
        if not self.hwp:
            if not self.initialize_hwp():
                return False

        abs_path = os.path.abspath(file_path)
        base_name = os.path.splitext(os.path.basename(abs_path))[0]
        pdf_filename = f"{base_name}.pdf"
        pdf_path = os.path.join(output_dir, pdf_filename)
        abs_pdf_path = os.path.abspath(pdf_path)

        self.log(f"[진행] 파일 여는 중: {os.path.basename(abs_path)}")

        try:
            # Open the file
            opened = self.hwp.Open(abs_path)
            if not opened:
                self.log(f"[오류] 파일을 열 수 없습니다: {os.path.basename(abs_path)} (보안 경고 팝업이 활성화되었거나 파일이 손상되었을 수 있습니다.)")
                return False

            self.log(f"[진행] PDF 변환 및 저장 중... -> {pdf_filename}")
            
            # Save as PDF
            try:
                self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
                self.hwp.HParameterSet.HFileOpenSave.filename = abs_pdf_path
                self.hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                success = self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
            except Exception as e_action:
                self.log(f"[알림] HAction 변환 실패, SaveAs 메서드로 재시도합니다. ({e_action})")
                success = self.hwp.SaveAs(abs_pdf_path, "PDF", "")

            if success:
                self.log(f"[완료] PDF 변환 완료: {pdf_filename}")
                return True
            else:
                self.log(f"[오류] PDF 변환 실패: {os.path.basename(abs_path)}")
                return False

        except Exception as e:
            self.log(f"[오류] 파일 처리 중 예외 발생 ({os.path.basename(abs_path)}): {e}")
            return False
        finally:
            try:
                self.hwp.Clear(1)
            except Exception:
                pass

    def shutdown(self):
        """Quit HWP process."""
        if self.hwp:
            try:
                self.log("한글 프로그램을 종료합니다.")
                self.hwp.Quit()
            except Exception:
                pass
            self.hwp = None


# ==========================================
# 2. Modern Tkinter GUI Application
# ==========================================

class ModernHwpPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HWP / HWPX 일괄 인쇄 유틸리티")
        self.root.geometry("900x720")
        self.root.minsize(800, 600)
        
        self.engine = HwpPrinterEngine(self.write_log)
        self.selected_directory = ""
        self.files_list = []  # List of dicts: {"path": str, "name": str, "size": str, "checked": bool, "status": str, "id": str}
        self.is_printing = False
        # False = Ascending, True = Descending
        self.sort_directions = {"filename": False, "mtime": False, "ctime": False, "size": False}
        
        # Configure Styles
        self.setup_styles()
        
        # Build UI layout
        self.create_widgets()
        
        # Detect printers on startup
        self.load_printers()
        
        # Initial log
        self.write_log("시스템 준비 완료. 출력할 HWP/HWPX 파일이 있는 폴더를 선택해 주세요.")

    def setup_styles(self):
        """Define colors and widget styles using ttk."""
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Color Palette - Modern Navy & Slate Blue
        self.bg_color = "#f8fafc"       # Soft background
        self.card_color = "#ffffff"     # White cards
        self.primary_color = "#4f46e5"  # Indigo primary
        self.text_dark = "#1e293b"      # Dark text
        self.accent_color = "#10b981"    # Green accent for success
        
        self.root.configure(bg=self.bg_color)
        
        # Frame styles
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("Card.TFrame", background=self.card_color, relief="solid", borderwidth=1)
        
        # Label styles
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_dark, font=("Malgun Gothic", 10))
        self.style.configure("Header.TLabel", font=("Malgun Gothic", 14, "bold"), background=self.bg_color)
        self.style.configure("CardHeader.TLabel", font=("Malgun Gothic", 11, "bold"), background=self.card_color)
        self.style.configure("CardText.TLabel", background=self.card_color, font=("Malgun Gothic", 10))
        
        # Button styles
        self.style.configure("TButton", font=("Malgun Gothic", 10, "bold"), padding=6)
        self.style.configure("Primary.TButton", background=self.primary_color, foreground="white")
        self.style.map("Primary.TButton",
            background=[("active", "#4338ca"), ("disabled", "#cbd5e1")],
            foreground=[("disabled", "#94a3b8")]
        )
        
        self.style.configure("Accent.TButton", background=self.accent_color, foreground="white")
        self.style.map("Accent.TButton",
            background=[("active", "#059669"), ("disabled", "#cbd5e1")],
            foreground=[("disabled", "#94a3b8")]
        )
        
        # Combobox / Radiobutton styles
        self.style.configure("TRadiobutton", background=self.card_color, font=("Malgun Gothic", 10))
        self.style.configure("TCheckbutton", background=self.bg_color, font=("Malgun Gothic", 10))

        # Scrollbar / Treeview styles
        self.style.configure("Treeview", 
            font=("Malgun Gothic", 9), 
            rowheight=25, 
            background="#ffffff", 
            fieldbackground="#ffffff"
        )
        self.style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"), background="#e2e8f0", foreground=self.text_dark)

    def create_widgets(self):
        # Grid weights to make it responsive
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)  # File list expands
        
        # ------------------
        # Header Area
        # ------------------
        header_frame = ttk.Frame(self.root, padding=(20, 15, 20, 10))
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.columnconfigure(0, weight=1)
        
        header_label = ttk.Label(header_frame, text="📄 HWP / HWPX 일괄 인쇄 매니저", style="Header.TLabel")
        header_label.grid(row=0, column=0, sticky="w")
        
        subtitle_label = ttk.Label(header_frame, text="폴더 내의 한글 파일들을 선택하고 일괄 인쇄 및 단면/양면 설정을 조정할 수 있습니다.", font=("Malgun Gothic", 9))
        subtitle_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        # User Manual Button for beginners
        btn_help = ttk.Button(header_frame, text="💡 사용 방법 (도움말)", command=self.show_help_dialog)
        btn_help.grid(row=0, column=1, rowspan=2, sticky="e", padx=(10, 0))
        
        # ------------------
        # Security Warning Banner (Row 1)
        # ------------------
        banner_frame = tk.Frame(self.root, bg="#fef3c7", bd=1, relief="flat", highlightbackground="#f59e0b", highlightthickness=1)
        banner_frame.grid(row=1, column=0, padx=20, pady=(5, 5), sticky="ew")
        
        banner_label = tk.Label(
            banner_frame, 
            text="⚠️ [필수 설정] 한글 외부 앱 보안 경고 팝업 제거 방법\n"
                 "한글 실행 후 [도구] ➡️ [스크립트 매크로] ➡️ [매크로 보안 설정]에서 보안 수준을 \"낮음\"으로 설정해야 팝업 경고창 없이 자동 인쇄가 정상 작동합니다.", 
            font=("Malgun Gothic", 9, "bold"), 
            bg="#fef3c7", 
            fg="#b45309",
            justify="left",
            anchor="w",
            padx=15,
            pady=10
        )
        banner_label.pack(fill="x", expand=True)
        
        # ------------------
        # Top Panel: Folder Selection (Card 1, Row 2)
        # ------------------
        folder_card = ttk.Frame(self.root, padding=15, relief="solid", borderwidth=1)
        folder_card.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        folder_card.columnconfigure(1, weight=1)
        
        folder_label = ttk.Label(folder_card, text="대상 폴더 :", font=("Malgun Gothic", 10, "bold"))
        folder_label.grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        self.folder_path_var = tk.StringVar()
        self.folder_entry = ttk.Entry(folder_card, textvariable=self.folder_path_var, font=("Malgun Gothic", 10))
        self.folder_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        btn_browse = ttk.Button(folder_card, text="폴더 선택 🔍", command=self.browse_folder)
        btn_browse.grid(row=0, column=2, sticky="e")
        
        # ------------------
        # Middle Panel: Split Content (File Options & Printer Settings, Row 3)
        # ------------------
        options_frame = ttk.Frame(self.root, padding=(20, 5, 20, 5))
        options_frame.grid(row=3, column=0, sticky="nsew")
        options_frame.columnconfigure(0, weight=1)  # File table takes all remaining space
        options_frame.columnconfigure(1, weight=0, minsize=290)  # Options panel has a fixed minimum width to prevent text clipping
        options_frame.rowconfigure(0, weight=1)
        
        # Left Side of Middle Panel: File Table
        table_container = ttk.LabelFrame(options_frame, text=" 인쇄 대상 파일 목록 ", padding=10)
        table_container.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(1, weight=1)
        
        # Selection helper buttons & File count
        select_btns_frame = ttk.Frame(table_container)
        select_btns_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        btn_select_all = ttk.Button(select_btns_frame, text="전체 선택", width=10, command=self.select_all_files)
        btn_select_all.grid(row=0, column=0, padx=(0, 5))
        
        btn_deselect_all = ttk.Button(select_btns_frame, text="전체 해제", width=10, command=self.deselect_all_files)
        btn_deselect_all.grid(row=0, column=1, padx=(0, 5))
        
        self.file_count_var = tk.StringVar(value="검색된 파일: 0개")
        lbl_file_count = ttk.Label(select_btns_frame, textvariable=self.file_count_var, font=("Malgun Gothic", 9, "bold"))
        lbl_file_count.grid(row=0, column=2, sticky="e", padx=(10, 0))
        select_btns_frame.columnconfigure(2, weight=1)

        # Treeview Scrollbar
        tree_scroll_y = ttk.Scrollbar(table_container, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(table_container, orient="horizontal")
        
        # Treeview Table
        # Columns: Select, File Name, Modified Date, Created Date, Size, Status
        self.tree = ttk.Treeview(
            table_container, 
            columns=("checked", "filename", "mtime", "ctime", "size", "status"), 
            show="headings",
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        
        self.tree.heading("checked", text="선택")
        self.tree.heading("filename", text="파일명 ▲▼", command=lambda: self.sort_by_column("filename"))
        self.tree.heading("mtime", text="수정일 ▲▼", command=lambda: self.sort_by_column("mtime"))
        self.tree.heading("ctime", text="생성일 ▲▼", command=lambda: self.sort_by_column("ctime"))
        self.tree.heading("size", text="파일 크기 ▲▼", command=lambda: self.sort_by_column("size"))
        self.tree.heading("status", text="상태")
        
        self.tree.column("checked", width=50, minwidth=40, anchor="center")
        self.tree.column("filename", width=200, minwidth=150, anchor="w")
        self.tree.column("mtime", width=130, minwidth=100, anchor="center")
        self.tree.column("ctime", width=130, minwidth=100, anchor="center")
        self.tree.column("size", width=80, minwidth=60, anchor="e")
        self.tree.column("status", width=90, minwidth=80, anchor="center")
        
        self.tree.grid(row=1, column=0, sticky="nsew")
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_y.grid(row=1, column=1, sticky="ns")
        tree_scroll_x.config(command=self.tree.xview)
        tree_scroll_x.grid(row=2, column=0, sticky="ew")
        
        # Treeview double click or click to toggle checkbox
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        
        # Right Side of Middle Panel: Printer & Print Options
        settings_container = ttk.LabelFrame(options_frame, text=" 인쇄 및 프린터 설정 ", padding=15)
        settings_container.grid(row=0, column=1, sticky="nsew")
        settings_container.columnconfigure(0, weight=1)
        
        # Printer selection Dropdown
        ttk.Label(settings_container, text="🖨️ 프린터 선택", font=("Malgun Gothic", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.printer_var = tk.StringVar()
        self.printer_combobox = ttk.Combobox(settings_container, textvariable=self.printer_var, state="readonly", font=("Malgun Gothic", 9))
        self.printer_combobox.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        
        # Duplex settings (Simplex / Duplex options using ComboBox to save vertical space)
        ttk.Label(settings_container, text="📖 인쇄 모드 (단면/양면)", font=("Malgun Gothic", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(0, 5))
        
        self.duplex_modes = {
            "기본 프린터 설정 사용": "default",
            "단면 인쇄 (Simplex)": "simplex",
            "양면 인쇄 (긴 쪽 넘김)": "long",
            "양면 인쇄 (짧은 쪽 넘김)": "short"
        }
        
        self.duplex_var = tk.StringVar(value="기본 프린터 설정 사용")
        self.duplex_combobox = ttk.Combobox(settings_container, textvariable=self.duplex_var, values=list(self.duplex_modes.keys()), state="readonly", font=("Malgun Gothic", 9))
        self.duplex_combobox.grid(row=3, column=0, sticky="ew", pady=(0, 15))
        
        # Number of Copies
        ttk.Label(settings_container, text="👥 인쇄 부수 설정", font=("Malgun Gothic", 10, "bold")).grid(row=4, column=0, sticky="w", pady=(5, 5))
        
        copies_frame = ttk.Frame(settings_container)
        copies_frame.grid(row=5, column=0, sticky="w", pady=(0, 15))
        
        self.copies_var = tk.IntVar(value=1)
        self.copies_spinbox = ttk.Spinbox(copies_frame, from_=1, to=99, textvariable=self.copies_var, width=5, font=("Malgun Gothic", 9))
        self.copies_spinbox.grid(row=0, column=0, padx=(0, 5))
        ttk.Label(copies_frame, text="부").grid(row=0, column=1)

        # Action Buttons
        self.btn_start_print = ttk.Button(settings_container, text="인쇄 시작 ▶", style="Primary.TButton", command=self.start_print_job)
        self.btn_start_print.grid(row=6, column=0, sticky="ew", pady=(15, 5))
        
        self.btn_save_pdf = ttk.Button(settings_container, text="하위 output 폴더에 PDF로 저장 💾", style="Accent.TButton", command=self.start_pdf_job)
        self.btn_save_pdf.grid(row=7, column=0, sticky="ew", pady=(5, 5))
        
        # System Warning Note
        warning_note = "※ 시작 전 한글 프로그램의 경고 팝업창을 모두 닫아주세요."
        lbl_warning = ttk.Label(settings_container, text=warning_note, font=("Malgun Gothic", 8), foreground="#64748b", justify="left", wraplength=220)
        lbl_warning.grid(row=8, column=0, sticky="w", pady=(5, 0))

        # ------------------
        # Bottom Panel: Console Log & Progress Bar (Row 4)
        # ------------------
        bottom_frame = ttk.Frame(self.root, padding=(20, 5, 20, 20))
        bottom_frame.grid(row=4, column=0, sticky="ew")
        bottom_frame.columnconfigure(0, weight=1)
        
        # Progress Bar & Progress Label
        progress_info_frame = ttk.Frame(bottom_frame)
        progress_info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        progress_info_frame.columnconfigure(1, weight=1)
        
        self.progress_label_var = tk.StringVar(value="대기 중...")
        lbl_progress_status = ttk.Label(progress_info_frame, textvariable=self.progress_label_var, font=("Malgun Gothic", 9, "bold"))
        lbl_progress_status.grid(row=0, column=0, sticky="w")
        
        self.progress_percentage_var = tk.StringVar(value="0%")
        lbl_percentage = ttk.Label(progress_info_frame, textvariable=self.progress_percentage_var, font=("Malgun Gothic", 9, "bold"))
        lbl_percentage.grid(row=0, column=1, sticky="e")
        
        self.progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        # Log console
        ttk.Label(bottom_frame, text="실시간 진행 로그", font=("Malgun Gothic", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(0, 3))
        
        log_scroll = ttk.Scrollbar(bottom_frame)
        self.log_text = tk.Text(bottom_frame, height=4, bg="#0f172a", fg="#e2e8f0", 
                                insertbackground="white", font=("Consolas", 9), yscrollcommand=log_scroll.set)
        self.log_text.grid(row=3, column=0, sticky="ew")
        
        log_scroll.config(command=self.log_text.yview)
        log_scroll.grid(row=3, column=1, sticky="ns")
        
        # Prevent manual input in log window
        self.log_text.config(state="disabled")

    def show_help_dialog(self):
        """Show a friendly help popup window for beginners."""
        help_win = tk.Toplevel(self.root)
        help_win.title("💡 사용 방법 안내 (초보자용)")
        help_win.geometry("650x640")
        help_win.resizable(False, False)
        help_win.configure(bg="#f8fafc")
        
        # Center the popup window relative to root
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        x = root_x + (root_w - 650) // 2
        y = root_y + (root_h - 640) // 2
        help_win.geometry(f"650x640+{x}+{y}")
        
        # Make the dialog modal
        help_win.transient(self.root)
        help_win.grab_set()
        
        # Main Frame inside dialog
        main_frame = ttk.Frame(help_win, padding=20)
        main_frame.grid(row=0, column=0, sticky="nsew")
        help_win.columnconfigure(0, weight=1)
        help_win.rowconfigure(0, weight=1)
        
        # Title inside dialog
        title_lbl = ttk.Label(main_frame, text="💡 한글 일괄 인쇄 매니저 사용 방법", font=("Malgun Gothic", 12, "bold"))
        title_lbl.grid(row=0, column=0, sticky="w", pady=(0, 15))
        
        # Instructions list (Step Card)
        steps_card = ttk.LabelFrame(main_frame, text=" 단계별 사용 안내 ", padding=15)
        steps_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        steps_card.columnconfigure(0, weight=1)
        
        steps_text = (
            "1. 폴더 선택:\n"
            "   - 우측 상단의 [폴더 선택 🔍] 버튼을 누른 뒤, 출력하려는 한글 파일들이 모여있는 폴더를 선택합니다.\n\n"
            "2. 인쇄 대상 확인:\n"
            "   - 선택된 폴더 안의 한글 파일들이 왼쪽 목록에 표로 로드됩니다.\n"
            "   - 출력하지 않을 파일이 있다면 파일명 왼쪽의 체크박스(☑)를 클릭하여 해제(☐)할 수 있습니다.\n\n"
            "3. 프린터 및 옵션 선택:\n"
            "   - 우측 패널에서 사용할 프린터를 선택하고, 인쇄 모드(단면 / 양면 여부)와 인쇄할 부수를 지정합니다.\n\n"
            "4. 인쇄 시작:\n"
            "   - 파란색 [인쇄 시작 ▶] 버튼을 누르면 차례대로 인쇄가 진행됩니다. 진행 상황은 하단 로그창에서 확인 가능합니다."
        )
        
        lbl_steps = ttk.Label(steps_card, text=steps_text, font=("Malgun Gothic", 9), justify="left", wraplength=570)
        lbl_steps.grid(row=0, column=0, sticky="w")
        
        # Warning Card
        warning_card = ttk.LabelFrame(main_frame, text=" ⚠️ 중요: 외부 앱 보안 승인 팝업 해결 방법 ", padding=15)
        warning_card.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        warning_card.columnconfigure(0, weight=1)
        
        warning_intro = (
            "자동 인쇄를 시작할 때 한글 프로그램 창이 잠깐 켜지며 파일마다 '보안 경고' 또는\n"
            "'외부 앱 접근 허용 여부'를 묻는 팝업창이 뜰 수 있습니다.\n"
            "매번 [허용]을 누르지 않고 팝업창을 완전히 제거하려면 아래 설정을 적용해 주세요.\n\n"
            "📌 매크로 보안 수준 변경 방법 (추천)\n"
            '(이 팝업을 아예 안 뜨게 만들고 싶으시다면 한글 프로그램을 여신 뒤 '
            '[도구] ➡️ [스크립트 매크로] ➡️ [매크로 보안 설정]에서 '
            '보안 수준을 "낮음"으로 설정해 주시면 됩니다.)'
        )
        
        lbl_warning = ttk.Label(warning_card, text=warning_intro, font=("Malgun Gothic", 9), justify="left", foreground="#b45309", wraplength=570)
        lbl_warning.grid(row=0, column=0, sticky="w")
        
        # Close Button
        btn_close = ttk.Button(main_frame, text="확인 (닫기)", command=help_win.destroy, style="Primary.TButton")
        btn_close.grid(row=3, column=0, pady=(10, 0))

    def load_printers(self):
        """Detect and load available system printers into Combobox."""
        printers = self.engine.get_available_printers()
        if printers:
            self.printer_combobox['values'] = printers
            try:
                default_p = win32print.GetDefaultPrinter()
                if default_p in printers:
                    self.printer_combobox.set(default_p)
                else:
                    self.printer_combobox.current(0)
            except Exception:
                self.printer_combobox.current(0)
        else:
            self.printer_combobox['values'] = ["사용 가능한 프린터 없음"]
            self.printer_combobox.current(0)
            self.write_log("[경고] 연결된 프린터를 찾을 수 없습니다.")

    def browse_folder(self):
        """Open Folder selection dialog and search for HWP/HWPX files."""
        dir_path = filedialog.askdirectory()
        if not dir_path:
            return
            
        self.selected_directory = os.path.normpath(dir_path)
        self.folder_path_var.set(self.selected_directory)
        self.scan_files_in_folder()

    def sort_by_column(self, col_id):
        """Sort file list when a treeview column header is clicked."""
        if not self.files_list:
            return

        # Toggle direction (False = Ascending, True = Descending)
        reverse = self.sort_directions[col_id]
        self.sort_directions[col_id] = not reverse

        # Define sort key mapping
        if col_id == "filename":
            self.files_list.sort(key=lambda x: x["name"].lower(), reverse=reverse)
            sort_name = "파일명"
        elif col_id == "mtime":
            self.files_list.sort(key=lambda x: os.path.getmtime(x["path"]), reverse=reverse)
            sort_name = "수정일"
        elif col_id == "ctime":
            self.files_list.sort(key=lambda x: os.path.getctime(x["path"]), reverse=reverse)
            sort_name = "생성일"
        elif col_id == "size":
            self.files_list.sort(key=lambda x: x["size_bytes"], reverse=reverse)
            sort_name = "파일 크기"
        else:
            return

        # Update Treeview headers to show current direction
        self.update_header_indicators(col_id, reverse)

        # Clear existing items in Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Re-insert in sorted order
        for f in self.files_list:
            checked_char = "☑" if f["checked"] else "☐"
            item_id = self.tree.insert(
                "", 
                "end", 
                values=(checked_char, f["name"], f["mtime"], f["ctime"], f["size"], f["status"])
            )
            f["id"] = item_id

        dir_str = "내림차순" if reverse else "오름차순"
        self.write_log(f"파일 정렬 기준 적용: {sort_name} {dir_str}")

    def update_header_indicators(self, active_col, reverse):
        """Update column header text to show sorting direction arrows."""
        headers = {
            "filename": "파일명",
            "mtime": "수정일",
            "ctime": "생성일",
            "size": "파일 크기"
        }
        for col_id, base_text in headers.items():
            if col_id == active_col:
                arrow = " ▼" if reverse else " ▲"
                self.tree.heading(col_id, text=base_text + arrow)
            else:
                self.tree.heading(col_id, text=base_text + " ▲▼")

    def scan_files_in_folder(self):
        """Scan selected folder for HWP & HWPX files and update Treeview."""
        if not self.selected_directory or not os.path.isdir(self.selected_directory):
            self.write_log("[오류] 유효한 폴더 경로가 아닙니다.")
            return

        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.files_list.clear()

        # Reset sorting directions
        self.sort_directions = {"filename": False, "mtime": False, "ctime": False, "size": False}

        try:
            self.write_log(f"폴더 스캔 중: {self.selected_directory}")
            files = os.listdir(self.selected_directory)
            hwp_files = [f for f in files if f.lower().endswith(('.hwp', '.hwpx'))]
            
            for filename in hwp_files:
                full_path = os.path.join(self.selected_directory, filename)
                size_bytes = os.path.getsize(full_path)
                
                # Format file size for readability
                if size_bytes < 1024 * 1024:
                    size_str = f"{size_bytes / 1024:.1f} KB"
                else:
                    size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
                
                # Get and format creation and modification times
                mtime_epoch = os.path.getmtime(full_path)
                ctime_epoch = os.path.getctime(full_path)
                mtime_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime_epoch))
                ctime_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(ctime_epoch))
                
                # Default values
                status = "대기 중"
                
                self.files_list.append({
                    "id": "",  # Will be assigned by sort_by_column
                    "name": filename,
                    "path": full_path,
                    "mtime": mtime_str,
                    "ctime": ctime_str,
                    "size": size_str,
                    "size_bytes": size_bytes,
                    "checked": True,
                    "status": status
                })

            # Sort and populate Treeview (filename ascending by default)
            self.sort_by_column("filename")

            self.file_count_var.set(f"검색된 파일: {len(self.files_list)}개")
            self.write_log(f"스캔 완료: 총 {len(self.files_list)}개의 HWP/HWPX 파일을 발견했습니다.")
            self.reset_progress()
            
        except Exception as e:
            self.write_log(f"[오류] 폴더 스캔 중 오류 발생: {e}")

    def on_tree_click(self, event):
        """Toggle checkbox state when clicking the checkbox column."""
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        
        if not row_id or col != "#1":
            return

        # Find matching file in our model
        for f in self.files_list:
            if f["id"] == row_id:
                new_state = not f["checked"]
                f["checked"] = new_state
                check_char = "☑" if new_state else "☐"
                self.tree.set(row_id, "checked", check_char)
                break

    def select_all_files(self):
        """Check all files in the list."""
        for f in self.files_list:
            f["checked"] = True
            self.tree.set(f["id"], "checked", "☑")

    def deselect_all_files(self):
        """Uncheck all files in the list."""
        for f in self.files_list:
            f["checked"] = False
            self.tree.set(f["id"], "checked", "☐")

    def write_log(self, message):
        """Thread-safe log printing to the console text widget."""
        def append_text():
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
            
        if threading.current_thread() is threading.main_thread():
            append_text()
        else:
            self.root.after(0, append_text)

    def reset_progress(self):
        self.progress_bar["value"] = 0
        self.progress_label_var.set("대기 중...")
        self.progress_percentage_var.set("0%")

    def start_print_job(self):
        """Trigger print process. Runs in background thread to prevent UI freezing."""
        if self.is_printing:
            return

        selected_files = [f for f in self.files_list if f["checked"]]
        
        if not selected_files:
            messagebox.showwarning("경고", "인쇄할 파일을 하나 이상 선택해 주세요.")
            return

        printer_name = self.printer_var.get()
        if not printer_name or printer_name == "사용 가능한 프린터 없음":
            messagebox.showerror("오류", "유효한 프린터를 선택해 주세요.")
            return

        duplex_val = self.duplex_var.get()
        duplex_mode = self.duplex_modes.get(duplex_val, "default")
        copies = self.copies_var.get()
        
        if copies < 1:
            messagebox.showwarning("경고", "부수는 1부 이상이어야 합니다.")
            return

        confirm = messagebox.askyesno(
            "인쇄 시작 확인", 
            f"선택한 {len(selected_files)}개의 한글 파일을 인쇄하시겠습니까?\n"
            f"프린터: {printer_name}\n"
            f"인쇄 부수: {copies}부\n"
            f"인쇄 모드: {self.get_duplex_korean_name(duplex_mode)}"
        )
        if not confirm:
            return

        self.is_printing = True
        self.btn_start_print.config(state="disabled")
        self.btn_save_pdf.config(state="disabled")
        
        # Start background printing thread
        threading.Thread(
            target=self.bg_print_process,
            args=(selected_files, printer_name, duplex_mode, copies),
            daemon=True
        ).start()

    def get_duplex_korean_name(self, mode_str):
        if mode_str == "simplex": return "단면 인쇄"
        if mode_str == "long": return "양면 인쇄 (긴 쪽 넘김)"
        if mode_str == "short": return "양면 인쇄 (짧은 쪽 넘김)"
        return "기본 프린터 설정"

    def bg_print_process(self, selected_files, printer_name, duplex_mode, copies):
        """Background print handler run in thread."""
        total_files = len(selected_files)
        success_count = 0
        failed_count = 0
        
        self.write_log("=========================================")
        self.write_log(f"일괄 인쇄 작업을 시작합니다. (대상 파일 수: {total_files}개)")
        self.write_log("=========================================")

        # 1. Temporarily configure printer duplex settings
        duplex_modified = self.engine.configure_duplex(printer_name, duplex_mode)
        
        # 2. Launch Hancom Office Automation
        init_ok = self.engine.initialize_hwp()
        if not init_ok:
            self.write_log("[오류] 한글 자동화 엔진을 불러오지 못해 인쇄가 취소되었습니다.")
            if duplex_modified:
                self.engine.restore_printer_settings()
            
            self.root.after(0, self.on_print_job_finished, False, "한글 프로그램 로드 실패")
            return

        try:
            for idx, file_info in enumerate(selected_files):
                # Update UI status to "Printing"
                self.root.after(0, self.update_file_status, file_info["id"], "인쇄 중...")
                self.progress_label_var.set(f"인쇄 진행 중... ({idx+1}/{total_files})")
                
                # Execute printing
                ok = self.engine.print_file(file_info["path"], printer_name, copies)
                
                if ok:
                    success_count += 1
                    self.root.after(0, self.update_file_status, file_info["id"], "완료")
                else:
                    failed_count += 1
                    self.root.after(0, self.update_file_status, file_info["id"], "실패")
                
                # Update progress bar
                progress = int(((idx + 1) / total_files) * 100)
                self.progress_bar["value"] = progress
                self.progress_percentage_var.set(f"{progress}%")
                
                # Wait briefly between prints to let spooler handle the job
                time.sleep(0.5)

            # Wait a few seconds for final spooling before restoring printer settings
            if duplex_modified and self.engine.printer_restore_info:
                self.write_log("[대기] 남은 인쇄 데이터가 스풀러로 정상 전송되도록 잠시 대기합니다 (3초)...")
                time.sleep(3.0)

        finally:
            # 3. Restore printer settings
            if duplex_modified:
                self.engine.restore_printer_settings()
            
            # 4. Quit Hwp
            self.engine.shutdown()

        self.write_log("=========================================")
        self.write_log(f"일괄 인쇄 완료: 성공 {success_count}개, 실패 {failed_count}개")
        self.write_log("=========================================")
        
        self.root.after(0, self.on_print_job_finished, True, f"인쇄 완료: 성공 {success_count}, 실패 {failed_count}")

    def update_file_status(self, item_id, status_str):
        """Update the file status column in Treeview."""
        self.tree.set(item_id, "status", status_str)
        # Also update files_list
        for f in self.files_list:
            if f["id"] == item_id:
                f["status"] = status_str
                break

    def on_print_job_finished(self, success, summary_msg):
        self.is_printing = False
        self.btn_start_print.config(state="normal")
        self.btn_save_pdf.config(state="normal")
        self.progress_label_var.set("작업 완료")
        
        if success:
            messagebox.showinfo("인쇄 완료", f"일괄 인쇄 작업이 완료되었습니다.\n\n{summary_msg}")
        else:
            messagebox.showerror("인쇄 오류", f"인쇄 작업 중 오류가 발생했습니다.\n\n{summary_msg}")

    def start_pdf_job(self):
        """Trigger PDF conversion process. Runs in background thread to prevent UI freezing."""
        if self.is_printing:
            return

        selected_files = [f for f in self.files_list if f["checked"]]
        
        if not selected_files:
            messagebox.showwarning("경고", "PDF로 변환할 파일을 하나 이상 선택해 주세요.")
            return

        confirm = messagebox.askyesno(
            "PDF 변환 시작 확인", 
            f"선택한 {len(selected_files)}개의 한글 파일을 PDF로 변환하여 저장하시겠습니까?\n"
            f"저장 위치: 선택된 폴더 하위의 'output' 폴더"
        )
        if not confirm:
            return

        self.is_printing = True
        self.btn_start_print.config(state="disabled")
        self.btn_save_pdf.config(state="disabled")
        
        # Start background PDF conversion thread
        threading.Thread(
            target=self.bg_pdf_process,
            args=(selected_files,),
            daemon=True
        ).start()

    def bg_pdf_process(self, selected_files):
        """Background PDF conversion handler run in thread."""
        total_files = len(selected_files)
        success_count = 0
        failed_count = 0
        
        self.write_log("=========================================")
        self.write_log(f"PDF 일괄 변환 작업을 시작합니다. (대상 파일 수: {total_files}개)")
        self.write_log("=========================================")

        # 1. Create output folder
        output_dir = os.path.join(self.selected_directory, "output")
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.write_log(f"[설정] 변환 출력 폴더를 생성했습니다: {output_dir}")
        except Exception as e:
            self.write_log(f"[오류] 변환 출력 폴더 생성 실패: {e}")
            self.root.after(0, self.on_pdf_job_finished, False, f"출력 폴더 생성 실패 ({e})")
            return

        # 2. Launch Hancom Office Automation
        init_ok = self.engine.initialize_hwp()
        if not init_ok:
            self.write_log("[오류] 한글 자동화 엔진을 불러오지 못해 PDF 변환이 취소되었습니다.")
            self.root.after(0, self.on_pdf_job_finished, False, "한글 프로그램 로드 실패")
            return

        try:
            for idx, file_info in enumerate(selected_files):
                # Update UI status to "Converting"
                self.root.after(0, self.update_file_status, file_info["id"], "PDF 변환 중...")
                self.progress_label_var.set(f"PDF 변환 진행 중... ({idx+1}/{total_files})")
                
                # Execute conversion
                ok = self.engine.save_as_pdf(file_info["path"], output_dir)
                
                if ok:
                    success_count += 1
                    self.root.after(0, self.update_file_status, file_info["id"], "완료")
                else:
                    failed_count += 1
                    self.root.after(0, self.update_file_status, file_info["id"], "실패")
                
                # Update progress bar
                progress = int(((idx + 1) / total_files) * 100)
                self.progress_bar["value"] = progress
                self.progress_percentage_var.set(f"{progress}%")
                
                # Brief delay between jobs
                time.sleep(0.5)

        finally:
            # Quit Hwp
            self.engine.shutdown()

        self.write_log("=========================================")
        self.write_log(f"PDF 일괄 변환 완료: 성공 {success_count}개, 실패 {failed_count}개")
        self.write_log("=========================================")
        
        self.root.after(0, self.on_pdf_job_finished, True, f"변환 완료: 성공 {success_count}, 실패 {failed_count}")

    def on_pdf_job_finished(self, success, summary_msg):
        self.is_printing = False
        self.btn_start_print.config(state="normal")
        self.btn_save_pdf.config(state="normal")
        self.progress_label_var.set("작업 완료")
        
        if success:
            messagebox.showinfo("변환 완료", f"PDF 일괄 변환 작업이 완료되었습니다.\n\n{summary_msg}")
        else:
            messagebox.showerror("변환 오류", f"PDF 변환 작업 중 오류가 발생했습니다.\n\n{summary_msg}")


# ==========================================
# 3. Headless CLI mode and Script Main
# ==========================================

def run_cli(args):
    """Executes HWP print/PDF tasks in command line mode."""
    print("=========================================")
    if args.pdf:
        print(" HWP / HWPX 일괄 PDF 변환 CLI 도구")
    else:
        print(" HWP / HWPX 일괄 인쇄 CLI 도구")
    print("=========================================")
    
    dir_path = os.path.abspath(args.dir)
    if not os.path.isdir(dir_path):
        print(f"[오류] 폴더를 찾을 수 없습니다: {dir_path}")
        sys.exit(1)
        
    engine = HwpPrinterEngine()
    
    # Gather HWP / HWPX files
    files = [f for f in os.listdir(dir_path) if f.lower().endswith(('.hwp', '.hwpx'))]
    if not files:
        print(f"[알림] 폴더 내에 HWP/HWPX 파일이 존재하지 않습니다: {dir_path}")
        sys.exit(0)
        
    # Sort files according to CLI arguments
    if args.sort == "name_asc":
        files.sort(key=lambda x: x.lower())
    elif args.sort == "name_desc":
        files.sort(key=lambda x: x.lower(), reverse=True)
    elif args.sort == "mtime_asc":
        files.sort(key=lambda x: os.path.getmtime(os.path.join(dir_path, x)))
    elif args.sort == "mtime_desc":
        files.sort(key=lambda x: os.path.getmtime(os.path.join(dir_path, x)), reverse=True)
    elif args.sort == "ctime_asc":
        files.sort(key=lambda x: os.path.getctime(os.path.join(dir_path, x)))
    elif args.sort == "ctime_desc":
        files.sort(key=lambda x: os.path.getctime(os.path.join(dir_path, x)), reverse=True)

    if args.pdf:
        print(f"[정보] 총 {len(files)}개의 파일을 PDF로 변환합니다.")
        
        # 1. Create output folder
        output_dir = os.path.join(dir_path, "output")
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f"[설정] 변환 출력 폴더를 생성했습니다: {output_dir}")
        except Exception as e:
            print(f"[오류] 변환 출력 폴더 생성 실패: {e}")
            sys.exit(1)

        # 2. Initialize HWP
        if not engine.initialize_hwp():
            print("[오류] 한글 자동화 엔진을 초기화할 수 없습니다.")
            sys.exit(1)

        success_count = 0
        failed_count = 0

        try:
            for f in files:
                full_path = os.path.join(dir_path, f)
                print(f"-> [{files.index(f)+1}/{len(files)}] {f} PDF 변환 진행 중...")
                ok = engine.save_as_pdf(full_path, output_dir)
                if ok:
                    success_count += 1
                else:
                    failed_count += 1
                # Brief delay between jobs
                time.sleep(0.5)
        finally:
            engine.shutdown()

        print("=========================================")
        print(f"PDF 변환 완료: 성공 {success_count}개, 실패 {failed_count}개")
        print("=========================================")
        return

    # 1. Check printer
    printer_name = args.printer
    available_printers = engine.get_available_printers()
    
    if not available_printers:
        print("[오류] 시스템에 연결된 프린터가 없습니다.")
        sys.exit(1)
        
    if not printer_name:
        try:
            printer_name = win32print.GetDefaultPrinter()
            print(f"[설정] 기본 프린터를 사용합니다: {printer_name}")
        except Exception:
            printer_name = available_printers[0]
            print(f"[설정] 감지된 첫 번째 프린터를 사용합니다: {printer_name}")
    else:
        if printer_name not in available_printers:
            print(f"[경고] 지정된 프린터 '{printer_name}'를 시스템에서 찾을 수 없습니다.")
            print(f"       가용한 프린터 목록: {', '.join(available_printers)}")
            print("인쇄를 취소합니다.")
            sys.exit(1)

    print(f"[정보] 총 {len(files)}개의 파일을 인쇄합니다.")

    # 3. Apply Duplex settings
    duplex_modified = engine.configure_duplex(printer_name, args.duplex)

    # 4. Initialize HWP
    if not engine.initialize_hwp():
        print("[오류] 한글 자동화 엔진을 초기화할 수 없습니다.")
        if duplex_modified:
            engine.restore_printer_settings()
        sys.exit(1)

    success_count = 0
    failed_count = 0

    try:
        for f in files:
            full_path = os.path.join(dir_path, f)
            print(f"-> [{files.index(f)+1}/{len(files)}] {f} 인쇄 진행 중...")
            ok = engine.print_file(full_path, printer_name, args.copies)
            if ok:
                success_count += 1
            else:
                failed_count += 1
            # Brief delay between jobs
            time.sleep(0.5)
            
        if duplex_modified and engine.printer_restore_info:
            print("[대기] 인쇄 데이터가 스풀러로 전송 완료될 때까지 잠시 대기합니다 (3초)...")
            time.sleep(3.0)
            
    finally:
        # 5. Restore printer settings
        if duplex_modified:
            engine.restore_printer_settings()
        # 6. Shutdown engine
        engine.shutdown()

    print("=========================================")
    print(f"인쇄 완료: 성공 {success_count}개, 실패 {failed_count}개")
    print("=========================================")


def main():
    parser = argparse.ArgumentParser(description="HWP/HWPX Batch Printer Tool")
    parser.add_argument("--dir", type=str, help="HWP/HWPX 파일이 있는 폴더 경로 (CLI 모드 실행용)")
    parser.add_argument("--printer", type=str, help="사용할 프린터 이름 (기본값: 시스템 기본 프린터)")
    parser.add_argument("--duplex", type=str, choices=["default", "simplex", "long", "short"], default="default",
                        help="단면/양면 설정: default(기본값), simplex(단면), long(긴쪽양면), short(짧은쪽양면)")
    parser.add_argument("--copies", type=int, default=1, help="인쇄 부수 (기본값: 1)")
    parser.add_argument("--pdf", action="store_true", help="인쇄 대신 PDF 파일로 변환하여 'output' 폴더에 저장")
    parser.add_argument("--sort", type=str, choices=["name_asc", "name_desc", "mtime_asc", "mtime_desc", "ctime_asc", "ctime_desc"], default="name_asc",
                        help="CLI 출력 파일 정렬 기준: name_asc(이름 오름차순), name_desc(이름 내림차순), mtime_asc(수정일 오름차순), mtime_desc(수정일 내림차순), ctime_asc(생성일 오름차순), ctime_desc(생성일 내림차순)")
    
    args = parser.parse_args()
    
    # If a directory argument is provided, run in headless CLI mode.
    # Otherwise, launch the GUI.
    if args.dir:
        run_cli(args)
    else:
        root = tk.Tk()
        app = ModernHwpPrinterApp(root)
        
        # When GUI is closed, verify printer settings are clean
        def on_closing():
            if app.is_printing:
                if not messagebox.askyesno("경고", "인쇄가 진행 중입니다. 정말 종료하시겠습니까? (프린터 설정 복구가 실행되지 않을 수 있습니다)"):
                    return
            try:
                # Clean up printer and Hwp in case they are running
                if app.engine.printer_restore_info:
                    app.engine.restore_printer_settings()
                app.engine.shutdown()
            except Exception:
                pass
            root.destroy()
            
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()

if __name__ == "__main__":
    main()
