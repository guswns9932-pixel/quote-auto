"""Tkinter 기반 GUI: 검색 조건 입력, 진행 상태 표시, 결과 테이블을 담당한다."""

import csv
import math
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .scanner import scan
from .text_utils import build_pattern, parse_extension_filter, parse_keywords

APP_TITLE = "폴더 내부 키워드 검색기"
APP_VERSION = "2.0"

ACCENT_COLOR = "#2563eb"
ACCENT_COLOR_ACTIVE = "#1d4ed8"
MUTED_TEXT = "#6b7280"
STRIPE_COLOR = "#f3f4f6"
STATUS_COLORS = {
    "idle": "#374151",
    "running": ACCENT_COLOR,
    "done": "#15803d",
    "cancelled": "#b45309",
    "error": "#b91c1c",
}


def build_app_icon(size=32):
    """외부 이미지 파일 없이 돋보기 모양 아이콘을 직접 그린다."""
    image = tk.PhotoImage(width=size, height=size)
    image.put(ACCENT_COLOR, to=(0, 0, size, size))

    cx, cy = size * 0.42, size * 0.42
    outer_radius = size * 0.30
    inner_radius = size * 0.20

    for y in range(size):
        for x in range(size):
            distance = math.hypot(x - cx, y - cy)
            if inner_radius <= distance <= outer_radius:
                image.put("#ffffff", (x, y))

    handle_length = size * 0.34
    thickness = max(1, size // 16)
    angle = math.radians(45)
    steps = int(handle_length)
    for step in range(steps):
        radius = outer_radius + step
        hx = cx + radius * math.cos(angle)
        hy = cy + radius * math.sin(angle)
        for dx in range(-thickness, thickness + 1):
            for dy in range(-thickness, thickness + 1):
                px, py = int(hx + dx), int(hy + dy)
                if 0 <= px < size and 0 <= py < size:
                    image.put("#ffffff", (px, py))

    return image


def _segment_distance(px, py, x1, y1, x2, y2):
    dx, dy = x2 - x1, y2 - y1
    length_sq = dx * dx + dy * dy
    if length_sq == 0:
        return math.hypot(px - x1, py - y1)
    t = max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / length_sq))
    proj_x, proj_y = x1 + t * dx, y1 + t * dy
    return math.hypot(px - proj_x, py - proj_y)


def build_checkbox_images(size=15):
    """
    ttk 'clam' 테마의 체크박스가 체크 시 'X' 모양으로 그려지는 것을 대신할,
    실제 체크마크(✓) 모양의 인디케이터 이미지 두 장(선택 안 됨/선택됨)을 그린다.
    """
    border = "#9ca3af"
    unchecked = tk.PhotoImage(width=size, height=size)
    checked = tk.PhotoImage(width=size, height=size)

    unchecked.put("#ffffff", to=(0, 0, size, size))
    checked.put(ACCENT_COLOR, to=(0, 0, size, size))

    for i in range(size):
        for image in (unchecked, checked):
            image.put(border, (i, 0))
            image.put(border, (i, size - 1))
            image.put(border, (0, i))
            image.put(border, (size - 1, i))

    p1 = (size * 0.22, size * 0.52)
    p2 = (size * 0.42, size * 0.74)
    p3 = (size * 0.80, size * 0.26)
    thickness = max(1.2, size * 0.12)

    for y in range(size):
        for x in range(size):
            distance = min(
                _segment_distance(x + 0.5, y + 0.5, *p1, *p2),
                _segment_distance(x + 0.5, y + 0.5, *p2, *p3),
            )
            if distance <= thickness:
                checked.put("#ffffff", (x, y))

    return unchecked, checked


def open_path(path):
    """OS별 기본 연결 프로그램으로 파일을 연다."""
    path = str(path)
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class Tooltip:
    """마우스를 올리면 위젯 아래에 설명 박스를 띄우고, 벗어나면 없앤다.

    창 테두리/제목표시줄이 없는 Toplevel(overrideredirect)로 그려서
    일반 윈도우 창이 아니라 말풍선 같은 플로팅 박스처럼 보이게 한다.
    """

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<Destroy>", self._hide)

    def _show(self, event=None):
        if self.tip_window is not None or not self.text:
            return

        x = self.widget.winfo_rootx()
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        try:
            tw.attributes("-topmost", True)
        except tk.TclError:
            pass

        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#111827",
            foreground="#f9fafb",
            relief="solid",
            borderwidth=1,
            font=("Helvetica", 9),
            padx=10,
            pady=8,
        )
        label.pack()

    def _hide(self, event=None):
        if self.tip_window is not None:
            self.tip_window.destroy()
            self.tip_window = None


class ContentSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1280x760")
        self.minsize(980, 600)

        self._icon_image = build_app_icon()
        self.iconphoto(True, self._icon_image)

        self.task_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker = None
        self.result_paths = {}
        self._row_parity = 0
        self.folders = []

        self.keyword_var = tk.StringVar()
        self.exclude_keyword_var = tk.StringVar()
        self.extension_var = tk.StringVar()
        self.subfolder_var = tk.BooleanVar(value=True)
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.wildcard_var = tk.BooleanVar(value=False)
        self.filename_var = tk.BooleanVar(value=True)
        self.contents_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="폴더와 키워드를 입력하세요.")

        self._apply_style()
        self._build_menu()
        self._build_ui()
        self.after(100, self._process_queue)

    # ------------------------------------------------------------------
    # 스타일 / 레이아웃
    # ------------------------------------------------------------------
    def _apply_style(self):
        style = ttk.Style(self)
        available = style.theme_names()
        if "clam" in available:
            style.theme_use("clam")

        self.configure(background="#ffffff")

        style.configure("TFrame", background="#ffffff")
        style.configure("TLabelframe", background="#ffffff", bordercolor="#d1d5db")
        style.configure("TLabelframe.Label", background="#ffffff", font=("Helvetica", 10, "bold"))
        style.configure("TLabel", background="#ffffff", font=("Helvetica", 10))
        style.configure("Muted.TLabel", background="#ffffff", foreground=MUTED_TEXT, font=("Helvetica", 9))
        style.configure("TCheckbutton", background="#ffffff", font=("Helvetica", 10))
        style.configure("Header.TLabel", background="#ffffff", font=("Helvetica", 16, "bold"))
        style.configure("Sub.TLabel", background="#ffffff", foreground=MUTED_TEXT, font=("Helvetica", 10))

        style.configure("TButton", font=("Helvetica", 10), padding=6)
        style.configure(
            "Accent.TButton",
            font=("Helvetica", 10, "bold"),
            foreground="#ffffff",
            background=ACCENT_COLOR,
            padding=8,
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_COLOR_ACTIVE), ("disabled", "#93c5fd")],
            foreground=[("disabled", "#eff6ff")],
        )

        style.configure(
            "Treeview",
            font=("Helvetica", 10),
            rowheight=26,
            background="#ffffff",
            fieldbackground="#ffffff",
        )
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

        style.configure("TProgressbar", background=ACCENT_COLOR, troughcolor="#e5e7eb")

        self._checkbox_images = build_checkbox_images()
        unchecked_img, checked_img = self._checkbox_images
        style.element_create("Check.Checkbutton.indicator", "image", unchecked_img, ("selected", checked_img))
        style.layout("TCheckbutton", [
            ("Checkbutton.padding", {"sticky": "nswe", "children": [
                ("Check.Checkbutton.indicator", {"side": "left", "sticky": ""}),
                ("Checkbutton.focus", {"side": "left", "sticky": "", "children": [
                    ("Checkbutton.label", {"sticky": "nswe"}),
                ]}),
            ]}),
        ])

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="종료", command=self.destroy)
        menubar.add_cascade(label="파일", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="정보", command=self._show_about)
        menubar.add_cascade(label="도움말", menu=help_menu)

        self.config(menu=menubar)

    def _show_about(self):
        messagebox.showinfo(
            "프로그램 정보",
            f"{APP_TITLE}\n버전 {APP_VERSION}\n\n"
            "지정한 폴더의 파일명과 파일 내부 텍스트를 키워드 또는 와일드카드(*, ?)로 검색합니다."
        )

    def _build_ui(self):
        outer = ttk.Frame(self, padding=(16, 12))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="폴더 안의 파일명과 파일 내부 텍스트를 키워드 또는 와일드카드(*, ?)로 검색합니다.",
            style="Sub.TLabel",
        ).pack(anchor="w")

        search_frame = ttk.LabelFrame(outer, text="검색 조건", padding=12)
        search_frame.pack(fill="x")

        ttk.Label(search_frame, text="검색 폴더").grid(row=0, column=0, sticky="nw", padx=(0, 8), pady=6)

        folder_list_frame = ttk.Frame(search_frame)
        folder_list_frame.grid(row=0, column=1, sticky="ew", pady=6)

        self.folder_listbox = tk.Listbox(
            folder_list_frame,
            height=3,
            exportselection=False,
            background="#ffffff",
            foreground="#111827",
            selectbackground=ACCENT_COLOR,
            selectforeground="#ffffff",
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            font=("Helvetica", 10),
        )
        self.folder_listbox.pack(side="left", fill="both", expand=True)

        folder_scroll = ttk.Scrollbar(folder_list_frame, orient="vertical", command=self.folder_listbox.yview)
        folder_scroll.pack(side="left", fill="y")
        self.folder_listbox.configure(yscrollcommand=folder_scroll.set)

        folder_button_frame = ttk.Frame(search_frame)
        folder_button_frame.grid(row=0, column=2, sticky="n", padx=(8, 0), pady=6)
        ttk.Button(folder_button_frame, text="폴더 추가", command=self._add_folder).pack(fill="x")
        ttk.Button(folder_button_frame, text="선택 삭제", command=self._remove_selected_folder).pack(
            fill="x", pady=(4, 0)
        )
        ttk.Button(folder_button_frame, text="전체 삭제", command=self._clear_folders).pack(fill="x", pady=(4, 0))

        ttk.Label(search_frame, text="검색 키워드").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        keyword_entry = ttk.Entry(search_frame, textvariable=self.keyword_var)
        keyword_entry.grid(row=1, column=1, sticky="ew", pady=6)
        keyword_entry.bind("<Return>", lambda event: self.start_search())
        ttk.Label(search_frame, text="쉼표(,)로 구분해 여러 개 입력 가능", style="Muted.TLabel").grid(
            row=1, column=2, sticky="w", padx=(8, 0), pady=6
        )

        ttk.Label(search_frame, text="제외 키워드").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        exclude_keyword_entry = ttk.Entry(search_frame, textvariable=self.exclude_keyword_var)
        exclude_keyword_entry.grid(row=2, column=1, sticky="ew", pady=6)
        exclude_keyword_entry.bind("<Return>", lambda event: self.start_search())
        ttk.Label(
            search_frame,
            text="검색 키워드가 일치해도 이 단어가 같이 있으면 결과에서 제외",
            style="Muted.TLabel",
        ).grid(row=2, column=2, sticky="w", padx=(8, 0), pady=6)

        ttk.Label(search_frame, text="확장자 필터").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
        extension_entry = ttk.Entry(search_frame, textvariable=self.extension_var)
        extension_entry.grid(row=3, column=1, sticky="ew", pady=6)
        extension_entry.bind("<Return>", lambda event: self.start_search())
        ttk.Label(search_frame, text="예: txt, pdf (비워두면 전체)", style="Muted.TLabel").grid(
            row=3, column=2, sticky="w", padx=(8, 0), pady=6
        )

        option_frame = ttk.Frame(search_frame)
        option_frame.grid(row=4, column=1, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Checkbutton(option_frame, text="하위 폴더 포함", variable=self.subfolder_var).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(option_frame, text="대소문자 구분", variable=self.case_sensitive_var).pack(side="left", padx=(0, 14))
        wildcard_check = ttk.Checkbutton(
            option_frame, text="와일드카드 사용 (*, ?)", variable=self.wildcard_var
        )
        wildcard_check.pack(side="left", padx=(0, 14))
        Tooltip(
            wildcard_check,
            "* : 임의 길이의 문자열과 일치\n"
            "? : 문자 1개와 일치\n\n"
            "예시\n"
            "  invoice*2024  →  \"invoice 12월 2024\"에 일치\n"
            "  202?년  →  \"2024년\"에는 일치, \"20244년\"에는 불일치",
        )
        ttk.Checkbutton(option_frame, text="파일명 검색", variable=self.filename_var).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(option_frame, text="파일 내부 검색", variable=self.contents_var).pack(side="left")

        search_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(outer)
        button_frame.pack(fill="x", pady=12)

        self.search_button = ttk.Button(
            button_frame, text="검색 시작", command=self.start_search, style="Accent.TButton"
        )
        self.search_button.pack(side="left")

        self.stop_button = ttk.Button(button_frame, text="검색 중지", command=self.stop_search, state="disabled")
        self.stop_button.pack(side="left", padx=(8, 0))

        ttk.Button(button_frame, text="결과 초기화", command=self.clear_results).pack(side="left", padx=(8, 0))
        ttk.Button(button_frame, text="CSV 저장", command=self.export_csv).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(button_frame, mode="determinate")
        self.progress.pack(side="right", fill="x", expand=True, padx=(20, 0))

        result_frame = ttk.LabelFrame(outer, text="검색 결과", padding=8)
        result_frame.pack(fill="both", expand=True)

        columns = ("keyword", "name", "type", "location", "snippet", "path")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("keyword", text="키워드")
        self.tree.heading("name", text="파일명")
        self.tree.heading("type", text="형식")
        self.tree.heading("location", text="일치 위치")
        self.tree.heading("snippet", text="일치 내용")
        self.tree.heading("path", text="전체 경로")

        self.tree.column("keyword", width=110, minwidth=80)
        self.tree.column("name", width=220, minwidth=140)
        self.tree.column("type", width=70, minwidth=60, anchor="center")
        self.tree.column("location", width=150, minwidth=100)
        self.tree.column("snippet", width=430, minwidth=220)
        self.tree.column("path", width=420, minwidth=220)

        self.tree.tag_configure("odd", background=STRIPE_COLOR)
        self.tree.tag_configure("even", background="#ffffff")

        y_scroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda event: self.open_selected_file())
        self.tree.bind("<Return>", lambda event: self.open_selected_file())

        action_frame = ttk.Frame(outer)
        action_frame.pack(fill="x", pady=(10, 0))

        ttk.Button(action_frame, text="선택 파일 열기", command=self.open_selected_file).pack(side="left")
        ttk.Button(action_frame, text="선택 파일 폴더 열기", command=self.open_selected_folder).pack(side="left", padx=(8, 0))

        self.status_label = ttk.Label(action_frame, textvariable=self.status_var)
        self.status_label.pack(side="right")

    def _set_status(self, message, state="idle"):
        self.status_var.set(message)
        self.status_label.configure(foreground=STATUS_COLORS.get(state, STATUS_COLORS["idle"]))

    # ------------------------------------------------------------------
    # 동작
    # ------------------------------------------------------------------
    @staticmethod
    def _is_covered_by(path, ancestor):
        return path == ancestor or ancestor in path.parents

    def _add_folder(self):
        folder = filedialog.askdirectory(title="검색할 폴더 선택")
        if not folder:
            return

        new_path = Path(folder).resolve()

        if any(self._is_covered_by(new_path, existing) for existing in self.folders):
            messagebox.showinfo("이미 포함됨", "이미 선택된 폴더에 포함되는 하위 폴더입니다.")
            return

        # 새로 추가하는 폴더가 기존에 선택된 폴더의 상위 폴더라면, 그 하위 폴더는 중복이니 뺀다.
        self.folders = [f for f in self.folders if not self._is_covered_by(f, new_path)]
        self.folders.append(new_path)
        self._refresh_folder_listbox()

    def _remove_selected_folder(self):
        selection = self.folder_listbox.curselection()
        for index in reversed(selection):
            del self.folders[index]
        self._refresh_folder_listbox()

    def _clear_folders(self):
        self.folders = []
        self._refresh_folder_listbox()

    def _refresh_folder_listbox(self):
        self.folder_listbox.delete(0, "end")
        for folder in self.folders:
            self.folder_listbox.insert("end", str(folder))

    def clear_results(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("검색 중", "검색을 중지한 뒤 결과를 초기화하세요.")
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.result_paths.clear()
        self._row_parity = 0
        self._set_status("결과를 초기화했습니다.", "idle")
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress["value"] = 0

    def start_search(self):
        folders = list(self.folders)
        keywords = parse_keywords(self.keyword_var.get())

        if not folders:
            messagebox.showwarning("입력 확인", "검색할 폴더를 하나 이상 추가하세요.")
            return

        missing = [str(folder) for folder in folders if not folder.is_dir()]
        if missing:
            messagebox.showerror(
                "폴더 오류",
                "다음 폴더가 더 이상 존재하지 않습니다:\n\n" + "\n".join(missing),
            )
            return

        if not keywords:
            messagebox.showwarning("입력 확인", "검색 키워드를 입력하세요.")
            return

        if not self.filename_var.get() and not self.contents_var.get():
            messagebox.showwarning("검색 범위", "파일명 검색 또는 파일 내부 검색 중 하나 이상을 선택하세요.")
            return

        exclude_keywords = parse_keywords(self.exclude_keyword_var.get())

        try:
            pattern = build_pattern(
                keywords,
                self.wildcard_var.get(),
                self.case_sensitive_var.get()
            )
            exclude_pattern = None
            if exclude_keywords:
                exclude_pattern = build_pattern(
                    exclude_keywords,
                    self.wildcard_var.get(),
                    self.case_sensitive_var.get()
                )
        except re.error as exc:
            messagebox.showerror("검색어 오류", f"검색어를 처리할 수 없습니다.\n\n{exc}")
            return

        extensions = parse_extension_filter(self.extension_var.get())

        self.clear_results()
        self.cancel_event.clear()
        self.search_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self._set_status("검색을 시작합니다...", "running")
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)

        options = {
            "folders": folders,
            "pattern": pattern,
            "exclude_pattern": exclude_pattern,
            "include_subfolders": self.subfolder_var.get(),
            "search_filename": self.filename_var.get(),
            "search_contents": self.contents_var.get(),
            "extensions": extensions,
        }

        self.worker = threading.Thread(
            target=self._search_worker,
            args=(options,),
            daemon=True
        )
        self.worker.start()

    def stop_search(self):
        self.cancel_event.set()
        self._set_status("검색 중지 요청을 처리하는 중입니다...", "running")

    def _search_worker(self, options):
        errors = []

        def on_file_start(processed, path):
            self.task_queue.put(("progress", processed, str(path)))

        def on_file_result(path, location, keyword, snippet):
            self.task_queue.put((
                "result",
                path.name,
                path.suffix.lower() or "(없음)",
                location,
                keyword,
                snippet,
                str(path),
            ))

        def on_error(message):
            errors.append(message)

        try:
            processed, total_hits, file_count = scan(
                folders=options["folders"],
                pattern=options["pattern"],
                exclude_pattern=options["exclude_pattern"],
                include_subfolders=options["include_subfolders"],
                search_filename=options["search_filename"],
                search_contents=options["search_contents"],
                extensions=options["extensions"],
                is_cancelled=self.cancel_event.is_set,
                on_file_start=on_file_start,
                on_file_result=on_file_result,
                on_error=on_error,
            )

            if self.cancel_event.is_set():
                self.task_queue.put(("cancelled", processed, total_hits, file_count, errors))
            else:
                self.task_queue.put(("done", processed, total_hits, file_count, errors))

        except Exception as exc:
            self.task_queue.put(("fatal", str(exc), traceback.format_exc()))

    def _process_queue(self):
        try:
            while True:
                event = self.task_queue.get_nowait()
                kind = event[0]

                if kind == "progress":
                    index, path = event[1], event[2]
                    self._set_status(f"{index:,}개 확인 중: {Path(path).name}", "running")

                elif kind == "result":
                    _, name, ext, location, keyword, snippet, path = event
                    self._row_parity += 1
                    tag = "odd" if self._row_parity % 2 else "even"
                    iid = self.tree.insert(
                        "",
                        "end",
                        values=(keyword or "", name, ext, location, snippet, path),
                        tags=(tag,),
                    )
                    self.result_paths[iid] = path

                elif kind == "done":
                    processed, hit_count, file_count, errors = event[1], event[2], event[3], event[4]
                    self.search_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress["value"] = self.progress["maximum"]
                    self._set_status(
                        f"완료: {processed:,}개 파일 검사 / {file_count:,}개 파일 일치 / {hit_count:,}건 발견",
                        "done",
                    )
                    if errors:
                        self._show_error_summary(errors)

                elif kind == "cancelled":
                    processed, hit_count, file_count, errors = event[1], event[2], event[3], event[4]
                    self.search_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress["value"] = 0
                    self._set_status(
                        f"중지됨: {processed:,}개 검사 / {file_count:,}개 파일 일치 / {hit_count:,}건 발견",
                        "cancelled",
                    )
                    if errors:
                        self._show_error_summary(errors)

                elif kind == "fatal":
                    message, detail = event[1], event[2]
                    self.search_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress["value"] = 0
                    self._set_status("오류로 검색이 중단되었습니다.", "error")
                    messagebox.showerror("검색 오류", f"{message}\n\n{detail}")

        except queue.Empty:
            pass

        self.after(100, self._process_queue)

    def _show_error_summary(self, errors):
        preview = "\n".join(errors[:10])
        more = ""
        if len(errors) > 10:
            more = f"\n\n외 {len(errors) - 10:,}건"
        messagebox.showwarning(
            "일부 파일 검색 실패",
            "일부 파일은 손상, 암호화, 권한 또는 형식 문제로 검색하지 못했습니다.\n\n"
            + preview
            + more
        )

    def _selected_path(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("선택 필요", "검색 결과에서 파일을 선택하세요.")
            return None

        item_id = selection[0]
        return self.result_paths.get(item_id)

    def open_selected_file(self):
        path = self._selected_path()
        if not path:
            return

        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror("파일 열기 실패", str(exc))

    def open_selected_folder(self):
        path = self._selected_path()
        if not path:
            return

        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            else:
                open_path(Path(path).parent)
        except Exception as exc:
            messagebox.showerror("폴더 열기 실패", str(exc))

    def export_csv(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("저장할 결과 없음", "먼저 검색을 실행하세요.")
            return

        save_path = filedialog.asksaveasfilename(
            title="검색 결과 CSV 저장",
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv")]
        )
        if not save_path:
            return

        try:
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["키워드", "파일명", "형식", "일치 위치", "일치 내용", "전체 경로"])
                for item in items:
                    writer.writerow(self.tree.item(item, "values"))

            messagebox.showinfo("저장 완료", f"CSV 파일을 저장했습니다.\n\n{save_path}")
        except Exception as exc:
            messagebox.showerror("저장 실패", str(exc))
