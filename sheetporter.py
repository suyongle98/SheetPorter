# -*- coding: utf-8 -*-
"""SheetPorter 表格搬运工 · 图形界面

把 Excel 源表按列名搬进 Excel 模板：自动换算单位、自动加行、保护表头表尾。

用法：双击 SheetPorter.exe，或运行 python sheetporter.py
规则保存在程序目录的 rules 文件夹，每个 .json 一套规则。
"""
import os
import sys
import json
import queue
import shutil
import threading
import datetime
import traceback
from copy import deepcopy

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont

import fillcore

APP_NAME = "表格搬运工"
APP_EN = "SheetPorter"
APP_VERSION = "1.1.1"
APP_TITLE = "%s %s v%s" % (APP_NAME, APP_EN, APP_VERSION)

# ---------------------------------------------------------------- 路径

if getattr(sys, "frozen", False):
    # 打包后：exe 所在目录
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    # 源码运行：脚本所在目录
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RULES_DIR = os.path.join(BASE_DIR, "rules")


def ensure_examples():
    """打包版第一次运行时，把示例文件释放到程序目录，保证示例规则能直接跑。"""
    dst = os.path.join(BASE_DIR, "examples")
    src = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "examples")
    if os.path.isdir(dst) or os.path.abspath(src) == os.path.abspath(dst):
        return
    if os.path.isdir(src):
        shutil.copytree(src, dst)


def ensure_rules():
    """确保 rules 目录存在；第一次运行时放一套示例规则进去。"""
    ensure_examples()
    os.makedirs(RULES_DIR, exist_ok=True)
    if any(f.lower().endswith(".json") for f in os.listdir(RULES_DIR)):
        return
    src = os.path.join(BASE_DIR, "examples", "demo_rule.json")
    if os.path.exists(src):
        shutil.copy(src, os.path.join(RULES_DIR, "示例规则.json"))


# ---------------------------------------------------------------- 常量

TRANSFORMS = [
    ("直接搬", "direct"),
    ("取斜杠后（1.0000袋/盒 → 盒）", "after_slash"),
    ("取斜杠前", "before_slash"),
    ("去掉首尾空格", "strip"),
    ("固定值", "fixed"),
]
TRANSFORM_LABEL = {k: v for v, k in TRANSFORMS}
TRANSFORM_VALUE = {v: k for k, v in TRANSFORMS}

# 配色：浅底卡片 + 一个主色，办公软件里最不容易出错的一套
C_BG = "#F1F4F9"
C_CARD = "#FFFFFF"
C_BORDER = "#DCE3EE"
C_LINE = "#EDF1F7"
C_TEXT = "#1F2937"
C_TEXT2 = "#475569"
C_MUTED = "#8492A6"
C_PRIMARY = "#2563EB"
C_PRIMARY_D = "#1D4ED8"
C_PRIMARY_L = "#EFF4FF"
C_HEAD_BG = "#F8FAFC"
C_DARK = "#111C2E"
C_DARK2 = "#1B2A44"
C_OK = "#15803D"
C_WARN = "#B45309"
C_ERR = "#B91C1C"

# 目标列名 → 常见源列名（自动匹配时用；只放确实同义的）
ALIASES = {
    "药品编码": ["医保代码", "药品编码", "编码", "商品编码"],
    "生产厂家": ["生产企业", "生产厂家", "厂家", "生产厂商"],
    "单价": ["采购价", "单价", "价格", "进价"],
    "数量": ["采购数量", "数量"],
    "金额": ["采购金额", "金额", "总金额"],
    "单位": ["包装信息", "单位", "包装"],
    "规格": ["规格", "规格型号"],
    "药品名称": ["药品名称", "品名", "名称"],
}

# 运行期再按实际字体填充
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_SMALL = ("Microsoft YaHei UI", 9)


# ---------------------------------------------------------------- 工具函数

def pick_font(root):
    """挑一个本机存在的中文界面字体。"""
    try:
        fams = set(tkfont.families(root))
    except Exception:
        return "Microsoft YaHei UI"
    for f in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Arial"):
        if f in fams:
            return f
    return "Microsoft YaHei UI"


def read_headers(path, sheet, header_row):
    """读一张表的表头名列表。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        if isinstance(sheet, int):
            ws = wb.worksheets[sheet]
        elif sheet in (None, ""):
            ws = wb.worksheets[0]
        else:
            ws = wb[sheet]
        out = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(header_row, c).value
            if v is not None and str(v).strip():
                out.append(str(v).strip())
        return out
    finally:
        wb.close()


def as_sheet(v):
    v = str(v).strip()
    if v == "":
        return 0
    try:
        return int(v)
    except ValueError:
        return v


def to_rel(p):
    if not p:
        return ""
    try:
        r = os.path.relpath(p, BASE_DIR)
        if r.startswith(".."):
            return os.path.normpath(p)
        return r.replace("\\", "/")
    except Exception:
        return p


def open_folder(path):
    try:
        os.startfile(path)
    except Exception:
        pass


def enable_dpi_awareness():
    """让窗口在高分屏上清晰，而不是被系统整体拉伸变糊。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# ---------------------------------------------------------------- 主界面

class App:
    def __init__(self, root):
        self.root = root
        self.fam = pick_font(root)
        FONT_UI = (self.fam, 10)
        FONT_SMALL = (self.fam, 9)

        self.profile = None
        self.rule_path = None
        self.tasks = []
        self.rule_files = []
        self.dirty = False
        self.log_queue = queue.Queue()
        self.running = False
        self._loading = False

        root.title(APP_TITLE)
        root.configure(bg=C_BG)
        self._center(1300, 960)
        root.minsize(1100, 700)

        ensure_rules()
        self._setup_style()
        self._build_ui()
        self._bind_keys()
        self._reload_rules()
        root.after(200, self._poll_log)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- 窗口

    def _center(self, w, h):
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            # 小屏幕上不要超出可用区域
            w = min(w, sw - 60)
            h = min(h, sh - 90)
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2 - 24)
            self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
        except Exception:
            self.root.geometry("%dx%d" % (w, h))

    def _bind_keys(self):
        self.root.bind("<Control-s>", lambda e: (self._save_rule(), "break")[1])
        self.root.bind("<Control-n>", lambda e: (self._new_rule(), "break")[1])
        self.root.bind("<F5>", lambda e: (self._run(), "break")[1])

    def _on_close(self):
        if self.running and not messagebox.askyesno("正在生成", "还在生成文件，确定要退出吗？"):
            return
        if self.dirty:
            ans = messagebox.askyesnocancel("有未保存的修改", "规则「%s」有改动还没保存。\n\n要先保存吗？" % self.var_name.get())
            if ans is None:
                return
            if ans:
                self._save_rule(silent=True)
        self.root.destroy()

    # ---------------- 样式

    def _setup_style(self):
        fam = self.fam
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", font=(fam, 10), background=C_BG, foreground=C_TEXT)

        # 框架 / 标签
        style.configure("TFrame", background=C_BG)
        style.configure("Card.TFrame", background=C_CARD)
        style.configure("TLabel", background=C_BG, foreground=C_TEXT, font=(fam, 10))
        style.configure("Card.TLabel", background=C_CARD, foreground=C_TEXT, font=(fam, 10))
        style.configure("Muted.TLabel", background=C_CARD, foreground=C_MUTED, font=(fam, 9))
        style.configure("Field.TLabel", background=C_CARD, foreground=C_TEXT2, font=(fam, 9))
        style.configure("Tiny.TLabel", background=C_CARD, foreground=C_MUTED, font=(fam, 9))

        # 按钮：默认白底细边，主按钮蓝底白字
        style.configure(
            "TButton",
            background="#F7F9FC", foreground=C_TEXT2, bordercolor=C_BORDER,
            lightcolor="#F7F9FC", darkcolor="#F7F9FC", focuscolor="#F7F9FC",
            relief="flat", padding=(12, 6), font=(fam, 9),
        )
        style.map(
            "TButton",
            background=[("pressed", "#E6EDF8"), ("active", "#EFF4FB"), ("disabled", "#FBFCFE")],
            bordercolor=[("active", "#C8D5E8"), ("disabled", "#EAEFF6")],
            foreground=[("disabled", "#AFBACB")],
        )
        style.configure(
            "Accent.TButton",
            background=C_PRIMARY, foreground="#FFFFFF", bordercolor=C_PRIMARY,
            lightcolor=C_PRIMARY, darkcolor=C_PRIMARY, focuscolor=C_PRIMARY,
            relief="flat", padding=(22, 9), font=(fam, 10, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", "#A7C0F0"), ("pressed", C_PRIMARY_D), ("active", C_PRIMARY_D)],
            bordercolor=[("disabled", "#A7C0F0"), ("active", C_PRIMARY_D)],
            foreground=[("disabled", "#F2F6FF")],
        )
        style.configure("Ghost.TButton", background="#EAF1FF", foreground=C_PRIMARY_D,
                        bordercolor="#C9DAF8", lightcolor="#EAF1FF", darkcolor="#EAF1FF",
                        padding=(12, 6), font=(fam, 9))
        style.map("Ghost.TButton",
                  background=[("active", "#DCE8FF"), ("pressed", "#D0E0FF")],
                  bordercolor=[("active", "#A9C4F2")])
        style.configure("Danger.TButton", foreground="#B03A3A", padding=(12, 6), font=(fam, 9))
        style.map("Danger.TButton", foreground=[("active", C_ERR), ("pressed", C_ERR)])
        style.configure("Small.TButton", padding=(10, 4), font=(fam, 9))
        style.configure("Icon.TButton", padding=(6, 4), font=(fam, 10))

        # 输入框
        style.configure(
            "TEntry",
            fieldbackground=C_CARD, foreground=C_TEXT, bordercolor=C_BORDER,
            lightcolor=C_BORDER, darkcolor=C_BORDER, insertcolor=C_TEXT,
            relief="flat", padding=(8, 5),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", C_PRIMARY), ("disabled", "#EDF1F7")],
            lightcolor=[("focus", C_PRIMARY)],
            darkcolor=[("focus", C_PRIMARY)],
            fieldbackground=[("disabled", "#F7F9FC")],
        )
        style.configure(
            "TCombobox",
            fieldbackground=C_CARD, background=C_CARD, foreground=C_TEXT,
            bordercolor=C_BORDER, lightcolor=C_BORDER, darkcolor=C_BORDER,
            arrowcolor=C_TEXT2, padding=(8, 4), relief="flat",
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", C_CARD), ("disabled", "#F7F9FC")],
            bordercolor=[("focus", C_PRIMARY), ("hover", "#C8D5E8")],
            arrowcolor=[("active", C_PRIMARY)],
        )

        # 表格
        style.configure(
            "Treeview",
            background=C_CARD, fieldbackground=C_CARD, foreground=C_TEXT,
            bordercolor=C_CARD, lightcolor=C_CARD, darkcolor=C_CARD,
            borderwidth=0, relief="flat", rowheight=27, font=(fam, 9),
        )
        style.map(
            "Treeview",
            background=[("selected", C_PRIMARY_L)],
            foreground=[("selected", C_PRIMARY_D)],
        )
        style.configure(
            "Treeview.Heading",
            background="#F1F5FB", foreground="#5B6B82", relief="flat",
            borderwidth=0, padding=(9, 7), font=(fam, 9, "bold"),
        )
        style.map("Treeview.Heading", background=[("active", "#EFF3F9")])

        # 滚动条 / 进度条 / 分隔条
        style.configure(
            "Vertical.TScrollbar",
            background="#D3DBE8", troughcolor="#F6F8FC", bordercolor="#F6F8FC",
            arrowcolor="#8B99AD", gripcount=0, relief="flat", arrowsize=12,
        )
        style.map("Vertical.TScrollbar", background=[("active", "#B9C5D8")])
        style.configure(
            "Thin.Horizontal.TProgressbar",
            background=C_PRIMARY, troughcolor="#E6ECF5", bordercolor="#E6ECF5",
            lightcolor=C_PRIMARY, darkcolor=C_PRIMARY, thickness=5,
        )
        style.configure("TPanedwindow", background=C_BG)
        style.configure("Sash", sashthickness=10, gripcount=0, background=C_BG)

        # 下拉列表（原生 Listbox）
        self.root.option_add("*TCombobox*Listbox.background", C_CARD)
        self.root.option_add("*TCombobox*Listbox.foreground", C_TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", C_PRIMARY)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
        self.root.option_add("*TCombobox*Listbox.font", (fam, 9))

    # ---------------- 界面骨架

    def _card(self, parent, title, hint=""):
        """白色卡片：标题行 + 内容区。返回 (卡片, 标题行, 内容区)。"""
        card = tk.Frame(parent, bg=C_CARD, highlightbackground=C_BORDER,
                        highlightthickness=1, bd=0)
        head = tk.Frame(card, bg=C_CARD)
        head.pack(fill="x", padx=16, pady=(12, 0))
        tk.Label(head, text=title, bg=C_CARD, fg=C_TEXT,
                 font=(self.fam, 10, "bold")).pack(side="left")
        if hint:
            tk.Label(head, text=hint, bg=C_CARD, fg=C_MUTED,
                     font=(self.fam, 9)).pack(side="left", padx=(10, 0))
        body = tk.Frame(card, bg=C_CARD)
        body.pack(fill="both", expand=True, padx=16, pady=(6, 10))
        return card, head, body

    def _build_ui(self):
        root = self.root
        fam = self.fam

        # ===== 底部状态栏（先占位，保证永远贴底） =====
        self.status_var = tk.StringVar(value="就绪")
        status = tk.Frame(root, bg=C_BG)
        status.pack(fill="x", side="bottom")
        tk.Label(status, textvariable=self.status_var, bg=C_BG, fg=C_MUTED,
                 font=(fam, 9)).pack(side="left", padx=18, pady=(3, 5))
        tk.Label(status, text="本地处理 · 不联网 · 不上传", bg=C_BG, fg="#A6B2C4",
                 font=(fam, 9)).pack(side="right", padx=18, pady=(3, 5))

        # ===== 顶部品牌栏 =====
        header = tk.Frame(root, bg=C_DARK, height=68)
        header.pack(fill="x")
        header.pack_propagate(False)
        hb = tk.Frame(header, bg=C_DARK)
        hb.pack(fill="both", expand=True, padx=20)

        tk.Label(hb, text="⇄", bg=C_DARK, fg="#5EA0FF",
                 font=(fam, 19, "bold")).pack(side="left")
        title_box = tk.Frame(hb, bg=C_DARK)
        title_box.pack(side="left", padx=(10, 0))
        tk.Label(title_box, text=APP_NAME, bg=C_DARK, fg="#FFFFFF",
                 font=(fam, 15, "bold")).pack(anchor="w")
        tk.Label(title_box, text="按列名把 Excel 源表搬进模板 · 纯本地", bg=C_DARK,
                 fg="#8AA0C0", font=(fam, 9)).pack(anchor="w")

        self.pill = tk.Label(hb, text="●  就绪", bg=C_DARK2, fg="#8FD3FF",
                             font=(fam, 9), padx=14, pady=5)
        self.pill.pack(side="right")
        tk.Label(hb, text="v%s" % APP_VERSION, bg=C_DARK, fg="#5C7295",
                 font=(fam, 9)).pack(side="right", padx=(0, 14))

        # ===== 规则工具条 =====
        bar = tk.Frame(root, bg=C_CARD)
        bar.pack(fill="x")
        bb = tk.Frame(bar, bg=C_CARD)
        bb.pack(fill="x", padx=16, pady=9)

        tk.Label(bb, text="当前规则", bg=C_CARD, fg=C_MUTED,
                 font=(fam, 9)).pack(side="left", padx=(0, 8))
        self.rule_var = tk.StringVar()
        self.rule_box = ttk.Combobox(bb, textvariable=self.rule_var,
                                     state="readonly", width=38, font=(fam, 9))
        self.rule_box.pack(side="left")
        self.rule_box.bind("<<ComboboxSelected>>", self._on_rule_change)

        tk.Frame(bb, bg=C_BORDER, width=1, height=22).pack(side="left", padx=14)

        for text, cmd, st in [
            ("＋ 新建", self._new_rule, "TButton"),
            ("复制", self._copy_rule, "TButton"),
            ("导入", self._import_rule, "TButton"),
            ("删除", self._delete_rule, "Danger.TButton"),
        ]:
            ttk.Button(bb, text=text, command=cmd, style=st).pack(side="left", padx=(0, 6))

        ttk.Button(bb, text="保存规则", command=self._save_rule,
                   style="Ghost.TButton").pack(side="right")

        tk.Frame(root, bg=C_BORDER, height=1).pack(fill="x")

        # ===== 主体（最后 pack，占满中间） =====
        main = tk.Frame(root, bg=C_BG)

        # ---- 规则设置卡片：左右两栏，每栏两行 ----
        set_card, _, set_body = self._card(main, "规则设置", "模板、输出位置与生成日期")
        set_card.pack(fill="x")

        cols = tk.Frame(set_body, bg=C_CARD)
        cols.pack(fill="x")
        cols.columnconfigure(0, weight=1, uniform="col")
        cols.columnconfigure(1, weight=1, uniform="col")
        lcol = tk.Frame(cols, bg=C_CARD)
        lcol.grid(row=0, column=0, sticky="we", padx=(0, 22))
        rcol = tk.Frame(cols, bg=C_CARD)
        rcol.grid(row=0, column=1, sticky="we")

        def row(parent, label, width=11):
            f = tk.Frame(parent, bg=C_CARD)
            f.pack(fill="x", pady=1)
            ttk.Label(f, text=label, style="Field.TLabel",
                      width=width, anchor="w").pack(side="left")
            return f

        self.var_name = tk.StringVar()
        self.var_tpl = tk.StringVar()
        self.var_outdir = tk.StringVar()
        self.var_outname = tk.StringVar()
        self.var_date = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))

        f = row(lcol, "规则名称")
        ttk.Entry(f, textvariable=self.var_name).pack(side="left", fill="x", expand=True)

        f = row(lcol, "模板文件")
        ttk.Entry(f, textvariable=self.var_tpl).pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="浏览…", style="Small.TButton",
                   command=self._pick_template).pack(side="left", padx=(8, 0))

        f = row(rcol, "输出文件名")
        ttk.Entry(f, textvariable=self.var_outname).pack(side="left", fill="x", expand=True)
        ttk.Label(f, text="日期", style="Field.TLabel").pack(side="left", padx=(12, 6))
        ttk.Entry(f, textvariable=self.var_date, width=12).pack(side="left")

        f = row(rcol, "输出目录")
        ttk.Entry(f, textvariable=self.var_outdir).pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="浏览…", style="Small.TButton",
                   command=self._pick_outdir).pack(side="left", padx=(8, 0))

        # ---- 任务 / 详情 ----
        paned = ttk.PanedWindow(main, orient="horizontal")
        paned.pack(fill="both", expand=True, pady=(14, 0))

        left_card, left_head, left_body = self._card(paned, "任务列表")
        self.task_count = tk.Label(left_head, text="", bg=C_CARD, fg=C_MUTED, font=(fam, 9))
        self.task_count.pack(side="right")
        paned.add(left_card, weight=2)

        tbtn = tk.Frame(left_body, bg=C_CARD)
        tbtn.pack(side="bottom", fill="x", pady=(10, 0))
        ttk.Button(tbtn, text="＋ 添加", style="Small.TButton",
                   command=self._add_task).pack(side="left")
        ttk.Button(tbtn, text="－ 删除", style="Small.TButton",
                   command=self._del_task).pack(side="left", padx=(6, 0))
        ttk.Button(tbtn, text="↑", style="Icon.TButton", width=3,
                   command=lambda: self._move_task(-1)).pack(side="left", padx=(14, 0))
        ttk.Button(tbtn, text="↓", style="Icon.TButton", width=3,
                   command=lambda: self._move_task(1)).pack(side="left", padx=(4, 0))

        tv_wrap = tk.Frame(left_body, bg=C_CARD)
        tv_wrap.pack(fill="both", expand=True)
        self.task_tree = ttk.Treeview(tv_wrap, columns=("name", "src", "dst"),
                                      show="headings", height=6)
        self.task_tree.heading("name", text="任务")
        self.task_tree.heading("src", text="源文件")
        self.task_tree.heading("dst", text="目标 sheet")
        self.task_tree.column("name", width=100, minwidth=70)
        self.task_tree.column("src", width=150, minwidth=90)
        self.task_tree.column("dst", width=100, minwidth=70)
        self.task_tree.tag_configure("odd", background="#FAFBFE")
        tsb = ttk.Scrollbar(tv_wrap, orient="vertical", command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=tsb.set)
        tsb.pack(side="right", fill="y")
        self.task_tree.pack(side="left", fill="both", expand=True)
        self.task_tree.bind("<<TreeviewSelect>>", self._on_task_select)

        right = ttk.Frame(paned, style="TFrame")
        paned.add(right, weight=5)

        det_card, _, det_body = self._card(
            right, "任务详情", "一个任务 = 一张源表 + 一个 sheet · 固定单元格写 A1=内容")
        det_card.pack(fill="x")
        form = det_body

        self.v_task_name = tk.StringVar()
        self.v_src_dir = tk.StringVar()
        self.v_src_pattern = tk.StringVar()
        self.v_src_sheet = tk.StringVar(value="0")
        self.v_src_header = tk.StringVar(value="1")
        self.v_src_data = tk.StringVar(value="2")
        self.v_dst_sheet = tk.StringVar()
        self.v_dst_header = tk.StringVar(value="2")
        self.v_dst_data = tk.StringVar(value="3")
        self.v_cells = tk.StringVar()

        dgrid = tk.Frame(form, bg=C_CARD)
        dgrid.pack(fill="x")
        dgrid.columnconfigure(0, weight=1, uniform="dcol")
        dgrid.columnconfigure(1, weight=1, uniform="dcol")
        c1 = tk.Frame(dgrid, bg=C_CARD)
        c1.grid(row=0, column=0, sticky="nwe", padx=(0, 20))
        c2 = tk.Frame(dgrid, bg=C_CARD)
        c2.grid(row=0, column=1, sticky="nwe")

        f = row(c1, "任务名称", 10)
        ttk.Entry(f, textvariable=self.v_task_name).pack(side="left", fill="x", expand=True)

        f = row(c1, "文件名匹配", 10)
        ttk.Entry(f, textvariable=self.v_src_pattern).pack(side="left", fill="x", expand=True)

        f = row(c1, "源表 sheet", 10)
        ttk.Entry(f, textvariable=self.v_src_sheet, width=6).pack(side="left")
        ttk.Label(f, text="表头行", style="Field.TLabel").pack(side="left", padx=(16, 5))
        ttk.Entry(f, textvariable=self.v_src_header, width=4).pack(side="left")
        ttk.Label(f, text="数据起始行", style="Field.TLabel").pack(side="left", padx=(16, 5))
        ttk.Entry(f, textvariable=self.v_src_data, width=4).pack(side="left")

        f = row(c2, "源文件夹", 10)
        ttk.Entry(f, textvariable=self.v_src_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(f, text="浏览…", style="Small.TButton",
                   command=self._pick_src_dir).pack(side="left", padx=(6, 0))

        f = row(c2, "目标 sheet", 10)
        self.dst_sheet_box = ttk.Combobox(f, textvariable=self.v_dst_sheet,
                                          width=12, font=(fam, 9))
        self.dst_sheet_box.pack(side="left")
        ttk.Label(f, text="表头行", style="Field.TLabel").pack(side="left", padx=(16, 5))
        ttk.Entry(f, textvariable=self.v_dst_header, width=4).pack(side="left")
        ttk.Label(f, text="数据起始行", style="Field.TLabel").pack(side="left", padx=(16, 5))
        ttk.Entry(f, textvariable=self.v_dst_data, width=4).pack(side="left")

        f = row(c2, "固定单元格", 10)
        ttk.Entry(f, textvariable=self.v_cells).pack(side="left", fill="x", expand=True)

        # ---- 列映射 ----
        map_card, map_head, map_body = self._card(right, "列映射", "双击一行修改")
        map_card.pack(fill="both", expand=True, pady=(14, 0))

        mbtn = tk.Frame(map_head, bg=C_CARD)
        mbtn.pack(side="right")
        ttk.Button(mbtn, text="自动匹配", style="Small.TButton",
                   command=self._auto_match).pack(side="left", padx=(0, 6))
        ttk.Button(mbtn, text="添加", style="Small.TButton",
                   command=self._add_mapping).pack(side="left", padx=(0, 6))
        ttk.Button(mbtn, text="修改", style="Small.TButton",
                   command=self._edit_mapping).pack(side="left", padx=(0, 6))
        ttk.Button(mbtn, text="删除", style="Small.TButton",
                   command=self._del_mapping).pack(side="left")

        mw = tk.Frame(map_body, bg=C_CARD)
        mw.pack(fill="both", expand=True)
        self.map_tree = ttk.Treeview(mw, columns=("dst", "src", "tf"),
                                     show="headings", height=5)
        self.map_tree.heading("dst", text="目标列")
        self.map_tree.heading("src", text="源列")
        self.map_tree.heading("tf", text="转换")
        self.map_tree.column("dst", width=170, minwidth=100)
        self.map_tree.column("src", width=200, minwidth=100)
        self.map_tree.column("tf", width=240, minwidth=120)
        self.map_tree.tag_configure("odd", background="#FAFBFE")
        msb = ttk.Scrollbar(mw, orient="vertical", command=self.map_tree.yview)
        self.map_tree.configure(yscrollcommand=msb.set)
        msb.pack(side="right", fill="y")
        self.map_tree.pack(side="left", fill="both", expand=True)
        self.map_tree.bind("<Double-1>", lambda e: self._edit_mapping())

        # ===== 操作条 =====
        action = tk.Frame(root, bg=C_BG)
        self.run_btn = ttk.Button(action, text="▶   开始生成", style="Accent.TButton",
                                  command=self._run)
        self.run_btn.pack(side="left")
        ttk.Button(action, text="打开输出文件夹", style="TButton",
                   command=self._open_out).pack(side="left", padx=(10, 0))
        self.progress = ttk.Progressbar(action, mode="determinate", value=0,
                                        style="Thin.Horizontal.TProgressbar", length=200)
        self.progress.pack(side="left", padx=(18, 0), pady=(6, 0))
        tk.Label(action, text="快捷键 F5 生成 · Ctrl+S 保存", bg=C_BG, fg=C_MUTED,
                 font=(fam, 9)).pack(side="right")

        # ===== 日志 =====
        log_card, log_head, log_body = self._card(root, "运行日志")
        ttk.Button(log_head, text="清空", style="Small.TButton",
                   command=self._clear_log).pack(side="right")

        lw = tk.Frame(log_body, bg=C_CARD)
        lw.pack(fill="both", expand=True)
        self.log_text = tk.Text(lw, height=4, font=(fam, 9), wrap="word",
                                bg="#FBFCFE", fg=C_TEXT2, relief="flat",
                                highlightthickness=1, highlightbackground=C_LINE,
                                highlightcolor=C_LINE, padx=10, pady=7,
                                insertbackground=C_TEXT, spacing1=1, spacing3=2)
        lsb = ttk.Scrollbar(lw, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_text.tag_configure("time", foreground="#A6B2C4")
        self.log_text.tag_configure("info", foreground=C_TEXT2)
        self.log_text.tag_configure("detail", foreground="#7B8AA0")
        self.log_text.tag_configure("ok", foreground=C_OK)
        self.log_text.tag_configure("warn", foreground=C_WARN)
        self.log_text.tag_configure("error", foreground=C_ERR)
        self.log_text.configure(state="disabled")

        # 底部向上依次：日志 → 操作条 → 主体
        log_card.pack(side="bottom", fill="x", padx=16, pady=(12, 0))
        action.pack(side="bottom", fill="x", padx=16, pady=(10, 0))
        main.pack(fill="both", expand=True, padx=16, pady=(12, 0))

        # 任何字段改动都标记为「未保存」
        for v in (self.var_name, self.var_tpl, self.var_outdir, self.var_outname,
                  self.var_date, self.v_task_name, self.v_src_dir, self.v_src_pattern,
                  self.v_src_sheet, self.v_src_header, self.v_src_data,
                  self.v_dst_sheet, self.v_dst_header, self.v_dst_data, self.v_cells):
            v.trace_add("write", self._mark_dirty)

    # ---------------- 规则管理

    def _reload_rules(self, select=None):
        os.makedirs(RULES_DIR, exist_ok=True)
        self.rule_files = sorted(
            [os.path.join(RULES_DIR, f) for f in os.listdir(RULES_DIR) if f.lower().endswith(".json")]
        )
        names = []
        for p in self.rule_files:
            try:
                with open(p, encoding="utf-8") as f:
                    d = json.load(f)
                names.append(d.get("name") or os.path.splitext(os.path.basename(p))[0])
            except Exception:
                names.append(os.path.splitext(os.path.basename(p))[0])
        self.rule_box["values"] = names
        if not self.rule_files:
            self.profile = self._blank_profile()
            self._fill_form()
            self._refresh_tasks()
            return
        idx = 0
        if select:
            for i, p in enumerate(self.rule_files):
                if os.path.normcase(p) == os.path.normcase(select):
                    idx = i
                    break
        self.rule_box.current(idx)
        self._load_profile(self.rule_files[idx])

    def _blank_profile(self):
        return {
            "name": "新规则",
            "template": "",
            "output_dir": "",
            "output_name": "输出-{年}年{月}月{日}.xlsx",
            "tasks": [],
        }

    def _load_profile(self, path):
        try:
            with open(path, encoding="utf-8") as f:
                self.profile = json.load(f)
        except Exception as e:
            messagebox.showerror("读取规则失败", "%s\n%s" % (path, e))
            return
        self.rule_path = path
        self.profile.setdefault("tasks", [])
        self.tasks = self.profile["tasks"]
        self._loading = True
        self._fill_form()
        self._loading = False
        self._refresh_tasks()
        self._clear_dirty()
        self._set_status("已加载：%s" % os.path.basename(path), "info")

    def _fill_form(self):
        prev = self._loading
        self._loading = True
        try:
            self.var_name.set(self.profile.get("name", ""))
            self.var_tpl.set(self.profile.get("template", ""))
            self.var_outdir.set(self.profile.get("output_dir", ""))
            self.var_outname.set(self.profile.get("output_name", ""))
        finally:
            self._loading = prev
        self._refresh_sheet_names()

    def _collect_profile(self):
        self.profile["name"] = self.var_name.get().strip()
        self.profile["template"] = self.var_tpl.get().strip()
        self.profile["output_dir"] = self.var_outdir.get().strip()
        self.profile["output_name"] = self.var_outname.get().strip()
        self.profile["tasks"] = self.tasks
        return self.profile

    def _mark_dirty(self, *_):
        if self._loading:
            return
        self.dirty = True
        if not self.root.title().endswith(" *"):
            self.root.title(APP_TITLE + " *")

    def _clear_dirty(self):
        self.dirty = False
        if self.root.title().endswith(" *"):
            self.root.title(APP_TITLE)

    def _persist(self):
        """把当前界面内容写回规则文件（不重载界面）。"""
        if self.rule_path is None:
            return
        self._collect_task()
        profile = self._collect_profile()
        with open(self.rule_path, "w", encoding="utf-8") as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        self._clear_dirty()

    def _save_rule(self, silent=False):
        self._collect_task()
        profile = self._collect_profile()
        if self.rule_path is None:
            self.rule_path = os.path.join(RULES_DIR, "%s.json" % self._safe_name(profile["name"] or "新规则"))
        try:
            with open(self.rule_path, "w", encoding="utf-8") as f:
                json.dump(profile, f, ensure_ascii=False, indent=2)
            self._clear_dirty()
            cur_idx = self._current_task_index()
            self._reload_rules(select=self.rule_path)
            self._refresh_tasks(keep_index=cur_idx)
            if not silent:
                self._set_status("已保存：%s" % os.path.basename(self.rule_path), "ok")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _new_rule(self):
        name = self._ask_text("新建规则", "规则名称", "新规则")
        if not name:
            return
        path = os.path.join(RULES_DIR, self._safe_name(name) + ".json")
        p = self._blank_profile()
        p["name"] = name
        with open(path, "w", encoding="utf-8") as f:
            json.dump(p, f, ensure_ascii=False, indent=2)
        self.rule_path = path
        self._reload_rules(select=path)
        self._set_status("已新建规则：%s" % name, "ok")

    def _copy_rule(self):
        if not self.rule_path:
            return
        name = self._ask_text("复制规则", "新规则名称", self.var_name.get() + " 副本")
        if not name:
            return
        self._collect_task()
        p = json.loads(json.dumps(self._collect_profile()))
        p["name"] = name
        path = os.path.join(RULES_DIR, self._safe_name(name) + ".json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(p, f, ensure_ascii=False, indent=2)
        self.rule_path = path
        self._reload_rules(select=path)
        self._set_status("已复制为：%s" % name, "ok")

    def _import_rule(self):
        """从别的电脑 / 备份文件导入一套规则。"""
        p = filedialog.askopenfilename(
            title="选择要导入的规则文件",
            filetypes=[("规则文件", "*.json"), ("所有文件", "*.*")],
            initialdir=BASE_DIR,
        )
        if p:
            self._do_import(p)

    def _do_import(self, path, ask=True):
        """把 path 这个规则文件复制进 rules 目录，返回是否成功。"""
        try:
            with open(path, encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception as e:
            if ask:
                messagebox.showerror("导入失败", "这个文件不是有效的规则文件：\n%s" % e)
            return False

        if not isinstance(data, dict) or not isinstance(data.get("tasks"), list):
            if ask:
                messagebox.showerror("导入失败", "这个 JSON 不像规则文件：缺少 tasks 列表。")
            return False

        name = str(data.get("name") or os.path.splitext(os.path.basename(path))[0]).strip() or "导入规则"
        base = self._safe_name(name)
        target = os.path.join(RULES_DIR, base + ".json")

        # 如果选的就是规则目录里自己，直接重载
        try:
            if os.path.exists(target) and os.path.samefile(path, target):
                self._reload_rules(select=target)
                return True
        except Exception:
            pass

        if os.path.exists(target):
            if not ask:
                k = 2
                while os.path.exists(os.path.join(RULES_DIR, "%s-%d.json" % (base, k))):
                    k += 1
                target = os.path.join(RULES_DIR, "%s-%d.json" % (base, k))
                data["name"] = "%s (%d)" % (name, k)
            else:
                ans = messagebox.askyesnocancel(
                    "规则已存在",
                    "已经有一套叫「%s」的规则了。\n\n"
                    "是：覆盖它\n否：另存为副本\n取消：不导入" % name,
                )
                if ans is None:
                    return False
                if not ans:
                    k = 2
                    while os.path.exists(os.path.join(RULES_DIR, "%s-%d.json" % (base, k))):
                        k += 1
                    target = os.path.join(RULES_DIR, "%s-%d.json" % (base, k))
                    data["name"] = "%s (%d)" % (name, k)

        try:
            with open(target, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            if ask:
                messagebox.showerror("导入失败", str(e))
            return False

        self._reload_rules(select=target)
        tpl = fillcore.resolve_path(BASE_DIR, data.get("template", ""))
        if not tpl or not os.path.exists(tpl):
            self._set_status("已导入「%s」，但模板路径不存在，请重新选择" % data["name"], "warn")
        else:
            self._set_status("已导入：%s" % os.path.basename(target), "ok")
        if ask:
            tip = "规则「%s」已导入。" % data["name"]
            if not tpl or not os.path.exists(tpl):
                tip += "\n\n注意：规则里的模板路径在当前电脑上不存在，请在「规则设置」里重新选一下模板和文件夹。"
            messagebox.showinfo("导入成功", tip)
        return True

    def _delete_rule(self):
        if not self.rule_path:
            return
        if not messagebox.askyesno("删除规则", "确定删除「%s」吗？" % self.var_name.get()):
            return
        try:
            os.remove(self.rule_path)
        except Exception as e:
            messagebox.showerror("删除失败", str(e))
            return
        self.rule_path = None
        self._reload_rules()
        self._set_status("已删除规则", "info")

    def _on_rule_change(self, _evt=None):
        i = self.rule_box.current()
        if i < 0 or i >= len(self.rule_files):
            return
        try:
            self._persist()
        except Exception:
            pass
        self._load_profile(self.rule_files[i])

    def _safe_name(self, s):
        for ch in '\\/:*?"<>|':
            s = s.replace(ch, "_")
        return s.strip() or "规则"

    def _ask_text(self, title, label, default=""):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.configure(bg=C_CARD)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        wrap = tk.Frame(win, bg=C_CARD, padx=22, pady=18)
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text=label, bg=C_CARD, fg=C_TEXT2,
                 font=(self.fam, 9)).pack(anchor="w", pady=(0, 8))
        var = tk.StringVar(value=default)
        ent = ttk.Entry(wrap, textvariable=var, width=44)
        ent.pack(fill="x")
        ent.focus_set()
        ent.select_range(0, "end")

        result = {"v": None}

        def ok():
            result["v"] = var.get().strip()
            win.destroy()

        bar = tk.Frame(wrap, bg=C_CARD)
        bar.pack(fill="x", pady=(16, 0))
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="right")
        ttk.Button(bar, text="确定", style="Accent.TButton", command=ok).pack(side="right", padx=(0, 8))
        ent.bind("<Return>", lambda e: ok())
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_window(win, 420, 170)
        self.root.wait_window(win)
        return result["v"]

    def _center_window(self, win, w, h):
        try:
            win.update_idletasks()
            x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
            y = self.root.winfo_rooty() + (self.root.winfo_height() - h) // 2
            win.geometry("%dx%d+%d+%d" % (w, h, max(0, x), max(0, y)))
        except Exception:
            win.geometry("%dx%d" % (w, h))

    # ---------------- 任务

    def _refresh_tasks(self, keep_index=None):
        cur = self._current_task_index() if keep_index is None else keep_index
        self.task_tree.delete(*self.task_tree.get_children())
        for i, t in enumerate(self.tasks):
            src = t.get("source", {})
            self.task_tree.insert("", "end", values=(
                t.get("name", ""),
                src.get("pattern", ""),
                t.get("target", {}).get("sheet", ""),
            ), tags=("odd",) if i % 2 else ())
        self.task_count.configure(text="共 %d 个" % len(self.tasks))
        if self.tasks:
            idx = min(cur or 0, len(self.tasks) - 1)
            self.task_tree.selection_set(self.task_tree.get_children()[idx])
            self._show_task(self.tasks[idx])
        else:
            self._clear_task_form()

    def _current_task_index(self):
        sel = self.task_tree.selection()
        if not sel:
            return None
        return self.task_tree.index(sel[0])

    def _on_task_select(self, _evt=None):
        i = self._current_task_index()
        if i is None or i >= len(self.tasks):
            return
        self._show_task(self.tasks[i])

    def _clear_task_form(self):
        self._loading = True
        for v in (self.v_task_name, self.v_src_dir, self.v_src_pattern,
                  self.v_dst_sheet, self.v_cells):
            v.set("")
        self.v_src_sheet.set("0")
        self.v_src_header.set("1")
        self.v_src_data.set("2")
        self.v_dst_header.set("2")
        self.v_dst_data.set("3")
        self._loading = False
        self._refresh_mapping()

    def _show_task(self, t):
        src = t.get("source", {})
        dst = t.get("target", {})
        self._loading = True
        self.v_task_name.set(t.get("name", ""))
        self.v_src_dir.set(src.get("dir", ""))
        self.v_src_pattern.set(src.get("pattern", ""))
        self.v_src_sheet.set(str(src.get("sheet", 0)))
        self.v_src_header.set(str(src.get("header_row", 1)))
        self.v_src_data.set(str(src.get("data_start", 2)))
        self.v_dst_sheet.set(dst.get("sheet", ""))
        self.v_dst_header.set(str(dst.get("header_row", 2)))
        self.v_dst_data.set(str(dst.get("data_start", 3)))
        cells = t.get("cells", {})
        self.v_cells.set("; ".join("%s=%s" % (k, v) for k, v in cells.items()))
        self._loading = False
        self._refresh_sheet_names()
        self._refresh_mapping()

    def _refresh_sheet_names(self):
        tpl = self.var_tpl.get().strip()
        p = fillcore.resolve_path(BASE_DIR, tpl) if tpl else ""
        names = []
        if p and os.path.exists(p):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(p, read_only=True)
                names = wb.sheetnames
                wb.close()
            except Exception:
                names = []
        self.dst_sheet_box["values"] = names

    def _int_value(self, var, default):
        try:
            s = str(var.get()).strip()
            return int(s) if s else default
        except Exception:
            return default

    def _collect_task(self):
        i = self._current_task_index()
        if i is None or i >= len(self.tasks):
            return
        cells = {}
        for part in self.v_cells.get().split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                if k.strip():
                    cells[k.strip()] = v.strip()
        self.tasks[i] = {
            "name": self.v_task_name.get().strip(),
            "source": {
                "dir": self.v_src_dir.get().strip(),
                "pattern": self.v_src_pattern.get().strip(),
                "sheet": as_sheet(self.v_src_sheet.get()),
                "header_row": self._int_value(self.v_src_header, 1),
                "data_start": self._int_value(self.v_src_data, 2),
            },
            "target": {
                "sheet": self.v_dst_sheet.get().strip(),
                "header_row": self._int_value(self.v_dst_header, 2),
                "data_start": self._int_value(self.v_dst_data, 3),
            },
            "cells": cells,
            "columns": self.tasks[i].get("columns", []),
        }
        return self.tasks[i]

    def _add_task(self):
        self._collect_task()
        t = {
            "name": "新任务",
            "source": {"dir": "", "pattern": "*.xlsx", "sheet": 0, "header_row": 1, "data_start": 2},
            "target": {"sheet": "", "header_row": 2, "data_start": 3},
            "cells": {},
            "columns": [],
        }
        self.tasks.append(t)
        self._refresh_tasks(keep_index=len(self.tasks) - 1)
        self._show_task(t)
        self._mark_dirty()
        self._set_status("已添加任务，记得点「保存规则」", "info")

    def _del_task(self):
        i = self._current_task_index()
        if i is None:
            return
        if not messagebox.askyesno("删除任务", "确定删除任务「%s」吗？" % self.tasks[i].get("name", "")):
            return
        self.tasks.pop(i)
        self._mark_dirty()
        self._refresh_tasks(keep_index=max(0, i - 1))
        if self.tasks:
            self._on_task_select()
        else:
            self._clear_task_form()

    def _move_task(self, delta):
        i = self._current_task_index()
        if i is None:
            return
        j = i + delta
        if j < 0 or j >= len(self.tasks):
            return
        self._collect_task()
        self.tasks[i], self.tasks[j] = self.tasks[j], self.tasks[i]
        self._mark_dirty()
        self._refresh_tasks(keep_index=j)
        self._show_task(self.tasks[j])

    # ---------------- 列映射

    def _refresh_mapping(self):
        self.map_tree.delete(*self.map_tree.get_children())
        i = self._current_task_index()
        if i is None:
            return
        for k, col in enumerate(self.tasks[i].get("columns", [])):
            self.map_tree.insert("", "end", values=(
                col.get("target", ""),
                col.get("source", ""),
                TRANSFORM_VALUE.get(col.get("transform", "direct"), col.get("transform", "")),
            ), tags=("odd",) if k % 2 else ())

    def _task_source_path(self):
        i = self._current_task_index()
        if i is None:
            return ""
        import glob
        t = self.tasks[i]
        src = t.get("source", {})
        d = fillcore.resolve_path(BASE_DIR, src.get("dir", ""))
        pat = src.get("pattern", "")
        if not d or not pat:
            return ""
        if any(ch in pat for ch in "*?["):
            m = sorted(glob.glob(os.path.join(d, pat)))
            return m[-1] if m else ""
        return os.path.join(d, pat)

    def _auto_match(self):
        i = self._current_task_index()
        if i is None:
            return
        tpl = fillcore.resolve_path(BASE_DIR, self.var_tpl.get().strip())
        if not tpl or not os.path.exists(tpl):
            messagebox.showwarning("提示", "请先选好模板文件")
            return
        src_path = self._task_source_path()
        if not src_path or not os.path.exists(src_path):
            messagebox.showwarning("提示", "找不到源文件，请先设置源文件夹和文件名匹配")
            return
        try:
            src_headers = read_headers(src_path, as_sheet(self.v_src_sheet.get()), self._int_value(self.v_src_header, 1))
            dst_headers = read_headers(tpl, self.v_dst_sheet.get().strip(), self._int_value(self.v_dst_header, 2))
        except Exception as e:
            messagebox.showerror("读取表头失败", str(e))
            return
        cols = []
        for h in dst_headers:
            src_h = None
            if h in src_headers:
                src_h = h
            else:
                for cand in ALIASES.get(h, []):
                    if cand in src_headers:
                        src_h = cand
                        break
            if not src_h:
                continue
            tf = "direct"
            if h == "单位" and src_h == "包装信息":
                tf = "after_slash"
            cols.append({"target": h, "source": src_h, "transform": tf})
        self.tasks[i]["columns"] = cols
        self._mark_dirty()
        self._refresh_mapping()
        self._set_status("自动匹配了 %d 列，可手工补充" % len(cols), "ok")

    def _add_mapping(self):
        self._edit_mapping(add=True)

    def _edit_mapping(self, add=False):
        i = self._current_task_index()
        if i is None:
            return
        cols = self.tasks[i].setdefault("columns", [])
        sel = self.map_tree.selection()
        idx = self.map_tree.index(sel[0]) if sel else None
        if add or idx is None:
            cur = {"target": "", "source": "", "transform": "direct", "arg": ""}
            idx = None
        else:
            cur = dict(cols[idx])

        tpl = fillcore.resolve_path(BASE_DIR, self.var_tpl.get().strip())
        src_path = self._task_source_path()
        try:
            dst_headers = read_headers(tpl, self.v_dst_sheet.get().strip(), self._int_value(self.v_dst_header, 2)) if tpl and os.path.exists(tpl) else []
            src_headers = read_headers(src_path, as_sheet(self.v_src_sheet.get()), self._int_value(self.v_src_header, 1)) if src_path and os.path.exists(src_path) else []
        except Exception:
            dst_headers, src_headers = [], []

        win = tk.Toplevel(self.root)
        win.title("列映射")
        win.configure(bg=C_CARD)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        wrap = tk.Frame(win, bg=C_CARD, padx=22, pady=18)
        wrap.pack(fill="both", expand=True)
        wrap.columnconfigure(1, weight=1)

        v_dst = tk.StringVar(value=cur.get("target", ""))
        v_src = tk.StringVar(value=cur.get("source", ""))
        v_tf = tk.StringVar(value=TRANSFORM_VALUE.get(cur.get("transform", "direct"), "直接搬"))
        v_arg = tk.StringVar(value=cur.get("arg", "") or "")

        def lab(row, text):
            tk.Label(wrap, text=text, bg=C_CARD, fg=C_TEXT2, font=(self.fam, 9)).grid(
                row=row, column=0, sticky="w", pady=5, padx=(0, 14))

        lab(0, "目标列")
        ttk.Combobox(wrap, textvariable=v_dst, values=dst_headers, width=36,
                     font=(self.fam, 9)).grid(row=0, column=1, sticky="we", pady=5)

        lab(1, "源列")
        ttk.Combobox(wrap, textvariable=v_src, values=src_headers, width=36,
                     font=(self.fam, 9)).grid(row=1, column=1, sticky="we", pady=5)

        lab(2, "转换方式")
        ttk.Combobox(wrap, textvariable=v_tf, state="readonly",
                     values=[x[0] for x in TRANSFORMS], width=36,
                     font=(self.fam, 9)).grid(row=2, column=1, sticky="we", pady=5)

        lab(3, "固定值内容")
        ttk.Entry(wrap, textvariable=v_arg).grid(row=3, column=1, sticky="we", pady=5)
        tk.Label(wrap, text="只有转换方式选「固定值」时才用这一格",
                 bg=C_CARD, fg=C_MUTED, font=(self.fam, 9)).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 2))

        def ok():
            item = {
                "target": v_dst.get().strip(),
                "source": v_src.get().strip(),
                "transform": TRANSFORM_VALUE.get(v_tf.get(), "direct"),
                "arg": v_arg.get().strip(),
            }
            if not item["target"]:
                messagebox.showwarning("提示", "目标列不能为空", parent=win)
                return
            if idx is None:
                cols.append(item)
            else:
                cols[idx] = item
            win.destroy()
            self._mark_dirty()
            self._refresh_mapping()
            self._set_status("列映射已更新，记得点「保存规则」", "info")

        bar = tk.Frame(wrap, bg=C_CARD)
        bar.grid(row=5, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(bar, text="取消", command=win.destroy).pack(side="right")
        ttk.Button(bar, text="确定", style="Accent.TButton", command=ok).pack(side="right", padx=(0, 8))
        win.bind("<Escape>", lambda e: win.destroy())
        self._center_window(win, 520, 300)
        self.root.wait_window(win)

    def _del_mapping(self):
        i = self._current_task_index()
        if i is None:
            return
        sel = self.map_tree.selection()
        if not sel:
            return
        idx = self.map_tree.index(sel[0])
        cols = self.tasks[i].get("columns", [])
        if idx < len(cols):
            cols.pop(idx)
        self._mark_dirty()
        self._refresh_mapping()

    # ---------------- 浏览

    def _pick_template(self):
        p = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
            initialdir=BASE_DIR,
        )
        if p:
            self.var_tpl.set(to_rel(p))
            self._refresh_sheet_names()
            self._set_status("模板：%s" % os.path.basename(p), "info")

    def _pick_outdir(self):
        p = filedialog.askdirectory(title="选择输出文件夹", initialdir=BASE_DIR)
        if p:
            self.var_outdir.set(to_rel(p))

    def _pick_src_dir(self):
        p = filedialog.askdirectory(title="选择源文件所在文件夹", initialdir=BASE_DIR)
        if p:
            self.v_src_dir.set(to_rel(p))

    def _open_out(self):
        d = fillcore.resolve_path(BASE_DIR, self.var_outdir.get().strip())
        if d and os.path.isdir(d):
            open_folder(d)
        else:
            messagebox.showinfo("提示", "输出文件夹还不存在，生成一次就会自动建好")

    # ---------------- 运行

    def _parse_date(self):
        s = self.var_date.get().strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
        raise ValueError("日期格式不对，应写成 2026-11-01")

    def _run(self):
        if self.running:
            return
        try:
            d = self._parse_date()
        except ValueError as e:
            messagebox.showwarning("日期有误", str(e))
            return
        self._collect_task()
        profile = deepcopy(self._collect_profile())
        if not profile.get("template"):
            messagebox.showwarning("提示", "请先选择模板文件")
            return
        if not profile.get("tasks"):
            messagebox.showwarning("提示", "还没有任务，请先添加")
            return
        self.running = True
        self.run_btn.configure(state="disabled", text="生成中…")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._set_status("正在生成…", "busy")
        self._clear_log()
        self._log("开始生成，日期 %s" % d.strftime("%Y-%m-%d"))
        threading.Thread(target=self._run_worker, args=(profile, fillcore.date_context(d)), daemon=True).start()

    def _run_worker(self, profile, ctx):
        try:
            out = fillcore.run_profile(profile, BASE_DIR, ctx, log=self._log)
            self._log("完成：%s" % out)
            self.log_queue.put(("done", "生成完成", "ok"))
            try:
                os.startfile(os.path.dirname(out))
            except Exception:
                pass
        except Exception as e:
            self._log("出错了：%s" % e)
            traceback.print_exc()
            self.log_queue.put(("done", "生成失败，请看日志", "err"))

    def _run_done(self):
        self.running = False
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.run_btn.configure(state="normal", text="▶   开始生成")

    # ---------------- 日志

    def _log(self, msg):
        self.log_queue.put(str(msg))

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _log_tag(self, msg):
        if any(k in msg for k in ("出错", "失败", "错误", "找不到", "不存在", "没有")):
            return "error"
        if "警告" in msg:
            return "warn"
        if msg.startswith("    "):
            return "detail"
        if any(k in msg for k in ("完成", "已保存")):
            return "ok"
        return "info"

    def _poll_log(self):
        try:
            pending = []
            while True:
                pending.append(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        if pending:
            self.log_text.configure(state="normal")
            for msg in pending:
                if isinstance(msg, tuple) and msg[0] == "done":
                    self._set_status(msg[1], msg[2])
                    self._run_done()
                    continue
                stamp = datetime.datetime.now().strftime("%H:%M:%S")
                self.log_text.insert("end", stamp + "  ", "time")
                self.log_text.insert("end", msg + "\n", self._log_tag(msg))
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(200, self._poll_log)

    # ---------------- 状态

    def _set_status(self, s, level="info"):
        s = str(s)
        self.status_var.set(s)
        colors = {
            "info": (C_DARK2, "#8FD3FF"),
            "busy": ("#173A6B", "#BFDBFE"),
            "ok": ("#12351F", "#86EFAC"),
            "warn": ("#3A2A10", "#FCD34D"),
            "err": ("#3B1A1A", "#FCA5A5"),
        }
        bg, fg = colors.get(level, colors["info"])
        short = s if len(s) <= 26 else s[:25] + "…"
        self.pill.configure(text="●  " + short, bg=bg, fg=fg)


def main():
    enable_dpi_awareness()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
