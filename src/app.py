# -*- coding: utf-8 -*-
"""建筑设计周期定额计算软件 - GUI 主程序（customtkinter）。

风格参考 Marvis 主界面：黑白灰主色、圆角、高饱和度点缀、简约高级。
以「设计阶段 + 项目 + 多个单体子项」为主线组织输入，支持单体增删/挂起，
调整因素勾选与参数输入，输出 Word 报告（自定义保存位置）与可选 Excel。
"""
from __future__ import annotations

import os
import sys
import json
import traceback
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from version import APP_VERSION, APP_NAME, APP_FILENAME_PREFIX, REPORT_TITLE
from quota_data import (
    QUOTA_CATEGORIES, RENOVATION_LEVELS, TWO_STAGE_TEXT,
    DAYS_PER_MONTH, DAYS_PER_WEEK, COMPRESS_RISK_TEXT, COMPRESS_WARN_TEXT,
)
from quota_engine import (
    ProjectSpec, UnitSpec, AdjustmentSpec, NewTechItem, ReworkItem,
    calculate, months_to_weeks, months_to_days, months_to_days_int,
)
from report_docx import generate_report
from export_excel import export_excel

# ---------- 配置持久化（v1.2 第 1 项：参数记忆） ----------
def _config_path():
    """exe 打包后 APPDATA 存在，否则退到用户目录，避免写入程序目录被拒。"""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "design_cycle_quota_config.json")

CONFIG_PATH = _config_path()
REPORT_DEFAULT_NAME = "设计周期定额计算报告"

# ---------- 主题 ----------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# v1.1：莫兰迪色系（低饱和度灰调：米灰 / 灰蓝 / 灰粉 / 沙橙），文字统一为黑
COLOR_BG = "#EDEAE4"          # 米灰背景
COLOR_PANEL = "#E2DED6"       # 浅米灰面板
COLOR_CARD = "#F6F3ED"        # 米白卡片
COLOR_DARK = "#1A1A1A"        # 黑（文字主色）
COLOR_NAVY = "#33415C"        # 深蓝（顶栏标题背景 / 下拉箭头）
COLOR_NAVY_HOVER = "#41506F"  # 深蓝 hover
COLOR_NAVY_DARK = "#2A364C"   # 深蓝变体（徽标背景）
COLOR_ACCENT = "#A9716B"      # 莫兰迪陶土红（点缀）
COLOR_ACCENT_HOVER = "#8F5F58"
COLOR_ACCENT2 = "#C9A37E"     # 莫兰迪沙橙
COLOR_ACCENT2_HOVER = "#B28962"
COLOR_GRAY = "#1A1A1A"        # 最浅灰文字统一改为黑色（确保清晰可读）
COLOR_GREEN = "#7D9A7A"       # 莫兰迪绿
COLOR_BORDER = "#CFC9BE"      # 灰米边框

RADIUS = 10

CODE_ITEMS = []
for cat in QUOTA_CATEGORIES:
    for it in cat["items"]:
        label = f'{it["code"]}  {it["name"]}（{cat["cat"]}'
        if it.get("ref"):
            ref_label = it["ref"]
            for c in QUOTA_CATEGORIES:
                for x in c["items"]:
                    if x["code"] == it["ref"]:
                        ref_label = f'{it["ref"]} {x["name"]}'
                        break
            label += "，参考 " + ref_label
        CODE_ITEMS.append(label + "）")


def code_of(label: str) -> str:
    return label.split(" ")[0] if label else "01-1"


def name_of(code: str) -> str:
    for cat in QUOTA_CATEGORIES:
        for it in cat["items"]:
            if it["code"] == code:
                return it["name"]
    return code


# ---------- 卡片组件 ----------
class Card(ctk.CTkFrame):
    """圆角白色卡片容器。"""

    def __init__(self, master, title=None, **kwargs):
        kwargs.setdefault("fg_color", COLOR_CARD)
        kwargs.setdefault("corner_radius", RADIUS)
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", COLOR_BORDER)
        super().__init__(master, **kwargs)
        self.pack(fill="x", padx=12, pady=6)
        if title:
            lbl = ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=15, weight="bold"),
                               text_color=COLOR_DARK, anchor="w")
            lbl.pack(fill="x", padx=14, pady=(10, 4))


def field_row(master, label, widget):
    row = ctk.CTkFrame(master, fg_color="transparent")
    row.pack(fill="x", padx=14, pady=3)
    lbl = ctk.CTkLabel(row, text=label, width=120, anchor="w", font=ctk.CTkFont(size=12),
                       text_color=COLOR_GRAY)
    lbl.pack(side="left")
    widget.pack(side="left", fill="x", expand=True, padx=(6, 0))
    return row


class Entry(ctk.CTkEntry):
    def __init__(self, master, placeholder="", width=200, numeric=False, **kw):
        kw.setdefault("height", 30)
        kw.setdefault("corner_radius", 6)
        kw.setdefault("fg_color", "#FBF9F5")
        kw.setdefault("border_color", COLOR_BORDER)
        kw.setdefault("text_color", COLOR_DARK)
        kw.setdefault("placeholder_text_color", COLOR_DARK)
        super().__init__(master, placeholder_text=placeholder, width=width, **kw)
        if numeric:
            vcmd = (self.register(self._check_numeric), "%P")
            self.configure(validate="key", validatecommand=vcmd)

    @staticmethod
    def _check_numeric(text):
        """仅允许数字/小数点/负号输入，拒绝字母等异常字符（修复输入框出现 a'a）。"""
        if text in ("", "-", ".", "-."):
            return True
        try:
            float(text)
            return True
        except ValueError:
            return False


# ---------- 单体子项行 ----------
class UnitRow(ctk.CTkFrame):
    def __init__(self, master, on_delete, on_copy=None, index_label="单体1"):
        super().__init__(master, fg_color="#FBF9F5", corner_radius=8,
                         border_width=1, border_color=COLOR_BORDER)
        self.pack(fill="x", padx=10, pady=4)
        self.var_name = tk.StringVar(value=index_label)
        self.var_code = tk.StringVar(value=CODE_ITEMS[0])
        self.var_scale = tk.StringVar(value="10000")
        self.var_custom = tk.StringVar(value="")
        self.var_suspend = tk.BooleanVar(value=False)
        self.on_copy = on_copy

        self.lbl_name = ctk.CTkLabel(self, text="名称", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY)
        self.lbl_name.grid(row=0, column=0, padx=(10, 2), pady=6)
        self.ent_name = Entry(self, width=110)
        self.ent_name.grid(row=0, column=1, padx=2)
        self.ent_name.configure(textvariable=self.var_name)

        self.lbl_code = ctk.CTkLabel(self, text="定额子项", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY)
        self.lbl_code.grid(row=0, column=2, padx=(8, 2), pady=6)
        self.opt_code = ctk.CTkOptionMenu(self, values=CODE_ITEMS, variable=self.var_code,
                                          width=360, height=30, corner_radius=6,
                                          fg_color="#FBF9F5", button_color=COLOR_NAVY,
                                          button_hover_color=COLOR_NAVY_HOVER,
                                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11))
        self.opt_code.grid(row=0, column=3, padx=2)

        self.lbl_scale = ctk.CTkLabel(self, text="规模㎡", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY)
        self.lbl_scale.grid(row=0, column=4, padx=(8, 2), pady=6)
        self.ent_scale = Entry(self, width=80, numeric=True)
        self.ent_scale.grid(row=0, column=5, padx=2)
        self.ent_scale.configure(textvariable=self.var_scale)

        self.chk_under = ctk.CTkCheckBox(self, text="200床以下(疗养/养老)", variable=tk.BooleanVar(value=False),
                                          width=170, font=ctk.CTkFont(size=10), text_color=COLOR_GRAY,
                                          fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT2)
        self.chk_under.grid(row=0, column=6, padx=(8, 2))
        self.chk_under_var = self.chk_under._check_state if hasattr(self.chk_under, "_check_state") else None

        self.chk_suspend = ctk.CTkCheckBox(self, text="挂起", variable=self.var_suspend,
                                           width=60, font=ctk.CTkFont(size=10), text_color=COLOR_GRAY,
                                           fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT2,
                                           command=self._on_suspend_toggle)
        self.chk_suspend.grid(row=0, column=7, padx=(8, 2))

        self.btn_copy = ctk.CTkButton(self, text="复制上一行", width=84, height=28, corner_radius=6,
                                      fg_color="#FBF9F5", hover_color="#E8E4DA",
                                      text_color=COLOR_DARK, border_width=1, border_color=COLOR_BORDER,
                                      command=self._copy_prev, font=ctk.CTkFont(size=10))
        self.btn_copy.grid(row=0, column=8, padx=(8, 2), pady=6)

        self.btn_del = ctk.CTkButton(self, text="删除", width=54, height=28, corner_radius=6,
                                     fg_color="#FBF9F5", hover_color="#F3E8E5",
                                     text_color=COLOR_ACCENT, border_width=1, border_color=COLOR_ACCENT,
                                     command=on_delete, font=ctk.CTkFont(size=11))
        self.btn_del.grid(row=0, column=9, padx=(8, 10), pady=6)

    def _copy_prev(self):
        if self.on_copy:
            self.on_copy(self)

    def _on_suspend_toggle(self):
        if self.var_suspend.get():
            if not messagebox.askyesno("确认挂起", f"确认挂起单体「{self.var_name.get()}」？\n挂起后该单体保留数据但不参与计算。"):
                self.var_suspend.set(False)
                return
        self._toggle_suspend()

    def _toggle_suspend(self):
        state = "disabled" if self.var_suspend.get() else "normal"
        for w in (self.ent_name, self.opt_code, self.ent_scale, self.chk_under):
            w.configure(state=state)

    def get_unit(self):
        try:
            scale = float(self.var_scale.get() or 0)
        except ValueError:
            scale = 0.0
        custom = None
        if self.var_custom.get().strip():
            try:
                parts = [float(x) for x in self.var_custom.get().replace("，", ",").split(",")]
                if len(parts) >= 3:
                    custom = {"scheme": parts[0], "preliminary": parts[1], "construction": parts[2]}
            except ValueError:
                custom = None
        code = code_of(self.var_code.get())
        under = False
        try:
            under = self.chk_under._check_state.get() if hasattr(self.chk_under, "_check_state") else False
        except Exception:
            under = False
        return UnitSpec(name=self.var_name.get().strip() or "单体",
                        code=code,
                        scale=scale,
                        suspended=self.var_suspend.get(),
                        under_200_beds=under and code == "03-4",
                        custom_periods=custom)


# ---------- 新技术行 ----------
class NewTechRow(ctk.CTkFrame):
    def __init__(self, master, on_delete, idx=1):
        super().__init__(master, fg_color="#FBF9F5", corner_radius=8,
                         border_width=1, border_color=COLOR_BORDER)
        self.pack(fill="x", padx=10, pady=3)
        self.var_name = tk.StringVar(value=f"新技术{idx}")
        self.var_desc = tk.StringVar(value="")
        self.var_months = tk.StringVar(value="1.0")
        self.var_pool = tk.BooleanVar(value=False)
        self.var_group = tk.StringVar(value="1")

        ctk.CTkLabel(self, text="名称", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=0, padx=(10, 2), pady=5)
        e1 = Entry(self, width=110)
        e1.grid(row=0, column=1, padx=2)
        e1.configure(textvariable=self.var_name)
        ctk.CTkLabel(self, text="概述", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=2, padx=(8, 2))
        e2 = Entry(self, width=180)
        e2.grid(row=0, column=3, padx=2)
        e2.configure(textvariable=self.var_desc)
        ctk.CTkLabel(self, text="增加(月)", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=4, padx=(8, 2))
        e3 = Entry(self, width=60, numeric=True)
        e3.grid(row=0, column=5, padx=2)
        e3.configure(textvariable=self.var_months)
        self.chk = ctk.CTkCheckBox(self, text="可统筹", variable=self.var_pool, width=70,
                                   font=ctk.CTkFont(size=10), text_color=COLOR_GRAY,
                                   fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT2)
        self.chk.grid(row=0, column=6, padx=(8, 2))
        ctk.CTkLabel(self, text="组号", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=7, padx=(8, 2))
        e4 = Entry(self, width=44, numeric=True)
        e4.grid(row=0, column=8, padx=2)
        e4.configure(textvariable=self.var_group)
        ctk.CTkButton(self, text="删除", width=50, height=26, corner_radius=6,
                      fg_color="#FBF9F5", hover_color="#F3E8E5", text_color=COLOR_ACCENT,
                      border_width=1, border_color=COLOR_ACCENT, command=on_delete,
                      font=ctk.CTkFont(size=10)).grid(row=0, column=9, padx=(8, 10), pady=5)

    def get_item(self):
        try:
            months = float(self.var_months.get() or 0)
        except ValueError:
            months = 0.0
        try:
            group = int(self.var_group.get() or 0)
        except ValueError:
            group = 0
        return NewTechItem(name=self.var_name.get().strip() or "新技术",
                           desc=self.var_desc.get().strip(),
                           months=months,
                           suspended=False,
                           poolable=self.var_pool.get(),
                           pool_group=group)


# ---------- 返工行 ----------
class ReworkRow(ctk.CTkFrame):
    def __init__(self, master, on_delete):
        super().__init__(master, fg_color="#FBF9F5", corner_radius=8,
                         border_width=1, border_color=COLOR_BORDER)
        self.pack(fill="x", padx=10, pady=3)
        self.var_stage = tk.StringVar(value="construction")
        self.var_months = tk.StringVar(value="1.0")
        self.var_note = tk.StringVar(value="")

        ctk.CTkLabel(self, text="阶段", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=0, padx=(10, 2), pady=5)
        self.opt = ctk.CTkOptionMenu(self, values=["construction 施工图", "preliminary 初步", "scheme 方案"],
                                     variable=self.var_stage, width=140, height=28, corner_radius=6,
                                     fg_color="#FBF9F5", button_color=COLOR_NAVY,
                                     button_hover_color=COLOR_NAVY_HOVER,
                                     text_color=COLOR_DARK, font=ctk.CTkFont(size=10))
        self.opt.grid(row=0, column=1, padx=2)
        ctk.CTkLabel(self, text="增加(月)", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=2, padx=(8, 2))
        e1 = Entry(self, width=60, numeric=True)
        e1.grid(row=0, column=3, padx=2)
        e1.configure(textvariable=self.var_months)
        ctk.CTkLabel(self, text="注释", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).grid(row=0, column=4, padx=(8, 2))
        e2 = Entry(self, width=220)
        e2.grid(row=0, column=5, padx=2)
        e2.configure(textvariable=self.var_note)
        ctk.CTkButton(self, text="删除", width=50, height=26, corner_radius=6,
                      fg_color="#FBF9F5", hover_color="#F3E8E5", text_color=COLOR_ACCENT,
                      border_width=1, border_color=COLOR_ACCENT, command=on_delete,
                      font=ctk.CTkFont(size=10)).grid(row=0, column=6, padx=(8, 10), pady=5)

    def get_item(self):
        try:
            months = float(self.var_months.get() or 0)
        except ValueError:
            months = 0.0
        stage = self.var_stage.get().split(" ")[0]
        return ReworkItem(stage=stage, months=months, note=self.var_note.get().strip())


# ---------- 主窗口 ----------
class QuotaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("建筑设计周期定额计算软件（2016版）")
        self.geometry("1360x840")
        self.minsize(1100, 720)
        self.configure(fg_color=COLOR_BG)

        self.unit_rows = []
        self.tech_rows = []
        self.rework_rows = []
        self._busy = False                 # v1.2 第4项：生成过程防误操作
        self._output_name_auto = True      # v1.2 第3项：文件名未被手动修改
        self._last_docx_path = ""
        self._last_xlsx_path = ""

        self._build_topbar()
        self._build_body()

        # 默认值
        self.var_project.set("")
        self.var_output_dir.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        self._add_unit()
        # v1.2 第1项：启动恢复上次配置；第3项：项目名联动文件名；第8项：快捷键
        self._load_config()
        self.var_project.trace_add("write", self._on_project_changed)
        self.var_output_name.trace_add("write", self._on_output_name_changed)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-Return>", lambda e: self._on_calculate())
        self.bind("<Control-s>", lambda e: self._on_calculate())
        self.bind("<Control-S>", lambda e: self._on_calculate())

    # ---------- 顶栏 ----------
    def _build_topbar(self):
        bar = ctk.CTkFrame(self, fg_color=COLOR_NAVY, corner_radius=0, height=64)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        title = ctk.CTkLabel(bar, text=f"{APP_NAME} v{APP_VERSION}", font=ctk.CTkFont(size=20, weight="bold"),
                             text_color="#FFFFFF")
        title.pack(side="left", padx=20, pady=16)
        badge = ctk.CTkLabel(bar, text="2016版定额 · Marvis Style", font=ctk.CTkFont(size=11),
                             text_color=COLOR_ACCENT2, fg_color=COLOR_NAVY_DARK, corner_radius=6)
        badge.pack(side="left", padx=(4, 0), pady=16)
        self.btn_calc = ctk.CTkButton(bar, text="开始计算并生成报告", width=220, height=36, corner_radius=8,
                                      fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF",
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._on_calculate)
        self.btn_calc.pack(side="right", padx=20)

    # ---------- 主体 ----------
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color=COLOR_BG)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, weight=1, uniform="col")
        body.grid_columnconfigure(1, weight=0, uniform="col")
        body.grid_rowconfigure(0, weight=1)

        # 左侧输入（可滚动）
        left = ctk.CTkFrame(body, fg_color=COLOR_PANEL, corner_radius=0)
        left.grid(row=0, column=0, sticky="nsew")
        self.canvas = ctk.CTkScrollableFrame(left, fg_color=COLOR_PANEL, corner_radius=0)
        self.canvas.pack(fill="both", expand=True)

        self._build_project_card()
        self._build_units_card()
        self._build_adjust_card()

        # 右侧输出设置 + 结果
        right = ctk.CTkFrame(body, fg_color=COLOR_BG, width=330, corner_radius=0)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_propagate(False)

        out_card = Card(right, title="输出设置")
        self.var_output_dir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.var_output_name = tk.StringVar(value="设计周期定额计算报告")
        self.var_excel = tk.BooleanVar(value=False)
        row_dir = ctk.CTkFrame(out_card, fg_color="transparent")
        row_dir.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row_dir, text="Word 保存位置", width=120, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=COLOR_GRAY).pack(side="left")
        ent_dir = Entry(row_dir, width=150)
        ent_dir.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ent_dir.configure(textvariable=self.var_output_dir)
        ctk.CTkButton(out_card, text="选择文件夹…", width=120, height=28, corner_radius=6,
                      fg_color="#FBF9F5", hover_color="#DCD8CF", text_color=COLOR_DARK,
                      border_width=1, border_color=COLOR_BORDER,
                      command=self._choose_dir, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14, pady=(0, 4))
        row_name = ctk.CTkFrame(out_card, fg_color="transparent")
        row_name.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row_name, text="文件名", width=120, anchor="w",
                     font=ctk.CTkFont(size=12), text_color=COLOR_GRAY).pack(side="left")
        ent_name = Entry(row_name, width=150)
        ent_name.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ent_name.configure(textvariable=self.var_output_name)
        chk = ctk.CTkCheckBox(out_card, text="同时导出 Excel（供客户调整）", variable=self.var_excel,
                              font=ctk.CTkFont(size=12), text_color=COLOR_DARK,
                              fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT2)
        chk.pack(anchor="w", padx=14, pady=6)

        res_card = Card(right, title="计算结果")
        self.res_text = ctk.CTkTextbox(res_card, height=300, corner_radius=8,
                                       fg_color="#FDFBF7", border_color=COLOR_BORDER, border_width=1,
                                       font=ctk.CTkFont(size=12), text_color=COLOR_DARK)
        self.res_text.pack(fill="x", padx=12, pady=6)
        self.res_text.insert("1.0", "填写左侧信息后点击「开始计算并生成报告」。\n\n"
                                    "结果将显示各阶段周期（月/周/天）、分设计阶段单体统计、调整明细与风险警告。")
        self.res_text.configure(state="disabled")
        btn_row = ctk.CTkFrame(res_card, fg_color="transparent")
        btn_row.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkButton(btn_row, text="复制结果", width=110, height=28, corner_radius=6,
                      fg_color="#FBF9F5", hover_color="#E8E4DA", text_color=COLOR_DARK,
                      border_width=1, border_color=COLOR_BORDER,
                      command=self._copy_result, font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkLabel(btn_row, text="快捷键：Ctrl+Enter 计算 / Ctrl+S 生成",
                     font=ctk.CTkFont(size=10), text_color=COLOR_GRAY).pack(side="left", padx=10)

        note = ctk.CTkLabel(right, text="时间单位：1 月 = 4 周 = 28 天（天保留整数）",
                            font=ctk.CTkFont(size=10), text_color=COLOR_GRAY)
        note.pack(pady=(0, 12))

    # ---------- 项目卡片 ----------
    def _build_project_card(self):
        card = Card(self.canvas, title="项目信息")
        self.var_project = tk.StringVar()
        self.var_location = tk.StringVar()
        self.var_designer = tk.StringVar()
        self.var_start = tk.StringVar(value=datetime.date.today().isoformat())
        self.var_stage_mode = tk.StringVar(value="three  三阶段（方案→初步→施工图）")
        self.var_round_mode = tk.StringVar(value="round  四舍五入")
        self.var_digits = tk.StringVar(value="2")
        self.var_multi_mode = tk.StringVar(value="none  不折减（默认）")
        self.var_multi_value = tk.StringVar(value="0")

        def make_row(label, var, placeholder="", width=260):
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=3)
            ctk.CTkLabel(row, text=label, width=120, anchor="w", font=ctk.CTkFont(size=12),
                         text_color=COLOR_GRAY).pack(side="left")
            e = Entry(row, width=width, placeholder=placeholder)
            e.pack(side="left", fill="x", expand=True, padx=(6, 0))
            e.configure(textvariable=var)
            return e

        make_row("项目名称", self.var_project)
        make_row("项目地点", self.var_location)
        make_row("设计单位", self.var_designer)
        make_row("开始时间", self.var_start, placeholder="YYYY-MM-DD")
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="阶段模式", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(row, values=["three  三阶段（方案→初步→施工图）",
                                       "two  两阶段（初步并入施工图）"],
                          variable=self.var_stage_mode, width=320, height=30,
                          corner_radius=6, fg_color="#FBF9F5",
                          button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="精度规则", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(row, values=["round  四舍五入", "ceil  向上进位", "floor  向下取整"],
                          variable=self.var_round_mode, width=200, height=30,
                          corner_radius=6, fg_color="#FBF9F5",
                          button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(row, text="小数位数", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(16, 0))
        e_digits = Entry(row, width=56, numeric=True)
        e_digits.pack(side="left", padx=4)
        e_digits.configure(textvariable=self.var_digits)
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text="多单体折减", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(row, values=["none  不折减（默认）", "days  减少天数", "ratio  按比例折减"],
                          variable=self.var_multi_mode, width=200, height=30,
                          corner_radius=6, fg_color="#FBF9F5",
                          button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(row, text="折减参数", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(16, 0))
        e_multi = Entry(row, width=72, placeholder="天数或比例0~1", numeric=True)
        e_multi.pack(side="left", padx=4)
        e_multi.configure(textvariable=self.var_multi_value)

    # ---------- 单体卡片 ----------
    def _build_units_card(self):
        card = Card(self.canvas, title="单体子项（多单体周期相加，支持增删/挂起）")
        self.units_container = ctk.CTkFrame(card, fg_color="transparent")
        self.units_container.pack(fill="x", padx=8, pady=4)
        bar = ctk.CTkFrame(card, fg_color="transparent")
        bar.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkButton(bar, text="+ 添加单体", width=110, height=30, corner_radius=6,
                      fg_color=COLOR_ACCENT2, hover_color=COLOR_ACCENT2_HOVER, text_color="#FFFFFF",
                      font=ctk.CTkFont(size=12), command=self._add_unit).pack(side="left")
        ctk.CTkLabel(bar, text="挂起：保留数据但不参与计算（Q6 决策）", font=ctk.CTkFont(size=10),
                     text_color=COLOR_GRAY).pack(side="left", padx=12)

    def _add_unit(self):
        idx = len(self.unit_rows) + 1
        row = UnitRow(self.units_container, on_delete=lambda: self._del_unit(row),
                      on_copy=self._copy_unit,
                      index_label=f"单体{idx}")
        self.unit_rows.append(row)

    def _del_unit(self, row):
        if len(self.unit_rows) <= 1:
            messagebox.showinfo("提示", "至少保留一个单体子项。")
            return
        if not messagebox.askyesno("确认删除", f"确认删除单体「{row.var_name.get()}」？该操作不可恢复。"):
            return
        row.destroy()
        self.unit_rows.remove(row)

    def _copy_unit(self, row):
        """复制上一行（名称外的参数）到当前行（v1.2 第7项）。"""
        idx = self.unit_rows.index(row)
        if idx <= 0:
            messagebox.showinfo("提示", "已是第一行，没有可复制的上一行。")
            return
        prev = self.unit_rows[idx - 1]
        row.var_code.set(prev.var_code.get())
        row.var_scale.set(prev.var_scale.get())
        row.var_custom.set(prev.var_custom.get())
        if prev.chk_under.get():
            row.chk_under.select()
        else:
            row.chk_under.deselect()
        row.var_suspend.set(prev.var_suspend.get())
        row._toggle_suspend()

    # ---------- 调整因素卡片 ----------
    def _build_adjust_card(self):
        card = Card(self.canvas, title="调整因素（依据总说明五，14 条规则；默认全部叠加，可分组统筹）")

        # 复用
        self.var_reuse = tk.StringVar(value="none")
        self.var_reuse_scheme = tk.StringVar(value="normal")
        self.var_reuse_prelim = tk.StringVar(value="ratio")
        self.var_reuse_ratio = tk.StringVar(value="0")
        r1 = ctk.CTkFrame(card, fg_color="transparent")
        r1.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(r1, text="复用套图", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(r1, values=["none  不适用", "1a  全部复用(施工图×10%)", "1b  上部复用(施工图×20%)"],
                          variable=self.var_reuse, width=230, height=30, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(r1, text="方案:", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        ctk.CTkOptionMenu(r1, values=["normal  照常", "cancel  取消", "ratio  按比例"],
                          variable=self.var_reuse_scheme, width=100, height=28, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=10)).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(r1, text="初步:", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        ctk.CTkOptionMenu(r1, values=["normal  照常", "cancel  取消", "ratio  按比例(默认)"],
                          variable=self.var_reuse_prelim, width=110, height=28, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=10)).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(r1, text="自定义比例%", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        e = Entry(r1, width=56, numeric=True)
        e.pack(side="left", padx=(4, 0))
        e.configure(textvariable=self.var_reuse_ratio)

        # 改扩建
        self.var_renovation = tk.StringVar(value="none")
        self.var_renovation_custom = tk.StringVar(value="30")
        r2 = ctk.CTkFrame(card, fg_color="transparent")
        r2.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(r2, text="改扩建", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(r2, values=["none  不适用", "simple  简单(+10%)", "general  一般(+20%)",
                                      "complex  复杂(+30%)", "special  特别复杂(+40%)", "custom  自定义"],
                          variable=self.var_renovation, width=230, height=30, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(r2, text="自定义%", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        e = Entry(r2, width=56, numeric=True)
        e.pack(side="left", padx=(4, 0))
        e.configure(textvariable=self.var_renovation_custom)

        # 勾选类
        self.chk_vars = {}
        r3 = ctk.CTkFrame(card, fg_color="transparent")
        r3.pack(fill="x", padx=14, pady=3)
        checks = [
            ("separate_entrust", "分阶段分别委托(+10% 初步/施工图)"),
            ("co_design", "多单位共同设计(主体+10%)"),
            ("civil_air_defense", "五级及以上附建式人防(+15% 初步/施工图)"),
            ("scenic", "风景区建筑(+30%，适用所有酒店类)"),
        ]
        for i, (key, text) in enumerate(checks):
            v = tk.BooleanVar(value=False)
            self.chk_vars[key] = v
            cb = ctk.CTkCheckBox(r3, text=text, variable=v, font=ctk.CTkFont(size=11),
                                 text_color=COLOR_DARK, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT2)
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=4, pady=3)

        # 特殊建筑 / 涉外 / 工作量变化
        self.var_special = tk.StringVar(value="0")
        self.var_foreign = tk.StringVar(value="0")
        self.var_workchange = tk.StringVar(value="0")
        r4 = ctk.CTkFrame(card, fg_color="transparent")
        r4.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(r4, text="特殊建筑(古建等)增加月", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left")
        e1 = Entry(r4, width=60, numeric=True)
        e1.pack(side="left", padx=4)
        e1.configure(textvariable=self.var_special)
        ctk.CTkLabel(r4, text="涉外工程增加月(默认0)", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(14, 0))
        e2 = Entry(r4, width=60, numeric=True)
        e2.pack(side="left", padx=4)
        e2.configure(textvariable=self.var_foreign)
        ctk.CTkLabel(r4, text="工作量变化增加月", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(14, 0))
        e3 = Entry(r4, width=60, numeric=True)
        e3.pack(side="left", padx=4)
        e3.configure(textvariable=self.var_workchange)

        # 节假日
        self.var_holiday_mode = tk.StringVar(value="none")
        self.var_holiday_input = tk.StringVar(value="0")
        self.var_holiday_override = tk.StringVar(value="0")
        self.var_holiday_note = tk.StringVar(value="")
        r5 = ctk.CTkFrame(card, fg_color="transparent")
        r5.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(r5, text="节假日", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(r5, values=["none  不增加", "auto  自动统计(不含周六周日,推演至2039)", "input  输入天数"],
                          variable=self.var_holiday_mode, width=200, height=30, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(r5, text="天数", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        e1 = Entry(r5, width=48, numeric=True)
        e1.pack(side="left", padx=4)
        e1.configure(textvariable=self.var_holiday_input)
        ctk.CTkLabel(r5, text="覆写", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(10, 0))
        e2 = Entry(r5, width=48, numeric=True)
        e2.pack(side="left", padx=4)
        e2.configure(textvariable=self.var_holiday_override)
        ctk.CTkLabel(r5, text="注释", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(10, 0))
        e3 = Entry(r5, width=180)
        e3.pack(side="left", padx=4)
        e3.configure(textvariable=self.var_holiday_note)

        # 压缩周期
        self.var_compress_mode = tk.StringVar(value="none")
        self.var_compress_value = tk.StringVar(value="0")
        r6 = ctk.CTkFrame(card, fg_color="transparent")
        r6.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(r6, text="压缩周期", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkOptionMenu(r6, values=["none  不压缩", "target  压缩后总周期(月)", "days  压缩天数"],
                          variable=self.var_compress_mode, width=200, height=30, corner_radius=6,
                          fg_color="#FBF9F5", button_color=COLOR_NAVY, button_hover_color=COLOR_NAVY_HOVER,
                          text_color=COLOR_DARK, font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 0))
        ctk.CTkLabel(r6, text="值", font=ctk.CTkFont(size=11), text_color=COLOR_GRAY).pack(side="left", padx=(12, 0))
        e = Entry(r6, width=60, numeric=True)
        e.pack(side="left", padx=4)
        e.configure(textvariable=self.var_compress_value)

        # 新技术
        tech_card = ctk.CTkFrame(card, fg_color="transparent")
        tech_card.pack(fill="x", padx=10, pady=6)
        bar = ctk.CTkFrame(tech_card, fg_color="transparent")
        bar.pack(fill="x", padx=4)
        ctk.CTkLabel(bar, text="新技术/特殊要求（规则9）", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLOR_DARK).pack(side="left")
        ctk.CTkButton(bar, text="+ 添加新技术项", width=110, height=26, corner_radius=6,
                      fg_color=COLOR_ACCENT2, hover_color=COLOR_ACCENT2_HOVER, text_color="#FFFFFF",
                      font=ctk.CTkFont(size=10), command=self._add_tech).pack(side="right")
        self.tech_container = ctk.CTkFrame(tech_card, fg_color="transparent")
        self.tech_container.pack(fill="x", padx=4, pady=4)

        # 返工
        rework_card = ctk.CTkFrame(card, fg_color="transparent")
        rework_card.pack(fill="x", padx=10, pady=6)
        bar = ctk.CTkFrame(rework_card, fg_color="transparent")
        bar.pack(fill="x", padx=4)
        ctk.CTkLabel(bar, text="返工/修改（规则11）", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLOR_DARK).pack(side="left")
        ctk.CTkButton(bar, text="+ 添加返工项", width=110, height=26, corner_radius=6,
                      fg_color=COLOR_ACCENT2, hover_color=COLOR_ACCENT2_HOVER, text_color="#FFFFFF",
                      font=ctk.CTkFont(size=10), command=self._add_rework).pack(side="right")
        self.rework_container = ctk.CTkFrame(rework_card, fg_color="transparent")
        self.rework_container.pack(fill="x", padx=4, pady=4)

        # 分组统筹
        self.var_pool_groups = tk.StringVar(value="")
        r7 = ctk.CTkFrame(card, fg_color="transparent")
        r7.pack(fill="x", padx=14, pady=(6, 10))
        ctk.CTkLabel(r7, text="分组统筹", width=120, anchor="w", font=ctk.CTkFont(size=12),
                     text_color=COLOR_GRAY).pack(side="left")
        ctk.CTkLabel(r7, text="格式：规则key=组号,如 9=1,11=1（同组取最长）", font=ctk.CTkFont(size=10),
                     text_color=COLOR_GRAY).pack(side="left", padx=(6, 0))
        e = Entry(r7, width=260)
        e.pack(side="left", padx=(10, 0))
        e.configure(textvariable=self.var_pool_groups)

    def _add_tech(self):
        row = NewTechRow(self.tech_container, on_delete=lambda: self._del_row(self.tech_rows, row),
                         idx=len(self.tech_rows) + 1)
        self.tech_rows.append(row)

    def _add_rework(self):
        row = ReworkRow(self.rework_container, on_delete=lambda: self._del_row(self.rework_rows, row))
        self.rework_rows.append(row)

    def _del_row(self, rows, row):
        row.destroy()
        rows.remove(row)

    # ---------- 输出目录 ----------
    def _choose_dir(self):
        d = filedialog.askdirectory(title="选择 Word/Excel 输出文件夹")
        if d:
            self.var_output_dir.set(d)

    @staticmethod
    def _is_file_locked(path):
        """检测目标文件是否被其他进程占用（只读探测，不写入、不破坏文件）。"""
        try:
            with open(path, "a"):
                pass
            return False
        except PermissionError:
            return True
        except OSError:
            return False

    @staticmethod
    def _unique_path(out_dir, base, ext):
        """生成不冲突的路径：base_2.ext、base_3.ext …"""
        n = 2
        while True:
            cand = os.path.join(out_dir, f"{base}_{n}{ext}")
            if not os.path.exists(cand):
                return cand
            n += 1

    # ---------- 组装输入并计算 ----------
    def _collect(self):
        def f(v):
            try:
                return float(v.get().strip() or 0)
            except ValueError:
                return 0.0

        units = []
        for row in self.unit_rows:
            u = row.get_unit()
            if not u.name:
                messagebox.showerror("输入错误", "存在名称为空的单体，请填写名称。")
                raise ValueError("单体名称不完整")
            if not u.suspended and (not u.scale or u.scale <= 0):
                messagebox.showerror("输入错误", f"单体「{u.name}」规模必填且需大于 0（㎡）。")
                raise ValueError("单体规模不完整")
            units.append(u)
        if not units:
            raise ValueError("至少需要一个单体")

        # v1.2 第5项：开始时间格式校验（YYYY-MM-DD）
        start_raw = self.var_start.get().strip()
        if start_raw:
            try:
                datetime.date.fromisoformat(start_raw)
            except ValueError:
                messagebox.showerror("输入错误", f"开始时间格式应为 YYYY-MM-DD（当前：{start_raw}）。")
                raise ValueError("开始时间格式错误")

        adj = AdjustmentSpec()

        # 复用
        adj.reuse = self.var_reuse.get().split(" ")[0]
        adj.reuse_scheme_mode = self.var_reuse_scheme.get().split(" ")[0]
        adj.reuse_prelim_mode = self.var_reuse_prelim.get().split(" ")[0]
        adj.reuse_ratio = f(self.var_reuse_ratio) / 100.0 if f(self.var_reuse_ratio) > 0 else 0.0

        # 改扩建
        ren = self.var_renovation.get().split(" ")[0]
        adj.renovation = ren
        adj.renovation_custom = f(self.var_renovation_custom) / 100.0

        # 勾选
        adj.separate_entrust = self.chk_vars["separate_entrust"].get()
        adj.co_design = self.chk_vars["co_design"].get()
        adj.civil_air_defense = self.chk_vars["civil_air_defense"].get()
        adj.scenic = self.chk_vars["scenic"].get()

        adj.special_months = f(self.var_special)
        adj.foreign_months = f(self.var_foreign)
        adj.work_change_months = f(self.var_workchange)

        # 节假日
        adj.holiday_mode = self.var_holiday_mode.get().split(" ")[0]
        adj.holiday_input_days = f(self.var_holiday_input)
        adj.holiday_override_days = f(self.var_holiday_override)
        adj.holiday_note = self.var_holiday_note.get().strip()

        # 压缩
        adj.compress_mode = self.var_compress_mode.get().split(" ")[0]
        adj.compress_value = f(self.var_compress_value)

        # 新技术 / 返工
        adj.new_tech_items = [r.get_item() for r in self.tech_rows]
        adj.rework_items = [r.get_item() for r in self.rework_rows]

        # 分组统筹
        pool_groups = {}
        raw = self.var_pool_groups.get().strip()
        if raw:
            for part in raw.replace("，", ",").split(","):
                part = part.strip()
                if not part:
                    continue
                if "=" in part:
                    k, v = part.split("=", 1)
                    try:
                        pool_groups[k.strip()] = int(v.strip())
                    except ValueError:
                        pass
        adj.pool_groups = pool_groups

        # 项目
        stage_mode = self.var_stage_mode.get().split(" ")[0]
        proj = ProjectSpec(
            project_name=self.var_project.get().strip(),
            location=self.var_location.get().strip(),
            designer=self.var_designer.get().strip(),
            start_date=self.var_start.get().strip(),
            stage_mode=stage_mode,
            units=units,
            multi_unit_reduction=self.var_multi_mode.get().split(" ")[0],
            multi_unit_value=f(self.var_multi_value),
            rounding_mode=self.var_round_mode.get().split(" ")[0],
            digits=max(0, int(f(self.var_digits) or 2)),
            adjustment=adj,
        )
        return proj

    def _on_calculate(self):
        if self._busy:
            return
        try:
            proj = self._collect()
        except ValueError as e:
            messagebox.showerror("输入有误", str(e))
            return
        # v1.2 第4项：生成过程防误操作（禁用计算按钮与关键输入，防止连点重复计算）
        self._set_busy(True)
        try:
            result = calculate(proj)
        except Exception as e:
            self._set_busy(False)
            messagebox.showerror("计算失败", str(e))
            return

        # 结果展示（兜底保护：任何异常均明确提示，避免 tkinter 静默吞掉异常导致"按钮无效"）
        try:
            self.res_text.configure(state="normal")
            self.res_text.delete("1.0", "end")
            lines = []
            lines.append(f"项目：{result.project.project_name or '未命名项目'}")
            lines.append(f"阶段模式：{'三阶段' if result.project.stage_mode == 'three' else '两阶段（初步并入施工图）'}")
            lines.append("")
            lines.append("各阶段周期（月 / 周 / 天）：")
            lines.append(f"  方案设计：{result.stage_scheme:.2f}（{months_to_weeks(result.stage_scheme):.2f} 周 / {months_to_days_int(result.stage_scheme)} 天）")
            lines.append(f"  初步设计：{result.stage_preliminary:.2f}（{months_to_weeks(result.stage_preliminary):.2f} 周 / {months_to_days_int(result.stage_preliminary)} 天）")
            lines.append(f"  施工图设计：{result.stage_construction:.2f}（{months_to_weeks(result.stage_construction):.2f} 周 / {months_to_days_int(result.stage_construction)} 天）")
            lines.append(f"  调整后合计：{result.regular_total:.2f} 月（{months_to_days_int(result.regular_total)} 天）")
            lines.append(f"  最终周期：{result.final_total:.2f} 月（{months_to_weeks(result.final_total):.2f} 周 / {months_to_days_int(result.final_total)} 天）")
            if result.compress.active:
                lines.append(f"  压缩后占比：{result.compress.ratio:.1f}%" + (" ⚠ 低于50%" if result.compress.warn else ""))
            # v1.2 第9项：GUI 结果区同步起止日期
            if result.stage_dates:
                lines.append("")
                lines.append("各阶段起止日期（1 月 = 28 天推算）：")
                for sd in result.stage_dates:
                    lines.append(f"  {sd.stage_name}：{sd.start} ~ {sd.end}（{sd.days} 天 / {sd.months:.2f} 月）")
            else:
                lines.append("")
                lines.append("各阶段起止日期：未填写开始时间，未推算")
            lines.append("")
            lines.append("分设计阶段统计（各单体 × 阶段，调整后，月 / 天）：")
            for ur in result.unit_results:
                lines.append(
                    f"  {ur.spec.name}：方案 {ur.adjusted_scheme:.2f}/{months_to_days_int(ur.adjusted_scheme)}天，"
                    f"初步 {ur.adjusted_preliminary:.2f}/{months_to_days_int(ur.adjusted_preliminary)}天，"
                    f"施工图 {ur.adjusted_construction:.2f}/{months_to_days_int(ur.adjusted_construction)}天，"
                    f"小计 {ur.adjusted_total:.2f}/{months_to_days_int(ur.adjusted_total)}天"
                )
            lines.append(
                f"  项目合计：方案 {result.stage_scheme:.2f}/{months_to_days_int(result.stage_scheme)}天，"
                f"初步 {result.stage_preliminary:.2f}/{months_to_days_int(result.stage_preliminary)}天，"
                f"施工图 {result.stage_construction:.2f}/{months_to_days_int(result.stage_construction)}天，"
                f"小计 {result.regular_total:.2f}/{months_to_days_int(result.regular_total)}天"
            )
            # v1.2 第10项：GUI 结果区同步计算链等式
            if result.equation_text:
                lines.append("")
                lines.append("计算链等式：")
                lines.append("  " + result.equation_text)
            if result.warnings:
                lines.append("")
                lines.append("风险提示：")
                for w in result.warnings:
                    lines.append("  · " + w)
            self.res_text.insert("1.0", "\n".join(lines))
            self.res_text.configure(state="disabled")
        except Exception as e:
            self._set_busy(False)
            messagebox.showerror("结果展示失败", f"{e}\n{traceback.format_exc()}")
            return

        # 压缩归 0/过小警告（v1.1.2）：最终周期不足 1 个月时明确提示建议减小压缩值
        if result.compress.active and result.final_total < 1.0:
            messagebox.showwarning(
                "压缩警告",
                f"压缩后最终周期仅 {result.final_total:.2f} 月（不足 1 个月），周期过小可能不满足实际设计需要。\n"
                "建议减小压缩天数或改用「压缩到目标周期」模式后重新计算。")

        # 保存报告
        out_dir = self.var_output_dir.get().strip() or os.path.expanduser("~")
        base = self._sanitize_filename(self.var_output_name.get().strip() or REPORT_DEFAULT_NAME).replace(".docx", "")
        if not base:
            base = REPORT_DEFAULT_NAME
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except OSError:
                self._set_busy(False)
                messagebox.showerror("输出路径错误", f"无法创建目录：{out_dir}")
                return
        docx_path = os.path.join(out_dir, f"{base}.docx")
        locked_notice = ""
        try:
            if os.path.exists(docx_path) and self._is_file_locked(docx_path):
                # 目标文件正被其他进程（如 Word/WPS）占用：自动另存为带序号的新文件，避免再次保存失败
                alt = self._unique_path(out_dir, base, ".docx")
                generate_report(result, alt)
                docx_path = alt
                locked_notice = f"\n（原文件 {base}.docx 正被占用，已自动另存为 {os.path.basename(alt)}，请先关闭占用该文件的程序）"
            else:
                generate_report(result, docx_path)
        except Exception as e:
            self._set_busy(False)
            messagebox.showerror("报告生成失败", f"{e}\n{traceback.format_exc()}")
            return
        msg = f"Word 报告已生成：\n{docx_path}\n{locked_notice}".rstrip()
        self._last_docx_path = docx_path
        self._last_xlsx_path = ""
        if self.var_excel.get():
            xlsx_path = os.path.join(out_dir, f"{base}.xlsx")
            try:
                if os.path.exists(xlsx_path) and self._is_file_locked(xlsx_path):
                    alt = self._unique_path(out_dir, base, ".xlsx")
                    export_excel(result, alt)
                    self._last_xlsx_path = alt
                    msg += f"\nExcel 已导出（原文件被占用，另存为 {os.path.basename(alt)}）：\n{alt}"
                else:
                    export_excel(result, xlsx_path)
                    self._last_xlsx_path = xlsx_path
                    msg += f"\nExcel 已导出：\n{xlsx_path}"
            except Exception as e:
                msg += f"\nExcel 导出失败：{e}\n"
        self._set_busy(False)
        # v1.2 第1项：生成成功后保存配置；第2项：一键打开（自定义结果弹窗）
        self._save_config()
        self._show_result_dialog("完成", msg)

    # ---------- v1.2 新增：配置记忆 / 防误操作 / 一键打开 / 文件名联动 / 复制（第1-8项） ----------
    @staticmethod
    def _sanitize_filename(name):
        """过滤文件名非法字符（第5项）。"""
        for ch in '\\/:*?"<>|':
            name = name.replace(ch, "")
        return name.strip()

    def _all_widgets(self):
        out = []
        stack = [self]
        while stack:
            w = stack.pop()
            out.append(w)
            try:
                stack.extend(w.winfo_children())
            except Exception:
                pass
        return out

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.btn_calc.configure(text="正在生成…" if busy else "开始计算并生成报告")
        for w in self._all_widgets():
            try:
                if isinstance(w, (ctk.CTkEntry, ctk.CTkOptionMenu, ctk.CTkCheckBox, ctk.CTkButton)):
                    w.configure(state=state)
            except Exception:
                pass
        if not busy:
            for row in self.unit_rows:
                row._toggle_suspend()

    def _save_config(self):
        """保存全部 GUI 参数到用户级配置文件（第1项）。"""
        try:
            cfg = {
                "project": self.var_project.get(),
                "location": self.var_location.get(),
                "designer": self.var_designer.get(),
                "start": self.var_start.get(),
                "stage_mode": self.var_stage_mode.get(),
                "round_mode": self.var_round_mode.get(),
                "digits": self.var_digits.get(),
                "multi_mode": self.var_multi_mode.get(),
                "multi_value": self.var_multi_value.get(),
                "output_dir": self.var_output_dir.get(),
                "output_name": self.var_output_name.get(),
                "excel": self.var_excel.get(),
                "output_name_auto": self._output_name_auto,
                "reuse": self.var_reuse.get(),
                "reuse_scheme": self.var_reuse_scheme.get(),
                "reuse_prelim": self.var_reuse_prelim.get(),
                "reuse_ratio": self.var_reuse_ratio.get(),
                "renovation": self.var_renovation.get(),
                "renovation_custom": self.var_renovation_custom.get(),
                "checks": {k: v.get() for k, v in self.chk_vars.items()},
                "special": self.var_special.get(),
                "foreign": self.var_foreign.get(),
                "workchange": self.var_workchange.get(),
                "holiday_mode": self.var_holiday_mode.get(),
                "holiday_input": self.var_holiday_input.get(),
                "holiday_override": self.var_holiday_override.get(),
                "holiday_note": self.var_holiday_note.get(),
                "compress_mode": self.var_compress_mode.get(),
                "compress_value": self.var_compress_value.get(),
                "pool_groups": self.var_pool_groups.get(),
                "units": [
                    {"name": r.var_name.get(), "code": r.var_code.get(), "scale": r.var_scale.get(),
                     "custom": r.var_custom.get(), "under": bool(r.chk_under.get()),
                     "suspend": r.var_suspend.get()}
                    for r in self.unit_rows
                ],
                "techs": [
                    {"name": r.var_name.get(), "desc": r.var_desc.get(), "months": r.var_months.get(),
                     "pool": r.var_pool.get(), "group": r.var_group.get()}
                    for r in self.tech_rows
                ],
                "reworks": [
                    {"stage": r.var_stage.get(), "months": r.var_months.get(), "note": r.var_note.get()}
                    for r in self.rework_rows
                ],
            }
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_config(self):
        """启动时恢复上次配置（第1项）。"""
        if not os.path.exists(CONFIG_PATH):
            return
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return
        try:
            self.var_project.set(cfg.get("project", ""))
            self.var_location.set(cfg.get("location", ""))
            self.var_designer.set(cfg.get("designer", ""))
            if cfg.get("start"):
                self.var_start.set(cfg["start"])
            self.var_stage_mode.set(cfg.get("stage_mode", self.var_stage_mode.get()))
            self.var_round_mode.set(cfg.get("round_mode", self.var_round_mode.get()))
            self.var_digits.set(cfg.get("digits", self.var_digits.get()))
            self.var_multi_mode.set(cfg.get("multi_mode", self.var_multi_mode.get()))
            self.var_multi_value.set(cfg.get("multi_value", self.var_multi_value.get()))
            if cfg.get("output_dir"):
                self.var_output_dir.set(cfg["output_dir"])
            if cfg.get("output_name"):
                self.var_output_name.set(cfg["output_name"])
            self.var_excel.set(cfg.get("excel", False))
            self._output_name_auto = cfg.get("output_name_auto", True)
            self.var_reuse.set(cfg.get("reuse", self.var_reuse.get()))
            self.var_reuse_scheme.set(cfg.get("reuse_scheme", self.var_reuse_scheme.get()))
            self.var_reuse_prelim.set(cfg.get("reuse_prelim", self.var_reuse_prelim.get()))
            self.var_reuse_ratio.set(cfg.get("reuse_ratio", self.var_reuse_ratio.get()))
            self.var_renovation.set(cfg.get("renovation", self.var_renovation.get()))
            self.var_renovation_custom.set(cfg.get("renovation_custom", self.var_renovation_custom.get()))
            for k, v in cfg.get("checks", {}).items():
                if k in self.chk_vars:
                    self.chk_vars[k].set(v)
            self.var_special.set(cfg.get("special", self.var_special.get()))
            self.var_foreign.set(cfg.get("foreign", self.var_foreign.get()))
            self.var_workchange.set(cfg.get("workchange", self.var_workchange.get()))
            self.var_holiday_mode.set(cfg.get("holiday_mode", self.var_holiday_mode.get()))
            self.var_holiday_input.set(cfg.get("holiday_input", self.var_holiday_input.get()))
            self.var_holiday_override.set(cfg.get("holiday_override", self.var_holiday_override.get()))
            self.var_holiday_note.set(cfg.get("holiday_note", self.var_holiday_note.get()))
            self.var_compress_mode.set(cfg.get("compress_mode", self.var_compress_mode.get()))
            self.var_compress_value.set(cfg.get("compress_value", self.var_compress_value.get()))
            self.var_pool_groups.set(cfg.get("pool_groups", self.var_pool_groups.get()))
            # 单体/新技术/返工行
            for row in list(self.unit_rows):
                row.destroy()
                self.unit_rows.remove(row)
            for u in cfg.get("units", []):
                self._add_unit()
                r = self.unit_rows[-1]
                r.var_name.set(u.get("name", r.var_name.get()))
                if u.get("code"):
                    r.var_code.set(u["code"])
                r.var_scale.set(u.get("scale", r.var_scale.get()))
                r.var_custom.set(u.get("custom", ""))
                if u.get("under"):
                    r.chk_under.select()
                r.var_suspend.set(u.get("suspend", False))
                r._toggle_suspend()
            for row in list(self.tech_rows):
                row.destroy()
                self.tech_rows.remove(row)
            for t in cfg.get("techs", []):
                self._add_tech()
                r = self.tech_rows[-1]
                r.var_name.set(t.get("name", r.var_name.get()))
                r.var_desc.set(t.get("desc", ""))
                r.var_months.set(t.get("months", r.var_months.get()))
                r.var_pool.set(t.get("pool", False))
                r.var_group.set(t.get("group", r.var_group.get()))
            for row in list(self.rework_rows):
                row.destroy()
                self.rework_rows.remove(row)
            for rw in cfg.get("reworks", []):
                self._add_rework()
                r = self.rework_rows[-1]
                r.var_stage.set(rw.get("stage", r.var_stage.get()))
                r.var_months.set(rw.get("months", r.var_months.get()))
                r.var_note.set(rw.get("note", ""))
        except Exception:
            pass

    def _on_project_changed(self, *_):
        """项目名变化时自动联动输出文件名（第3项，不覆盖手动改过的文件名）。"""
        if not self._output_name_auto:
            return
        proj = self.var_project.get().strip()
        self._auto_setting = True
        try:
            if proj:
                self.var_output_name.set(f"{proj}_{REPORT_DEFAULT_NAME}")
            else:
                self.var_output_name.set(REPORT_DEFAULT_NAME)
        finally:
            self._auto_setting = False

    def _on_output_name_changed(self, *_):
        if getattr(self, "_auto_setting", False):
            return
        self._output_name_auto = False

    def _copy_result(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.res_text.get("1.0", "end").strip())
            messagebox.showinfo("已复制", "计算结果已复制到剪贴板。")
        except Exception as e:
            messagebox.showerror("复制失败", str(e))

    def _show_result_dialog(self, title, msg):
        """自定义结果弹窗：含「打开文件夹」「打开报告」「复制路径」按钮（第2项）。"""
        dlg = ctk.CTkToplevel(self)
        dlg.title(title)
        dlg.geometry("640x420")
        dlg.configure(fg_color=COLOR_BG)
        dlg.transient(self)
        dlg.grab_set()
        dlg.attributes("-topmost", True)
        ctk.CTkLabel(dlg, text="生成完成", font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=COLOR_DARK).pack(anchor="w", padx=18, pady=(16, 6))
        box = ctk.CTkTextbox(dlg, height=210, corner_radius=8, fg_color="#FDFBF7",
                             border_color=COLOR_BORDER, border_width=1,
                             font=ctk.CTkFont(size=12), text_color=COLOR_DARK)
        box.pack(fill="x", padx=18, pady=6)
        box.insert("1.0", msg)
        box.configure(state="disabled")
        row = ctk.CTkFrame(dlg, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=12)
        btn_cfg = dict(height=32, corner_radius=6, font=ctk.CTkFont(size=12))
        ctk.CTkButton(row, text="打开文件夹", width=110, fg_color=COLOR_NAVY,
                      hover_color=COLOR_NAVY_HOVER, text_color="#FFFFFF", **btn_cfg,
                      command=lambda: self._open_folder(os.path.dirname(self._last_docx_path))).pack(side="left", padx=2)
        if self._last_docx_path:
            ctk.CTkButton(row, text="打开报告", width=110, fg_color=COLOR_ACCENT,
                          hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", **btn_cfg,
                          command=lambda: self._open_file(self._last_docx_path)).pack(side="left", padx=2)
        if self._last_xlsx_path:
            ctk.CTkButton(row, text="打开Excel", width=110, fg_color=COLOR_ACCENT2,
                          hover_color=COLOR_ACCENT2_HOVER, text_color="#FFFFFF", **btn_cfg,
                          command=lambda: self._open_file(self._last_xlsx_path)).pack(side="left", padx=2)
        ctk.CTkButton(row, text="复制路径", width=100, fg_color="#FBF9F5", hover_color="#E8E4DA",
                      text_color=COLOR_DARK, border_width=1, border_color=COLOR_BORDER, **btn_cfg,
                      command=lambda: self._copy_paths()).pack(side="left", padx=2)
        ctk.CTkButton(row, text="关闭", width=80, fg_color="#FBF9F5", hover_color="#E8E4DA",
                      text_color=COLOR_DARK, border_width=1, border_color=COLOR_BORDER, **btn_cfg,
                      command=dlg.destroy).pack(side="right", padx=2)

    @staticmethod
    def _open_folder(path):
        try:
            os.startfile(path)
        except Exception:
            messagebox.showerror("打开失败", f"无法打开文件夹：{path}")

    @staticmethod
    def _open_file(path):
        try:
            os.startfile(path)
        except Exception:
            messagebox.showerror("打开失败", f"无法打开文件：{path}")

    def _copy_paths(self):
        parts = [p for p in (self._last_docx_path, self._last_xlsx_path) if p]
        if not parts:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append("\n".join(parts))
            messagebox.showinfo("已复制", "文件路径已复制到剪贴板。")
        except Exception as e:
            messagebox.showerror("复制失败", str(e))

    def _on_close(self):
        self._save_config()
        self.destroy()

    # ---------- 运行 ----------
    def run(self):
        self.mainloop()


def main():
    # 两阶段提示（Q14 决策：由用户判断）
    app = QuotaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
