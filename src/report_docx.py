# -*- coding: utf-8 -*-
"""Word 报告生成：基于 CalculationResult 生成带边框表格、浅彩计算依据标注的 .docx。

- 所有表格均带边框（Table Grid）
- 计算依据（定额原文表格取值、14 条规则原文、方法说明）以浅彩色底纹标注
- 最终周期给出 月 / 周 / 天 换算
"""
from __future__ import annotations

import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from quota_data import (
    QUOTA_TABLES, RULES_TEXT, TWO_STAGE_TEXT, EXCLUDED_ITEMS,
    DAYS_PER_MONTH, DAYS_PER_WEEK,
    NOTE_ABOVE_MAX, NOTE_BELOW_COVERED, NOTE_USER_CUSTOM,
    NOTE_INTERPOLATION, NOTE_AT_LEVEL, NOTE_BELOW_MIN,
    COMPRESS_RISK_TEXT, COMPRESS_WARN_TEXT,
)
from quota_engine import (CalculationResult, months_to_weeks, months_to_days,
                          months_to_days_int)
from version import APP_VERSION, APP_NAME, REPORT_TITLE, REPORT_FOOTER

# 浅彩色主题（黑白灰 + 高饱和点缀的浅色变体）
C_BG_HEAD = "D9D9D9"      # 表头灰
C_BG_ACCENT = "FFF2CC"    # 浅黄（计算依据）
C_BG_BLUE = "DEEBF7"      # 浅蓝（取值依据）
C_BG_GREEN = "E2EFDA"     # 浅绿（结果）
C_BG_RED = "FCE4E4"       # 浅红（警告）
C_BG_ORANGE = "FDE9D9"    # 浅橙（调整明细）
C_TXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)

METHOD_NOTES = {
    "at_level": NOTE_AT_LEVEL,
    "interpolation": NOTE_INTERPOLATION,
    "below_min": NOTE_BELOW_MIN,
    "above_max": NOTE_ABOVE_MAX,
    "below_covered": NOTE_BELOW_COVERED,
    "custom": NOTE_USER_CUSTOM,
}

STAGE_LABELS = {
    "scheme": "方案设计",
    "preliminary": "初步设计",
    "construction": "施工图设计",
}


def _set_cell_bg(cell, hex_color):
    """设置单元格底纹颜色（浅彩色标注）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _style_table(table):
    """统一表格样式：边框、表头加粗灰底。"""
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "微软雅黑"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                    run.font.size = Pt(10)
                    run.font.color.rgb = C_TXT_DARK
            if i == 0:
                _set_cell_bg(cell, C_BG_HEAD)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True


def _add_field(paragraph, field_code):
    """在段落中插入 Word 域（用于页码等动态字段）。"""
    run = paragraph.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = field_code
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    return run


def _setup_header_footer(doc):
    """页眉：报告标题；页脚：居中页码“第 X 页 共 Y 页”（v1.2 第 15 项）。"""
    for section in doc.sections:
        # 页眉
        hp = section.header.paragraphs[0]
        hp.text = ""
        hr = hp.add_run(f"{REPORT_TITLE}（{APP_NAME} v{APP_VERSION}）")
        hr.font.name = "微软雅黑"
        hr._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        hr.font.size = Pt(9)
        hr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 页脚
        fp = section.footer.paragraphs[0]
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = fp.add_run("第 ")
        r1.font.name = "微软雅黑"
        r1._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        r1.font.size = Pt(9)
        r1.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        _add_field(fp, "PAGE")
        r2 = fp.add_run(" 页 共 ")
        r2.font.name = "微软雅黑"
        r2._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        _add_field(fp, "NUMPAGES")
        r3 = fp.add_run(" 页")
        r3.font.name = "微软雅黑"
        r3._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        r3.font.size = Pt(9)
        r3.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        for r in (r1, r2, r3):
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")


def _heading(doc, text, level=1, color="1A1A1A"):
    """带色标题（避开默认主题色，保持黑白灰基调 + 点缀）。"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.color.rgb = RGBColor.from_string(color)
    return h


def _para(doc, text, bold=False, size=10, color=None, italic=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return p


def _table(doc, headers, rows, bg_map=None):
    """创建带边框表格；bg_map: {(row_idx, col_idx): hex} 或 {row_idx: hex}。"""
    t = doc.add_table(rows=1, cols=len(headers))
    hdr = t.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = str(htext)
    for r, row in enumerate(rows):
        cells = t.add_row().cells
        for c, val in enumerate(row):
            cells[c].text = "" if val is None else str(val)
    _style_table(t)
    if bg_map:
        for key, color in bg_map.items():
            if isinstance(key, tuple):
                _set_cell_bg(t.rows[key[0]].cells[key[1]], color)
            else:
                for c in range(len(t.rows[key].cells)):
                    _set_cell_bg(t.rows[key].cells[c], color)
    return t


def _fmt(v):
    """格式化周期数值：整数去尾零，保留 2 位小数。"""
    if v is None:
        return ""
    if isinstance(v, float) and abs(v - round(v)) < 1e-9:
        return f"{v:g}"
    return f"{v:g}"


def _fmt2(v):
    return f"{v:.2f}" if isinstance(v, float) else f"{v:g}"


def generate_report(result: CalculationResult, output_path: str,
                    show_quota_tables: bool = True) -> str:
    """生成 Word 报告并写入 output_path，返回路径。

    show_quota_tables：是否在“各单体规模取值”中附定额档位表并高亮当前档位（v1.2 第 13 项，默认开启）。
    """
    doc = Document()
    # 页面边距
    for section in doc.sections:
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
    _setup_header_footer(doc)

    # 默认字体
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    style.font.size = Pt(10)

    proj = result.project

    # ============ 1 标题 ============
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(REPORT_TITLE)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(20)
    run.bold = True
    run.font.color.rgb = RGBColor.from_string("1A1A1A")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run(proj.project_name or "未命名项目")
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(14)
    run.bold = True
    run.font.color.rgb = RGBColor(0xE8, 0x4B, 0x3C)  # 高饱和点缀

    _para(doc, f"报告生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
          size=9, color="808080").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ============ 2 项目信息 ============
    _heading(doc, "一、项目信息", 1, "1A1A1A")
    stage_mode_label = "三阶段（方案设计 → 初步设计 → 施工图设计）" if proj.stage_mode == "three" else "两阶段（方案设计 → 施工图设计，初步周期并入施工图）"
    info_rows = [
        ("项目名称", proj.project_name or ""),
        ("项目地点", proj.location or ""),
        ("设计单位", proj.designer or ""),
        ("委托阶段模式", stage_mode_label),
        ("计划开始时间", proj.start_date or ""),
        ("精度规则", {"round": "四舍五入", "ceil": "向上进位", "floor": "向下取整"}.get(proj.rounding_mode, proj.rounding_mode) + f"（保留 {proj.digits} 位小数）"),
        ("时间单位", "1 日历月 = 4 日历周 = 28 天；0.25 月 = 1 周 = 7 天"),
    ]
    t = _table(doc, ["项目信息", "内容"], info_rows, bg_map={0: C_BG_HEAD})
    t.columns[0].width = Cm(4)
    t.columns[1].width = Cm(12)

    if proj.stage_mode == "two":
        p = doc.add_paragraph()
        run = p.add_run("依据：总说明（二）" + TWO_STAGE_TEXT)
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(9)
        run.italic = True
        run.font.color.rgb = RGBColor(0x7F, 0x60, 0x00)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), C_BG_ACCENT)
        pPr.append(shd)

    # ============ 3 各单体规模取值与定额依据 ============
    _heading(doc, "二、各单体规模取值（定额依据）", 1, "1A1A1A")
    _para(doc, "规模取值规则：工程规模与定额表规模档位一致时直接取该档周期；位于两档之间时采用线性插值；"
               "低于起档规模时取最低档周期；高于最高档时按“T = T最高档 × S / S最高档”等比推算并标注；"
               "定额未涵盖的规模区间按疗养类 10000 ㎡档向下等比推算并标注。", size=9, color="808080")

    unit_rows = []
    bg_map = {0: C_BG_HEAD}
    for idx, ur in enumerate(result.unit_results):
        sr = ur.scale
        table = QUOTA_TABLES.get(sr.code)
        basis = ""
        if table is not None:
            basis = f"{sr.code} {table['name']}（{table['category']}）定额表"
        row = [ur.spec.name, sr.code, table["name"] if table else sr.code,
               _fmt(sr.scale), sr.method, _fmt(sr.scheme), _fmt(sr.preliminary),
               _fmt(sr.construction), _fmt(sr.total)]
        unit_rows.append(row)
        # 浅蓝标注取值依据行
        bg_map[idx + 1] = C_BG_BLUE
    _table(doc,
           ["单体名称", "定额编号", "定额子项", "规模（㎡）", "取值方法", "方案", "初步", "施工图", "合计（月）"],
           unit_rows, bg_map=bg_map)

    # v1.2 第 13 项：定额档位原表引用（精简列：规模/方案/初步/施工图/合计），高亮当前取值档位
    if show_quota_tables:
        for ur in result.unit_results:
            sr = ur.scale
            table = QUOTA_TABLES.get(sr.code)
            if table is None:
                continue
            hdr_row = [f"定额原表：{sr.code} {table['name']}（{table['category']}）——规模档位",
                       "方案", "初步", "施工图", "合计"]
            rows = []
            bg = {0: C_BG_HEAD}
            target_row = None
            for i, sc in enumerate(table["scales"]):
                rows.append([f"{sc:g} ㎡", _fmt(table["scheme"][i]), _fmt(table["preliminary"][i]),
                             _fmt(table["construction"][i]), _fmt(table["total"][i])])
                # 高亮当前取值所在档位：命中档 / 低于起档（第 1 档）/ 插值落在相邻两档之间 / 超档（最高档）
                if sr.method == "at_level" and abs(sc - sr.scale) < 1e-9:
                    target_row = i
                elif sr.method == "below_min" and i == 0:
                    target_row = i
                elif sr.method == "interpolation":
                    next_sc = table["scales"][i + 1] if i + 1 < len(table["scales"]) else None
                    if next_sc is not None and sc <= sr.scale < next_sc:
                        target_row = i
                elif sr.method in ("above_max", "below_covered") and i == len(table["scales"]) - 1:
                    target_row = i
            if target_row is not None:
                bg[target_row + 1] = C_BG_GREEN
            _table(doc, hdr_row, rows, bg_map=bg)
            _para(doc, f"↑ 绿色行为“{ur.spec.name}”规模 {_fmt(sr.scale)} ㎡ 对应的取值档位（方法：{sr.method}）。",
                  size=8, color="808080")

    # 依据说明（浅黄底纹段落）
    for ur in result.unit_results:
        sr = ur.scale
        note = METHOD_NOTES.get(sr.method, sr.note or "")
        p = doc.add_paragraph()
        run = p.add_run(f"依据说明（{ur.spec.name}，{sr.code}）：{note}")
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x7F, 0x60, 0x00)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), C_BG_ACCENT)
        pPr.append(shd)

    # ============ 4 各单体与阶段周期（调整后，分阶段统计） ============
    _heading(doc, "三、各单体与阶段周期（调整后分阶段统计）", 1, "1A1A1A")
    stage_rows = []
    bg_map = {0: C_BG_HEAD}
    for idx, ur in enumerate(result.unit_results):
        def _cell(months_val):
            return f"{_fmt2(months_val)} / {months_to_days_int(months_val)}天"
        stage_rows.append([ur.spec.name, _cell(ur.adjusted_scheme), _cell(ur.adjusted_preliminary),
                           _cell(ur.adjusted_construction), _cell(ur.adjusted_total)])
        bg_map[idx + 1] = C_BG_GREEN
    stage_rows.append(["项目合计（各单体调整后相加）",
                       _cell(result.stage_scheme), _cell(result.stage_preliminary),
                       _cell(result.stage_construction), _cell(result.stage_total)])
    bg_map[len(stage_rows)] = C_BG_GREEN
    _table(doc, ["单体", "方案（月/天）", "初步（月/天）", "施工图（月/天）", "阶段小计（月/天）"],
           stage_rows, bg_map=bg_map)
    _para(doc, "说明：阶段调整（复用套图、分阶段委托、人防、返工等）作用于阶段总值，表中各单体调整后值"
               "按各单体基础占比分摊并做末位校正，阶段小计与调整明细中的阶段总值一致；"
               "天 = 月 × 28，天数保留整数。", size=9, color="808080")

    if proj.stage_mode == "two":
        _para(doc, "两阶段模式：施工图设计周期 = 初步设计定额值 + 施工图设计定额值（依据总说明（二））。",
              size=9, color="808080")

    # ============ 5 调整明细 ============
    _heading(doc, "四、调整因素明细（规则原文 + 数值）", 1, "1A1A1A")
    _para(doc, "以下规则默认全部叠加；数字分组统筹的同组调整取最长周期，不可统筹项与各组结果继续叠加。",
          size=9, color="808080")

    adj_rows = []
    bg_map = {0: C_BG_HEAD}
    adj_keys = set()
    for ad in result.adjustments:
        # 跳过无实际影响的零记录（base=0 且 delta=0），避免报告噪音
        if abs(ad.base_value) < 1e-9 and abs(ad.delta) < 1e-9:
            continue
        adj_keys.add(ad.key)
        adj_rows.append([ad.key, ad.label, ad.applied_on, _fmt2(ad.base_value),
                         _fmt2(ad.delta), _fmt2(ad.final_value), ad.note])
        bg_map[len(adj_rows)] = C_BG_ORANGE
    _table(doc, ["规则", "名称", "作用对象", "调整前（月）", "增量（月）", "调整后（月）", "说明"],
           adj_rows, bg_map=bg_map)

    # 规则原文引用（浅黄）：仅引用有实际影响的规则
    seen = set()
    for ad in result.adjustments:
        if ad.key not in adj_keys or ad.key in seen:
            continue
        seen.add(ad.key)
        p = doc.add_paragraph()
        run = p.add_run(f"规则 {ad.key} 原文：{ad.rule_text}")
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x7F, 0x60, 0x00)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), C_BG_ACCENT)
        pPr.append(shd)

    # 规则2 改扩建 / 多单体折减 / 节假日 / 新技术等附加明细
    extra_rows = []
    extra_bg = {0: C_BG_HEAD}
    for it in getattr(result, "additive_items", []):
        if it.key in adj_keys:   # 已在调整明细表中展示，避免重复
            continue
        extra_rows.append([it.key, it.label, _fmt2(it.delta_months), it.rule_text])
        extra_bg[len(extra_rows)] = C_BG_ORANGE
    if extra_rows:
        _para(doc, "总周期附加项（含新技术、涉外、特殊建筑、工作量变化等）：", size=9, color="808080")
        _table(doc, ["规则", "名称", "增量（月）", "规则原文"], extra_rows, bg_map=extra_bg)

    # 多单体折减
    if proj.multi_unit_reduction == "days" and proj.multi_unit_value > 0:
        _para(doc, f"多单体折减：总周期减少 {_fmt2(proj.multi_unit_value)} 天（约 {_fmt2(proj.multi_unit_value / DAYS_PER_MONTH)} 月）。", size=9, color="808080")
    elif proj.multi_unit_reduction == "ratio" and proj.multi_unit_value > 0:
        _para(doc, f"多单体折减：总周期按 {proj.multi_unit_value * 100:.0f}% 比例折减。", size=9, color="808080")

    # 节假日
    if result.holiday_count > 0 or proj.adjustment.holiday_mode != "none":
        p = doc.add_paragraph()
        run = p.add_run(f"节假日增加：{_fmt2(result.holiday_count)} 天（国家法定节假日，不含周六、周日）" +
                        (f"，其中覆写 {_fmt2(result.holiday_override)} 天" if result.holiday_override > 0 else "") +
                        (f"。注释：{proj.adjustment.holiday_note}" if proj.adjustment.holiday_note else ""))
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x7F, 0x60, 0x00)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), C_BG_ACCENT)
        pPr.append(shd)

    # v1.2 第 11 项：节假日明细表（auto 自动统计时列出区间内每个法定节假日；否则列出输入/覆写天数）
    hd_rows = []
    hd_bg = {0: C_BG_HEAD}
    if result.holiday_details:
        for dstr, name, counted in result.holiday_details:
            hd_rows.append([dstr, name, "计入 +1 天" if counted else "不计入（周六/周日）"])
            hd_bg[len(hd_rows)] = C_BG_BLUE if counted else C_BG_HEAD
        _para(doc, "法定节假日明细（自动统计，口径：仅计入落在周一至周五的法定节假日）：", size=9, color="808080")
        _table(doc, ["日期", "节日", "是否计入"], hd_rows, bg_map=hd_bg)
        _para(doc, f"统计区间内共 {len([1 for _, _, c in result.holiday_details if c])} 个计入日，与上述“节假日增加 {_fmt2(result.holiday_count)} 天”一致。",
              size=9, color="808080")
    elif proj.adjustment.holiday_mode in ("input", "override") and (result.holiday_count > 0 or result.holiday_override > 0):
        _para(doc, "节假日设置（手动输入/覆写模式）：", size=9, color="808080")
        rows2 = []
        if proj.adjustment.holiday_mode == "input" and result.holiday_count > 0:
            rows2.append(["手动输入天数", f"{_fmt2(result.holiday_count)} 天",
                          f"由用户在导出设置中手工指定（{proj.adjustment.holiday_note or '无注释'}）"])
        if result.holiday_override > 0:
            rows2.append(["覆写天数", f"{_fmt2(result.holiday_override)} 天", "在自动/输入天数之外额外覆写增加的天数"])
        _table(doc, ["项目", "天数", "说明"], rows2, bg_map={0: C_BG_HEAD, 1: C_BG_BLUE, 2: C_BG_BLUE})

    # v1.2 第 12 项：默认推算项汇总清单（非“直接命中档位”的取值项）
    inferred = [ur for ur in result.unit_results if ur.scale.method in
                ("interpolation", "below_min", "above_max", "below_covered", "custom")]
    if inferred:
        _para(doc, "默认推算项汇总（非定额档位直接取值，便于复核）：", size=9, color="808080")
        inf_rows = []
        inf_bg = {0: C_BG_HEAD}
        for ur in inferred:
            sr = ur.scale
            inf_rows.append([ur.spec.name, sr.code, sr.method, _fmt(sr.scale), _fmt(sr.total),
                             METHOD_NOTES.get(sr.method, sr.note or "")])
            inf_bg[len(inf_rows)] = C_BG_ORANGE
        _table(doc, ["单体", "定额编号", "取值方法", "规模（㎡）", "采用值（合计月）", "说明"],
               inf_rows, bg_map=inf_bg)

    # ============ v1.2 第 9 项：各阶段起止日期推算 ============
    _heading(doc, "五、各阶段起止日期（推算）", 1, "1A1A1A")
    if not result.has_start_date:
        _para(doc, "未填写开始时间，未推算。", size=10)
    else:
        _para(doc, "推算口径：1 日历月 = 28 天；各阶段起止日期由计划开始时间顺推，阶段间按整月衔接。",
              size=9, color="808080")
        sd_rows = []
        sd_bg = {0: C_BG_HEAD}
        for sd in result.stage_dates:
            sd_rows.append([sd.stage_name, _fmt2(sd.months), sd.start, sd.end, sd.end])
            sd_bg[len(sd_rows)] = C_BG_BLUE
        _table(doc, ["阶段", "周期（月）", "开始日期", "结束日期", "完成节点"], sd_rows, bg_map=sd_bg)

    # ============ 6 最终周期换算 ============
    _heading(doc, "六、最终设计周期", 1, "1A1A1A")
    final_rows = [
        ("调整后总周期（压缩前）", _fmt2(result.regular_total), _fmt2(months_to_weeks(result.regular_total)),
         f"{months_to_days_int(result.regular_total)}"),
        ("最终设计周期", _fmt2(result.final_total), _fmt2(months_to_weeks(result.final_total)),
         f"{months_to_days_int(result.final_total)}"),
    ]
    if result.compress.active:
        final_rows.append(("压缩后周期占定额周期比例", f"{result.compress.ratio:.1f}%", "", ""))
    _table(doc, ["项目", "月", "周（×4）", "天（×28，取整）"], final_rows, bg_map={0: C_BG_HEAD, 1: C_BG_GREEN, 2: C_BG_GREEN})
    _para(doc, "换算公式：周 = 月 × 4；天 = 月 × 28（依据总说明时间单位），天数保留整数。", size=9, color="808080")

    # v1.2 第 10 项：完整计算链等式
    if result.equation_text:
        _para(doc, "计算链等式（定额合计 ± 各调整增量 - 压缩 = 最终周期）：", size=10, bold=True)
        _para(doc, result.equation_text, size=10, color="1A1A1A")
        _para(doc, "其中：定额合计 = 各单体按定额表取值之和；各调整增量 = 阶段调整、附加项、多单体折减、"
                   "节假日、新技术/返工等逐项累计；压缩 = 目标周期/压缩天数设定。",
              size=9, color="808080")

    # v1.2 第 14 项：多单体最终周期总览表（按规模降序）
    if len(result.unit_results) > 0:
        compressed_months = max(0.0, result.regular_total - result.final_total)
        stage_total = max(1e-9, result.stage_total)

        def _unit_final(ur):
            if compressed_months <= 0:
                return ur.adjusted_total
            return ur.adjusted_total - compressed_months * ur.adjusted_total / stage_total

        ov_rows = []
        ov_bg = {0: C_BG_HEAD}
        for ur in sorted(result.unit_results, key=lambda u: (u.scale.scale or 0), reverse=True):
            uf = _unit_final(ur)
            ov_rows.append([ur.spec.name,
                            _fmt(ur.scale.scale),
                            ur.scale.method,
                            f"{_fmt2(ur.adjusted_total)} / {months_to_days_int(ur.adjusted_total)}天",
                            f"{_fmt2(uf)} / {months_to_days_int(uf)}天"])
            ov_bg[len(ov_rows)] = C_BG_GREEN
        _para(doc, "各单体最终周期总览（按规模降序；压缩量按各单体调整后占比分摊）：", size=10, bold=True)
        _table(doc, ["单体", "规模（㎡）", "取值方法", "调整后总周期（月/天）", "最终周期（月/天）"],
               ov_rows, bg_map=ov_bg)

    # ============ 7 特殊说明与风险警告 ============
    _heading(doc, "七、特殊说明与风险警告", 1, "1A1A1A")
    warns = list(result.warnings)
    if result.compress.active and result.compress.warn:
        warns.insert(0, COMPRESS_WARN_TEXT)
    if result.compress.active:
        warns.insert(0, COMPRESS_RISK_TEXT)
    if warns:
        for w in warns:
            p = doc.add_paragraph()
            run = p.add_run("【警告】" + w)
            run.font.name = "微软雅黑"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
            run.font.size = Pt(10)
            run.bold = True
            run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
            pPr = p._p.get_or_add_pPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), C_BG_RED)
            pPr.append(shd)

    _para(doc, REPORT_FOOTER, size=9, color="808080")

    # ============ 附录：不计入周期的 10 类工作 ============
    _heading(doc, "附录：不计入设计周期的 10 类工作（总说明（四））", 2, "808080")
    for i, item in enumerate(EXCLUDED_ITEMS, 1):
        _para(doc, f"{i}. {item}", size=9, color="606060")

    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    # 冒烟测试
    from quota_engine import (ProjectSpec, UnitSpec, AdjustmentSpec,
                              NewTechItem, ReworkItem, calculate)
    p = ProjectSpec(project_name="某商业综合体项目", location="广东省深圳市", designer="示例设计院",
                    start_date="2026-11-02",
                    units=[UnitSpec(name="办公楼", code="01-1", scale=12000),
                           UnitSpec(name="商业裙房", code="01-2", scale=5000, suspended=True),
                           UnitSpec(name="酒店", code="01-5-2", scale=15000)],
                    adjustment=AdjustmentSpec(reuse="1a", reuse_scheme_mode="normal",
                                              reuse_prelim_mode="ratio", scenic=True,
                                              separate_entrust=True,
                                              new_tech_items=[NewTechItem(name="BIM正向设计", desc="全过程应用BIM", months=2.0, poolable=True, pool_group=1),
                                                              NewTechItem(name="装配式深化", desc="PC深化", months=1.5, poolable=True, pool_group=1)],
                                              compress_mode="target", compress_value=8.0))
    r = calculate(p)
    out = generate_report(r, r"C:\Users\2013083\Desktop\design-cycle-quota\temp\sample_report.docx")
    print("OK", out)
