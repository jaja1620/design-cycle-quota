# -*- coding: utf-8 -*-
"""Excel 导出（可选）：输出带数据关联计算的 xlsx，供客户调整后联动刷新。

工作簿结构：
  Sheet1 项目信息（含时间单位、精度规则等说明）
  Sheet2 单体周期（规模取值 + 阶段周期，含公式 = 合计）
  Sheet3 调整明细（规则原文 + 数值）
  Sheet4 最终周期换算（月 / 周 / 天，公式联动）
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from quota_data import DAYS_PER_MONTH, DAYS_PER_WEEK
from quota_engine import CalculationResult

HEAD_FILL = PatternFill("solid", fgColor="D9D9D9")
ACCENT_FILL = PatternFill("solid", fgColor="FFF2CC")
BLUE_FILL = PatternFill("solid", fgColor="DEEBF7")
GREEN_FILL = PatternFill("solid", fgColor="E2EFDA")
ORANGE_FILL = PatternFill("solid", fgColor="FDE9D9")
RED_FILL = PatternFill("solid", fgColor="FCE4E4")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="1A1A1A")
BODY_FONT = Font(name="微软雅黑", size=10, color="1A1A1A")


def _style_sheet(ws, headers):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22


def _apply_border(ws, max_row, max_col):
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = BORDER


def export_excel(result: CalculationResult, output_path: str) -> str:
    """生成 xlsx 并返回路径。"""
    proj = result.project
    wb = Workbook()

    # ---------- Sheet1 项目信息 ----------
    ws1 = wb.active
    ws1.title = "项目信息"
    stage_mode_label = "三阶段" if proj.stage_mode == "three" else "两阶段（初步并入施工图）"
    info = [
        ("项目名称", proj.project_name or ""),
        ("项目地点", proj.location or ""),
        ("设计单位", proj.designer or ""),
        ("委托阶段模式", stage_mode_label),
        ("计划开始时间", proj.start_date or ""),
        ("精度规则", {"round": "四舍五入", "ceil": "向上进位", "floor": "向下取整"}.get(proj.rounding_mode, proj.rounding_mode)),
        ("时间单位", "1 日历月 = 4 日历周 = 28 天"),
        ("多单体折减", {"none": "不折减", "days": f"减少 {proj.multi_unit_value:g} 天",
                          "ratio": f"按 {proj.multi_unit_value * 100:.0f}% 折减"}.get(proj.multi_unit_reduction, "不折减")),
    ]
    for i, (k, v) in enumerate(info, 1):
        ws1.cell(row=i, column=1, value=k).font = HEAD_FONT
        ws1.cell(row=i, column=1).fill = HEAD_FILL
        ws1.cell(row=i, column=1).border = BORDER
        vcell = ws1.cell(row=i, column=2, value=v)
        vcell.font = BODY_FONT
        vcell.border = BORDER
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 60

    # ---------- Sheet2 单体周期 ----------
    ws2 = wb.create_sheet("单体周期")
    headers2 = ["单体名称", "定额编号", "定额子项", "规模（㎡）", "取值方法",
                "方案（月）", "方案（天）", "初步（月）", "初步（天）", "施工图（月）", "施工图（天）",
                "合计（月）", "合计（天）", "状态"]
    _style_sheet(ws2, headers2)
    for i, ur in enumerate(result.unit_results, 2):
        sr = ur.scale
        from quota_data import QUOTA_TABLES
        table = QUOTA_TABLES.get(sr.code)
        row = [ur.spec.name, sr.code, table["name"] if table else sr.code, sr.scale, sr.method,
               sr.scheme, None, sr.preliminary, None, sr.construction, None, None, None,
               "挂起" if ur.spec.suspended else "参与计算"]
        for c, v in enumerate(row, 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.font = BODY_FONT
        # 天列：=ROUND(月*28, 0)，天数保留整数
        ws2.cell(row=i, column=7, value=f"=ROUND(F{i}*{DAYS_PER_MONTH},0)")
        ws2.cell(row=i, column=9, value=f"=ROUND(H{i}*{DAYS_PER_MONTH},0)")
        ws2.cell(row=i, column=11, value=f"=ROUND(J{i}*{DAYS_PER_MONTH},0)")
        # 合计（月）= 方案 + 初步 + 施工图（公式联动，客户可改阶段值自动重算）
        ws2.cell(row=i, column=12, value=f"=F{i}+H{i}+J{i}").font = Font(name="微软雅黑", size=10, bold=True, color="1A1A1A")
        ws2.cell(row=i, column=13, value=f"=ROUND(L{i}*{DAYS_PER_MONTH},0)")
        if ur.spec.suspended:
            for c in range(1, 15):
                ws2.cell(row=i, column=c).fill = PatternFill("solid", fgColor="EFEFEF")
        else:
            for c in range(1, 15):
                ws2.cell(row=i, column=c).fill = GREEN_FILL
    total_row = len(result.unit_results) + 2
    for c, v in enumerate(["项目合计", "", "", "", "", result.base_scheme, None,
                           result.base_preliminary, None, result.base_construction, None,
                           None, None, ""], 1):
        cell = ws2.cell(row=total_row, column=c, value=v)
        cell.border = BORDER
        cell.font = HEAD_FONT
        cell.fill = GREEN_FILL
    ws2.cell(row=total_row, column=7, value=f"=ROUND(F{total_row}*{DAYS_PER_MONTH},0)")
    ws2.cell(row=total_row, column=9, value=f"=ROUND(H{total_row}*{DAYS_PER_MONTH},0)")
    ws2.cell(row=total_row, column=11, value=f"=ROUND(J{total_row}*{DAYS_PER_MONTH},0)")
    ws2.cell(row=total_row, column=12, value=f"=SUM(L2:L{total_row - 1})")
    ws2.cell(row=total_row, column=13, value=f"=ROUND(L{total_row}*{DAYS_PER_MONTH},0)")
    widths2 = [16, 10, 24, 12, 14, 10, 10, 10, 10, 12, 10, 10, 10, 10]
    for c, w in enumerate(widths2, 1):
        ws2.column_dimensions[get_column_letter(c)].width = w

    # ---------- Sheet3 调整明细 ----------
    ws3 = wb.create_sheet("调整明细")
    headers3 = ["规则", "名称", "作用对象", "调整前（月）", "增量（月）", "调整后（月）", "说明"]
    _style_sheet(ws3, headers3)
    for i, ad in enumerate(result.adjustments, 2):
        if abs(ad.base_value) < 1e-9 and abs(ad.delta) < 1e-9:
            continue
        row = [ad.key, ad.label, ad.applied_on, ad.base_value, ad.delta, ad.final_value, ad.note]
        for c, v in enumerate(row, 1):
            cell = ws3.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.font = BODY_FONT
            cell.fill = ORANGE_FILL
    adj_keys = {ad.key for ad in result.adjustments}
    for i, it in enumerate(getattr(result, "additive_items", []), len(result.adjustments) + 2):
        if it.key in adj_keys:
            continue
        row = [it.key, it.label, "", "", it.delta_months, "", it.rule_text]
        for c, v in enumerate(row, 1):
            cell = ws3.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.font = BODY_FONT
            cell.fill = ACCENT_FILL
    widths3 = [8, 26, 24, 12, 12, 12, 46]
    for c, w in enumerate(widths3, 1):
        ws3.column_dimensions[get_column_letter(c)].width = w

    # ---------- Sheet4 最终周期 ----------
    ws4 = wb.create_sheet("最终周期")
    headers4 = ["项目", "月", "周（×4）", "天（×28，取整）"]
    _style_sheet(ws4, headers4)
    rows4 = [
        ("调整后总周期（压缩前）", result.regular_total, None, None),
        ("最终设计周期", result.final_total, None, None),
    ]
    for i, (name, months, _, _) in enumerate(rows4, 2):
        ws4.cell(row=i, column=1, value=name).font = HEAD_FONT
        ws4.cell(row=i, column=1).fill = GREEN_FILL
        ws4.cell(row=i, column=1).border = BORDER
        mcell = ws4.cell(row=i, column=2, value=months)
        mcell.border = BORDER
        mcell.font = BODY_FONT
        mcell.fill = GREEN_FILL
        # 周/天公式联动（天保留整数）
        wcell = ws4.cell(row=i, column=3, value=f"=B{i}*{DAYS_PER_WEEK}")
        wcell.border = BORDER
        wcell.font = BODY_FONT
        wcell.fill = GREEN_FILL
        dcell = ws4.cell(row=i, column=4, value=f"=ROUND(B{i}*{DAYS_PER_MONTH},0)")
        dcell.border = BORDER
        dcell.font = BODY_FONT
        dcell.fill = GREEN_FILL
    if result.compress.active:
        ws4.cell(row=len(rows4) + 2, column=1, value="压缩后周期占定额周期比例").font = HEAD_FONT
        ws4.cell(row=len(rows4) + 2, column=1).border = BORDER
        ws4.cell(row=len(rows4) + 2, column=2, value=result.compress.ratio / 100).number_format = "0.0%"
        ws4.cell(row=len(rows4) + 2, column=2).border = BORDER
        ws4.cell(row=len(rows4) + 2, column=2).fill = RED_FILL if result.compress.warn else GREEN_FILL
        ws4.cell(row=len(rows4) + 2, column=2).font = BODY_FONT
    for c, w in enumerate([26, 10, 12, 14], 1):
        ws4.column_dimensions[get_column_letter(c)].width = w

    # ---------- Sheet5 分阶段统计（调整后单体 × 阶段） ----------
    ws5 = wb.create_sheet("分阶段统计")
    headers5 = ["单体", "方案（月）", "方案（天）", "初步（月）", "初步（天）",
                "施工图（月）", "施工图（天）", "阶段小计（月）", "阶段小计（天）"]
    _style_sheet(ws5, headers5)
    for i, ur in enumerate(result.unit_results, 2):
        row = [ur.spec.name, ur.adjusted_scheme, None, ur.adjusted_preliminary, None,
               ur.adjusted_construction, None, ur.adjusted_total, None]
        for c, v in enumerate(row, 1):
            cell = ws5.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.font = BODY_FONT
            cell.fill = GREEN_FILL
        ws5.cell(row=i, column=3, value=f"=ROUND(B{i}*{DAYS_PER_MONTH},0)")
        ws5.cell(row=i, column=5, value=f"=ROUND(D{i}*{DAYS_PER_MONTH},0)")
        ws5.cell(row=i, column=7, value=f"=ROUND(F{i}*{DAYS_PER_MONTH},0)")
        ws5.cell(row=i, column=9, value=f"=ROUND(H{i}*{DAYS_PER_MONTH},0)")
    total_row5 = len(result.unit_results) + 2
    for c, v in enumerate(["项目合计（调整后）", result.stage_scheme, None, result.stage_preliminary,
                           None, result.stage_construction, None, result.stage_total, None], 1):
        cell = ws5.cell(row=total_row5, column=c, value=v)
        cell.border = BORDER
        cell.font = HEAD_FONT
        cell.fill = GREEN_FILL
    ws5.cell(row=total_row5, column=3, value=f"=ROUND(B{total_row5}*{DAYS_PER_MONTH},0)")
    ws5.cell(row=total_row5, column=5, value=f"=ROUND(D{total_row5}*{DAYS_PER_MONTH},0)")
    ws5.cell(row=total_row5, column=7, value=f"=ROUND(F{total_row5}*{DAYS_PER_MONTH},0)")
    ws5.cell(row=total_row5, column=9, value=f"=ROUND(H{total_row5}*{DAYS_PER_MONTH},0)")
    ws5.cell(row=total_row5 + 1, column=1,
             value="说明：阶段调整作用于阶段总值，各单体调整后值按基础占比分摊并末位校正；天 = 月 × 28，保留整数。").font = BODY_FONT
    for c, w in enumerate([18, 10, 10, 10, 10, 12, 10, 12, 12], 1):
        ws5.column_dimensions[get_column_letter(c)].width = w

    # ---------- Sheet6 输入参数（v1.2 第 16 项：完整导出 GUI 参数，便于存档复现） ----------
    ws6 = wb.create_sheet("输入参数")
    ad = proj.adjustment
    param_rows = []  # (分类, 参数名, 值)

    def _add(cat, k, v):
        param_rows.append((cat, k, "" if v is None else str(v)))

    _add("项目信息", "项目名称", proj.project_name)
    _add("项目信息", "项目地点", proj.location)
    _add("项目信息", "设计单位", proj.designer)
    _add("项目信息", "计划开始时间", proj.start_date)
    _add("项目信息", "委托阶段模式", proj.stage_mode)
    _add("项目信息", "精度规则", proj.rounding_mode)
    _add("项目信息", "保留小数位", proj.digits)
    _add("项目信息", "多单体折减方式", proj.multi_unit_reduction)
    _add("项目信息", "多单体折减数值", proj.multi_unit_value)
    for i, u in enumerate(proj.units, 1):
        _add("单体列表", f"单体{i} 名称", u.name)
        _add("单体列表", f"单体{i} 定额编号", u.code)
        _add("单体列表", f"单体{i} 规模（㎡）", u.scale)
        _add("单体列表", f"单体{i} 挂起", u.suspended)
        _add("单体列表", f"单体{i} 200床以下(03-4)", u.under_200_beds)
        _add("单体列表", f"单体{i} 自定义周期", u.custom_periods if u.custom_periods else "")
        _add("单体列表", f"单体{i} 自定义说明", u.custom_note)
    _add("调整因素", "复用套图 reuse", ad.reuse)
    _add("调整因素", "方案阶段复用方式", ad.reuse_scheme_mode)
    _add("调整因素", "初步阶段复用方式", ad.reuse_prelim_mode)
    _add("调整因素", "自定义复用比例", ad.reuse_ratio)
    _add("调整因素", "改扩建 renovation", ad.renovation)
    _add("调整因素", "改扩建自定义比例", ad.renovation_custom)
    _add("调整因素", "分阶段分别委托", ad.separate_entrust)
    _add("调整因素", "多个设计单位共同设计", ad.co_design)
    _add("调整因素", "五级及以上附建式人防", ad.civil_air_defense)
    _add("调整因素", "建于风景区的宾馆", ad.scenic)
    _add("调整因素", "特殊建筑增加（月）", ad.special_months)
    _add("调整因素", "涉外工程增加（月）", ad.foreign_months)
    _add("调整因素", "工作量变化（月）", ad.work_change_months)
    for i, it in enumerate(ad.new_tech_items, 1):
        _add("新技术/返工", f"新技术{i} 名称", it.name)
        _add("新技术/返工", f"新技术{i} 说明", it.desc)
        _add("新技术/返工", f"新技术{i} 月数", it.months)
        _add("新技术/返工", f"新技术{i} 挂起", it.suspended)
        _add("新技术/返工", f"新技术{i} 可统筹", it.poolable)
        _add("新技术/返工", f"新技术{i} 统筹组", it.pool_group)
    for i, rw in enumerate(ad.rework_items, 1):
        _add("新技术/返工", f"返工{i} 阶段", rw.stage)
        _add("新技术/返工", f"返工{i} 月数", rw.months)
        _add("新技术/返工", f"返工{i} 说明", rw.note)
    _add("节假日", "节假日模式", ad.holiday_mode)
    _add("节假日", "手动输入天数", ad.holiday_input_days)
    _add("节假日", "覆写天数", ad.holiday_override_days)
    _add("节假日", "节假日注释", ad.holiday_note)
    _add("压缩", "压缩模式", ad.compress_mode)
    _add("压缩", "压缩数值", ad.compress_value)
    if ad.pool_groups:
        for k, g in ad.pool_groups.items():
            _add("分组统筹", f"规则 {k} 组号", g)
    headers6 = ["分类", "参数名", "值"]
    for c, h in enumerate(headers6, 1):
        cell = ws6.cell(row=1, column=c, value=h)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, (cat, k, v) in enumerate(param_rows, 2):
        for c, val in enumerate((cat, k, v), 1):
            cell = ws6.cell(row=i, column=c, value=val)
            cell.border = BORDER
            cell.font = BODY_FONT
        ws6.cell(row=i, column=1).fill = ACCENT_FILL
    for c, w in enumerate([14, 30, 60], 1):
        ws6.column_dimensions[get_column_letter(c)].width = w

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    from quota_engine import (ProjectSpec, UnitSpec, AdjustmentSpec,
                              NewTechItem, calculate)
    p = ProjectSpec(project_name="某商业综合体项目", units=[UnitSpec(name="办公楼", code="01-1", scale=12000)],
                    adjustment=AdjustmentSpec(separate_entrust=True, co_design=True,
                                              new_tech_items=[NewTechItem(name="BIM", months=2.0)]))
    r = calculate(p)
    out = export_excel(r, r"C:\Users\2013083\Desktop\design-cycle-quota\temp\sample_export.xlsx")
    print("OK", out)
