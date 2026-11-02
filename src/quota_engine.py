# -*- coding: utf-8 -*-
"""计算引擎：按《全国建筑设计周期定额（2016版）》+ 15 条决策规则计算定额周期。

流水线：
  Step1 定额表定位 → Step2 规模取值（档位/插值/低于起档/高于最高档等比推算/未覆盖区间）
  Step3 精度处理（四舍五入/进位/保留位数，默认四舍五入保留2位）
  Step4 两阶段合并（施工图 = 初步表值 + 施工图表值，Q5）
  Step5 多单体合并（各单体周期相加后折减，默认不折减，Q6）
  Step6 调整引擎（14 条规则，默认全叠加，可分组统筹，Q15）
  Step7 压缩周期（Q10）与汇总输出
"""
from __future__ import annotations

import math
from datetime import date, timedelta
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP, ROUND_CEILING, ROUND_FLOOR

from quota_data import (
    QUOTA_TABLES, RULES_TEXT, RENOVATION_LEVELS,
    DAYS_PER_MONTH, NOTE_ABOVE_MAX, NOTE_BELOW_COVERED,
    NOTE_USER_CUSTOM, NOTE_INTERPOLATION, NOTE_AT_LEVEL, NOTE_BELOW_MIN,
    COMPRESS_RISK_TEXT, COMPRESS_WARN_TEXT,
)
from holidays import holiday_summary, holiday_detail


# ---------------------------------------------------------------------------
# 输入数据结构
# ---------------------------------------------------------------------------
@dataclass
class UnitSpec:
    """单体子项（Q6 多单体支持增删/挂起）。"""
    name: str = "单体1"
    code: str = "01-1"                 # 定额编号
    scale: float = 0.0                 # 总建筑面积（㎡，含地下地上，Q12）
    suspended: bool = False            # 挂起：保留数据不参与计算
    under_200_beds: bool = False       # 03-4 疗养/养老院 200 床以下（Q13）
    custom_periods: dict = None        # Q1 用户自定义 {scheme, preliminary, construction}（月）
    custom_note: str = ""              # 自定义说明


@dataclass
class NewTechItem:
    """新技术/特殊要求项（Q9：按项输入，可增删/挂起/统筹）。"""
    name: str = ""
    desc: str = ""
    months: float = 0.0                # 需要增加的时间（月）
    suspended: bool = False
    poolable: bool = False             # 是否周期可统筹
    pool_group: int = 0                # 统筹组号（0=不可统筹）


@dataclass
class ReworkItem:
    """返工/修改（规则11：按返工量增加对应阶段周期）。"""
    stage: str = "construction"        # scheme / preliminary / construction
    months: float = 0.0
    note: str = ""


@dataclass
class AdjustmentSpec:
    """调整因素集合（14 条规则 + 统筹分组）。"""
    reuse: str = "none"                # none / 1a（全部复用）/ 1b（上部复用）
    reuse_scheme_mode: str = "normal"  # normal / cancel / ratio（Q3）
    reuse_prelim_mode: str = "ratio"   # normal / cancel / ratio（Q3）
    reuse_ratio: float = 0.0           # 自定义比例（0=使用规则默认 10%/20%）
    renovation: str = "none"           # none / simple / general / complex / special / custom（Q4）
    renovation_custom: float = 0.0     # 自定义 10%~40%
    separate_entrust: bool = False     # 规则3 分阶段分别委托
    co_design: bool = False            # 规则4 多个设计单位共同设计
    civil_air_defense: bool = False    # 规则5 五级及以上附建式人防
    scenic: bool = False               # 规则7 风景区宾馆（Q8：适用所有酒店类）
    special_months: float = 0.0        # 规则8 特殊建筑（双方商议）
    new_tech_items: list = field(default_factory=list)  # 规则9
    foreign_months: float = 0.0        # 规则10 涉外工程（默认不增加）
    rework_items: list = field(default_factory=list)    # 规则11 返工
    holiday_mode: str = "none"         # none / auto / input（Q7）
    holiday_input_days: float = 0.0    # 用户输入天数
    holiday_override_days: float = 0.0 # 覆写天数（Q7）
    holiday_note: str = ""             # 节假日注释
    compress_mode: str = "none"        # none / target（压缩后周期）/ days（压缩天数）（Q10）
    compress_value: float = 0.0
    work_change_months: float = 0.0    # 规则14 工作量变化（双方协商）
    pool_groups: dict = field(default_factory=dict)  # 规则key -> 组号（Q15）


@dataclass
class ProjectSpec:
    """项目级输入。"""
    project_name: str = ""
    location: str = ""
    designer: str = ""
    start_date: str = ""               # YYYY-MM-DD 各阶段预测开始时间
    stage_mode: str = "three"          # three / two（两阶段=初步并入施工图，Q5/Q14）
    units: list = field(default_factory=list)
    multi_unit_reduction: str = "none" # none / days / ratio（Q6 折减，默认不折减）
    multi_unit_value: float = 0.0      # 折减天数 或 折减比例（0~1）
    rounding_mode: str = "round"       # round / ceil / floor（Q11）
    digits: int = 2
    adjustment: AdjustmentSpec = field(default_factory=AdjustmentSpec)


# ---------------------------------------------------------------------------
# 结果数据结构
# ---------------------------------------------------------------------------
@dataclass
class ScaleResult:
    """单个单体的规模取值结果。"""
    code: str
    name: str
    scale: float
    method: str                       # at_level / interpolation / below_min / above_max / below_covered / custom
    note: str
    scheme: float
    preliminary: float
    construction: float
    total: float


@dataclass
class AdditiveItem:
    """总周期附加项（用于统筹）。"""
    key: str
    label: str
    rule_text: str
    delta_months: float
    pool_group: int = 0               # 0 = 不可统筹


@dataclass
class UnitResult:
    spec: UnitSpec
    scale: ScaleResult
    stage_mode: str
    scheme: float
    preliminary: float
    construction: float
    total: float
    # v1.1：调整后阶段值（按各单体基础占比分摊阶段调整，供「单体×阶段」统计）
    adjusted_scheme: float = 0.0
    adjusted_preliminary: float = 0.0
    adjusted_construction: float = 0.0
    adjusted_total: float = 0.0


@dataclass
class AdjustmentDetail:
    """单个调整明细（报告）。"""
    key: str
    label: str
    rule_text: str
    applied_on: str                   # 作用对象描述
    base_value: float
    delta: float                      # 周期增量（月）
    final_value: float
    note: str = ""
    is_pool_group: bool = False
    pool_group: int = 0


@dataclass
class CompressResult:
    active: bool
    regular_total: float
    final_total: float
    compressed_days: float
    ratio: float                      # 压缩后/定额周期 百分比
    warn: bool
    risk_text: str


@dataclass
class StageDate:
    """单个阶段起止日期（1 日历月 = 28 天推算，v1.2）。"""
    stage: str                        # scheme / preliminary / construction
    stage_name: str                   # 方案设计阶段 / 初步设计阶段 / 施工图设计阶段
    start: str                        # YYYY-MM-DD
    end: str                          # YYYY-MM-DD（含当日）
    days: int                         # 周期天数（月→天取整，至少 1 天）
    months: float                     # 周期（月）


@dataclass
class CalculationResult:
    project: ProjectSpec
    unit_results: list                 # [UnitResult]
    base_scheme: float
    base_preliminary: float
    base_construction: float
    base_total: float
    stage_scheme: float
    stage_preliminary: float
    stage_construction: float
    stage_total: float
    additive_items: list              # [AdditiveItem]
    adjustments: list                 # [AdjustmentDetail]（含阶段调整与附加明细）
    regular_total: float              # 调整后（压缩前）总周期
    final_total: float                # 最终总周期
    holiday_count: float
    holiday_override: float
    compress: CompressResult
    warnings: list
    # v1.2：各阶段起止日期推算（9）
    stage_dates: list = field(default_factory=list)
    has_start_date: bool = False
    # v1.2：计算链等式（10）
    equation_parts: list = field(default_factory=list)
    equation_text: str = ""
    # v1.2：节假日明细（11）
    holiday_details: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# 精度工具
# ---------------------------------------------------------------------------
def round_value(v, mode="round", digits=2):
    if mode == "ceil":
        return float(Decimal(str(v)).quantize(Decimal("1").scaleb(-digits), rounding=ROUND_CEILING))
    if mode == "floor":
        return float(Decimal(str(v)).quantize(Decimal("1").scaleb(-digits), rounding=ROUND_FLOOR))
    return float(Decimal(str(v)).quantize(Decimal("1").scaleb(-digits), rounding=ROUND_HALF_UP))


# ---------------------------------------------------------------------------
# Step2 规模取值
# ---------------------------------------------------------------------------
def scale_periods(table, scale, under_200_beds=False):
    """按规模取值，返回 ScaleResult。"""
    scales = table["scales"]
    ts = table["scheme"], table["preliminary"], table["construction"]
    name = table["name"]
    code = _code_of_table(table)

    # 未覆盖区间（Q13）：疗养/养老院 200 床以下，按 10000㎡ 档向下等比推算
    if under_200_beds:
        base_scale = 10000.0
        if scale < base_scale:
            r = scale / base_scale
            # ts[i][0] 是该表 10000 ㎡ 档（起档）值
            return ScaleResult(code, name, scale, "below_covered", NOTE_BELOW_COVERED,
                               round_value(ts[0][0] * r), round_value(ts[1][0] * r),
                               round_value(ts[2][0] * r), 0.0)
        # 200 床以上但小于 10000 ㎡：走常规路径
        pass

    # 低于起档
    if scale <= scales[0]:
        s0 = scales[0]
        if scale == s0:
            return ScaleResult(code, name, scale, "at_level", NOTE_AT_LEVEL,
                               float(ts[0][0]), float(ts[1][0]), float(ts[2][0]), 0.0)
        return ScaleResult(code, name, scale, "below_min", NOTE_BELOW_MIN,
                           float(ts[0][0]), float(ts[1][0]), float(ts[2][0]), 0.0)

    # 高于最高档（Q1 等比推算）
    if scale > scales[-1]:
        smax = scales[-1]
        r = scale / smax
        return ScaleResult(code, name, scale, "above_max", NOTE_ABOVE_MAX,
                           round_value(ts[0][-1] * r), round_value(ts[1][-1] * r),
                           round_value(ts[2][-1] * r), 0.0)

    # 档位命中
    for i, s in enumerate(scales):
        if scale == s:
            return ScaleResult(code, name, scale, "at_level", NOTE_AT_LEVEL,
                               float(ts[0][i]), float(ts[1][i]), float(ts[2][i]), 0.0)

    # 两档之间：线性插值（总说明（六））
    for i in range(len(scales) - 1):
        if scales[i] < scale < scales[i + 1]:
            s1, s2 = scales[i], scales[i + 1]
            t = (scale - s1) / (s2 - s1)
            return ScaleResult(code, name, scale, "interpolation", NOTE_INTERPOLATION,
                               round_value(ts[0][i] + (ts[0][i + 1] - ts[0][i]) * t),
                               round_value(ts[1][i] + (ts[1][i + 1] - ts[1][i]) * t),
                               round_value(ts[2][i] + (ts[2][i + 1] - ts[2][i]) * t), 0.0)

    raise ValueError(f"规模取值失败：{code} scale={scale}")


def _code_of_table(table):
    for c, t in QUOTA_TABLES.items():
        if t is table:
            return c
    return "?"


def scale_total(sr):
    return round_value(sr.scheme + sr.preliminary + sr.construction)


# ---------------------------------------------------------------------------
# 单体计算
# ---------------------------------------------------------------------------
def compute_unit(spec: UnitSpec, stage_mode: str, project: ProjectSpec):
    """计算单个单体（未挂起）的基础周期。"""
    table = QUOTA_TABLES[spec.code]
    # 用户自定义周期（Q1 第二选项）
    if spec.custom_periods and any(spec.custom_periods.get(k) is not None for k in ("scheme", "preliminary", "construction")):
        sr = ScaleResult(spec.code, table["name"], spec.scale, "custom", NOTE_USER_CUSTOM,
                         float(spec.custom_periods.get("scheme") or 0),
                         float(spec.custom_periods.get("preliminary") or 0),
                         float(spec.custom_periods.get("construction") or 0), 0.0)
    else:
        sr = scale_periods(table, spec.scale, spec.under_200_beds)
    sr.total = scale_total(sr)

    scheme, preliminary, construction = sr.scheme, sr.preliminary, sr.construction
    # 两阶段模式（Q5）：施工图 = 初步表值 + 施工图表值
    if stage_mode == "two":
        construction = round_value(preliminary + construction, project.rounding_mode, project.digits)
        preliminary = 0.0

    total = round_value(scheme + preliminary + construction, project.rounding_mode, project.digits)
    return UnitResult(spec, sr, stage_mode, scheme, preliminary, construction, total)


# ---------------------------------------------------------------------------
# 调整引擎
# ---------------------------------------------------------------------------
def _stage_coefficient(adj: AdjustmentSpec, stage: str, base_value: float, project: ProjectSpec):
    """计算阶段级调整后的阶段值（多条规则累计应用，返回 (new_value, details)）。"""
    value = base_value
    details = []

    # 规则1a/1b 复用套图：作用于施工图
    if adj.reuse in ("1a", "1b") and stage == "construction":
        ratio = adj.reuse_ratio if adj.reuse_ratio > 0 else (0.10 if adj.reuse == "1a" else 0.20)
        key = "1a" if adj.reuse == "1a" else "1b"
        mode_label = "全部复用" if key == "1a" else "上部复用（基础重新设计）"
        delta = round_value(value * ratio, project.rounding_mode, project.digits)
        value = round_value(value + delta, project.rounding_mode, project.digits)
        details.append(AdjustmentDetail(
            key=key, label=f"复用套图-{mode_label}", rule_text=RULES_TEXT[key],
            applied_on="施工图设计阶段", base_value=value - delta,
            delta=delta, final_value=value,
            note=f"施工图周期 ×{ratio*100:.0f}%"))

    # Q3：方案/初步阶段的复用处理（照常/取消/按比例）
    if adj.reuse in ("1a", "1b"):
        key = "1a" if adj.reuse == "1a" else "1b"
        mode_label = "全部复用" if key == "1a" else "上部复用（基础重新设计）"
        ratio = adj.reuse_ratio if adj.reuse_ratio > 0 else (0.10 if key == "1a" else 0.20)
        if stage == "scheme":
            if adj.reuse_scheme_mode == "cancel":
                details.append(AdjustmentDetail(key=key, label=f"复用套图-{mode_label}（方案）", rule_text="Q3 决策",
                    applied_on="方案设计阶段", base_value=value, delta=-value, final_value=0.0,
                    note="方案阶段取消"))
                return 0.0, details
            if adj.reuse_scheme_mode == "ratio":
                delta = round_value(value * ratio, project.rounding_mode, project.digits)
                value = round_value(value + delta, project.rounding_mode, project.digits)
                details.append(AdjustmentDetail(key=key, label=f"复用套图-{mode_label}（方案）", rule_text="Q3 决策",
                    applied_on="方案设计阶段", base_value=value - delta, delta=delta, final_value=value,
                    note=f"方案阶段按比例 ×{ratio*100:.0f}%"))
        if stage == "preliminary":
            if adj.reuse_prelim_mode == "cancel":
                details.append(AdjustmentDetail(key=key, label=f"复用套图-{mode_label}（初步）", rule_text="Q3 决策",
                    applied_on="初步设计阶段", base_value=value, delta=-value, final_value=0.0,
                    note="初步阶段取消"))
                return 0.0, details
            if adj.reuse_prelim_mode == "ratio":
                delta = round_value(value * ratio, project.rounding_mode, project.digits)
                value = round_value(value + delta, project.rounding_mode, project.digits)
                details.append(AdjustmentDetail(key=key, label=f"复用套图-{mode_label}（初步）", rule_text="Q3 决策",
                    applied_on="初步设计阶段", base_value=value - delta, delta=delta, final_value=value,
                    note=f"初步阶段按比例 ×{ratio*100:.0f}%"))

    # 规则3 分阶段分别委托：初步、施工图各 +10%
    if adj.separate_entrust and stage in ("preliminary", "construction"):
        ratio = 0.10
        delta = round_value(value * ratio, project.rounding_mode, project.digits)
        details.append(AdjustmentDetail(
            key="3", label="分阶段分别委托设计", rule_text=RULES_TEXT["3"],
            applied_on="初步设计、施工图设计阶段", base_value=value,
            delta=delta, final_value=round_value(value + delta, project.rounding_mode, project.digits),
            note=f"×{ratio*100:.0f}%"))
        value = round_value(value + delta, project.rounding_mode, project.digits)

    # 规则5 五级及以上附建式人防：初步、施工图各 +15%
    if adj.civil_air_defense and stage in ("preliminary", "construction"):
        ratio = 0.15
        delta = round_value(value * ratio, project.rounding_mode, project.digits)
        details.append(AdjustmentDetail(
            key="5", label="五级及以上附建式人防设计", rule_text=RULES_TEXT["5"],
            applied_on="初步设计、施工图设计阶段", base_value=value,
            delta=delta, final_value=round_value(value + delta, project.rounding_mode, project.digits),
            note=f"×{ratio*100:.0f}%"))
        value = round_value(value + delta, project.rounding_mode, project.digits)

    # 规则11 返工：对应阶段增加
    for rw in adj.rework_items:
        if rw.stage == stage and rw.months > 0:
            delta = round_value(rw.months, project.rounding_mode, project.digits)
            details.append(AdjustmentDetail(
                key="11", label="建设方原因一般性返工、修改", rule_text=RULES_TEXT["11"],
                applied_on=f"{stage}设计阶段", base_value=value, delta=delta,
                final_value=round_value(value + delta, project.rounding_mode, project.digits),
                note=f"返工增加 {rw.months:g} 月" + (f"（{rw.note}）" if rw.note else "")))
            value = round_value(value + delta, project.rounding_mode, project.digits)

    return value, details


def _build_additive_items(project: ProjectSpec, stage_total: float, holiday_count: float):
    """构建总周期附加项（含统筹信息）。返回 (items, extra_details)。"""
    adj = project.adjustment
    items = []
    extra = []
    pg = adj.pool_groups or {}

    # 规则2 改扩建
    if adj.renovation in RENOVATION_LEVELS and adj.renovation != "none":
        level_name, ratio = RENOVATION_LEVELS[adj.renovation]
        if ratio is None:  # custom
            ratio = adj.renovation_custom / 100.0 if adj.renovation_custom else 0.0
        delta = round_value(stage_total * ratio, project.rounding_mode, project.digits)
        items.append(AdditiveItem("2", f"改建、扩建工程（{level_name}）", RULES_TEXT["2"], delta, pg.get("2", 0)))

    # 规则4 共同设计
    if adj.co_design:
        delta = round_value(stage_total * 0.10, project.rounding_mode, project.digits)
        items.append(AdditiveItem("4", "多个设计单位共同设计", RULES_TEXT["4"], delta, pg.get("4", 0)))

    # 规则7 风景区宾馆
    if adj.scenic:
        delta = round_value(stage_total * 0.30, project.rounding_mode, project.digits)
        items.append(AdditiveItem("7", "建于风景区的宾馆", RULES_TEXT["7"], delta, pg.get("7", 0)))

    # 规则8 特殊建筑（双方商议）
    if adj.special_months > 0:
        items.append(AdditiveItem("8", "特殊建筑（古建筑/仿古/保护性/园林庭院式/宗教类）", RULES_TEXT["8"],
                                  round_value(adj.special_months, project.rounding_mode, project.digits), pg.get("8", 0)))

    # 规则9 新技术项（可统筹 / 不可统筹）
    for i, nt in enumerate(adj.new_tech_items):
        if nt.suspended or nt.months <= 0:
            continue
        group = nt.pool_group if (nt.poolable and nt.pool_group > 0) else 0
        items.append(AdditiveItem(f"9-{i}", f"采用先进技术：{nt.name}", RULES_TEXT["9"],
                                  round_value(nt.months, project.rounding_mode, project.digits), group))
        if nt.desc:
            extra.append(f"  新技术项「{nt.name}」概述：{nt.desc}")

    # 规则10 涉外工程
    if adj.foreign_months > 0:
        items.append(AdditiveItem("10", "涉外工程（国外境外/援外/对外总承包）", RULES_TEXT["10"],
                                  round_value(adj.foreign_months, project.rounding_mode, project.digits), pg.get("10", 0)))

    # 规则12 节假日
    if holiday_count > 0:
        delta = round_value(holiday_count / DAYS_PER_MONTH, project.rounding_mode, project.digits)
        items.append(AdditiveItem("12", "国家法定节假日", RULES_TEXT["12"], delta, pg.get("12", 0)))

    # 规则14 工作量变化
    if adj.work_change_months > 0:
        items.append(AdditiveItem("14", "特殊工程工作量较大变化（双方协商）", RULES_TEXT["14"],
                                  round_value(adj.work_change_months, project.rounding_mode, project.digits), pg.get("14", 0)))

    return items, extra


def _merge_additives(items, project: ProjectSpec):
    """Q15 统筹合并：不可统筹项逐一叠加；同组取组内最长。"""
    groups = {}
    standalone = 0.0
    for it in items:
        if it.pool_group and it.pool_group > 0:
            g = groups.setdefault(it.pool_group, 0.0)
            groups[it.pool_group] = max(g, it.delta_months)
        else:
            standalone += it.delta_months
    return round_value(standalone + sum(groups.values()), project.rounding_mode, project.digits), groups


# ---------------------------------------------------------------------------
# 主计算入口
# ---------------------------------------------------------------------------
def calculate(project: ProjectSpec):
    """完整计算流程。"""
    adj = project.adjustment
    rounding = project.rounding_mode
    digits = project.digits

    # Step1-2-3-4: 各单体计算
    unit_results = []
    for spec in project.units:
        if spec.suspended:
            continue
        ur = compute_unit(spec, project.stage_mode, project)
        unit_results.append(ur)

    if not unit_results:
        raise ValueError("没有可参与计算的单体子项（可能全部被挂起）")

    # Step5: 多单体合并（Q6：各单体周期相加后折减，默认不折减）
    base_scheme = round_value(sum(u.scheme for u in unit_results), rounding, digits)
    base_prelim = round_value(sum(u.preliminary for u in unit_results), rounding, digits)
    base_cons = round_value(sum(u.construction for u in unit_results), rounding, digits)
    base_total = round_value(base_scheme + base_prelim + base_cons, rounding, digits)

    # 多单体折减（规则6：各单体周期相加后折减；v1.2 起按占比分摊到各阶段值，
    # 使折减真正传导至阶段调整与最终周期，并保证计算链等式可勾稽）
    reduction_delta = 0.0
    reduction_note = ""
    if len(unit_results) > 1 and project.multi_unit_reduction != "none":
        if project.multi_unit_reduction == "ratio":
            r = min(max(project.multi_unit_value, 0.0), 1.0)
            reduction_delta = round_value(base_total * r, rounding, digits)
            reduction_note = f"多单体折减比例 {r*100:.0f}%，减少 {reduction_delta:g} 月"
        elif project.multi_unit_reduction == "days":
            d = max(project.multi_unit_value, 0.0)
            reduction_delta = round_value(d / DAYS_PER_MONTH, rounding, digits)
            reduction_note = f"多单体折减 {d:g} 天，折合 {reduction_delta:g} 月"

    base_scheme = round_value(base_scheme, rounding, digits)
    base_prelim = round_value(base_prelim, rounding, digits)
    base_cons = round_value(base_cons, rounding, digits)
    if reduction_delta > 0 and base_total > 0:
        base_scheme = round_value(base_scheme - reduction_delta * (base_scheme / base_total), rounding, digits)
        base_prelim = round_value(base_prelim - reduction_delta * (base_prelim / base_total), rounding, digits)
        base_cons = round_value(base_cons - reduction_delta * (base_cons / base_total), rounding, digits)
    base_total = round_value(base_total - reduction_delta, rounding, digits)

    # Step6: 阶段调整
    stage_details = []
    s_scheme, d1 = _stage_coefficient(adj, "scheme", base_scheme, project)
    stage_details.extend(d1)
    s_prelim, d2 = _stage_coefficient(adj, "preliminary", base_prelim, project)
    stage_details.extend(d2)
    s_cons, d3 = _stage_coefficient(adj, "construction", base_cons, project)
    stage_details.extend(d3)
    stage_total = round_value(s_scheme + s_prelim + s_cons, rounding, digits)

    # v1.1：单体×阶段统计（调整后各单体按基础占比分摊，末位校正保证阶段小计=阶段总值）
    for ur in unit_results:
        ur.adjusted_scheme = _alloc_share(ur.scheme, base_scheme, s_scheme, rounding, digits)
        ur.adjusted_preliminary = _alloc_share(ur.preliminary, base_prelim, s_prelim, rounding, digits)
        ur.adjusted_construction = _alloc_share(ur.construction, base_cons, s_cons, rounding, digits)
        ur.adjusted_total = round_value(ur.adjusted_scheme + ur.adjusted_preliminary
                                        + ur.adjusted_construction, rounding, digits)
    _fix_alloc_tail(unit_results, base_scheme, s_scheme, base_prelim, s_prelim,
                    base_cons, s_cons, rounding, digits)

    # 节假日统计（Q7）
    holiday_count = 0.0
    holiday_override = 0.0
    if adj.holiday_mode == "auto":
        holiday_count = float(holiday_summary(project.start_date, stage_total * DAYS_PER_MONTH))
    elif adj.holiday_mode == "input":
        holiday_count = max(adj.holiday_input_days, 0.0)
    holiday_override = max(adj.holiday_override_days, 0.0)
    total_holiday_days = holiday_count + holiday_override

    # 总周期附加项（含统筹）
    items, extra_notes = _build_additive_items(project, stage_total, total_holiday_days)
    additive_total, groups = _merge_additives(items, project)

    regular_total = round_value(stage_total + additive_total, rounding, digits)

    # Step7: 压缩周期（Q10）
    compress = CompressResult(False, regular_total, regular_total, 0.0, 100.0, False, "")
    if adj.compress_mode != "none" and adj.compress_value > 0:
        if adj.compress_mode == "target":
            target = min(adj.compress_value, regular_total)
            final_total = round_value(target, rounding, digits)
        else:  # days
            final_total = round_value(max(regular_total - adj.compress_value / DAYS_PER_MONTH, 0.0), rounding, digits)
        compressed_days = round_value((regular_total - final_total) * DAYS_PER_MONTH, rounding, 1)
        ratio = (final_total / regular_total * 100.0) if regular_total else 0.0
        warn = ratio < 50.0
        risk_text = (COMPRESS_WARN_TEXT + "\n" + COMPRESS_RISK_TEXT) if warn else COMPRESS_RISK_TEXT
        compress = CompressResult(True, regular_total, final_total, compressed_days, ratio, warn, risk_text)
    else:
        final_total = regular_total

    # 汇总明细（阶段调整 + 附加项）
    adjustments = list(stage_details)
    # v1.1.2：附加项明细从 stage_total 起逐项累计（Q15 统筹同组取最长，同组仅保留最大项）
    group_max = {}
    for it in items:
        if it.pool_group > 0:
            prev = group_max.get(it.pool_group)
            if prev is None or it.delta_months > prev.delta_months:
                group_max[it.pool_group] = it
    shown_items = []
    for it in items:
        if it.pool_group > 0:
            if group_max.get(it.pool_group) is it:
                shown_items.append(it)
        else:
            shown_items.append(it)
    running = round_value(stage_total, rounding, digits)
    for it in shown_items:
        grp_txt = f"，统筹组 {it.pool_group}" if it.pool_group > 0 else ""
        base_val = running
        final_val = round_value(running + it.delta_months, rounding, digits)
        adjustments.append(AdjustmentDetail(
            key=it.key, label=it.label, rule_text=it.rule_text, applied_on="总周期",
            base_value=base_val, delta=it.delta_months, final_value=final_val,
            note=f"+{it.delta_months:g} 月{grp_txt}",
            is_pool_group=it.pool_group > 0, pool_group=it.pool_group))
        running = final_val
    # 校验：逐项累加终值必须等于 regular_total（与统筹合计一致）
    if abs(running - regular_total) > 0.5 * 10 ** -digits:
        raise ValueError(f"附加项明细累加校验失败：{running} != {regular_total}")
    if reduction_delta > 0:
        adjustments.insert(0, AdjustmentDetail(
            key="6", label="多单体子项合并折减", rule_text=RULES_TEXT["6"],
            applied_on="总周期", base_value=base_total + reduction_delta, delta=-reduction_delta,
            final_value=base_total, note=reduction_note))

    warnings = []
    if compress.active:
        if compress.warn:
            warnings.append("压缩后周期不足定额周期 50%")
        if compress.final_total < 1.0:
            warnings.append("压缩后周期不足 1 个月（最终周期过小），建议减小压缩天数或改用目标周期模式")

    # v1.2：各阶段起止日期推算（9）
    stage_dates, has_start = build_stage_dates(project, s_scheme, s_prelim, s_cons)
    # v1.2：计算链等式（10）
    equation_parts = build_equation_parts(base_total, adjustments, compress, regular_total, final_total)
    equation_text = format_equation(equation_parts, final_total)
    # v1.2：节假日明细（11，auto 模式且填写开始时间时统计区间内法定节假日）
    holiday_details = []
    if adj.holiday_mode == "auto" and project.start_date:
        holiday_details = holiday_detail(project.start_date, stage_total * DAYS_PER_MONTH)

    return CalculationResult(
        project=project, unit_results=unit_results,
        base_scheme=base_scheme, base_preliminary=base_prelim, base_construction=base_cons, base_total=base_total,
        stage_scheme=s_scheme, stage_preliminary=s_prelim, stage_construction=s_cons, stage_total=stage_total,
        additive_items=items, adjustments=adjustments,
        regular_total=regular_total, final_total=final_total,
        holiday_count=holiday_count, holiday_override=holiday_override,
        compress=compress, warnings=warnings,
        stage_dates=stage_dates, has_start_date=has_start,
        equation_parts=equation_parts, equation_text=equation_text,
        holiday_details=holiday_details)


# ---------------------------------------------------------------------------
# 周期换算工具（报告用）
# ---------------------------------------------------------------------------
def months_to_weeks(months, digits=1):
    return round_value(months * 4.0, "round", digits)


def months_to_days(months, digits=1):
    return round_value(months * DAYS_PER_MONTH, "round", digits)


def months_to_days_int(months):
    """v1.1：月→天（1 月=28 天），天数保留整数（四舍五入）。"""
    return int(round(months * DAYS_PER_MONTH))


def _alloc_share(unit_base, stage_base, stage_adj, rounding, digits):
    """按基础占比将阶段调整后总值分摊到单体（阶段基础值为 0 时结果为 0）。"""
    if stage_base <= 0:
        return 0.0
    return round_value(unit_base / stage_base * stage_adj, rounding, digits)


def _fix_alloc_tail(unit_results, base_scheme, s_scheme, base_prelim, s_prelim,
                    base_cons, s_cons, rounding, digits):
    """末位校正：将分摊四舍五入产生的尾差并入最后一项，保证阶段小计=阶段总值。"""
    if not unit_results:
        return
    s_total = round_value(sum(u.adjusted_scheme for u in unit_results), rounding, digits)
    if abs(s_total - s_scheme) >= 0.5 * 10 ** -digits:
        unit_results[-1].adjusted_scheme = round_value(
            unit_results[-1].adjusted_scheme + round_value(s_scheme - s_total, rounding, digits),
            rounding, digits)
    p_total = round_value(sum(u.adjusted_preliminary for u in unit_results), rounding, digits)
    if abs(p_total - s_prelim) >= 0.5 * 10 ** -digits:
        unit_results[-1].adjusted_preliminary = round_value(
            unit_results[-1].adjusted_preliminary + round_value(s_prelim - p_total, rounding, digits),
            rounding, digits)
    c_total = round_value(sum(u.adjusted_construction for u in unit_results), rounding, digits)
    if abs(c_total - s_cons) >= 0.5 * 10 ** -digits:
        unit_results[-1].adjusted_construction = round_value(
            unit_results[-1].adjusted_construction + round_value(s_cons - c_total, rounding, digits),
            rounding, digits)
    for ur in unit_results:
        ur.adjusted_total = round_value(ur.adjusted_scheme + ur.adjusted_preliminary
                                        + ur.adjusted_construction, rounding, digits)


# ---------------------------------------------------------------------------
# v1.2：各阶段起止日期推算（1 日历月 = 28 天）
# ---------------------------------------------------------------------------
def build_stage_dates(project, stage_scheme, stage_prelim, stage_cons):
    """由计划开始时间与各阶段周期（月，1月=28天）推算阶段起止日期。

    返回 (list[StageDate], has_start_date)；未填写或无法解析开始时间返回 ([], False)。
    三阶段：方案 → 初步 → 施工图；两阶段：初步周期并入施工图，仅方案 + 施工图两段。
    """
    if not project.start_date:
        return [], False
    try:
        start = date.fromisoformat(project.start_date)
    except (ValueError, TypeError):
        return [], False
    stages = [("scheme", "方案设计阶段", stage_scheme),
              ("preliminary", "初步设计阶段", stage_prelim),
              ("construction", "施工图设计阶段", stage_cons)]
    out = []
    cur = start
    for key, name, months in stages:
        if months <= 0:
            continue
        days = max(int(round(months * DAYS_PER_MONTH)), 1)
        end = cur + timedelta(days=days - 1)
        out.append(StageDate(key, name, cur.isoformat(), end.isoformat(), days,
                             round_value(months, project.rounding_mode, project.digits)))
        cur = end + timedelta(days=1)
    return out, True


# ---------------------------------------------------------------------------
# v1.2：计算链等式（可勾稽：定额合计 ± 各调整增量 - 压缩 = 最终周期）
# ---------------------------------------------------------------------------
def _fmt_g(v):
    """紧凑数字格式（整数去尾零，最多保留原精度）。"""
    if v is None:
        return ""
    if isinstance(v, float) and abs(v - round(v)) < 1e-9:
        return f"{v:g}"
    return f"{v:g}"


def build_equation_parts(base_total, adjustments, compress, regular_total, final_total):
    """从调整明细构建等式片段 [(label, value, operator)]。

    operator: "+"/"-"（value 本身含符号，负增量统一展示为减号）。
    起点为定额合计（已含多单体折减）；跳过多单体折减记录（其语义已含在起点中）；
    压缩仅在生效且 final < regular 时追加。逐项累加：base_total + Σdelta = regular_total。
    """
    parts = [("定额合计（含多单体折减）", base_total, "+")]
    for ad in adjustments:
        if ad.key == "6":
            continue
        op = "-" if ad.delta < 0 else "+"
        parts.append((ad.label, ad.delta, op))
    if compress.active and final_total < regular_total:
        parts.append(("压缩设计周期", round_value(regular_total - final_total, "round", 2), "-"))
    return parts


def format_equation(parts, final_total):
    """把等式片段拼成一行可读字符串，如：
    10.00（定额合计） + 1.00（…） - 0.50（…） = 10.50（最终周期）
    """
    if not parts:
        return ""
    buf = []
    for i, (label, val, op) in enumerate(parts):
        if i == 0:
            buf.append(f"{_fmt_g(val)}（{label}）")
        else:
            buf.append(f"{op}{_fmt_g(abs(val))}（{label}）")
    buf.append(f"= {_fmt_g(final_total)}（最终周期）")
    return " ".join(buf)
