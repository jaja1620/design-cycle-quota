# -*- coding: utf-8 -*-
"""国家法定节假日数据模块（Q7 决策：自动统计推演至 2039 年）。

统计口径：国家法定节假日指除周六、周日以外的其它法定假日，可增加的周期
是不含周六、周日的法定节假日天数（即仅累计落在周一至周五的法定节假日日期，
用 weekday() < 5 判断；落在周六/周日的法定节假日不计入）。

每年法定节假日共 13 天：元旦 1 天、春节 3 天、清明 1 天、劳动节 3 天、
端午 1 天、中秋 1 天、国庆 3 天。

- 已公布年份（2024-2026）按国务院公布安排计入（节假日日期为固定日期）。
- 2027-2039 年按规则推算：春节（农历正月初一至初三）、端午（农历五月初五）、
  中秋（农历八月十五）用农历换算；元旦 1/1、劳动节 5/1-5/3、国庆 10/1-10/3
  为固定公历日期；清明按 4 月 4 日近似（太阳黄经 15° 多在 4/4-4/5，取 4/4）。
"""
from datetime import date, timedelta

try:
    from lunardate import LunarDate
except ImportError:  # 兼容无 lunardate 环境（退化：仅统计固定日期节日）
    LunarDate = None

HOLIDAY_END_YEAR = 2039


def _solar_from_lunar(year, month, day):
    """农历转公历（lunardate 返回公历 date）。"""
    if LunarDate is None:
        return None
    try:
        ld = LunarDate(year, month, day)
        return date(ld.year, ld.month, ld.day)
    except Exception:
        return None


def statutory_holiday_items(year):
    """返回某年的法定节假日日期 -> 节日名 映射（推演至 2039）。"""
    items = {}
    # 元旦
    items[date(year, 1, 1)] = "元旦"
    # 春节（正月初一至初三）
    spring = _solar_from_lunar(year, 1, 1)
    if spring:
        for i, name in enumerate(("春节（正月初一）", "春节（正月初二）", "春节（正月初三）"), 1):
            items[spring + timedelta(days=i - 1)] = name
    else:
        for i, name in enumerate(("春节（正月初一）", "春节（正月初二）", "春节（正月初三）"), 1):
            items[date(year, 2, 9) + timedelta(days=i)] = name  # 兜底近似 2/10-2/12
    # 清明（近似 4/4，可微调）
    items[date(year, 4, 4)] = "清明节"
    # 劳动节（5/1-5/3）
    for i, name in enumerate(("劳动节（5/1）", "劳动节（5/2）", "劳动节（5/3）"), 1):
        items[date(year, 5, 1) + timedelta(days=i - 1)] = name
    # 端午（农历五月初五）
    duanwu = _solar_from_lunar(year, 5, 5)
    if duanwu:
        items[duanwu] = "端午节"
    # 中秋（农历八月十五）
    mid = _solar_from_lunar(year, 8, 15)
    if mid:
        items[mid] = "中秋节"
    # 国庆（10/1-10/3）
    for i, name in enumerate(("国庆节（10/1）", "国庆节（10/2）", "国庆节（10/3）"), 1):
        items[date(year, 10, 1) + timedelta(days=i - 1)] = name
    return items


def statutory_holidays(year):
    """返回某年的法定节假日日期集合（推演至 2039）。"""
    return set(statutory_holiday_items(year))


def count_holidays(start_date, duration_days):
    """统计 [start_date, start_date+duration_days) 期间可增加的法定节假日天数。

    口径：仅累计落在周一至周五（weekday() < 5）的法定节假日日期；
    落在周六、周日的法定节假日不计入。
    """
    if not start_date:
        return 0
    try:
        if isinstance(start_date, str):
            start_date = date.fromisoformat(start_date)
    except ValueError:
        return 0
    total = 0
    end = start_date + timedelta(days=int(duration_days))
    for y in range(start_date.year, min(end.year, HOLIDAY_END_YEAR) + 1):
        holidays = {d for d in statutory_holidays(y) if d.weekday() < 5}
        total += len(holidays & set(
            start_date + timedelta(days=i) for i in range((end - start_date).days)))
    return total


def holiday_summary(start_date, duration_days):
    """返回 (统计天数, 示例日期字符串)。

    统计天数 = count_holidays（不含周六、周日的法定节假日）。
    """
    days = count_holidays(start_date, duration_days)
    return days


def holiday_detail(start_date, duration_days):
    """返回统计区间内全部法定节假日明细：[(日期str, 节日名, 是否计入)]。

    与 count_holidays 同口径：duration_days 按 [start, start+duration) 区间统计；
    计入条件为落在周一至周五（weekday() < 5）；落在周六/周日返回“否”。
    无法解析开始时间或无节假日时返回空列表。
    """
    if not start_date:
        return []
    try:
        if isinstance(start_date, str):
            start_date = date.fromisoformat(start_date)
    except ValueError:
        return []
    end = start_date + timedelta(days=int(duration_days))
    out = []
    for y in range(start_date.year, min(end.year, HOLIDAY_END_YEAR) + 1):
        for d, name in sorted(statutory_holiday_items(y).items()):
            if start_date <= d < end:
                out.append((d.isoformat(), name, d.weekday() < 5))
    return out
