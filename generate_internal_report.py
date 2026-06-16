#!/usr/bin/env python3
"""
Generate an internal AliExpress Moloco insight report as a standalone HTML file.

The report reuses generate_dashboard.py's loader/classification logic so the
numbers stay aligned with the public dashboard.
"""

from __future__ import annotations

import contextlib
import html
import io
import json
import math
from collections import defaultdict
from pathlib import Path

import generate_dashboard as gd


ROOT = Path(__file__).resolve().parent
OUTPUT_FILE = ROOT / "internal_roi_share_report.html"

CAMPAIGN_IDS = [
    "aG917NDz5gBw6VRf",
    "VnntC5cAaSM427Eb",
    "aPLehU0cbfzIVdy3",
    "aAYWAdtK8yzGtcMp",
    "v7QnRXIVonS5uomZ",
    "C8w6QMX7C18oiKUy",
    "tD7ZrHZpoCqTQwKE",
    "F4Uf3q8nCIrXk9pe",
    "NSzoa2gxgzVwWqtv",
    "C2yH6jzJz4IzQyzZ",
    "rtHodELwUPFTZOsK",
    "KQzASDugGlYNVUcL",
    "UwfFBJhBjyhbTpx2",
    "tvXgqWrgtHKyOBlL",
    "WGR7ynKlrWf9GQ0F",
    "KynO7ddluR5Ued5R",
    "uZYdUYjX7BfsIDxD",
]

REPLACEMENT_IDS = [
    "aG917NDz5gBw6VRf",
    "VnntC5cAaSM427Eb",
    "aPLehU0cbfzIVdy3",
    "aAYWAdtK8yzGtcMp",
    "v7QnRXIVonS5uomZ",
    "C8w6QMX7C18oiKUy",
    "tD7ZrHZpoCqTQwKE",
]


def load_records() -> list[dict]:
    with contextlib.redirect_stdout(io.StringIO()):
        return gd.load_data()


def esc(value) -> str:
    return html.escape("" if value is None else str(value))


def fmt_ds(ds: str) -> str:
    return f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"


def fmt_short_ds(ds: str) -> str:
    return f"{int(ds[4:6])}/{int(ds[6:8])}"


def fmt_money(value, digits: int = 0) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    return f"${value:,.{digits}f}"


def fmt_num(value, digits: int = 2, suffix: str = "") -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    return f"{value:,.{digits}f}{suffix}"


def fmt_pct(value, signed: bool = True) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "-"
    sign = "+" if signed and value > 0 else ""
    return f"{sign}{value:.1f}%"


def label_for_record(record: dict) -> str:
    return gd.classify(record.get("name", ""), record.get("geo"))[1]


def primary_name(counter: dict[str, float]) -> str:
    if not counter:
        return "-"
    name = max(counter.items(), key=lambda item: item[1])[0]
    return gd.short_name(name)


def combined_label(counter: dict[str, float]) -> str:
    labels = [label for label, spend in sorted(counter.items(), key=lambda item: -item[1]) if spend > 0]
    return "/".join(labels) if labels else "-"


def weighted_roi(bucket: dict) -> float | None:
    spend = bucket.get("roi_spend", 0.0)
    return bucket.get("roi_weight", 0.0) / spend if spend > 0 else None


def compute_selected_campaign_trend(records: list[dict], dates: list[str]) -> dict:
    ids = set(CAMPAIGN_IDS)
    daily = defaultdict(lambda: {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0})
    date_set = set(dates)
    for record in records:
        if record["ds"] not in date_set:
            continue
        if str(record.get("cid") or "").strip() not in ids:
            continue
        spend = record.get("spend") or 0.0
        roi = record.get("roi")
        bucket = daily[record["ds"]]
        bucket["spend"] += spend
        if roi is not None and roi > 0 and spend > 0:
            bucket["roi_weight"] += roi * spend
            bucket["roi_spend"] += spend

    spend = [daily[d]["spend"] for d in dates]
    roi = [weighted_roi(daily[d]) for d in dates]
    total_spend = sum(spend)
    total_roi_weight = sum(daily[d]["roi_weight"] for d in dates)
    total_roi_spend = sum(daily[d]["roi_spend"] for d in dates)
    return {
        "dates": dates,
        "spend": spend,
        "roi": roi,
        "summary": {
            "spend": total_spend,
            "roi": total_roi_weight / total_roi_spend if total_roi_spend > 0 else None,
            "latest_roi": roi[-1] if roi else None,
            "campaign_count": len(CAMPAIGN_IDS),
        },
    }


def compute_campaign_vs_country(records: list[dict], dates: list[str]) -> dict:
    ids = set(CAMPAIGN_IDS)
    date_set = set(dates)
    country_daily = defaultdict(lambda: defaultdict(lambda: {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0}))
    campaign_daily = defaultdict(lambda: defaultdict(lambda: {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0}))
    campaign_labels = defaultdict(lambda: defaultdict(float))
    campaign_names = defaultdict(lambda: defaultdict(float))

    for record in records:
        ds = record.get("ds")
        if ds not in date_set:
            continue
        label = label_for_record(record)
        country_bucket = country_daily[label][ds]
        spend = record.get("spend") or 0.0
        country_bucket["spend"] += spend
        gmv = record.get("gmv")
        if gmv is not None:
            country_bucket["gmv"] += gmv
            country_bucket["gmv_rows"] += 1

        cid = str(record.get("cid") or "").strip()
        if cid not in ids:
            continue
        campaign_names[cid][record.get("name", "")] += spend
        campaign_labels[cid][label] += spend
        camp_bucket = campaign_daily[cid][ds]
        camp_bucket["spend"] += spend
        roi = record.get("roi")
        if roi is not None and roi > 0 and spend > 0:
            camp_bucket["roi_weight"] += roi * spend
            camp_bucket["roi_spend"] += spend

    rows = []
    missing = []
    for cid in CAMPAIGN_IDS:
        day_map = campaign_daily.get(cid)
        if not day_map:
            missing.append(cid)
            continue
        labels = [label for label, spend in campaign_labels[cid].items() if spend > 0]
        labels = labels or ["OTHER"]
        total = {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0}
        country_total = {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0}
        better_days = 0
        compare_days = 0
        valid_days = 0
        daily_roi = []
        daily_country_roi = []
        for ds in dates:
            camp_bucket = day_map.get(ds, {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0})
            total["spend"] += camp_bucket["spend"]
            total["roi_weight"] += camp_bucket["roi_weight"]
            total["roi_spend"] += camp_bucket["roi_spend"]
            camp_roi = weighted_roi(camp_bucket)
            if camp_roi is not None:
                valid_days += 1

            day_country = {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0}
            for label in labels:
                c = country_daily[label].get(ds, {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0})
                day_country["spend"] += c["spend"]
                day_country["gmv"] += c["gmv"]
                day_country["gmv_rows"] += c["gmv_rows"]
                country_total["spend"] += c["spend"]
                country_total["gmv"] += c["gmv"]
                country_total["gmv_rows"] += c["gmv_rows"]
            country_roi = (
                day_country["gmv"] / day_country["spend"]
                if day_country["spend"] > 0 and day_country["gmv_rows"] > 0
                else None
            )
            if camp_roi is not None and country_roi is not None:
                compare_days += 1
                if camp_roi >= country_roi:
                    better_days += 1
            daily_roi.append(camp_roi)
            daily_country_roi.append(country_roi)

        camp_roi_total = weighted_roi(total)
        country_roi_total = (
            country_total["gmv"] / country_total["spend"]
            if country_total["spend"] > 0 and country_total["gmv_rows"] > 0
            else None
        )
        diff = camp_roi_total - country_roi_total if camp_roi_total is not None and country_roi_total is not None else None
        if diff is None:
            status = "数据不足"
        elif diff >= 0 and better_days >= max(1, math.ceil(compare_days * 0.5)):
            status = "强于国家"
        elif diff >= 0:
            status = "高但波动"
        else:
            status = "弱于国家"
        rows.append(
            {
                "cid": cid,
                "name": primary_name(campaign_names[cid]),
                "label": combined_label(campaign_labels[cid]),
                "spend": total["spend"],
                "camp_roi": camp_roi_total,
                "country_roi": country_roi_total,
                "diff": diff,
                "better_days": better_days,
                "compare_days": compare_days,
                "valid_days": valid_days,
                "latest_roi": daily_roi[-1] if daily_roi else None,
                "latest_country_roi": daily_country_roi[-1] if daily_country_roi else None,
                "daily_roi": daily_roi,
                "daily_country_roi": daily_country_roi,
                "status": status,
            }
        )

    rows.sort(key=lambda r: (r["status"] != "强于国家", -(r["diff"] or -999), -r["spend"]))
    return {"rows": rows, "missing": missing, "dates": dates}


def compute_replacement(records: list[dict]) -> dict:
    ids = set(REPLACEMENT_IDS)
    before_dates = {f"2026060{d}" for d in range(1, 5)} | {"20260531"}
    after_dates = {f"202606{d:02d}" for d in range(5, 14)}
    windows = {
        "before": before_dates,
        "after": after_dates,
    }
    agg = defaultdict(lambda: {
        "names": defaultdict(float),
        "labels": defaultdict(float),
        "before": {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0, "spend_days": set()},
        "after": {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0, "spend_days": set()},
    })

    for record in records:
        cid = str(record.get("cid") or "").strip()
        if cid not in ids:
            continue
        ds = record.get("ds")
        window = None
        for key, date_set in windows.items():
            if ds in date_set:
                window = key
                break
        if window is None:
            continue
        spend = record.get("spend") or 0.0
        bucket = agg[cid][window]
        bucket["spend"] += spend
        if spend > 0:
            bucket["spend_days"].add(ds)
        roi = record.get("roi")
        if roi is not None and roi > 0 and spend > 0:
            bucket["roi_weight"] += roi * spend
            bucket["roi_spend"] += spend
        agg[cid]["names"][record.get("name", "")] += spend
        agg[cid]["labels"][label_for_record(record)] += spend

    rows = []
    for cid in REPLACEMENT_IDS:
        item = agg[cid]
        before = item["before"]
        after = item["after"]
        before_roi = weighted_roi(before)
        after_roi = weighted_roi(after)
        before_daily = before["spend"] / len(before["spend_days"]) if before["spend_days"] else None
        after_daily = after["spend"] / len(after["spend_days"]) if after["spend_days"] else None
        roi_change = (after_roi / before_roi - 1) * 100 if before_roi and after_roi else None
        spend_change = (after_daily / before_daily - 1) * 100 if before_daily and after_daily else None
        rows.append(
            {
                "cid": cid,
                "name": primary_name(item["names"]),
                "label": combined_label(item["labels"]),
                "before_roi": before_roi,
                "after_roi": after_roi,
                "roi_change": roi_change,
                "before_daily": before_daily,
                "after_daily": after_daily,
                "spend_change": spend_change,
                "before_days": len(before["spend_days"]),
                "after_days": len(after["spend_days"]),
            }
        )
    return {"rows": rows, "before": "2026-05-31~2026-06-04", "after": "2026-06-05~2026-06-13"}


def compute_replacement_daily(records: list[dict], dates: list[str]) -> dict:
    ids = set(REPLACEMENT_IDS)
    date_set = set(dates)
    agg = defaultdict(lambda: {
        "names": defaultdict(float),
        "labels": defaultdict(float),
        "daily": defaultdict(lambda: {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0}),
    })

    for record in records:
        ds = record.get("ds")
        cid = str(record.get("cid") or "").strip()
        if ds not in date_set or cid not in ids:
            continue
        spend = record.get("spend") or 0.0
        roi = record.get("roi")
        agg[cid]["names"][record.get("name", "")] += spend
        agg[cid]["labels"][label_for_record(record)] += spend
        bucket = agg[cid]["daily"][ds]
        bucket["spend"] += spend
        if roi is not None and roi > 0 and spend > 0:
            bucket["roi_weight"] += roi * spend
            bucket["roi_spend"] += spend

    rows = []
    detail = []
    for cid in REPLACEMENT_IDS:
        item = agg[cid]
        spend_series = []
        roi_series = []
        before = {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0}
        after = {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0}
        for ds in dates:
            bucket = item["daily"].get(ds, {"spend": 0.0, "roi_weight": 0.0, "roi_spend": 0.0})
            roi = weighted_roi(bucket)
            spend = bucket["spend"]
            spend_series.append(spend)
            roi_series.append(roi)
            target = after if ds >= "20260605" else before
            target["spend"] += spend
            target["roi_weight"] += bucket["roi_weight"]
            target["roi_spend"] += bucket["roi_spend"]
            detail.append(
                {
                    "date": fmt_ds(ds),
                    "cid": f'<span class="mono">{esc(cid)}</span>',
                    "name": esc(primary_name(item["names"])),
                    "label": esc(combined_label(item["labels"])),
                    "spend": fmt_money(spend, 0),
                    "roi": fmt_num(roi, 2, "x"),
                }
            )
        before_roi = weighted_roi(before)
        after_roi = weighted_roi(after)
        rows.append(
            {
                "cid": cid,
                "name": primary_name(item["names"]),
                "label": combined_label(item["labels"]),
                "spend": spend_series,
                "roi": roi_series,
                "total_spend": sum(spend_series),
                "latest_spend": spend_series[-1] if spend_series else None,
                "latest_roi": roi_series[-1] if roi_series else None,
                "before_roi": before_roi,
                "after_roi": after_roi,
                "roi_change": (after_roi / before_roi - 1) * 100 if before_roi and after_roi else None,
            }
        )

    total_spend = sum(sum(row["spend"]) for row in rows)
    latest_spend = sum((row["latest_spend"] or 0) for row in rows)
    return {
        "dates": dates,
        "rows": rows,
        "detail": detail,
        "summary": {
            "total_spend": total_spend,
            "latest_spend": latest_spend,
            "latest_date": dates[-1] if dates else None,
        },
    }


def svg_roi_line(dates: list[str], roi: list[float | None], title: str) -> str:
    width, height = 1060, 360
    left, right, top, bottom = 64, 60, 42, 58
    plot_w = width - left - right
    plot_h = height - top - bottom
    roi_vals = [v for v in roi if v is not None]
    min_roi = min(roi_vals) if roi_vals else 0
    max_roi = max(roi_vals) if roi_vals else 1
    if math.isclose(min_roi, max_roi):
        min_roi = 0
    padding = max((max_roi - min_roi) * 0.12, 0.3)
    min_roi = max(0, min_roi - padding)
    max_roi = max_roi + padding

    def x_at(i: int) -> float:
        return left + (plot_w / max(len(dates) - 1, 1)) * i

    def y_roi(value: float) -> float:
        return top + plot_h - ((value - min_roi) / max(max_roi - min_roi, 0.001)) * plot_h

    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="{esc(title)}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff"/>',
        f'<text x="{left}" y="24" class="svg-title">{esc(title)}</text>',
        f'<text x="{left}" y="{height - 16}" class="svg-axis">日期</text>',
        f'<text x="{left}" y="54" class="svg-axis">24H ROI</text>',
    ]
    for t in range(5):
        value = min_roi + (max_roi - min_roi) * t / 4
        y = y_roi(value)
        parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" stroke="#e5eaf2"/>')
        parts.append(f'<text x="{left - 10}" y="{y + 4:.1f}" class="svg-tick" text-anchor="end">{value:.1f}x</text>')
    for i, ds in enumerate(dates):
        x = x_at(i)
        if i % 2 == 0 or len(dates) <= 10:
            parts.append(f'<text x="{x:.1f}" y="{height - 34}" class="svg-tick" text-anchor="middle">{fmt_short_ds(ds)}</text>')
    points = []
    for i, value in enumerate(roi):
        if value is None:
            continue
        points.append((x_at(i), y_roi(value), value))
    if points:
        polyline = " ".join(f"{x:.1f},{y:.1f}" for x, y, _ in points)
        parts.append(f'<polyline points="{polyline}" fill="none" stroke="#ea580c" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>')
        for x, y, value in points:
            parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4.5" fill="#ea580c"/>')
        for x, y, value in points[-3:]:
            parts.append(f'<text x="{x:.1f}" y="{y - 10:.1f}" class="svg-note" text-anchor="middle">{value:.2f}x</text>')
    parts.append('<line x1="838" y1="23" x2="858" y2="23" stroke="#ea580c" stroke-width="3.5"/><text x="866" y="27" class="svg-axis">列举campaign加权ROI</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def heat_color(value: float | None) -> tuple[str, str]:
    if value is None:
        return "#eef1f5", "#8791a2"
    if value < 2:
        return "#fee2e2", "#991b1b"
    if value < 4:
        return "#fef3c7", "#92400e"
    if value < 6:
        return "#dcfce7", "#166534"
    return "#14532d", "#ffffff"


def svg_roi_heatmap(rows: list[dict], dates: list[str]) -> str:
    width = 1160
    left, top = 250, 62
    cell_w, cell_h = 58, 28
    right, bottom = 28, 38
    height = top + cell_h * len(rows) + bottom
    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="Campaign每日24H ROI热力图">',
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff"/>',
        '<text x="24" y="30" class="svg-title">近14天逐 Campaign 24H ROI 热力图</text>',
        '<text x="24" y="52" class="svg-axis">颜色越深代表 ROI 越高；灰色为当日无有效 ROI</text>',
    ]
    for j, ds in enumerate(dates):
        x = left + j * cell_w + cell_w / 2
        parts.append(f'<text x="{x:.1f}" y="52" class="svg-tick" text-anchor="middle">{fmt_short_ds(ds)}</text>')
    for i, row in enumerate(rows):
        y = top + i * cell_h
        label = f'{row["label"]} · {row["cid"][:6]}'
        parts.append(f'<text x="24" y="{y + 19}" class="svg-label">{esc(label)}</text>')
        for j, value in enumerate(row["daily_roi"]):
            x = left + j * cell_w
            fill, text = heat_color(value)
            parts.append(f'<rect x="{x:.1f}" y="{y + 2}" width="{cell_w - 4}" height="{cell_h - 4}" rx="5" fill="{fill}"/>')
            label_text = "-" if value is None else f"{value:.1f}"
            parts.append(f'<text x="{x + (cell_w - 4) / 2:.1f}" y="{y + 18}" class="svg-note" text-anchor="middle" style="fill:{text}">{label_text}</text>')
    parts.append(f'<text x="{width - right - 310}" y="{height - 16}" class="svg-axis">阈值：<2x 红 · 2-4x 黄 · 4-6x 浅绿 · >6x 深绿</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def svg_bar_chart(rows: list[dict], value_key: str, label_key: str, title: str, value_formatter, limit: int | None = None) -> str:
    rows = [r for r in rows if r.get(value_key) is not None]
    if limit:
        rows = rows[:limit]
    width = 1060
    row_h = 34
    top = 54
    left = 178
    right = 110
    height = top + row_h * max(len(rows), 1) + 34
    values = [r[value_key] for r in rows]
    max_abs = max([abs(v) for v in values] + [1])
    zero_x = left + (width - left - right) / 2
    scale = (width - left - right) / 2 / max_abs
    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="{esc(title)}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff"/>',
        f'<text x="24" y="30" class="svg-title">{esc(title)}</text>',
        f'<line x1="{zero_x:.1f}" y1="44" x2="{zero_x:.1f}" y2="{height - 24}" stroke="#cbd5e1" stroke-dasharray="4 4"/>',
    ]
    for i, row in enumerate(rows):
        y = top + i * row_h
        value = row[value_key]
        x = zero_x if value >= 0 else zero_x + value * scale
        bar_w = abs(value) * scale
        color = "#16a34a" if value < 0 else "#dc2626"
        if row.get("status") in ("强于国家", "高但波动"):
            color = "#2563eb" if value >= 0 else "#dc2626"
        if value < 0 and row.get("status") in ("弱于国家",):
            color = "#dc2626"
        parts.append(f'<text x="24" y="{y + 20}" class="svg-label">{esc(row[label_key])}</text>')
        parts.append(f'<rect x="{x:.1f}" y="{y + 7}" width="{max(bar_w, 1):.1f}" height="18" rx="5" fill="{color}"/>')
        value_x = x + bar_w + 8 if value >= 0 else x - 8
        anchor = "start" if value >= 0 else "end"
        parts.append(f'<text x="{value_x:.1f}" y="{y + 21}" class="svg-value" text-anchor="{anchor}">{esc(value_formatter(value))}</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def svg_grouped_replacement(rows: list[dict]) -> str:
    width, height = 1060, 460
    left, right, top, bottom = 82, 36, 62, 76
    plot_w = width - left - right
    plot_h = height - top - bottom
    max_roi = max([r["before_roi"] or 0 for r in rows] + [r["after_roi"] or 0 for r in rows] + [1])
    scale = (plot_h - 16) / max_roi
    base_y = top + plot_h
    group_w = plot_w / len(rows)
    bar_w = min(34, group_w * 0.26)
    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="6月5日替换前后24H ROI变化">',
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff"/>',
        '<text x="24" y="30" class="svg-title">6/5 标注替换前后 24H ROI</text>',
        f'<line x1="{left}" y1="{base_y:.1f}" x2="{width - right}" y2="{base_y:.1f}" stroke="#cbd5e1"/>',
        '<rect x="760" y="18" width="16" height="10" rx="2" fill="#94a3b8"/><text x="782" y="27" class="svg-axis">替换前</text>',
        '<rect x="840" y="18" width="16" height="10" rx="2" fill="#2563eb"/><text x="862" y="27" class="svg-axis">替换后</text>',
    ]
    for i, row in enumerate(rows):
        center = left + group_w * i + group_w / 2
        before_value = row["before_roi"] or 0
        after_value = row["after_roi"] or 0
        for j, (value, color) in enumerate(((before_value, "#94a3b8"), (after_value, "#2563eb"))):
            x = center + (j - 0.5) * (bar_w + 4)
            h = value * scale
            y = base_y - h
            parts.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{max(h, 1):.1f}" rx="5" fill="{color}"/>')
            parts.append(f'<text x="{x + bar_w / 2:.1f}" y="{y - 7:.1f}" class="svg-note" text-anchor="middle">{value:.1f}x</text>')
        short_cid = row["cid"][:4]
        parts.append(f'<text x="{center:.1f}" y="{height - 42}" class="svg-tick" text-anchor="middle">{esc(row["label"])}</text>')
        parts.append(f'<text x="{center:.1f}" y="{height - 22}" class="svg-tick muted" text-anchor="middle">{esc(short_cid)}</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def svg_replacement_daily_trends(rows: list[dict], dates: list[str]) -> str:
    width = 1160
    left, right = 92, 86
    top = 54
    row_h = 176
    plot_h = 92
    plot_w = width - left - right
    height = top + row_h * len(rows) + 48
    marker_ds = "20260605"
    marker_idx = dates.index(marker_ds) if marker_ds in dates else None

    def x_at(i: int) -> float:
        return left + plot_w * i / max(len(dates) - 1, 1)

    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="6月5日替换campaign每日消耗和ROI趋势">',
        f'<rect x="0" y="0" width="{width}" height="{height}" rx="16" fill="#ffffff"/>',
        '<text x="24" y="30" class="svg-title">6/5 替换 Campaign：分天消耗与 24H ROI 趋势</text>',
        '<rect x="760" y="18" width="16" height="10" rx="2" fill="#bfdbfe"/><text x="782" y="27" class="svg-axis">消耗</text>',
        '<line x1="830" y1="23" x2="850" y2="23" stroke="#ea580c" stroke-width="3.5"/><text x="858" y="27" class="svg-axis">24H ROI</text>',
        '<line x1="930" y1="23" x2="950" y2="23" stroke="#dc2626" stroke-width="2.5" stroke-dasharray="4 4"/><text x="958" y="27" class="svg-axis">6/5 替换</text>',
    ]

    for i, row in enumerate(rows):
        y0 = top + i * row_h
        chart_top = y0 + 34
        chart_base = chart_top + plot_h
        max_spend = max(row["spend"]) if row["spend"] else 1
        roi_vals = [v for v in row["roi"] if v is not None]
        max_roi = max(roi_vals) if roi_vals else 1
        min_roi = min(roi_vals) if roi_vals else 0
        if math.isclose(min_roi, max_roi):
            min_roi = 0
        padding = max((max_roi - min_roi) * 0.16, 0.5)
        min_roi = max(0, min_roi - padding)
        max_roi += padding
        bar_w = max(6, min(28, plot_w / max(len(dates), 1) * 0.56))

        def y_roi(value: float) -> float:
            return chart_base - ((value - min_roi) / max(max_roi - min_roi, 0.001)) * plot_h

        label = f'{row["label"]} · {row["cid"]} · {row["name"]}'
        parts.append(f'<text x="24" y="{y0 + 18}" class="svg-label">{esc(label)}</text>')
        parts.append(f'<text x="{width - right}" y="{y0 + 18}" class="svg-note" text-anchor="end">最新: {fmt_money(row["latest_spend"], 0)} · ROI {fmt_num(row["latest_roi"], 2, "x")}</text>')
        parts.append(f'<line x1="{left}" y1="{chart_base:.1f}" x2="{width - right}" y2="{chart_base:.1f}" stroke="#d8dee8"/>')
        for t in range(3):
            rv = min_roi + (max_roi - min_roi) * t / 2
            y = y_roi(rv)
            parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width - right}" y2="{y:.1f}" stroke="#edf1f7"/>')
            parts.append(f'<text x="{width - right + 8}" y="{y + 4:.1f}" class="svg-tick">{rv:.1f}x</text>')
        if marker_idx is not None:
            mx = x_at(marker_idx)
            parts.append(f'<line x1="{mx:.1f}" y1="{chart_top - 16}" x2="{mx:.1f}" y2="{chart_base + 18}" stroke="#dc2626" stroke-width="2" stroke-dasharray="5 5"/>')
            parts.append(f'<text x="{mx + 5:.1f}" y="{chart_top - 5}" class="svg-note" style="fill:#dc2626">6/5替换</text>')
        for j, spend in enumerate(row["spend"]):
            x = x_at(j)
            h = (spend / max_spend) * (plot_h * 0.78) if max_spend > 0 else 0
            parts.append(f'<rect x="{x - bar_w / 2:.1f}" y="{chart_base - h:.1f}" width="{bar_w:.1f}" height="{max(h, 1):.1f}" rx="3" fill="#bfdbfe"/>')
        points = []
        for j, roi in enumerate(row["roi"]):
            if roi is not None:
                points.append((x_at(j), y_roi(roi), roi))
        if points:
            polyline = " ".join(f"{x:.1f},{y:.1f}" for x, y, _ in points)
            parts.append(f'<polyline points="{polyline}" fill="none" stroke="#ea580c" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>')
            for x, y, roi in points:
                parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#ea580c"/>')
        for j, ds in enumerate(dates):
            if j % 3 == 0 or j == len(dates) - 1 or ds == marker_ds:
                parts.append(f'<text x="{x_at(j):.1f}" y="{chart_base + 32}" class="svg-tick" text-anchor="middle">{fmt_short_ds(ds)}</text>')
    parts.append("</svg>")
    return "\n".join(parts)


def table(headers: list[tuple[str, str]], rows: list[dict], row_class=None) -> str:
    head = "".join(f"<th>{esc(title)}</th>" for _, title in headers)
    body_rows = []
    for row in rows:
        cls = f' class="{row_class(row)}"' if row_class else ""
        cells = "".join(f"<td>{row.get(key, '')}</td>" for key, _ in headers)
        body_rows.append(f"<tr{cls}>{cells}</tr>")
    return f'<div class="table-wrap"><table><thead><tr>{head}</tr></thead><tbody>{"".join(body_rows)}</tbody></table></div>'


def render_report(context: dict) -> str:
    campaign = context["campaign"]
    replacement = context["replacement"]
    replacement_daily = context["replacement_daily"]
    selected = context["selected"]
    latest_ds = context["latest_ds"]
    last14_label = f"{fmt_ds(context['last14'][0])} ~ {fmt_ds(context['last14'][-1])}"
    replacement_range = f"{fmt_ds(replacement_daily['dates'][0])} ~ {fmt_ds(replacement_daily['dates'][-1])}"

    campaign_rows_sorted = campaign["rows"]
    heatmap_rows = sorted(campaign_rows_sorted, key=lambda r: r["camp_roi"] if r["camp_roi"] is not None else -1, reverse=True)
    campaign_chart_rows = [
        {
            "label": f'{r["label"]} · {r["cid"][:6]}',
            "diff": r["diff"],
            "status": r["status"],
        }
        for r in sorted(campaign["rows"], key=lambda r: r["diff"] if r["diff"] is not None else -999, reverse=True)
        if r["diff"] is not None
    ]
    strong_rows = [r for r in campaign_rows_sorted if r["status"] == "强于国家"]
    weak_rows = [r for r in campaign_rows_sorted if r["status"] == "弱于国家"]
    latest_rows = [r for r in campaign_rows_sorted if r.get("latest_roi") is not None]
    top_latest = max(latest_rows, key=lambda r: r["latest_roi"]) if latest_rows else None
    top_diff = campaign_chart_rows[:4]
    bottom_diff = sorted(
        [r for r in campaign_rows_sorted if r["diff"] is not None],
        key=lambda r: r["diff"],
    )[:4]

    top_diff_text = "、".join(f'{r["label"]} {r["cid"]}' for r in sorted(campaign_rows_sorted, key=lambda r: r["diff"] if r["diff"] is not None else -999, reverse=True)[:4])
    bottom_diff_text = "、".join(f'{r["label"]} {r["cid"]}' for r in bottom_diff)

    campaign_rows = []
    for r in campaign["rows"]:
        campaign_rows.append(
            {
                "cid": f'<span class="mono">{esc(r["cid"])}</span>',
                "name": esc(r["name"]),
                "label": esc(r["label"]),
                "camp_roi": fmt_num(r["camp_roi"], 2, "x"),
                "country_roi": fmt_num(r["country_roi"], 2, "x"),
                "diff": fmt_num(r["diff"], 2, "x") if r["diff"] is not None and r["diff"] < 0 else f'+{fmt_num(r["diff"], 2, "x")}' if r["diff"] is not None else "-",
                "latest_roi": fmt_num(r["latest_roi"], 2, "x"),
                "days": f'{r["better_days"]}/{r["compare_days"]}',
                "status": esc(r["status"]),
            }
        )

    replacement_rows = []
    for r in replacement["rows"]:
        replacement_rows.append(
            {
                "cid": f'<span class="mono">{esc(r["cid"])}</span>',
                "name": esc(r["name"]),
                "label": esc(r["label"]),
                "before_roi": fmt_num(r["before_roi"], 2, "x"),
                "after_roi": fmt_num(r["after_roi"], 2, "x"),
                "roi_change": fmt_pct(r["roi_change"]),
                "days": f'{r["before_days"]}/{r["after_days"]}',
            }
        )

    replacement_daily_rows = []
    for r in replacement_daily["rows"]:
        replacement_daily_rows.append(
            {
                "cid": f'<span class="mono">{esc(r["cid"])}</span>',
                "name": esc(r["name"]),
                "label": esc(r["label"]),
                "total_spend": fmt_money(r["total_spend"], 0),
                "latest_spend": fmt_money(r["latest_spend"], 0),
                "latest_roi": fmt_num(r["latest_roi"], 2, "x"),
                "before_roi": fmt_num(r["before_roi"], 2, "x"),
                "after_roi": fmt_num(r["after_roi"], 2, "x"),
                "roi_change": fmt_pct(r["roi_change"]),
            }
        )
    replacement_detail_rows = sorted(replacement_daily["detail"], key=lambda r: (r["date"], r["cid"]))
    replacement_latest_rows = [r for r in replacement_daily["rows"] if r.get("latest_roi") is not None]
    replacement_latest_best = max(replacement_latest_rows, key=lambda r: r["latest_roi"]) if replacement_latest_rows else None

    css = """
    :root {
      --bg: #f6f7fb;
      --ink: #172033;
      --muted: #657084;
      --line: #dbe1ea;
      --card: #ffffff;
      --blue: #2563eb;
      --orange: #ea580c;
      --green: #16a34a;
      --red: #dc2626;
      --amber: #b7791f;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font: 15px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    main { width: min(1180px, calc(100vw - 40px)); margin: 0 auto; padding: 34px 0 54px; }
    header {
      padding: 34px 0 20px;
      border-bottom: 1px solid var(--line);
      margin-bottom: 24px;
    }
    h1 { margin: 0 0 10px; font-size: 34px; line-height: 1.12; letter-spacing: 0; }
    h2 { margin: 34px 0 14px; font-size: 22px; line-height: 1.2; letter-spacing: 0; }
    h3 { margin: 0 0 8px; font-size: 16px; letter-spacing: 0; }
    p { margin: 0 0 10px; }
    .sub { color: var(--muted); font-size: 15px; }
    .grid { display: grid; gap: 14px; }
    .kpis { grid-template-columns: repeat(4, minmax(0, 1fr)); margin: 22px 0 8px; }
    .kpi, .panel {
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 10px 24px rgba(15, 23, 42, 0.05);
    }
    .kpi { padding: 18px; min-height: 104px; }
    .kpi .label { color: var(--muted); font-size: 13px; margin-bottom: 8px; }
    .kpi .value { font-size: 26px; font-weight: 760; line-height: 1.05; }
    .kpi .note { color: var(--muted); font-size: 12px; margin-top: 8px; }
    .panel { padding: 18px; margin: 14px 0; }
    .summary { grid-template-columns: 1.2fr 1fr; }
    .tabbar {
      display: flex;
      gap: 8px;
      align-items: center;
      border-bottom: 1px solid var(--line);
      margin: 0 0 22px;
      padding-bottom: 8px;
    }
    .tab-btn {
      appearance: none;
      border: 1px solid var(--line);
      background: #fff;
      color: var(--muted);
      border-radius: 7px;
      padding: 10px 14px;
      font: 700 14px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      cursor: pointer;
    }
    .tab-btn.active {
      color: #fff;
      border-color: var(--blue);
      background: var(--blue);
    }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }
    .callout {
      border-left: 4px solid var(--blue);
      padding: 14px 16px;
      background: #eef5ff;
      border-radius: 6px;
    }
    .callout.warn { border-left-color: var(--amber); background: #fff7e6; }
    .bullets { margin: 8px 0 0; padding-left: 18px; }
    .bullets li { margin: 6px 0; }
    .chart { width: 100%; height: auto; display: block; border: 1px solid var(--line); border-radius: 8px; background: #fff; }
    .svg-title { font: 700 18px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; fill: #172033; }
    .svg-label { font: 600 13px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; fill: #263447; }
    .svg-value, .svg-note { font: 700 12px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; fill: #263447; }
    .svg-axis, .svg-tick { font: 12px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; fill: #657084; }
    .muted { fill: #8791a2; }
    .table-wrap { overflow-x: auto; background: #fff; border: 1px solid var(--line); border-radius: 8px; }
    .dense-table .table-wrap { max-height: 560px; overflow: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 860px; }
    th, td { padding: 10px 12px; border-bottom: 1px solid #edf0f5; text-align: left; vertical-align: top; }
    th { font-size: 12px; color: #526074; background: #f9fafc; white-space: nowrap; }
    td { font-size: 13px; }
    tr:last-child td { border-bottom: 0; }
    tr.good td:first-child, tr.strong td:first-child { border-left: 4px solid var(--green); }
    tr.bad td:first-child, tr.weak td:first-child { border-left: 4px solid var(--red); }
    tr.neutral td:first-child { border-left: 4px solid #94a3b8; }
    .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; }
    .tag {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 999px;
      font-size: 12px;
      background: #eef2ff;
      color: #3343a7;
      white-space: nowrap;
    }
    .foot {
      color: var(--muted);
      font-size: 12px;
      border-top: 1px solid var(--line);
      margin-top: 34px;
      padding-top: 16px;
    }
    @media (max-width: 860px) {
      main { width: min(100vw - 24px, 1180px); padding-top: 18px; }
      h1 { font-size: 26px; }
      .kpis, .summary { grid-template-columns: 1fr; }
      .tabbar { overflow-x: auto; }
      .kpi { min-height: auto; }
    }
    """

    def campaign_row_class(row):
        status = row["status"]
        if status == "强于国家":
            return "strong"
        if status == "弱于国家":
            return "weak"
        return "neutral"

    selected_chart = svg_roi_line(selected["dates"], selected["roi"], "列举 Campaign 组合 24H ROI 趋势")
    heatmap_chart = svg_roi_heatmap(heatmap_rows, campaign["dates"])
    campaign_chart = svg_bar_chart(campaign_chart_rows, "diff", "label", "Campaign ROI 相对所在国家整体 ROI 的差值", lambda v: fmt_num(v, 2, "x") if v < 0 else f"+{fmt_num(v, 2, 'x')}")
    replacement_chart = svg_grouped_replacement(replacement["rows"])
    replacement_daily_chart = svg_replacement_daily_trends(replacement_daily["rows"], replacement_daily["dates"])

    if campaign["missing"]:
        missing_campaign = ", ".join(f'<span class="mono">{esc(cid)}</span>' for cid in campaign["missing"])
        campaign_note = f'<div class="callout warn"><strong>缺失 campaign：</strong>{missing_campaign}。KynO7ddluR5Ued5R 有效天数较少，结论需要谨慎看。</div>'
    else:
        campaign_note = '<div class="callout warn"><strong>样本提醒：</strong>KynO7ddluR5Ued5R 有效天数较少，结论需要谨慎看。</div>'

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AliExpress Campaign 24H ROI 内部复盘</title>
  <style>{css}</style>
</head>
<body>
<main>
  <header>
    <h1>AliExpress Campaign 24H ROI 内部复盘</h1>
    <p class="sub">仅覆盖你列举的 campaign · 数据截至 {fmt_ds(latest_ds)} · 近 14 天窗口 {last14_label}</p>
    <div class="grid kpis">
      <div class="kpi"><div class="label">分析 Campaign 数</div><div class="value">{len(campaign_rows_sorted)} 个</div><div class="note">当前有效清单</div></div>
      <div class="kpi"><div class="label">近14天组合 24H ROI</div><div class="value">{fmt_num(selected["summary"]["roi"], 2, "x")}</div><div class="note">按 campaign 当日权重加权</div></div>
      <div class="kpi"><div class="label">强于所在国家 ROI</div><div class="value">{len(strong_rows)} 个</div><div class="note">近14天加权 ROI 对比</div></div>
      <div class="kpi"><div class="label">最新日最高 ROI</div><div class="value">{fmt_num(top_latest["latest_roi"] if top_latest else None, 2, "x")}</div><div class="note">{esc(top_latest["cid"] if top_latest else "-")}</div></div>
    </div>
  </header>

  <nav class="tabbar" aria-label="report tabs">
    <button class="tab-btn active" type="button" data-tab="overview" onclick="showTab('overview')">Campaign ROI 总览</button>
    <button class="tab-btn" type="button" data-tab="replacement-daily" onclick="showTab('replacement-daily')">6/5替换专项</button>
  </nav>

  <section id="tab-overview" class="tab-panel active">
  <section class="grid summary">
    <div class="panel">
      <h3>内部分享主结论</h3>
      <ul class="bullets">
        <li>近14天列举 campaign 组合 24H ROI 为 {fmt_num(selected["summary"]["roi"], 2, "x")}，最新日为 {fmt_num(selected["summary"]["latest_roi"], 2, "x")}。</li>
        <li>相对所在国家整体 ROI 表现最强的是 {esc(top_diff_text)}。</li>
        <li>弱于所在国家整体 ROI 的主要是 {esc(bottom_diff_text)}，这几条更适合先观察或压低优先级。</li>
        <li>6/5 标注替换的 7 个 campaign 已放到单独 tab，按 2026-05-23 到最新日期展示分天消耗和 ROI 趋势。</li>
      </ul>
    </div>
    <div class="panel">
      <h3>口径说明</h3>
      <p>Campaign ROI 使用数据表里的 campaign 24H ROI，并按该 campaign 当日权重加权。</p>
      <p>国家/项目整体 ROI 使用同 label 的 Σ24h_gmv / Σcost，用于判断 campaign 是否强于所在国家整体表现。</p>
      <p>6/5替换专项 tab 从 2026-05-23 到最新日期展示每日消耗与 24H ROI，并用红色虚线标注 2026-06-05。</p>
    </div>
  </section>

  <h2>1. 列举 Campaign 组合 ROI 趋势</h2>
  <div class="panel">{selected_chart}</div>

  <h2>2. 逐 Campaign 每日 24H ROI</h2>
  <div class="panel">{heatmap_chart}</div>

  <h2>3. Campaign 相对所在国家 ROI</h2>
  <div class="panel">{campaign_chart}</div>
  {campaign_note}
  {table([
      ("cid", "Campaign ID"),
      ("name", "Campaign Name"),
      ("label", "国家/项目"),
      ("camp_roi", "近14天 24H ROI"),
      ("country_roi", "国家整体ROI"),
      ("diff", "ROI差值"),
      ("latest_roi", f"{fmt_short_ds(latest_ds)} 24H ROI"),
      ("days", "优于国家天数"),
      ("status", "判断"),
  ], campaign_rows, campaign_row_class)}

  </section>

  <section id="tab-replacement-daily" class="tab-panel">
    <section class="grid kpis">
      <div class="kpi"><div class="label">替换 Campaign 数</div><div class="value">{len(replacement_daily["rows"])} 个</div><div class="note">均标注 2026-06-05 替换</div></div>
      <div class="kpi"><div class="label">趋势日期范围</div><div class="value">{fmt_short_ds(replacement_daily["dates"][0])} - {fmt_short_ds(replacement_daily["dates"][-1])}</div><div class="note">{esc(replacement_range)}</div></div>
      <div class="kpi"><div class="label">累计消耗</div><div class="value">{fmt_money(replacement_daily["summary"]["total_spend"], 0)}</div><div class="note">7个campaign合计</div></div>
      <div class="kpi"><div class="label">最新日最高ROI</div><div class="value">{fmt_num(replacement_latest_best["latest_roi"] if replacement_latest_best else None, 2, "x")}</div><div class="note">{esc(replacement_latest_best["cid"] if replacement_latest_best else "-")}</div></div>
    </section>

    <div class="callout"><strong>特殊标注：</strong>下面 7 个 campaign 均按 2026-06-05 替换处理，趋势图里的红色虚线就是替换日期。</div>

    <h2>6/5替换 Campaign 分天消耗与 ROI 趋势</h2>
    <div class="panel">{replacement_daily_chart}</div>

    <h2>替换 Campaign 汇总</h2>
    {table([
        ("cid", "Campaign ID"),
        ("name", "Campaign Name"),
        ("label", "国家/项目"),
        ("total_spend", "5/23至最新累计消耗"),
        ("latest_spend", f"{fmt_short_ds(latest_ds)} 消耗"),
        ("latest_roi", f"{fmt_short_ds(latest_ds)} 24H ROI"),
        ("before_roi", "6/5前ROI"),
        ("after_roi", "6/5后ROI"),
        ("roi_change", "ROI变化"),
    ], replacement_daily_rows)}

    <h2>分天明细</h2>
    <div class="dense-table">
    {table([
        ("date", "日期"),
        ("cid", "Campaign ID"),
        ("name", "Campaign Name"),
        ("label", "国家/项目"),
        ("spend", "消耗"),
        ("roi", "24H ROI"),
    ], replacement_detail_rows)}
    </div>
  </section>

  <div class="foot">
    数据源：Downloads/aliexpress_moloco_compaign_data.xlsx sheet“最新” + repo历史 AE 数据。生成脚本：generate_internal_report.py。报告用于内部复盘分享。
  </div>
  <script>
    function showTab(name, updateHash = true) {{
      document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.toggle('active', panel.id === 'tab-' + name));
      document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.tab === name));
      if (updateHash) window.location.hash = name;
    }}
    document.addEventListener('DOMContentLoaded', () => {{
      const initial = (window.location.hash || '').replace('#', '') || 'overview';
      showTab(initial === 'replacement-daily' ? initial : 'overview', false);
    }});
  </script>
</main>
</body>
</html>
"""


def main() -> None:
    records = load_records()
    dates = gd.get_all_dates(records)
    if not dates:
        raise RuntimeError("No valid dates found in source data.")
    latest_ds = dates[-1]
    last14 = gd.get_window(records, 14)
    replacement_dates = [d for d in dates if "20260523" <= d <= latest_ds]
    context = {
        "latest_ds": latest_ds,
        "last14": last14,
        "selected": compute_selected_campaign_trend(records, last14),
        "campaign": compute_campaign_vs_country(records, last14),
        "replacement": compute_replacement(records),
        "replacement_daily": compute_replacement_daily(records, replacement_dates),
    }
    html_text = render_report(context)
    OUTPUT_FILE.write_text(html_text, encoding="utf-8")
    manifest = {
        "output": str(OUTPUT_FILE),
        "latest_ds": latest_ds,
        "last14": [last14[0], last14[-1]],
        "replacement_daily": [replacement_dates[0], replacement_dates[-1]],
        "campaigns": len(context["campaign"]["rows"]),
        "replacement_campaigns": len(context["replacement_daily"]["rows"]),
        "selected_campaign_roi": context["selected"]["summary"]["roi"],
        "missing_campaigns": context["campaign"]["missing"],
    }
    print(json.dumps(manifest, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
