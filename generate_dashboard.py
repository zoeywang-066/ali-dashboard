#!/usr/bin/env python3
"""
阿里投放数据 Dashboard Generator
用法: python3 generate_dashboard.py
更新数据后重新运行即可刷新看板
"""

import json
import re
import subprocess
import sys
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path

DESKTOP      = Path.home() / "Desktop"
DOWNLOADS    = Path.home() / "Downloads"
REPO_DIR     = Path.home() / "ali-dashboard"
OUTPUT_FILE  = str(REPO_DIR / "index.html")
XLSX_FILE    = DOWNLOADS / "aliexpress_moloco_compaign_data.xlsx"
CSV_FILE     = DESKTOP / "阿里投放数据.csv"
NUMBERS_FILE = DESKTOP / "阿里投放数据.numbers"
SHEET_NAME   = "最新"

# ─── 数据过滤 ────────────────────────────────────────────────────────────────
def is_other_channel(name):
    """剔除其他渠道（criteo）和脏数据行。保留 copy of * 等仍属 Moloco 数据的项。"""
    if not name:
        return True
    if name in ("campaign_name", "campaign_namedsp"):
        return True
    if "criteo" in name.lower():
        return True
    return False

# ─── 分类 & 名称 ─────────────────────────────────────────────────────────────
def classify(name):
    if not name: return ("country", "未知")
    if "EU10"    in name: return ("project", "EU10")
    if "海托"    in name: return ("project", "海托")
    if "欧洲本地" in name: return ("project", "欧洲本地")
    m = re.search(r'[Aa]ndroid[_\s][Aa]pp[_\s]([A-Z]{1,5}(?:/[A-Z]{1,5})?)', name, re.IGNORECASE)
    if m: return ("country", m.group(1).upper())
    return ("country", "OTHER")

def short_name(name):
    for pfx in ("AliExpress_moloco_rmkt_rta_", "AliExpress_moloco_rmkt_null_"):
        if name.startswith(pfx): name = name[len(pfx):]
    for sfx in ("_RE_null",):
        if name.endswith(sfx): name = name[:-len(sfx)]
    name = re.sub(r'^[Aa]ndroid[_][Aa]pp[_]', '', name)
    return name.strip("_")

# ─── 读取 ────────────────────────────────────────────────────────────────────
def load_data():
    # 优先读取 xlsx，其次 csv，最后回落到 Numbers
    if XLSX_FILE.exists():
        records = _load_xlsx(XLSX_FILE)
    elif CSV_FILE.exists():
        records = _load_csv(CSV_FILE)
    elif NUMBERS_FILE.exists():
        records = _load_numbers(NUMBERS_FILE)
    else:
        print(f"错误：找不到数据文件\n  {XLSX_FILE}\n  {CSV_FILE}\n  {NUMBERS_FILE}")
        sys.exit(1)

    before = len(records)
    excluded = {r["name"] for r in records if is_other_channel(r["name"])}
    records = [r for r in records if not is_other_channel(r["name"])]
    if excluded:
        print(f"  过滤其他渠道(criteo)+脏数据: -{before - len(records)} 行 ({len(excluded)} 个 campaign)")
    return records

# 已知表头关键词集合（用于数据行污染检测）
HEADER_KEYWORDS = {
    "ds", "campaign_name", "campaign_namedsp", "campaign_id",
    "花费", "costdsp", "24h_dac", "session_dac", "dac成本",
    "24h-gmvroi", "24h_gmvroi",
}

def _load_xlsx(path):
    import pandas as pd
    print(f"读取 xlsx: {path.name}  sheet={SHEET_NAME}")
    raw = pd.read_excel(path, sheet_name=SHEET_NAME, dtype=str, header=None)

    # 找出所有 header 行（首列值为 'ds'）
    header_idxs = [i for i, row in raw.iterrows()
                   if str(row.iloc[0]).strip().lower() == "ds"]
    if not header_idxs:
        print("  ❌ 警告：未找到任何 header 行（首列='ds'）")
        return []

    print(f"  检测到 {len(header_idxs)} 个数据段")

    REQUIRED_FIELDS = {
        "spend(花费/costdsp)":         ["花费", "costdsp"],
        "campaign_name":               ["campaign_name", "campaign_namedsp"],
        "campaign_id":                 ["campaign_id"],
    }
    OPTIONAL_FIELDS = {
        "dac(24h_dac/session_dac)":    ["24h_dac", "session_dac"],
        "dac_cost":                    ["dac成本"],
        "roi(24h-gmvroi)":             ["24h-gmvroi", "24h_gmvroi"],
    }

    all_records = []
    warnings = []
    sec_stats = []  # [(sec_i, h_i, parsed, contaminated, missing)]

    for sec_i, h_i in enumerate(header_idxs):
        end_i = header_idxs[sec_i + 1] if sec_i + 1 < len(header_idxs) else len(raw)
        cols  = [str(v).strip() for v in raw.iloc[h_i].tolist()]
        cm    = {c: ci for ci, c in enumerate(cols) if c not in ("nan", "None", "")}

        # ★ 校验1：必需字段
        missing_required = [lbl for lbl, cands in REQUIRED_FIELDS.items()
                            if not any(c in cm for c in cands)]
        if missing_required:
            warnings.append(f"  ⚠️  段#{sec_i+1} (行{h_i+1}) 缺少必需字段: {', '.join(missing_required)}")

        # ★ 校验2：可选字段（仅提示，不阻塞）
        missing_optional = [lbl for lbl, cands in OPTIONAL_FIELDS.items()
                            if not any(c in cm for c in cands)]
        if missing_optional and sec_i < 3:  # 只对前 3 段报告可选字段缺失，避免噪音
            warnings.append(f"  ℹ️  段#{sec_i+1} (行{h_i+1}) 缺少可选字段: {', '.join(missing_optional)}")

        # 花费列：花费 > costdsp
        spend_i    = cm.get("花费",        cm.get("costdsp"))
        roi_i      = cm.get("24h-gmvroi",  cm.get("24h_gmvroi"))
        # DAC列：优先 24h_dac（新格式），其次 session_dac
        dac_i      = cm.get("24h_dac",     cm.get("session_dac"))
        # dac成本：新格式直接提供
        dac_cost_i = cm.get("dac成本")
        name_i     = cm.get("campaign_name", cm.get("campaign_namedsp"))
        cid_i      = cm.get("campaign_id")

        def _gv(row, idx):
            if idx is None: return None
            v = str(row.iloc[idx]).strip()
            return None if v in ("nan", "None", "", "-") else v

        def _tf(v):
            try: return float(str(v).replace(",", "")) if v is not None else None
            except: return None

        sec_parsed = 0
        sec_contaminated = 0
        for row_i in range(h_i + 1, end_i):
            row = raw.iloc[row_i]
            try:
                ds_raw = str(row.iloc[0]).strip().split(".")[0]
                if not ds_raw.isdigit() or len(ds_raw) != 8:
                    continue

                # ★ 校验3：字段污染检测
                # 检查 name 字段是否是表头关键词（说明这行是混入的字段行）
                name_val = _gv(row, name_i)
                if name_val and name_val.lower() in HEADER_KEYWORDS:
                    sec_contaminated += 1
                    warnings.append(
                        f"  ⚠️  段#{sec_i+1} 行{row_i+1}: name='{name_val}' 疑似表头污染，已剔除"
                    )
                    continue

                # 检查 spend 列是否是字符串关键词（说明列对错了或行错位）
                spend_raw = _gv(row, spend_i)
                if spend_raw and spend_raw.lower() in HEADER_KEYWORDS:
                    sec_contaminated += 1
                    warnings.append(
                        f"  ⚠️  段#{sec_i+1} 行{row_i+1}: spend='{spend_raw}' 列错位/污染，已剔除"
                    )
                    continue

                name  = name_val or ""
                cid   = _gv(row, cid_i) or ""
                spend = _tf(spend_raw) or 0.0
                roi   = _tf(_gv(row, roi_i))
                dac   = _tf(_gv(row, dac_i)) or 0.0
                # dac成本：新格式直接读取；旧格式留 None（由 build_dac_data 按汇总计算）
                dac_cost = _tf(_gv(row, dac_cost_i)) if dac_cost_i is not None else None
                all_records.append({
                    "ds": ds_raw, "name": name, "cid": cid,
                    "spend": spend, "roi": roi, "dac": dac, "dac_cost": dac_cost,
                })
                sec_parsed += 1
            except Exception as e:
                warnings.append(f"  ⚠️  段#{sec_i+1} 行{row_i+1} 解析异常: {e}")
                continue

        sec_stats.append((sec_i + 1, h_i + 1, sec_parsed, sec_contaminated, len(missing_required)))

    # ★ 打印校验报告
    if warnings:
        print("─── 数据校验警告 ───")
        # 限制告警条数，避免刷屏
        for w in warnings[:30]:
            print(w)
        if len(warnings) > 30:
            print(f"  ... 还有 {len(warnings) - 30} 条警告未显示")
        print("──────────────────")

    # 段级行数统计（若段间差异大可能是数据问题）
    if sec_stats:
        parsed_counts = [s[2] for s in sec_stats]
        total_contaminated = sum(s[3] for s in sec_stats)
        sections_with_issues = sum(1 for s in sec_stats if s[3] > 0 or s[4] > 0)
        if total_contaminated > 0 or sections_with_issues > 0:
            print(f"  校验汇总: {sections_with_issues}/{len(sec_stats)} 段有问题, "
                  f"共剔除 {total_contaminated} 行污染数据")
        # 异常段（行数远低于中位数）告警
        if len(parsed_counts) >= 3:
            sorted_counts = sorted(parsed_counts)
            median = sorted_counts[len(sorted_counts) // 2]
            outliers = [s for s in sec_stats if median > 0 and s[2] < median * 0.3 and s[2] > 0]
            if outliers:
                detail = ", ".join(f"段#{s[0]}={s[2]}行" for s in outliers)
                print(f"  ⚠️  段行数偏低(中位数={median}): {detail}")

    print(f"加载 {len(all_records)} 行数据")
    return all_records

def _load_csv(path):
    import pandas as pd
    print(f"读取 CSV: {path.name}")
    df = pd.read_csv(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return _df_to_records(df)

def _load_numbers(path):
    try:
        import numbers_parser
    except ImportError:
        print("pip3 install numbers-parser --break-system-packages")
        sys.exit(1)
    print(f"读取 Numbers: {path.name}")
    doc   = numbers_parser.Document(str(path))
    table = doc.sheets[0].tables[0]
    header = [table.cell(1, c).value for c in range(table.num_cols)]
    col    = {v: i for i, v in enumerate(header) if v}
    records = []
    for r in range(2, table.num_rows):
        ds_val = table.cell(r, col.get("ds", 0)).value
        if not ds_val or not isinstance(ds_val, (int, float)): continue
        ds      = str(int(ds_val))
        name    = table.cell(r, col.get("campaign_name", 2)).value or ""
        cid     = table.cell(r, col.get("campaign_id",   1)).value or ""
        spend_v = table.cell(r, col.get("花费",          5)).value
        roi_v   = table.cell(r, col.get("24h-gmvroi",   20)).value
        spend = float(spend_v) if isinstance(spend_v, (int, float)) else 0.0
        roi   = float(roi_v)   if isinstance(roi_v,   (int, float)) else None
        records.append({"ds": ds, "name": name, "cid": cid, "spend": spend, "roi": roi, "dac": 0, "dac_cost": None})
    print(f"加载 {len(records)} 行数据")
    return records

def _df_to_records(df):
    """CSV / Numbers 备用加载路径（xlsx 走 _load_xlsx 逐 section 解析）"""
    records = []
    roi_col      = next((c for c in df.columns if "24h-gmvroi" in c or "24h_gmvroi" in c), None)
    spend_col    = next((c for c in df.columns if c in ("花费", "spend", "costdsp")), None)
    dac_col      = next((c for c in df.columns if c in ("24h_dac", "session_dac")), None)
    dac_cost_col = next((c for c in df.columns if c == "dac成本"), None)
    for _, row in df.iterrows():
        try:
            ds_raw = str(row.get("ds", "")).strip().split(".")[0]
            if not ds_raw.isdigit() or len(ds_raw) != 8: continue
            name     = str(row.get("campaign_name", row.get("campaign_namedsp", "")) or "")
            cid      = str(row.get("campaign_id", "") or "")
            spend    = float(str(row.get(spend_col, 0) or 0).replace(",", "")) if spend_col else 0.0
            roi_raw  = row.get(roi_col) if roi_col else None
            roi      = float(roi_raw) if roi_raw not in (None, "", "nan", "None") else None
            dac_raw  = row.get(dac_col) if dac_col else None
            dac      = float(str(dac_raw).replace(",", "")) if dac_raw not in (None, "", "nan", "None") else 0.0
            dc_raw   = row.get(dac_cost_col) if dac_cost_col else None
            dac_cost = float(str(dc_raw).replace(",", "")) if dc_raw not in (None, "", "nan", "None") else None
            records.append({"ds": ds_raw, "name": name, "cid": cid,
                            "spend": spend, "roi": roi, "dac": dac, "dac_cost": dac_cost})
        except (ValueError, TypeError):
            continue
    print(f"加载 {len(records)} 行数据")
    return records

# ─── DAC 数据（来自最新 sheet dac成本字段，24h_dac 口径）──────────────────────
def build_dac_data(records):
    """按 3/7/14 天窗口聚合 dac成本，返回国家汇总 + 分天趋势 + Campaign 分天趋势"""
    if not records:
        return {}
    result = {}
    for n in (3, 7, 14):
        dates    = get_window(records, n)
        labels   = fmt_dates(dates)
        date_set = set(dates)

        # country → day → {cost, dac}
        c_daily = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0}))
        # country → camp_short → day → {cost, dac}
        c_camp  = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0})))
        # DAC 专项：整体 + 分 campaign
        sp_overall = defaultdict(lambda: {"cost": 0.0, "dac": 0.0})           # ds → {cost, dac}
        sp_camp    = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0}))  # camp_short → ds → {cost, dac}

        for r in records:
            if r["ds"] not in date_set:
                continue
            # 只用有 dac成本 直接字段的记录（24h_dac 口径），session_dac 不参与
            if r.get("dac_cost") is None:
                continue
            _, label = classify(r["name"])
            sn = short_name(r["name"])
            c_daily[label][r["ds"]]["cost"] += r["spend"]
            c_daily[label][r["ds"]]["dac"]  += r.get("dac") or 0
            c_camp[label][sn][r["ds"]]["cost"] += r["spend"]
            c_camp[label][sn][r["ds"]]["dac"]  += r.get("dac") or 0
            # DAC 专项独立聚合（不影响国家口径）
            if "DAC专项" in (r.get("name") or ""):
                sp_overall[r["ds"]]["cost"] += r["spend"]
                sp_overall[r["ds"]]["dac"]  += r.get("dac") or 0
                sp_camp[sn][r["ds"]]["cost"] += r["spend"]
                sp_camp[sn][r["ds"]]["dac"]  += r.get("dac") or 0

        # Country summary（窗口合计）
        country_summary = {}
        for c, ds_map in c_daily.items():
            tc = sum(b["cost"] for b in ds_map.values())
            td = sum(b["dac"]  for b in ds_map.values())
            if td > 0:
                country_summary[c] = {
                    "cost":        round(tc, 0),
                    "dac":         round(td, 0),
                    "dac_cost":    round(tc / td, 2),
                    "daily_spend": round(tc / n, 0),
                }

        # Country daily series（分天 DAC成本）
        country_daily = {}
        for c, ds_map in c_daily.items():
            series = [round(ds_map[d]["cost"] / ds_map[d]["dac"], 2)
                      if ds_map.get(d, {}).get("dac", 0) > 0 else None
                      for d in dates]
            if any(v is not None for v in series):
                country_daily[c] = series

        # Camp daily per country
        camp_daily = {}
        for c, camp_map in c_camp.items():
            camps = {}
            for camp, ds_map in camp_map.items():
                series = [round(ds_map[d]["cost"] / ds_map[d]["dac"], 2)
                          if ds_map.get(d, {}).get("dac", 0) > 0 else None
                          for d in dates]
                if any(v is not None for v in series):
                    camps[camp] = series
            if camps:
                camp_daily[c] = camps

        # DAC 专项：整体 + 分 campaign 的汇总和分天序列
        def _series(ds_map):
            spend_s = [round(ds_map.get(d, {}).get("cost", 0), 2) for d in dates]
            dac_s   = [int(round(ds_map.get(d, {}).get("dac", 0), 0)) for d in dates]
            cost_s  = [round(ds_map[d]["cost"] / ds_map[d]["dac"], 2)
                       if ds_map.get(d, {}).get("dac", 0) > 0 else None
                       for d in dates]
            return spend_s, dac_s, cost_s

        dac_special = {}
        if sp_overall:
            tc = sum(b["cost"] for b in sp_overall.values())
            td = sum(b["dac"]  for b in sp_overall.values())
            sp_s, dc_s, cs_s = _series(sp_overall)
            dac_special["overall"] = {
                "cost":         round(tc, 0),
                "dac":          int(round(td, 0)),
                "dac_cost":     round(tc / td, 2) if td > 0 else None,
                "daily_spend":  round(tc / n, 0),
                "spend_series": sp_s,
                "dac_series":   dc_s,
                "cost_series":  cs_s,
            }
            camps = {}
            for camp, ds_map in sp_camp.items():
                tc2 = sum(b["cost"] for b in ds_map.values())
                td2 = sum(b["dac"]  for b in ds_map.values())
                sp2, dc2, cs2 = _series(ds_map)
                camps[camp] = {
                    "cost":         round(tc2, 0),
                    "dac":          int(round(td2, 0)),
                    "dac_cost":     round(tc2 / td2, 2) if td2 > 0 else None,
                    "daily_spend":  round(tc2 / n, 0),
                    "spend_series": sp2,
                    "dac_series":   dc2,
                    "cost_series":  cs2,
                }
            dac_special["campaigns"] = camps

        result[str(n)] = {
            "labels":          labels,
            "country_summary": country_summary,
            "country_daily":   country_daily,
            "camp_daily":      camp_daily,
            "dac_special":     dac_special,
        }
        n_countries = len(country_summary)
        n_special   = len(dac_special.get("campaigns", {})) if dac_special else 0
        print(f"  [DAC] {n}天窗口: {n_countries} 个国家/项目有 dac成本 数据 · DAC专项 {n_special} 个 campaign")
    return result

# ─── 日期窗口 ─────────────────────────────────────────────────────────────────
def get_window(records, n):
    all_dates = sorted({r["ds"] for r in records if r["ds"].isdigit()})
    if not all_dates: return []
    end_dt   = datetime.strptime(all_dates[-1], "%Y%m%d")
    start_dt = end_dt - timedelta(days=n - 1)
    return [d for d in all_dates if datetime.strptime(d, "%Y%m%d") >= start_dt]

# ─── 聚合 ────────────────────────────────────────────────────────────────────
def aggregate(records, dates):
    date_set    = set(dates)
    proj_agg    = defaultdict(lambda: defaultdict(lambda: {"spend":0,"ws":0,"wsr":0}))
    country_agg = defaultdict(lambda: defaultdict(lambda: {"spend":0,"ws":0,"wsr":0}))
    camp_agg    = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"spend":0,"ws":0,"wsr":0})))
    cid_to_short = {}

    for r in records:
        if r["ds"] not in date_set: continue
        kind, label = classify(r["name"])
        target = proj_agg[label] if kind == "project" else country_agg[label]
        b = target[r["ds"]]
        b["spend"] += r["spend"]
        if r["roi"] and r["roi"] > 0 and r["spend"] > 0:
            b["ws"]  += r["spend"]
            b["wsr"] += r["roi"] * r["spend"]
        sn = cid_to_short.setdefault(r["cid"], short_name(r["name"]))
        cb = camp_agg[label][sn][r["ds"]]
        cb["spend"] += r["spend"]
        if r["roi"] and r["roi"] > 0 and r["spend"] > 0:
            cb["ws"]  += r["spend"]
            cb["wsr"] += r["roi"] * r["spend"]

    def to_series(agg_dict):
        result = {}
        for label, ds_map in agg_dict.items():
            spend_list, roi_list = [], []
            for d in dates:
                b = ds_map.get(d, {"spend":0,"ws":0,"wsr":0})
                spend_list.append(round(b["spend"], 2))
                roi_list.append(round(b["wsr"]/b["ws"], 2) if b["ws"] > 0 else None)
            if sum(s for s in spend_list if s) > 0:
                result[label] = {"spend": spend_list, "roi": roi_list}
        return result

    def to_camp_series():
        result = {}
        for group, camp_map in camp_agg.items():
            gcamps = {}
            for camp_label, ds_map in camp_map.items():
                spend_list, roi_list = [], []
                for d in dates:
                    b = ds_map.get(d, {"spend":0,"ws":0,"wsr":0})
                    spend_list.append(round(b["spend"], 2))
                    roi_list.append(round(b["wsr"]/b["ws"], 2) if b["ws"] > 0 else None)
                if sum(s for s in spend_list if s) > 0:
                    gcamps[camp_label] = {"spend": spend_list, "roi": roi_list}
            if gcamps:
                result[group] = gcamps
        return result

    return to_series(proj_agg), to_series(country_agg), to_camp_series()

def fmt_dates(dates):
    return [f"{d[4:6]}/{d[6:8]}" for d in dates]

# ─── 信号计算 ─────────────────────────────────────────────────────────────────
def _compute_signals_for_window(records, compare_days):
    """compare_days: 1 / 3 / 7 — 近N天 vs 前N天"""
    all_dates = sorted({r["ds"] for r in records if r["ds"].isdigit()})
    if not all_dates: return [], {}

    end_dt       = datetime.strptime(all_dates[-1], "%Y%m%d")
    recent_start = end_dt - timedelta(days=compare_days - 1)
    prev_end     = recent_start - timedelta(days=1)
    prev_start   = prev_end - timedelta(days=compare_days - 1)

    recent_dates = {d for d in all_dates if datetime.strptime(d, "%Y%m%d") >= recent_start}
    prev_dates   = {d for d in all_dates
                    if prev_start <= datetime.strptime(d, "%Y%m%d") <= prev_end}

    cid_to_short = {}
    agg = defaultdict(lambda: {"r_spend":0,"r_ws":0,"r_wsr":0,"p_spend":0,"p_ws":0,"p_wsr":0})

    for r in records:
        _, label = classify(r["name"])
        sn = cid_to_short.setdefault(r["cid"], short_name(r["name"]))
        camp_key = f"{label}|||{sn}"
        for key in (label, camp_key):
            if r["ds"] in recent_dates:
                agg[key]["r_spend"] += r["spend"]
                if r["roi"] and r["roi"]>0 and r["spend"]>0:
                    agg[key]["r_ws"]  += r["spend"]
                    agg[key]["r_wsr"] += r["roi"] * r["spend"]
            elif r["ds"] in prev_dates:
                agg[key]["p_spend"] += r["spend"]
                if r["roi"] and r["roi"]>0 and r["spend"]>0:
                    agg[key]["p_ws"]  += r["spend"]
                    agg[key]["p_wsr"] += r["roi"] * r["spend"]

    grp_ws  = sum(b["r_ws"]  for k,b in agg.items() if "|||" not in k)
    grp_wsr = sum(b["r_wsr"] for k,b in agg.items() if "|||" not in k)
    overall_roi = grp_wsr / grp_ws if grp_ws > 0 else 5.0

    # 日内波动（仅7天窗口有意义）
    daily_roi = defaultdict(lambda: defaultdict(lambda: {"ws":0,"wsr":0}))
    for r in records:
        if r["ds"] not in recent_dates: continue
        _, label = classify(r["name"])
        if r["roi"] and r["roi"]>0 and r["spend"]>0:
            daily_roi[label][r["ds"]]["ws"]  += r["spend"]
            daily_roi[label][r["ds"]]["wsr"] += r["roi"] * r["spend"]

    def daily_cv(label):
        vals = [b["wsr"]/b["ws"] for b in daily_roi[label].values() if b["ws"]>0]
        if len(vals) < 3: return None
        mean = sum(vals)/len(vals)
        std  = (sum((v-mean)**2 for v in vals)/len(vals))**0.5
        return round(std/mean*100, 1) if mean > 0 else None

    # 花费门槛按天数等比缩放
    SPEND_MIN_GROUP = max(50,  int(400 * compare_days / 7))
    SPEND_MIN_CAMP  = max(20,  int(150 * compare_days / 7))

    signals = []
    for key, b in agg.items():
        is_camp   = "|||" in key
        spend_min = SPEND_MIN_CAMP if is_camp else SPEND_MIN_GROUP
        if b["r_spend"] < spend_min: continue

        r_roi = b["r_wsr"] / b["r_ws"] if b["r_ws"] > 0 else None
        p_roi = b["p_wsr"] / b["p_ws"] if b["p_ws"] > 0 else None
        chg   = round((r_roi - p_roi) / p_roi * 100, 1) if (r_roi and p_roi) else None

        if is_camp:
            group_label, label = key.split("|||", 1)
        else:
            group_label = label = key

        cv = daily_cv(group_label) if (not is_camp and compare_days >= 7) else None

        sig = None
        if cv is not None and cv > 60:
            sig = {"level":"anomaly",      "tag":"日内ROI波动异常",
                   "action":"ROI日间波动剧烈，建议核查数据或投放稳定性","priority":2}
        elif chg is not None and chg <= -30:
            sig = {"level":"critical",     "tag":"ROI大幅下滑",
                   "action":"建议大幅削减预算或暂停，排查素材/受众/竞争变化","priority":1}
        elif chg is not None and chg <= -15:
            sig = {"level":"warning",      "tag":"ROI明显下滑",
                   "action":"建议适当降低预算，持续观察 2-3 天","priority":2}
        elif r_roi is not None and r_roi < overall_roi*0.65 and b["r_spend"] > spend_min*2:
            sig = {"level":"warning",      "tag":"ROI持续偏低",
                   "action":"ROI低于整体均值35%以上，建议审视投放策略","priority":3}
        elif chg is not None and chg >= 80:
            sig = {"level":"anomaly",      "tag":"ROI异常跳升",
                   "action":"涨幅超80%，可能是数据异常或偶发因素，需二次确认","priority":2}
        elif chg is not None and chg >= 20:
            sig = {"level":"opportunity",  "tag":"ROI显著提升",
                   "action":"表现持续改善，建议增加预算放量","priority":1}
        elif chg is not None and chg >= 10 and r_roi and r_roi > overall_roi:
            sig = {"level":"opportunity",  "tag":"ROI稳步提升",
                   "action":"高于均值且持续提升，可适当加量","priority":2}

        if sig:
            signals.append({
                "key": key, "label": label, "group": group_label,
                "is_campaign": is_camp,
                "r_spend": round(b["r_spend"], 0),
                "r_roi":   round(r_roi, 2) if r_roi else None,
                "p_roi":   round(p_roi, 2) if p_roi else None,
                "chg": chg, "cv": cv, **sig,
            })

    level_order = {"critical":0,"warning":1,"anomaly":2,"opportunity":3}
    signals.sort(key=lambda s: (level_order[s["level"]], s["priority"], -s["r_spend"]))

    meta = {
        "recent":      f"{recent_start.strftime('%m/%d')}–{end_dt.strftime('%m/%d')}",
        "prev":        f"{prev_start.strftime('%m/%d')}–{prev_end.strftime('%m/%d')}",
        "overall_roi": round(overall_roi, 2),
        "n_issues":    sum(1 for s in signals if s["level"] in ("critical","warning")),
        "n_opps":      sum(1 for s in signals if s["level"] == "opportunity"),
        "n_anomalies": sum(1 for s in signals if s["level"] == "anomaly"),
    }
    return signals, meta

def compute_all_signals(records):
    result = {}
    for n in (1, 3, 7):
        sigs, meta = _compute_signals_for_window(records, n)
        result[str(n)] = {"signals": sigs, "meta": meta}
    # 用7天的 overall_roi 作为全局参考
    overall_roi = result["7"]["meta"]["overall_roi"]
    return result, overall_roi

# ─── 颜色 ────────────────────────────────────────────────────────────────────
PROJECT_COLORS = {
    "EU10":    {"bar": "rgba(99,102,241,0.75)",  "line": "#4f46e5"},
    "海托":    {"bar": "rgba(239,68,68,0.75)",   "line": "#dc2626"},
    "欧洲本地": {"bar": "rgba(16,185,129,0.75)", "line": "#059669"},
}
PALETTE = [
    (59,130,246),(245,158,11),(139,92,246),(236,72,153),(20,184,166),
    (249,115,22),(132,204,22),(6,182,212),(168,85,247),(34,197,94),
    (234,179,8),(239,68,68),(14,165,233),(251,191,36),(16,185,129),(100,116,139),
]
def rgb(r,g,b,a=1): return f"rgba({r},{g},{b},{a})"
def country_color(i):
    r,g,b = PALETTE[i % len(PALETTE)]
    return rgb(r,g,b,0.75), rgb(r,g,b,1)

# ─── 构建完整数据包 ──────────────────────────────────────────────────────────
def build_data(records):
    all_countries, all_projects = set(), set()
    for r in records:
        kind, label = classify(r["name"])
        (all_projects if kind == "project" else all_countries).add(label)

    country_list = sorted(all_countries)
    c_colors = {c: country_color(i) for i, c in enumerate(country_list)}
    p_colors = {p: (PROJECT_COLORS.get(p,{"bar":rgb(100,100,100,.75),"line":"#555"})["bar"],
                    PROJECT_COLORS.get(p,{"bar":rgb(100,100,100,.75),"line":"#555"})["line"])
                for p in all_projects}
    all_colors = {**{k:{"bar":bc,"line":lc} for k,(bc,lc) in c_colors.items()},
                  **{k:{"bar":bc,"line":lc} for k,(bc,lc) in p_colors.items()}}

    result = {}
    for n in (7, 14, 30):
        dates  = get_window(records, n)
        labels = fmt_dates(dates)
        proj, country, campaigns = aggregate(records, dates)

        def make_ds(series, colors_map):
            bars, lines = [], []
            for lbl in sorted(series.keys()):
                bc = colors_map.get(lbl,{}).get("bar", rgb(100,100,100,.75))
                lc = colors_map.get(lbl,{}).get("line","#555")
                d  = series[lbl]
                bars.append({"label":lbl,"data":d["spend"],"backgroundColor":bc,
                             "borderColor":lc,"borderWidth":1,"yAxisID":"ySpend",
                             "type":"bar","stack":"spend"})
                lines.append({"label":lbl,"data":d["roi"],"borderColor":lc,
                              "backgroundColor":"transparent","borderWidth":2.5,
                              "pointRadius":4,"pointHoverRadius":6,"tension":0.3,
                              "yAxisID":"yROI","type":"line","spanGaps":True})
            return bars + lines

        def make_camp_ds(group_camps):
            bars, lines = [], []
            for i, (lbl, d) in enumerate(sorted(group_camps.items())):
                r2,g2,b2 = PALETTE[i % len(PALETTE)]
                bc,lc = rgb(r2,g2,b2,0.7), rgb(r2,g2,b2,1)
                bars.append({"label":lbl,"data":d["spend"],"backgroundColor":bc,
                             "borderColor":lc,"borderWidth":1,"yAxisID":"ySpend",
                             "type":"bar","stack":"spend"})
                lines.append({"label":lbl,"data":d["roi"],"borderColor":lc,
                              "backgroundColor":"transparent","borderWidth":2.5,
                              "pointRadius":4,"pointHoverRadius":6,"tension":0.3,
                              "yAxisID":"yROI","type":"line","spanGaps":True})
            return bars + lines

        def summ(series):
            return {lbl: {"spend": round(sum(s for s in d["spend"] if s),0),
                          "roi": round(sum(v for v in d["roi"] if v)/len([v for v in d["roi"] if v]),2)
                                 if any(v for v in d["roi"] if v) else None}
                    for lbl, d in series.items()}

        all_camp_ds = {g: make_camp_ds(gc) for g, gc in campaigns.items()}

        result[str(n)] = {
            "labels":          labels,
            "proj_ds":         make_ds(proj,    all_colors),
            "country_ds":      make_ds(country, all_colors),
            "proj_summary":    summ(proj),
            "country_summary": summ(country),
            "group_summary":   {**summ(proj), **summ(country)},
            "camp_ds":         all_camp_ds,
            "c_colors":        {k:{"bar":bc,"line":lc} for k,(bc,lc) in c_colors.items()},
            "p_colors":        {k:{"bar":bc,"line":lc} for k,(bc,lc) in p_colors.items()},
            "all_colors":      all_colors,
            "proj_raw":        dict(proj),
            "country_raw":     dict(country),
        }
    return result

# ─── HTML ─────────────────────────────────────────────────────────────────────
def generate_html(data_json, all_signals_json, dac_json, generated_at):
    return f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>阿里投放数据看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f1f5f9;color:#1e293b;min-height:100vh}}
.header{{background:linear-gradient(135deg,#1e293b,#334155);color:#fff;padding:18px 32px;display:flex;justify-content:space-between;align-items:center}}
.header h1{{font-size:19px;font-weight:700}}
.header .meta{{font-size:12px;opacity:.6}}
.topbar{{background:#fff;border-bottom:1px solid #e2e8f0;padding:0 32px;display:flex;align-items:center}}
.tab{{padding:13px 22px;cursor:pointer;font-size:14px;font-weight:500;border-bottom:3px solid transparent;color:#64748b;transition:all .15s;white-space:nowrap}}
.tab.active{{border-bottom-color:#4f46e5;color:#4f46e5}}
.tab:hover{{color:#4f46e5}}
.period-bar{{margin-left:auto;display:flex;gap:6px;padding:8px 0;flex-shrink:0}}
.pbtn{{padding:5px 14px;border-radius:99px;border:1.5px solid #e2e8f0;background:#fff;font-size:13px;cursor:pointer;color:#64748b;transition:all .15s}}
.pbtn.active{{border-color:#4f46e5;background:#4f46e5;color:#fff}}
.spbtn{{padding:4px 12px;border-radius:99px;border:1.5px solid #e2e8f0;background:#fff;font-size:12px;cursor:pointer;color:#64748b;transition:all .15s}}
.spbtn.active{{border-color:#0891b2;background:#0891b2;color:#fff}}
.panel{{display:none;padding:24px 32px}}
.panel.active{{display:block}}
.section-title{{font-size:15px;font-weight:600;margin-bottom:14px;color:#0f172a;display:flex;align-items:center;gap:8px}}
/* ── KPI bar ── */
.kpi-bar{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:28px}}
.kpi{{background:#fff;border-radius:12px;padding:16px 20px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
.kpi .k-label{{font-size:12px;color:#64748b;margin-bottom:4px}}
.kpi .k-val{{font-size:24px;font-weight:800;color:#0f172a}}
.kpi .k-sub{{font-size:12px;margin-top:3px}}
.kpi.red .k-val{{color:#dc2626}} .kpi.green .k-val{{color:#059669}}
.kpi.amber .k-val{{color:#d97706}}
/* ── Signal cards ── */
.sig-grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px}}
@media(max-width:900px){{.sig-grid{{grid-template-columns:1fr}}}}
.sig-col-title{{font-size:14px;font-weight:700;margin-bottom:12px;display:flex;align-items:center;gap:6px}}
.sig-list{{display:flex;flex-direction:column;gap:10px}}
.sig-card{{background:#fff;border-radius:10px;padding:14px 16px;border:1px solid #e2e8f0;
           border-left:4px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,.06);
           transition:box-shadow .15s}}
.sig-card:hover{{box-shadow:0 3px 10px rgba(0,0,0,.1)}}
.sig-card.critical{{border-left-color:#dc2626}}
.sig-card.warning{{border-left-color:#f59e0b}}
.sig-card.anomaly{{border-left-color:#7c3aed}}
.sig-card.opportunity{{border-left-color:#059669}}
.sig-top{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px}}
.sig-name{{font-size:14px;font-weight:700;color:#0f172a}}
.sig-tag{{font-size:11px;padding:2px 8px;border-radius:99px;font-weight:600;white-space:nowrap}}
.tag-critical{{background:#fee2e2;color:#b91c1c}}
.tag-warning{{background:#fef3c7;color:#b45309}}
.tag-anomaly{{background:#ede9fe;color:#6d28d9}}
.tag-opportunity{{background:#d1fae5;color:#065f46}}
.sig-meta{{display:flex;gap:14px;font-size:12px;color:#64748b;margin-bottom:6px;flex-wrap:wrap}}
.sig-meta span{{display:flex;align-items:center;gap:3px}}
.roi-arrow{{font-size:13px;font-weight:700}}
.arrow-up{{color:#059669}} .arrow-down{{color:#dc2626}} .arrow-warn{{color:#7c3aed}}
.sig-action{{font-size:12px;color:#475569;background:#f8fafc;border-radius:6px;padding:6px 10px;border:1px solid #e2e8f0}}
.sig-badge-camp{{font-size:10px;background:#f1f5f9;color:#64748b;padding:1px 6px;border-radius:4px;margin-left:4px;font-weight:500}}
/* ── 时间说明 ── */
.period-note{{font-size:12px;color:#94a3b8;margin-bottom:20px}}
/* ── 通用卡片 & 图表 ── */
.cards{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px}}
.card{{background:#fff;border-radius:10px;padding:12px 16px;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.07);border:1px solid #e2e8f0;cursor:pointer;transition:all .15s}}
.card:hover{{border-color:#a5b4fc;box-shadow:0 2px 8px rgba(99,102,241,.15)}}
.card.selected{{border-color:#4f46e5;background:#ede9fe}}
.card .name{{font-size:12px;color:#64748b;margin-bottom:3px}}
.card .spend{{font-size:16px;font-weight:700;color:#0f172a}}
.card .roi{{font-size:12px;color:#6366f1;margin-top:2px;font-weight:500}}
.chart-box{{background:#fff;border-radius:12px;padding:22px;box-shadow:0 1px 3px rgba(0,0,0,.07);border:1px solid #e2e8f0;margin-bottom:22px}}
.chart-hd{{font-size:13px;font-weight:600;color:#374151;margin-bottom:14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.badge{{font-size:11px;background:#ede9fe;color:#4f46e5;padding:2px 8px;border-radius:99px;font-weight:500;white-space:nowrap}}
canvas{{max-height:340px}}
.note{{font-size:11px;color:#94a3b8;margin-top:6px}}
.clear-btn{{font-size:12px;color:#6366f1;cursor:pointer;text-decoration:underline}}
.empty-hint{{color:#94a3b8;font-size:14px;padding:40px;text-align:center;background:#fff;border-radius:12px;border:1px dashed #e2e8f0}}
.group-chips{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px}}
.chip{{padding:6px 14px;border-radius:99px;border:1.5px solid #e2e8f0;background:#fff;font-size:13px;font-weight:500;cursor:pointer;color:#475569;transition:all .15s}}
.chip:hover{{border-color:#a5b4fc;color:#4f46e5}}
.chip.selected{{border-color:#4f46e5;background:#4f46e5;color:#fff}}
.chip.proj{{border-color:#fecaca}}
.chip.proj.selected{{background:#dc2626;border-color:#dc2626;color:#fff}}
.no-signals{{color:#94a3b8;font-size:13px;padding:20px;text-align:center;background:#f8fafc;border-radius:10px;border:1px dashed #e2e8f0}}
</style>
</head>
<body>

<!-- ══ 密码保护 ══ -->
<div id="lock-screen" style="
  position:fixed;inset:0;z-index:9999;
  background:linear-gradient(135deg,#1e293b,#334155);
  display:flex;align-items:center;justify-content:center;
">
  <div style="background:#fff;border-radius:16px;padding:40px 48px;
              box-shadow:0 20px 60px rgba(0,0,0,.4);text-align:center;width:320px">
    <div style="font-size:32px;margin-bottom:8px">🔒</div>
    <div style="font-size:18px;font-weight:700;color:#0f172a;margin-bottom:4px">投放数据看板</div>
    <div style="font-size:13px;color:#64748b;margin-bottom:24px">AliExpress Moloco</div>
    <input id="pwd-input" type="password" placeholder="请输入访问密码"
      style="width:100%;padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;
             font-size:14px;outline:none;box-sizing:border-box;margin-bottom:12px"
      onkeydown="if(event.key==='Enter')checkPwd()"
    >
    <div id="pwd-err" style="color:#dc2626;font-size:12px;margin-bottom:10px;display:none">密码错误，请重试</div>
    <button onclick="checkPwd()" style="
      width:100%;padding:10px;background:#4f46e5;color:#fff;border:none;
      border-radius:8px;font-size:14px;font-weight:600;cursor:pointer">
      进入
    </button>
  </div>
</div>

<script>
(function(){{
  const PWD = "moloco2026";
  const KEY = "ali_dash_auth";
  if (sessionStorage.getItem(KEY) === "1") {{
    document.getElementById("lock-screen").style.display = "none";
  }}
  window.checkPwd = function() {{
    const val = document.getElementById("pwd-input").value;
    if (val === PWD) {{
      sessionStorage.setItem(KEY, "1");
      document.getElementById("lock-screen").style.display = "none";
    }} else {{
      document.getElementById("pwd-err").style.display = "block";
      document.getElementById("pwd-input").value = "";
      document.getElementById("pwd-input").focus();
    }}
  }};
  setTimeout(() => document.getElementById("pwd-input").focus(), 100);
}})();
</script>

<div class="header">
  <h1>AliExpress Moloco 投放数据看板</h1>
  <div class="meta">数据更新: {generated_at}</div>
</div>

<div class="topbar">
  <div class="tab active" onclick="switchTab('home')">首页</div>
  <div class="tab" onclick="switchTab('project')">项目视图</div>
  <div class="tab" onclick="switchTab('country')">国家视图</div>
  <div class="tab" onclick="switchTab('campaign')">Campaign 视图</div>
  <div class="tab" onclick="switchTab('dac')">DAC成本</div>
  <div class="period-bar">
    <button class="pbtn" onclick="setPeriod(7)">过去 7 天</button>
    <button class="pbtn active" onclick="setPeriod(14)">过去 14 天</button>
    <button class="pbtn" onclick="setPeriod(30)">过去 30 天</button>
  </div>
</div>

<!-- ══ 首页 ══ -->
<div id="panel-home" class="panel active">
  <div id="kpi-bar" class="kpi-bar"></div>
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap">
    <span style="font-size:13px;color:#64748b;font-weight:500">信号对比：</span>
    <button class="spbtn" onclick="setSigPeriod(1)">对比昨天</button>
    <button class="spbtn" onclick="setSigPeriod(3)">对比最近3天</button>
    <button class="spbtn active" onclick="setSigPeriod(7)">对比最近7天</button>
    <span class="period-note" id="period-note" style="margin:0;font-size:12px"></span>
  </div>
  <div class="sig-grid">
    <div>
      <div class="sig-col-title">⚠️ 需要关注 &amp; 优化</div>
      <div id="sig-issues" class="sig-list"></div>
    </div>
    <div>
      <div class="sig-col-title">📈 建议放量</div>
      <div id="sig-opps" class="sig-list"></div>
    </div>
  </div>
  <div>
    <div class="sig-col-title">⚡ 异常波动</div>
    <div id="sig-anomalies" class="sig-list" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:10px"></div>
  </div>
</div>

<!-- ══ 项目面板 ══ -->
<div id="panel-project" class="panel">
  <div class="section-title">
    项目汇总
    <span id="clear-project" class="clear-btn" style="display:none" onclick="selectProject(null)">✕ 清除筛选</span>
  </div>
  <div id="proj-cards" class="cards"></div>
  <div class="chart-box">
    <div class="chart-hd" id="proj-chart-title">项目日消耗（柱）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span></div>
    <canvas id="projChart"></canvas>
    <div class="note">点击项目卡片单独查看；再次点击或点 ✕ 取消筛选</div>
  </div>
</div>

<!-- ══ 国家面板 ══ -->
<div id="panel-country" class="panel">
  <div class="section-title">
    国家汇总
    <span id="clear-country" class="clear-btn" style="display:none" onclick="selectCountry(null)">✕ 清除筛选</span>
  </div>
  <div id="country-cards" class="cards"></div>
  <div class="chart-box">
    <div class="chart-hd" id="country-chart-title">国家日消耗（柱堆叠）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span></div>
    <canvas id="countryChart"></canvas>
    <div class="note">点击国家卡片单独查看；再次点击或点 ✕ 取消筛选</div>
  </div>
  <div id="proj-ref-box" class="chart-box" style="display:none">
    <div class="chart-hd">项目整体趋势（参考）<span class="badge">同期</span></div>
    <canvas id="projRefChart"></canvas>
  </div>
</div>

<!-- ══ Campaign 面板 ══ -->
<div id="panel-campaign" class="panel">
  <div class="section-title">
    选择国家 / 项目
    <span id="clear-camp" class="clear-btn" style="display:none" onclick="selectGroup(null)">✕ 清除</span>
  </div>
  <div id="group-chips" class="group-chips"></div>
  <div id="camp-chart-area">
    <div class="empty-hint">👆 选择上方国家或项目，查看各 Campaign 趋势</div>
  </div>
</div>

<!-- ══ DAC成本 面板 ══ -->
<div id="panel-dac" class="panel">
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;flex-wrap:wrap">
    <span style="font-size:13px;color:#64748b;font-weight:500">数据窗口：</span>
    <button class="spbtn active" id="dac-btn-3"  onclick="setDacPeriod(3)">近3天</button>
    <button class="spbtn"        id="dac-btn-7"  onclick="setDacPeriod(7)">近7天</button>
    <button class="spbtn"        id="dac-btn-14" onclick="setDacPeriod(14)">近14天</button>
    <span style="font-size:12px;color:#94a3b8;margin-left:4px">数据源：最新 sheet · dac成本字段（24h_dac口径）</span>
  </div>
  <div class="section-title">
    各国 DAC成本（窗口合计）
    <span id="clear-dac-country" class="clear-btn" style="display:none" onclick="selectDacCountry(null)">✕ 清除</span>
  </div>
  <div id="dac-country-cards" class="cards"></div>
  <div id="dac-chart-area">
    <div class="empty-hint">👆 点击国家 / 项目卡片，查看日趋势和 Campaign 明细</div>
  </div>

  <div class="section-title" style="margin-top:32px">
    DAC 专项（campaign name 含「DAC专项」）
    <span id="clear-dac-special" class="clear-btn" style="display:none" onclick="selectDacSpecial(null)">✕ 清除</span>
  </div>
  <div id="dac-special-cards" class="cards"></div>
  <div id="dac-special-chart-area">
    <div class="empty-hint">👆 点击「整体」或单个 campaign 卡片，查看 DAC数 / 花费 / DAC成本 分天趋势</div>
  </div>
</div>

<script>
const ALL         = {data_json};
const ALL_SIGNALS = {all_signals_json};
const DAC_DATA    = {dac_json};

const DAC_PALETTE = [
  [59,130,246],[245,158,11],[139,92,246],[236,72,153],[20,184,166],
  [249,115,22],[132,204,22],[6,182,212],[168,85,247],[34,197,94],
  [234,179,8],[239,68,68],[14,165,233],[251,191,36],[16,185,129],[100,116,139],
];

let period     = 14;
let sigPeriod  = 7;  // 信号对比窗口: 1 / 3 / 7
let dacPeriod      = 3;
let activeTab      = "home";
let selProject     = null;
let selCountry     = null;
let selGroup       = null;
let selDacCountry  = null;
let selDacSpecial  = null;

let projChart=null, countryChart=null, projRefChart=null, campChart=null;
let dacTrendChart=null, dacCampChart=null;
let dacSpecSpendChart=null, dacSpecCostChart=null;

if (typeof ChartDataLabels !== "undefined") Chart.register(ChartDataLabels);

// ── Chart factory ────────────────────────────────────────────────────────────
function makeChart(id, datasets, labels, existing) {{
  if (existing) existing.destroy();
  return new Chart(document.getElementById(id), {{
    type:"bar", data:{{labels,datasets}},
    options:{{
      responsive:true, interaction:{{mode:"index",intersect:false}},
      scales:{{
        x:{{grid:{{color:"rgba(0,0,0,.05)"}},ticks:{{font:{{size:12}}}}}},
        ySpend:{{type:"linear",position:"left",
          title:{{display:true,text:"消耗 (USD)",font:{{size:12}}}},
          grid:{{color:"rgba(0,0,0,.06)"}},
          ticks:{{callback:v=>"$"+v.toLocaleString(),font:{{size:11}}}}}},
        yROI:{{type:"linear",position:"right",
          title:{{display:true,text:"ROI (x)",font:{{size:12}}}},
          grid:{{drawOnChartArea:false}},
          ticks:{{callback:v=>v+"x",font:{{size:11}}}},min:0}},
      }},
      plugins:{{
        legend:{{position:"top",labels:{{font:{{size:12}},usePointStyle:true,pointStyleWidth:12}}}},
        tooltip:{{mode:"index",intersect:false,
          callbacks:{{label:ctx=>{{
            const v=ctx.parsed.y;
            if(v===null||v===undefined) return null;
            return ctx.dataset.yAxisID==="ySpend"
              ? ctx.dataset.label+": $"+v.toLocaleString()
              : ctx.dataset.label+" ROI: "+v+"x";
          }}}}}},
        datalabels:{{
          display: ctx => ctx.dataset.yAxisID === "yROI",
          formatter: (v, ctx) => {{
            const val = ctx.dataset.data[ctx.dataIndex];
            return val != null ? val + "x" : null;
          }},
          color: ctx => ctx.dataset.borderColor || "#333",
          font: {{ size:11, weight:"700" }},
          anchor:"end",
          align:"top",
          offset:3,
          backgroundColor:"rgba(255,255,255,0.82)",
          borderRadius:3,
          padding:{{top:2,bottom:2,left:4,right:4}},
        }},
      }},
    }},
  }});
}}

// ── 首页 ─────────────────────────────────────────────────────────────────────
function setSigPeriod(n) {{
  sigPeriod = n;
  document.querySelectorAll(".spbtn").forEach((b,i)=>
    b.classList.toggle("active",[1,3,7][i]===n));
  renderSignals();
}}

function renderHome() {{
  const d = ALL[String(period)];
  const meta7 = ALL_SIGNALS["7"].meta;
  const allSpend = Object.values(d.group_summary).reduce((s,x)=>s+(x.spend||0),0);
  const dailyAvg = (allSpend / d.labels.length).toFixed(0);
  const issues7   = ALL_SIGNALS["7"].signals.filter(s=>["critical","warning"].includes(s.level));
  const opps7     = ALL_SIGNALS["7"].signals.filter(s=>s.level==="opportunity");
  const anomalies7= ALL_SIGNALS["7"].signals.filter(s=>s.level==="anomaly");

  document.getElementById("kpi-bar").innerHTML = `
    <div class="kpi"><div class="k-label">近 ${{period}} 天总消耗</div>
      <div class="k-val">$${{allSpend.toLocaleString()}}</div>
      <div class="k-sub" style="color:#64748b">USD</div></div>
    <div class="kpi"><div class="k-label">日均消耗</div>
      <div class="k-val">$${{Number(dailyAvg).toLocaleString()}}</div>
      <div class="k-sub" style="color:#64748b">USD / 天</div></div>
    <div class="kpi"><div class="k-label">整体平均 ROI</div>
      <div class="k-val">${{meta7.overall_roi}}x</div>
      <div class="k-sub" style="color:#64748b">近7天加权</div></div>
    <div class="kpi red"><div class="k-label">需关注项</div>
      <div class="k-val">${{issues7.length}}</div>
      <div class="k-sub" style="color:#dc2626">ROI下滑 / 持续偏低</div></div>
    <div class="kpi green"><div class="k-label">可放量项</div>
      <div class="k-val">${{opps7.length}}</div>
      <div class="k-sub" style="color:#059669">ROI显著提升</div></div>
    <div class="kpi amber"><div class="k-label">异常波动</div>
      <div class="k-val">${{anomalies7.length}}</div>
      <div class="k-sub" style="color:#d97706">需核实数据</div></div>`;

  renderSignals();
}}

function renderSignals() {{
  const sd   = ALL_SIGNALS[String(sigPeriod)];
  const sigs = sd.signals;
  const meta = sd.meta;
  const labelMap = {{1:"昨天", 3:"最近3天", 7:"最近7天"}};

  const issues    = sigs.filter(s=>["critical","warning"].includes(s.level));
  const opps      = sigs.filter(s=>s.level==="opportunity");
  const anomalies = sigs.filter(s=>s.level==="anomaly");

  document.getElementById("period-note").innerHTML =
    `📅 对比：<strong>${{meta.recent}}</strong> vs 前期 <strong>${{meta.prev}}</strong>`;

  function buildCard(s) {{
    const chgStr = s.chg != null
      ? (s.chg>0 ? `<span class="roi-arrow arrow-up">▲${{s.chg}}%</span>`
                 : `<span class="roi-arrow arrow-down">▼${{Math.abs(s.chg)}}%</span>`)
      : (s.cv != null ? `<span class="roi-arrow arrow-warn">波动 CV=${{s.cv}}%</span>` : "");
    const campBadge = s.is_campaign ? `<span class="sig-badge-camp">campaign</span>` : "";
    return `<div class="sig-card ${{s.level}}">
      <div class="sig-top">
        <div class="sig-name">${{s.label}}${{campBadge}}</div>
        <span class="sig-tag tag-${{s.level}}">${{s.tag}}</span>
      </div>
      <div class="sig-meta">
        <span>💰 $${{s.r_spend.toLocaleString()}}</span>
        ${{s.r_roi ? `<span>🎯 ROI: ${{s.r_roi}}x</span>` : ""}}
        ${{s.p_roi ? `<span>📊 前期: ${{s.p_roi}}x</span>` : ""}}
        <span>${{chgStr}}</span>
      </div>
      <div class="sig-action">💡 ${{s.action}}</div>
    </div>`;
  }}

  document.getElementById("sig-issues").innerHTML =
    issues.length ? issues.map(buildCard).join("") :
    `<div class="no-signals">✅ 暂无需要关注的项目</div>`;
  document.getElementById("sig-opps").innerHTML =
    opps.length ? opps.map(buildCard).join("") :
    `<div class="no-signals">暂无显著放量机会</div>`;
  document.getElementById("sig-anomalies").innerHTML =
    anomalies.length ? anomalies.map(buildCard).join("") :
    `<div class="no-signals">暂无异常波动</div>`;
}}

// ── 项目面板 ────────────────────────────────────────────────────────────────
function renderProject() {{
  const d = ALL[String(period)];
  const cards = document.getElementById("proj-cards");
  cards.innerHTML = "";
  const days = d.labels.length;
  Object.entries(d.proj_summary).sort((a,b)=>(b[1].spend||0)-(a[1].spend||0))
    .forEach(([lbl,s])=>{{
      const daily = days > 0 ? Math.round((s.spend||0)/days).toLocaleString() : "-";
      const div=document.createElement("div");
      div.className="card"+(selProject===lbl?" selected":"");
      div.innerHTML=`<div class="name">${{lbl}}</div>
        <div class="spend">$${{(s.spend||0).toLocaleString()}}</div>
        <div class="roi">日均: $${{daily}} · ROI: ${{s.roi!=null?s.roi+"x":"-"}}</div>`;
      div.onclick=()=>selectProject(lbl===selProject?null:lbl);
      cards.appendChild(div);
    }});

  const titleEl = document.getElementById("proj-chart-title");
  if (selProject && d.proj_raw && d.proj_raw[selProject]) {{
    const raw = d.proj_raw[selProject];
    const pc  = (d.p_colors||{{}})[selProject] || {{bar:"rgba(99,102,241,.75)",line:"#4f46e5"}};
    const ds  = [
      {{label:selProject, data:raw.spend, backgroundColor:pc.bar, borderColor:pc.line,
        borderWidth:1, yAxisID:"ySpend", type:"bar", stack:"spend"}},
      {{label:selProject+" ROI", data:raw.roi, borderColor:pc.line, backgroundColor:"transparent",
        borderWidth:2.5, pointRadius:4, pointHoverRadius:6, tension:0.3,
        yAxisID:"yROI", type:"line", spanGaps:true}},
    ];
    if (titleEl) titleEl.innerHTML=`${{selProject}} 日消耗（柱）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span>`;
    document.getElementById("clear-project").style.display="inline";
    projChart = makeChart("projChart", ds, d.labels, projChart);
  }} else {{
    if (titleEl) titleEl.innerHTML=`项目日消耗（柱）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span>`;
    document.getElementById("clear-project").style.display="none";
    projChart = makeChart("projChart", d.proj_ds, d.labels, projChart);
  }}
}}

// ── 国家面板 ────────────────────────────────────────────────────────────────
function renderCountry() {{
  const d = ALL[String(period)];
  const cards = document.getElementById("country-cards");
  cards.innerHTML = "";
  const days = d.labels.length;
  Object.entries(d.country_summary).sort((a,b)=>(b[1].spend||0)-(a[1].spend||0))
    .forEach(([lbl,s])=>{{
      const daily = days > 0 ? Math.round((s.spend||0)/days).toLocaleString() : "-";
      const div=document.createElement("div");
      div.className="card"+(selCountry===lbl?" selected":"");
      div.innerHTML=`<div class="name">${{lbl}}</div>
        <div class="spend">$${{(s.spend||0).toLocaleString()}}</div>
        <div class="roi">日均: $${{daily}} · ROI: ${{s.roi!=null?s.roi+"x":"-"}}</div>`;
      div.onclick=()=>selectCountry(lbl===selCountry?null:lbl);
      cards.appendChild(div);
    }});

  if (selCountry && d.country_raw[selCountry]) {{
    const raw=d.country_raw[selCountry];
    const cc=d.c_colors[selCountry]||{{bar:"rgba(99,102,241,.75)",line:"#4f46e5"}};
    const ds=[
      {{label:selCountry,data:raw.spend,backgroundColor:cc.bar,borderColor:cc.line,
        borderWidth:1,yAxisID:"ySpend",type:"bar",stack:"spend"}},
      {{label:selCountry+" ROI",data:raw.roi,borderColor:cc.line,backgroundColor:"transparent",
        borderWidth:2.5,pointRadius:4,pointHoverRadius:6,tension:0.3,
        yAxisID:"yROI",type:"line",spanGaps:true}},
    ];
    document.getElementById("country-chart-title").innerHTML=
      `${{selCountry}} 日消耗（柱）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span>`;
    countryChart=makeChart("countryChart",ds,d.labels,countryChart);
    document.getElementById("clear-country").style.display="inline";
    document.getElementById("proj-ref-box").style.display="block";
    projRefChart=makeChart("projRefChart",d.proj_ds,d.labels,projRefChart);
  }} else {{
    document.getElementById("country-chart-title").innerHTML=
      `国家日消耗（柱堆叠）& ROI（线）<span class="badge">左轴: 消耗 USD · 右轴: ROI</span>`;
    countryChart=makeChart("countryChart",d.country_ds,d.labels,countryChart);
    document.getElementById("clear-country").style.display="none";
    document.getElementById("proj-ref-box").style.display="none";
  }}
}}

// ── Campaign 面板 ────────────────────────────────────────────────────────────
function renderCampaign() {{
  const d=ALL[String(period)];
  const chips=document.getElementById("group-chips");
  chips.innerHTML="";
  Object.keys(d.proj_summary).sort().forEach(lbl=>chips.appendChild(makeChip(lbl,d.proj_summary[lbl],true)));
  Object.entries(d.country_summary).sort((a,b)=>(b[1].spend||0)-(a[1].spend||0))
    .forEach(([lbl,s])=>chips.appendChild(makeChip(lbl,s,false)));

  const area=document.getElementById("camp-chart-area");
  if (!selGroup||!d.camp_ds[selGroup]) {{
    area.innerHTML=`<div class="empty-hint">👆 选择上方国家或项目，查看各 Campaign 趋势</div>`;
    campChart=null;
    document.getElementById("clear-camp").style.display="none";
    return;
  }}
  document.getElementById("clear-camp").style.display="inline";
  area.innerHTML=`<div class="chart-box">
    <div class="chart-hd">${{selGroup}} — Campaign 日消耗（柱）& ROI（线）
      <span class="badge">左轴: 消耗 USD · 右轴: ROI</span></div>
    <canvas id="campChart"></canvas>
    <div class="note">图例可点击隐藏/显示单个 Campaign</div>
  </div>`;
  campChart=makeChart("campChart",d.camp_ds[selGroup],d.labels,null);
}}

function makeChip(lbl,s,isProj) {{
  const chip=document.createElement("div");
  chip.className="chip"+(isProj?" proj":"")+(selGroup===lbl?" selected":"");
  const spend=(s.spend||0)>0?"  $"+(s.spend||0).toLocaleString():"";
  chip.innerHTML=`<strong>${{lbl}}</strong>${{spend}}`;
  chip.title=`均 ROI: ${{s.roi!=null?s.roi+"x":"-"}}`;
  chip.onclick=()=>selectGroup(lbl===selGroup?null:lbl);
  return chip;
}}

// ── DAC成本 面板 ─────────────────────────────────────────────────────────────
function setDacPeriod(n) {{
  dacPeriod = n;
  [3, 7, 14].forEach(d => {{
    const btn = document.getElementById('dac-btn-' + d);
    if (btn) btn.classList.toggle('active', d === n);
  }});
  renderDac();
}}

function selectDacCountry(c) {{
  selDacCountry = c;
  renderDac();
}}

function selectDacSpecial(k) {{
  selDacSpecial = k;
  renderDacSpecial();
}}

function renderDac() {{
  if (!DAC_DATA || !DAC_DATA[String(dacPeriod)]) return;
  renderDacSpecial();
  const d       = DAC_DATA[String(dacPeriod)];
  const summary = d.country_summary || {{}};
  const labels  = d.labels || [];

  // 国家 / 项目卡片
  const cardsEl = document.getElementById('dac-country-cards');
  if (cardsEl) {{
    cardsEl.innerHTML = '';
    Object.entries(summary)
      .sort((a, b) => a[1].dac_cost - b[1].dac_cost)
      .forEach(([c, s]) => {{
        const sel = selDacCountry === c;
        const div = document.createElement('div');
        div.className = 'card' + (sel ? ' selected' : '');
        div.innerHTML = `<div class="name">${{c}}</div>
          <div class="spend">$${{s.dac_cost}}</div>
          <div class="roi">日均消耗: $${{(s.daily_spend||0).toLocaleString()}} · 总花费: $${{s.cost.toLocaleString()}}</div>
          <div class="roi">DAC: ${{s.dac}}</div>`;
        div.onclick = () => selectDacCountry(c === selDacCountry ? null : c);
        cardsEl.appendChild(div);
      }});
  }}
  const clearBtn = document.getElementById('clear-dac-country');
  if (clearBtn) clearBtn.style.display = selDacCountry ? 'inline' : 'none';

  // 图表区
  const area = document.getElementById('dac-chart-area');
  if (!selDacCountry || !(d.country_daily || {{}})[selDacCountry]) {{
    area.innerHTML = '<div class="empty-hint">👆 点击国家 / 项目卡片，查看日趋势和 Campaign 明细</div>';
    if (dacTrendChart) {{ dacTrendChart.destroy(); dacTrendChart = null; }}
    if (dacCampChart)  {{ dacCampChart.destroy();  dacCampChart  = null; }}
    return;
  }}

  area.innerHTML = `
    <div class="chart-box">
      <div class="chart-hd">${{selDacCountry}} — DAC成本日趋势
        <span class="badge">dac成本字段 · 24h_dac口径</span></div>
      <canvas id="dacTrendChart"></canvas>
      <div class="note">数据来源：dac成本字段（costdsp / 24h_dac）</div>
    </div>
    <div class="chart-box">
      <div class="chart-hd">${{selDacCountry}} — Campaign DAC成本
        <span class="badge">分 Campaign 分天</span></div>
      <canvas id="dacCampChart"></canvas>
      <div class="note">图例可点击隐藏 / 显示单个 Campaign</div>
    </div>`;

  // 国家日趋势
  const cData = d.country_daily[selDacCountry];
  if (dacTrendChart) dacTrendChart.destroy();
  dacTrendChart = new Chart(document.getElementById('dacTrendChart'), {{
    type: 'line',
    data: {{ labels, datasets: [{{
      label: selDacCountry + ' DAC成本',
      data: cData,
      borderColor: '#4f46e5',
      backgroundColor: 'rgba(79,70,229,0.07)',
      fill: true,
      borderWidth: 2.5, pointRadius: 5, pointHoverRadius: 7,
      tension: 0.3, spanGaps: true,
    }}] }},
    options: {{
      responsive: true,
      interaction: {{ mode: 'index', intersect: false }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,.05)' }}, ticks: {{ font: {{ size: 12 }} }} }},
        y: {{
          title: {{ display: true, text: 'DAC成本 (USD)', font: {{ size: 12 }} }},
          ticks: {{ callback: v => '$' + v, font: {{ size: 11 }} }}, min: 0,
        }},
      }},
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': $' + ctx.parsed.y }} }},
        datalabels: {{
          display: true,
          formatter: v => v != null ? '$' + v : null,
          color: '#4f46e5', font: {{ size: 11, weight: '700' }},
          anchor: 'top', align: 'top', offset: 3,
          backgroundColor: 'rgba(255,255,255,0.85)', borderRadius: 3,
          padding: {{ top: 2, bottom: 2, left: 4, right: 4 }},
        }},
      }},
    }},
  }});

  // Campaign 明细
  const campData = ((d.camp_daily || {{}})[selDacCountry]) || {{}};
  const campDs = Object.entries(campData).map(([camp, vals], i) => {{
    const [rv, gv, bv] = DAC_PALETTE[i % DAC_PALETTE.length];
    return {{
      label: camp, data: vals,
      borderColor: `rgb(${{rv}},${{gv}},${{bv}})`,
      backgroundColor: 'transparent',
      borderWidth: 2, pointRadius: 4, pointHoverRadius: 6,
      tension: 0.3, spanGaps: true,
    }};
  }});
  if (dacCampChart) dacCampChart.destroy();
  dacCampChart = new Chart(document.getElementById('dacCampChart'), {{
    type: 'line',
    data: {{ labels, datasets: campDs }},
    options: {{
      responsive: true,
      interaction: {{ mode: 'index', intersect: false }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,.05)' }}, ticks: {{ font: {{ size: 12 }} }} }},
        y: {{
          title: {{ display: true, text: 'DAC成本 (USD)', font: {{ size: 12 }} }},
          ticks: {{ callback: v => '$' + v, font: {{ size: 11 }} }}, min: 0,
        }},
      }},
      plugins: {{
        legend: {{ position: 'top', labels: {{ font: {{ size: 12 }}, usePointStyle: true, pointStyleWidth: 12 }} }},
        tooltip: {{ callbacks: {{ label: ctx => {{
          const v = ctx.parsed.y;
          return v != null ? ctx.dataset.label + ': $' + v : null;
        }} }} }},
        datalabels: {{
          display: true,
          formatter: v => v != null ? '$' + v : null,
          color: ctx => ctx.dataset.borderColor || '#333',
          font: {{ size: 10, weight: '700' }},
          anchor: 'top', align: 'top', offset: 2,
          backgroundColor: 'rgba(255,255,255,0.82)', borderRadius: 3,
          padding: {{ top: 1, bottom: 1, left: 3, right: 3 }},
        }},
      }},
    }},
  }});
}}

// ── DAC 专项板块 ─────────────────────────────────────────────────────────────
function renderDacSpecial() {{
  if (!DAC_DATA || !DAC_DATA[String(dacPeriod)]) return;
  const d  = DAC_DATA[String(dacPeriod)];
  const sp = d.dac_special || {{}};
  const labels = d.labels || [];
  const cardsEl = document.getElementById('dac-special-cards');
  const area    = document.getElementById('dac-special-chart-area');
  if (!cardsEl || !area) return;

  cardsEl.innerHTML = '';
  if (!sp.overall) {{
    cardsEl.innerHTML = '<div class="empty-hint" style="width:100%">当前窗口暂无 DAC 专项数据</div>';
    area.innerHTML = '';
    if (dacSpecSpendChart) {{ dacSpecSpendChart.destroy(); dacSpecSpendChart = null; }}
    if (dacSpecCostChart)  {{ dacSpecCostChart.destroy();  dacSpecCostChart  = null; }}
    const clr = document.getElementById('clear-dac-special');
    if (clr) clr.style.display = 'none';
    return;
  }}

  // 整体卡片（高亮金色边框区分）
  const ov = sp.overall;
  const ovSel = selDacSpecial === '__overall__';
  const ovDiv = document.createElement('div');
  ovDiv.className = 'card' + (ovSel ? ' selected' : '');
  ovDiv.style.borderColor = ovSel ? '#4f46e5' : '#fbbf24';
  ovDiv.style.background  = ovSel ? '#ede9fe' : '#fffbeb';
  ovDiv.innerHTML = `<div class="name">🎯 整体</div>
    <div class="spend">$${{ov.dac_cost ?? '-'}}</div>
    <div class="roi">日均消耗: $${{(ov.daily_spend||0).toLocaleString()}} · 总花费: $${{ov.cost.toLocaleString()}}</div>
    <div class="roi">DAC: ${{ov.dac}}</div>`;
  ovDiv.onclick = () => selectDacSpecial(ovSel ? null : '__overall__');
  cardsEl.appendChild(ovDiv);

  // 各 campaign 卡片
  const camps = sp.campaigns || {{}};
  Object.entries(camps)
    .sort((a, b) => (b[1].cost || 0) - (a[1].cost || 0))
    .forEach(([camp, s]) => {{
      const sel = selDacSpecial === camp;
      const div = document.createElement('div');
      div.className = 'card' + (sel ? ' selected' : '');
      div.innerHTML = `<div class="name">${{camp}}</div>
        <div class="spend">$${{s.dac_cost ?? '-'}}</div>
        <div class="roi">日均消耗: $${{(s.daily_spend||0).toLocaleString()}} · 总花费: $${{s.cost.toLocaleString()}}</div>
        <div class="roi">DAC: ${{s.dac}}</div>`;
      div.onclick = () => selectDacSpecial(sel ? null : camp);
      cardsEl.appendChild(div);
    }});

  const clearBtn = document.getElementById('clear-dac-special');
  if (clearBtn) clearBtn.style.display = selDacSpecial ? 'inline' : 'none';

  // 趋势图
  if (!selDacSpecial) {{
    area.innerHTML = '<div class="empty-hint">👆 点击「整体」或单个 campaign 卡片，查看 DAC数 / 花费 / DAC成本 分天趋势</div>';
    if (dacSpecSpendChart) {{ dacSpecSpendChart.destroy(); dacSpecSpendChart = null; }}
    if (dacSpecCostChart)  {{ dacSpecCostChart.destroy();  dacSpecCostChart  = null; }}
    return;
  }}
  const target = selDacSpecial === '__overall__' ? ov : (camps[selDacSpecial] || null);
  if (!target) return;
  const title = selDacSpecial === '__overall__' ? 'DAC 专项整体' : selDacSpecial;

  area.innerHTML = `
    <div class="chart-box">
      <div class="chart-hd">${{title}} — 花费 & DAC 数
        <span class="badge">左轴: 花费 USD · 右轴: DAC 数</span></div>
      <canvas id="dacSpecSpendChart"></canvas>
    </div>
    <div class="chart-box">
      <div class="chart-hd">${{title}} — DAC 成本
        <span class="badge">花费 / DAC（按天）</span></div>
      <canvas id="dacSpecCostChart"></canvas>
    </div>`;

  if (dacSpecSpendChart) dacSpecSpendChart.destroy();
  dacSpecSpendChart = new Chart(document.getElementById('dacSpecSpendChart'), {{
    type: 'bar',
    data: {{ labels, datasets: [
      {{label:'花费', data: target.spend_series, backgroundColor:'rgba(79,70,229,0.7)',
        borderColor:'#4f46e5', borderWidth:1, yAxisID:'ySpend', type:'bar'}},
      {{label:'DAC 数', data: target.dac_series, borderColor:'#f59e0b',
        backgroundColor:'transparent', borderWidth:2.5, pointRadius:4, pointHoverRadius:6,
        tension:0.3, yAxisID:'yDac', type:'line', spanGaps:true}},
    ]}},
    options: {{
      responsive:true, interaction:{{mode:'index',intersect:false}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{font:{{size:12}}}}}},
        ySpend:{{type:'linear',position:'left',
          title:{{display:true,text:'花费 (USD)',font:{{size:12}}}},
          grid:{{color:'rgba(0,0,0,.06)'}},
          ticks:{{callback:v=>'$'+v,font:{{size:11}}}}, min:0}},
        yDac:{{type:'linear',position:'right',
          title:{{display:true,text:'DAC 数',font:{{size:12}}}},
          grid:{{drawOnChartArea:false}},
          ticks:{{font:{{size:11}}}}, min:0}},
      }},
      plugins:{{
        legend:{{position:'top',labels:{{font:{{size:12}},usePointStyle:true,pointStyleWidth:12}}}},
        tooltip:{{mode:'index',intersect:false,
          callbacks:{{label:ctx=>{{
            const v=ctx.parsed.y;
            if(v===null||v===undefined) return null;
            return ctx.dataset.yAxisID==='ySpend'
              ? ctx.dataset.label+': $'+v.toLocaleString()
              : ctx.dataset.label+': '+v;
          }}}}}},
        datalabels:{{display:false}},
      }},
    }},
  }});

  if (dacSpecCostChart) dacSpecCostChart.destroy();
  dacSpecCostChart = new Chart(document.getElementById('dacSpecCostChart'), {{
    type: 'line',
    data: {{ labels, datasets: [{{
      label: 'DAC 成本', data: target.cost_series,
      borderColor:'#4f46e5', backgroundColor:'rgba(79,70,229,0.07)', fill:true,
      borderWidth:2.5, pointRadius:5, pointHoverRadius:7, tension:0.3, spanGaps:true,
    }}] }},
    options: {{
      responsive:true, interaction:{{mode:'index',intersect:false}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{font:{{size:12}}}}}},
        y:{{title:{{display:true,text:'DAC 成本 (USD)',font:{{size:12}}}},
          ticks:{{callback:v=>'$'+v,font:{{size:11}}}}, min:0}},
      }},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:ctx=>ctx.parsed.y!=null?'DAC 成本: $'+ctx.parsed.y:null}}}},
        datalabels:{{display:true, formatter:v=>v!=null?'$'+v:null,
          color:'#4f46e5', font:{{size:11,weight:'700'}},
          anchor:'top', align:'top', offset:3,
          backgroundColor:'rgba(255,255,255,0.85)', borderRadius:3,
          padding:{{top:2,bottom:2,left:4,right:4}}}},
      }},
    }},
  }});
}}

// ── 交互 ────────────────────────────────────────────────────────────────────
function selectProject(p){{selProject=p;renderProject();}}
function selectCountry(c){{selCountry=c;renderCountry();}}
function selectGroup(g){{selGroup=g;renderCampaign();}}

function setPeriod(n){{
  period=n;
  document.querySelectorAll(".pbtn").forEach((b,i)=>b.classList.toggle("active",[7,14,30][i]===n));
  renderAll();
}}

function switchTab(name){{
  activeTab=name;
  const names=["home","project","country","campaign","dac"];
  document.querySelectorAll(".tab").forEach((t,i)=>t.classList.toggle("active",names[i]===name));
  document.querySelectorAll(".panel").forEach(p=>p.classList.remove("active"));
  document.getElementById("panel-"+name).classList.add("active");
  if(name==="dac") renderDac();
}}

function renderAll(){{
  renderHome();
  renderProject();
  renderCountry();
  renderCampaign();
}}

renderAll();
</script>
</body>
</html>"""

# ─── 主程序 ───────────────────────────────────────────────────────────────────
def main():
    records      = load_data()
    data         = build_data(records)
    dac_data     = build_dac_data(records)
    all_signals, overall_roi = compute_all_signals(records)

    data_json        = json.dumps(data,        ensure_ascii=False)
    all_signals_json = json.dumps(all_signals, ensure_ascii=False)
    dac_json         = json.dumps(dac_data,    ensure_ascii=False)
    generated        = datetime.now().strftime("%Y-%m-%d %H:%M")
    html             = generate_html(data_json, all_signals_json, dac_json, generated)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    for n in (7, 14, 30):
        dates = get_window(records, n)
        print(f"  {n}天窗口: {dates[0] if dates else '?'} ~ {dates[-1] if dates else '?'} ({len(dates)}天)")
    m7 = all_signals["7"]["meta"]
    print(f"  信号(7天): {m7['n_issues']} 个问题, {m7['n_opps']} 个放量机会")
    print(f"\n✓ 看板已生成: {OUTPUT_FILE}")

    # ── 自动推送到 GitHub ──────────────────────────────────────────────────────
    repo = str(REPO_DIR)
    try:
        subprocess.run(["git", "-C", repo, "add", "index.html", "generate_dashboard.py"],
                       check=True, capture_output=True)
        msg = f"data: {datetime.now().strftime('%m/%d')} update"
        result = subprocess.run(["git", "-C", repo, "diff", "--cached", "--quiet"])
        if result.returncode != 0:  # 有变更才 commit
            subprocess.run(["git", "-C", repo, "commit", "-m", msg],
                           check=True, capture_output=True)
            subprocess.run(["git", "-C", repo, "push"], check=True, capture_output=True)
            print(f"✓ 已推送到 GitHub: https://zoeywang-066.github.io/ali-dashboard/")
        else:
            print("  (无变更，跳过推送)")
    except subprocess.CalledProcessError as e:
        print(f"  [警告] git 推送失败: {e.stderr.decode() if e.stderr else e}")

if __name__ == "__main__":
    main()
