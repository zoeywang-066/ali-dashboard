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

DOWNLOADS    = Path.home() / "Downloads"
REPO_DIR     = Path.home() / "ali-dashboard"
OUTPUT_FILE  = str(REPO_DIR / "index.html")
XLSX_FILE    = DOWNLOADS / "aliexpress_moloco_compaign_data.xlsx"
CSV_FILE     = DOWNLOADS / "aliexpress_moloco_compaign_data.csv"
NUMBERS_FILE = DOWNLOADS / "阿里投放数据.numbers"
SHEET_NAME   = "最新"
HISTORY_FILE = REPO_DIR / "data" / "ae_history_20251001_20260111.csv"
HISTORY_START = "20251001"
HISTORY_END   = "20260111"

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
def classify(name, geo=None):
    if not name: return ("country", "未知")
    if "EU10"    in name: return ("project", "EU10")
    if "海托"    in name: return ("project", "海托")
    if "欧洲本地" in name: return ("project", "欧洲本地")
    if re.search(r'Synapse[_\s]+全托', name, re.IGNORECASE):
        return ("project", "Synapse_全托")
    m = re.search(r'[Aa]ndroid[_\s][Aa]pp[_\s]([A-Z]{1,5}(?:/[A-Z]{1,5})?)', name, re.IGNORECASE)
    if m: return ("country", m.group(1).upper())
    m = re.search(r'全托[_\s]([A-Z]{1,5})(?:[_\s]|$)', name, re.IGNORECASE)
    if m: return ("country", m.group(1).upper())
    if geo:
        g = str(geo).strip().upper()
        if g and g not in ("-", "NAN", "NONE", "未匹配"):
            return ("country", g)
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

    current_count = len(records)
    if HISTORY_FILE.exists():
        history_records = _load_history_csv(HISTORY_FILE, HISTORY_START, HISTORY_END)
        records = _merge_records(history_records, records)
        print(f"  合并历史 AE数据: +{len(history_records)} 行历史, 当前下载 {current_count} 行, 合并后 {len(records)} 行")
    else:
        print(f"  未找到历史 AE数据文件，跳过: {HISTORY_FILE}")

    before = len(records)
    excluded = {r["name"] for r in records if is_other_channel(r["name"])}
    records = [r for r in records if not is_other_channel(r["name"])]
    if excluded:
        print(f"  过滤其他渠道(criteo)+脏数据: -{before - len(records)} 行 ({len(excluded)} 个 campaign)")
    return records

# 已知表头关键词集合（用于数据行污染检测）
HEADER_KEYWORDS = {
    "ds", "统计时间", "campaign_name", "campaign_namedsp", "campaign_id",
    "花费", "cost", "costdsp", "24h_dac", "session_dac", "session_dau", "dac成本",
    "24h-gmvroi", "24h_gmvroi", "sess-gmvroi",
}

def _norm_col(name):
    return re.sub(r"[\s_\-]+", "", str(name or "").strip().lower())

def _find_col(cm, candidates):
    """按原始表头或规范化表头查列，兼容空格/下划线/连字符变化。"""
    if not candidates:
        return None
    norm_map = {_norm_col(k): v for k, v in cm.items()}
    for cand in candidates:
        if cand in cm:
            return cm[cand]
        idx = norm_map.get(_norm_col(cand))
        if idx is not None:
            return idx
    return None

def _to_float(v):
    try:
        if v is None:
            return None
        s = str(v).strip().replace(",", "")
        if s in ("", "-", "nan", "NaN", "None", "NONE"):
            return None
        return float(s)
    except Exception:
        return None

def _merge_records(*record_sets):
    """合并历史和当前数据；同一日期/campaign/name 重叠时，后传入的数据优先。"""
    merged = {}
    order = []
    for records in record_sets:
        for r in records:
            key = (r.get("ds"), str(r.get("cid") or ""), str(r.get("name") or ""))
            if key not in merged:
                order.append(key)
            merged[key] = r
    return [merged[k] for k in order if k in merged]

def _resolve_gmv_roi_cols(raw, h_i, end_i, cm, spend_i):
    """每个数据段独立识别 GMV/ROI 列，防止同一 sheet 中表头位置变化或错位。"""
    gmv_names = [
        "24h_gmv", "24h-gmv", "24hgmv", "gmv24", "gmv_24h",
        "24h-gmvroi", "24h_gmvroi", "24h GMV",
    ]
    roi_names = [
        "24h-gmvroi", "24h_gmvroi", "24h_roi", "24hroi", "roi24",
        "sess-gmvroi", "sess_gmvroi", "session_gmvroi", "商业roi",
    ]
    gmv_candidates = []
    roi_candidates = []
    for name in gmv_names:
        idx = _find_col(cm, [name])
        if idx is not None and (name, idx) not in gmv_candidates:
            gmv_candidates.append((name, idx))
    for name in roi_names:
        idx = _find_col(cm, [name])
        if idx is not None and (name, idx) not in roi_candidates:
            roi_candidates.append((name, idx))

    def _tf_local(v):
        try:
            if v is None:
                return None
            s = str(v).strip().replace(",", "")
            if s in ("", "-", "nan", "None"):
                return None
            return float(s)
        except Exception:
            return None

    def _cell(row, idx):
        if idx is None or idx >= len(row):
            return None
        return _tf_local(row.iloc[idx])

    best = None
    sample_end = min(end_i, h_i + 101)
    for gmv_name, gmv_i in gmv_candidates:
        for roi_name, roi_i in roi_candidates:
            if gmv_i == roi_i:
                continue
            errors = []
            for row_i in range(h_i + 1, sample_end):
                row = raw.iloc[row_i]
                ds_raw = str(row.iloc[0]).strip().split(".")[0]
                if not ds_raw.isdigit() or len(ds_raw) != 8:
                    continue
                spend = _cell(row, spend_i)
                gmv = _cell(row, gmv_i)
                roi = _cell(row, roi_i)
                if not spend or spend <= 0 or gmv is None or roi is None or roi <= 0:
                    continue
                calc_roi = gmv / spend
                if calc_roi <= 0:
                    continue
                errors.append(abs(calc_roi - roi) / max(abs(roi), 0.01))
            if len(errors) < 3:
                continue
            errors.sort()
            median_err = errors[len(errors) // 2]
            score = (median_err, -len(errors))
            if best is None or score < best["score"]:
                best = {
                    "gmv_i": gmv_i, "roi_i": roi_i,
                    "gmv_name": gmv_name, "roi_name": roi_name,
                    "score": score, "median_err": median_err,
                    "n": len(errors),
                }

    if best and best["median_err"] <= 0.08:
        return best["gmv_i"], best["roi_i"], best

    gmv_i = _find_col(cm, ["24h_gmv", "24h-gmv", "24hgmv", "gmv24"])
    roi_i = _find_col(cm, ["24h-gmvroi", "24h_gmvroi", "24h_roi", "sess-gmvroi"])
    return gmv_i, roi_i, None

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
        "dau(session_dau)":            ["session_dau"],
        "dac(24h_dac/session_dac)":    ["24h_dac", "session_dac"],
        "dac_cost":                    ["dac成本"],
        "roi(24h/sess gmvroi)":        ["24h-gmvroi", "24h_gmvroi", "sess-gmvroi"],
        "gmv(24h_gmv/shifted)":        ["24h_gmv", "24h-gmvroi", "24h_gmvroi"],
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
        spend_i    = _find_col(cm, ["花费", "cost", "costdsp", "spend"])
        gmv_i, roi_i, metric_resolution = _resolve_gmv_roi_cols(raw, h_i, end_i, cm, spend_i)
        if metric_resolution:
            warnings.append(
                f"  ℹ️  段#{sec_i+1} (行{h_i+1}) 指标映射: "
                f"GMV={metric_resolution['gmv_name']} · ROI={metric_resolution['roi_name']}"
            )
        # DAC列：优先 24h_dac（新格式），其次 session_dac
        dac_i      = _find_col(cm, ["24h_dac", "session_dac", "dac"])
        # dac成本：新格式直接提供
        dac_cost_i = _find_col(cm, ["dac成本"])
        dau_i      = _find_col(cm, ["session_dau", "session dau", "dau"])
        name_i     = _find_col(cm, ["campaign_name", "campaign_namedsp"])
        cid_i      = _find_col(cm, ["campaign_id"])
        geo_i      = _find_col(cm, ["geo type", "geo_type", "country", "country_code"])

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
                geo   = _gv(row, geo_i) or ""
                spend = _tf(spend_raw) or 0.0
                roi   = _tf(_gv(row, roi_i))
                gmv   = _tf(_gv(row, gmv_i))   # 24h_gmv 金额，用于项目/国家聚合 ROI = Σgmv/Σ花费
                if roi is None and gmv is not None and spend > 0:
                    roi = gmv / spend
                dac   = _tf(_gv(row, dac_i)) or 0.0
                dau   = _tf(_gv(row, dau_i)) or 0.0
                # dac成本：新格式直接读取；旧格式留 None（由 build_dac_data 按汇总计算）
                dac_cost = _tf(_gv(row, dac_cost_i)) if dac_cost_i is not None else None
                all_records.append({
                    "ds": ds_raw, "name": name, "cid": cid,
                    "spend": spend, "roi": roi, "gmv": gmv, "dac": dac, "dac_cost": dac_cost,
                    "dau": dau, "geo": geo,
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
        if "sess-gmvroi" in col and "24h-gmvroi" in col:
            gmv_v = table.cell(r, col.get("24h-gmvroi")).value
            roi_v = table.cell(r, col.get("sess-gmvroi")).value
        else:
            gmv_v = table.cell(r, col.get("24h_gmv",    19)).value
            roi_v = table.cell(r, col.get("24h-gmvroi", 20)).value
        spend = float(spend_v) if isinstance(spend_v, (int, float)) else 0.0
        gmv   = float(gmv_v)   if isinstance(gmv_v,   (int, float)) else None
        roi   = float(roi_v)   if isinstance(roi_v,   (int, float)) else None
        if roi is None and gmv is not None and spend > 0:
            roi = gmv / spend
        dau_v = table.cell(r, col.get("session_dau", col.get("dau", 10))).value if ("session_dau" in col or "dau" in col) else 0
        dau   = float(dau_v)   if isinstance(dau_v,   (int, float)) else 0.0
        records.append({"ds": ds, "name": name, "cid": cid, "spend": spend, "roi": roi, "gmv": gmv, "dac": 0, "dac_cost": None, "dau": dau, "geo": ""})
    print(f"加载 {len(records)} 行数据")
    return records

def _load_history_csv(path, start_ds, end_ds):
    import pandas as pd
    print(f"读取历史 AE数据: {path.name}  {start_ds}~{end_ds}")
    df = pd.read_csv(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return _df_to_records(df, start_ds=start_ds, end_ds=end_ds)

def _resolve_df_gmv_roi_cols(df, cm, ds_i, spend_i):
    """单表 CSV/历史数据的 GMV/ROI 字段识别，兼容 ROI 列为空或 GMV/ROI 错位。"""
    gmv_names = [
        "24h_gmv", "24h-gmv", "24hgmv", "gmv24", "gmv_24h",
        "24h-gmvroi", "24h_gmvroi",
    ]
    roi_names = [
        "24h-gmvroi", "24h_gmvroi", "24h_roi", "24hroi", "roi24",
        "sess-gmvroi", "sess_gmvroi", "session_gmvroi", "商业roi",
    ]
    gmv_candidates = []
    roi_candidates = []
    for name in gmv_names:
        idx = _find_col(cm, [name])
        if idx is not None and (name, idx) not in gmv_candidates:
            gmv_candidates.append((name, idx))
    for name in roi_names:
        idx = _find_col(cm, [name])
        if idx is not None and (name, idx) not in roi_candidates:
            roi_candidates.append((name, idx))

    best = None
    sample_end = min(len(df), 200)
    for gmv_name, gmv_i in gmv_candidates:
        for roi_name, roi_i in roi_candidates:
            if gmv_i == roi_i:
                continue
            errors = []
            for row_i in range(sample_end):
                row = df.iloc[row_i]
                ds_raw = str(row.iloc[ds_i]).strip().split(".")[0] if ds_i is not None else ""
                if not ds_raw.isdigit() or len(ds_raw) != 8:
                    continue
                spend = _to_float(row.iloc[spend_i]) if spend_i is not None else None
                gmv = _to_float(row.iloc[gmv_i]) if gmv_i is not None else None
                roi = _to_float(row.iloc[roi_i]) if roi_i is not None else None
                if not spend or spend <= 0 or gmv is None or roi is None or roi <= 0:
                    continue
                errors.append(abs((gmv / spend) - roi) / max(abs(roi), 0.01))
            if len(errors) < 3:
                continue
            errors.sort()
            median_err = errors[len(errors) // 2]
            score = (median_err, -len(errors))
            if best is None or score < best["score"]:
                best = {"gmv_i": gmv_i, "roi_i": roi_i, "score": score, "median_err": median_err}

    if best and best["median_err"] <= 0.08:
        return best["gmv_i"], best["roi_i"]

    gmv_i = _find_col(cm, ["24h_gmv", "24h-gmv", "24hgmv", "gmv24", "gmv_24h"])
    if gmv_i is None:
        gmv_i = _find_col(cm, ["24h-gmvroi", "24h_gmvroi"])
    roi_i = _find_col(cm, ["24h-gmvroi", "24h_gmvroi", "24h_roi", "sess-gmvroi", "商业roi"])
    if roi_i == gmv_i:
        roi_i = _find_col(cm, ["sess-gmvroi", "sess_gmvroi", "session_gmvroi", "商业roi"])
    return gmv_i, roi_i

def _df_to_records(df, start_ds=None, end_ds=None):
    """CSV / Numbers 备用加载路径（xlsx 走 _load_xlsx 逐 section 解析）"""
    records = []
    cm = {str(c).strip(): i for i, c in enumerate(df.columns)}
    ds_i        = _find_col(cm, ["ds", "统计时间", "date", "local date", "local_date"])
    name_i      = _find_col(cm, ["campaign_name", "campaign_namedsp", "campaign name"])
    cid_i       = _find_col(cm, ["campaign_id", "campaignid", "campaign id"])
    spend_i     = _find_col(cm, ["花费", "spend", "costdsp", "cost"])
    dau_i       = _find_col(cm, ["session_dau", "session dau", "dau"])
    dac_i       = _find_col(cm, ["24h_dac", "session_dac", "dac"])
    dac_cost_i  = _find_col(cm, ["dac成本", "dac_cost", "dac cost"])
    geo_i       = _find_col(cm, ["geo type", "geo_type", "country", "country_code"])
    gmv_i, roi_i = _resolve_df_gmv_roi_cols(df, cm, ds_i, spend_i)

    def _value(row, idx):
        if idx is None:
            return ""
        v = row.iloc[idx]
        if v is None:
            return ""
        s = str(v).strip()
        return "" if s in ("nan", "NaN", "None") else s

    for _, row in df.iterrows():
        try:
            ds_raw = _value(row, ds_i).split(".")[0]
            if not ds_raw.isdigit() or len(ds_raw) != 8:
                continue
            if start_ds and ds_raw < start_ds:
                continue
            if end_ds and ds_raw > end_ds:
                continue
            name     = _value(row, name_i)
            cid      = _value(row, cid_i)
            geo      = _value(row, geo_i)
            spend    = _to_float(_value(row, spend_i)) or 0.0
            roi      = _to_float(_value(row, roi_i)) if roi_i is not None else None
            gmv      = _to_float(_value(row, gmv_i)) if gmv_i is not None else None
            if roi is None and gmv is not None and spend > 0:
                roi = gmv / spend
            dac      = _to_float(_value(row, dac_i)) or 0.0
            dac_cost = _to_float(_value(row, dac_cost_i)) if dac_cost_i is not None else None
            dau      = _to_float(_value(row, dau_i)) or 0.0
            records.append({"ds": ds_raw, "name": name, "cid": cid,
                            "spend": spend, "roi": roi, "gmv": gmv, "dac": dac, "dac_cost": dac_cost, "dau": dau, "geo": geo})
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
    windows = [("all", get_all_dates(records))] + [(str(n), get_window(records, n)) for n in (3, 7, 14)]
    for key, dates in windows:
        labels   = fmt_dates(dates)
        date_set = set(dates)
        window_days = max(len(dates), 1)

        # country → day → {cost, dac}
        c_daily = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0}))
        # country → camp_short → day → {cost, dac}
        c_camp  = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0})))
        # DAC 专项：整体 + 分 campaign
        sp_overall = defaultdict(lambda: {"cost": 0.0, "dac": 0.0, "gmv": 0.0})           # ds → {cost, dac, gmv}
        sp_camp    = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dac": 0.0, "gmv": 0.0}))  # camp_short → ds → {cost, dac, gmv}

        for r in records:
            if r["ds"] not in date_set:
                continue
            # 只用有 dac成本 直接字段的记录（24h_dac 口径），session_dac 不参与
            if r.get("dac_cost") is None:
                continue
            _, label = classify(r["name"], r.get("geo"))
            sn = short_name(r["name"])
            c_daily[label][r["ds"]]["cost"] += r["spend"]
            c_daily[label][r["ds"]]["dac"]  += r.get("dac") or 0
            c_camp[label][sn][r["ds"]]["cost"] += r["spend"]
            c_camp[label][sn][r["ds"]]["dac"]  += r.get("dac") or 0
            # DAC 专项独立聚合（不影响国家口径）
            if "DAC专项" in (r.get("name") or ""):
                sp_overall[r["ds"]]["cost"] += r["spend"]
                sp_overall[r["ds"]]["dac"]  += r.get("dac") or 0
                sp_overall[r["ds"]]["gmv"]  += r.get("gmv") or 0
                sp_camp[sn][r["ds"]]["cost"] += r["spend"]
                sp_camp[sn][r["ds"]]["dac"]  += r.get("dac") or 0
                sp_camp[sn][r["ds"]]["gmv"]  += r.get("gmv") or 0

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
                    "daily_spend": round(tc / window_days, 0),
                }

        # Country daily series（分天 DAC成本）
        country_daily = {}
        country_metrics = {}
        for c, ds_map in c_daily.items():
            spend_s = [round(ds_map.get(d, {}).get("cost", 0), 2) for d in dates]
            dac_s   = [int(round(ds_map.get(d, {}).get("dac", 0), 0)) for d in dates]
            cost_s  = [round(ds_map[d]["cost"] / ds_map[d]["dac"], 2)
                       if ds_map.get(d, {}).get("dac", 0) > 0 else None
                       for d in dates]
            if any(v is not None for v in cost_s):
                country_daily[c] = cost_s
                country_metrics[c] = {
                    "spend_series": spend_s,
                    "dac_series":   dac_s,
                    "cost_series":  cost_s,
                }

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
            gmv_s   = [round(ds_map.get(d, {}).get("gmv", 0), 2) for d in dates]
            roi_s   = [round(ds_map[d]["gmv"] / ds_map[d]["cost"], 2)
                       if ds_map.get(d, {}).get("cost", 0) > 0 else None
                       for d in dates]
            return spend_s, dac_s, cost_s, gmv_s, roi_s

        dac_special = {}
        if sp_overall:
            tc = sum(b["cost"] for b in sp_overall.values())
            td = sum(b["dac"]  for b in sp_overall.values())
            tg = sum(b["gmv"]  for b in sp_overall.values())
            sp_s, dc_s, cs_s, gm_s, ro_s = _series(sp_overall)
            dac_special["overall"] = {
                "cost":         round(tc, 0),
                "dac":          int(round(td, 0)),
                "dac_cost":     round(tc / td, 2) if td > 0 else None,
                "gmv":          round(tg, 0),
                "roi":          round(tg / tc, 2) if tc > 0 else None,
                "daily_spend":  round(tc / window_days, 0),
                "spend_series": sp_s,
                "dac_series":   dc_s,
                "cost_series":  cs_s,
                "gmv_series":   gm_s,
                "roi_series":   ro_s,
            }
            camps = {}
            for camp, ds_map in sp_camp.items():
                tc2 = sum(b["cost"] for b in ds_map.values())
                td2 = sum(b["dac"]  for b in ds_map.values())
                tg2 = sum(b["gmv"]  for b in ds_map.values())
                sp2, dc2, cs2, gm2, ro2 = _series(ds_map)
                camps[camp] = {
                    "cost":         round(tc2, 0),
                    "dac":          int(round(td2, 0)),
                    "dac_cost":     round(tc2 / td2, 2) if td2 > 0 else None,
                    "gmv":          round(tg2, 0),
                    "roi":          round(tg2 / tc2, 2) if tc2 > 0 else None,
                    "daily_spend":  round(tc2 / window_days, 0),
                    "spend_series": sp2,
                    "dac_series":   dc2,
                    "cost_series":  cs2,
                    "gmv_series":   gm2,
                    "roi_series":   ro2,
                }
            dac_special["campaigns"] = camps

        result[key] = {
            "dates":           dates,
            "labels":          labels,
            "country_summary": country_summary,
            "country_metrics": country_metrics,
            "country_daily":   country_daily,
            "camp_daily":      camp_daily,
            "dac_special":     dac_special,
        }
        n_countries = len(country_summary)
        n_special   = len(dac_special.get("campaigns", {})) if dac_special else 0
        print(f"  [DAC] {key}窗口: {n_countries} 个国家/项目有 dac成本 数据 · DAC专项 {n_special} 个 campaign")
    return result

# ─── DAU 成本数据（cost / session_dau）────────────────────────────────────────
def build_dau_cost_data(records):
    """按国家/项目、Campaign、日期聚合 DAU成本：Σcost / Σsession_dau。"""
    if not records:
        return {}
    dates  = get_all_dates(records)
    labels = fmt_dates(dates)
    date_set = set(dates)

    # country/project → day → {cost, dau}
    c_daily = defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dau": 0.0}))
    # country/project → campaign → day → {cost, dau}
    c_camp = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"cost": 0.0, "dau": 0.0})))
    cid_to_short = {}

    for r in records:
        if r["ds"] not in date_set:
            continue
        dau = r.get("dau") or 0
        spend = r.get("spend") or 0
        if dau <= 0 or spend <= 0:
            continue
        _, label = classify(r["name"], r.get("geo"))
        cid = str(r.get("cid") or "").strip()
        camp_label = cid_to_short.setdefault((label, cid), short_name(r["name"])) if cid else short_name(r["name"])

        c_daily[label][r["ds"]]["cost"] += spend
        c_daily[label][r["ds"]]["dau"]  += dau
        c_camp[label][camp_label][r["ds"]]["cost"] += spend
        c_camp[label][camp_label][r["ds"]]["dau"]  += dau

    country_summary = {}
    country_metrics = {}
    country_daily = {}
    for c, ds_map in c_daily.items():
        total_cost = sum(b["cost"] for b in ds_map.values())
        total_dau  = sum(b["dau"] for b in ds_map.values())
        if total_dau <= 0:
            continue
        spend_s = [round(ds_map.get(d, {}).get("cost", 0), 2) for d in dates]
        dau_s   = [int(round(ds_map.get(d, {}).get("dau", 0), 0)) for d in dates]
        cost_s  = [
            round(ds_map[d]["cost"] / ds_map[d]["dau"], 4)
            if ds_map.get(d, {}).get("dau", 0) > 0 else None
            for d in dates
        ]
        country_summary[c] = {
            "cost":        round(total_cost, 0),
            "dau":         int(round(total_dau, 0)),
            "dau_cost":    round(total_cost / total_dau, 4),
            "daily_spend": round(total_cost / max(len(dates), 1), 0),
        }
        country_metrics[c] = {
            "spend_series": spend_s,
            "dau_series":   dau_s,
            "cost_series":  cost_s,
        }
        country_daily[c] = cost_s

    camp_daily = {}
    for c, camp_map in c_camp.items():
        camps = {}
        for camp, ds_map in camp_map.items():
            series = [
                round(ds_map[d]["cost"] / ds_map[d]["dau"], 4)
                if ds_map.get(d, {}).get("dau", 0) > 0 else None
                for d in dates
            ]
            if any(v is not None for v in series):
                camps[camp] = series
        if camps:
            camp_daily[c] = camps

    print(f"  [DAU] all窗口: {len(country_summary)} 个国家/项目有 session_dau 数据")
    return {
        "all": {
            "dates":           dates,
            "labels":          labels,
            "country_summary": country_summary,
            "country_metrics": country_metrics,
            "country_daily":   country_daily,
            "camp_daily":      camp_daily,
        }
    }

# ─── 日期窗口 ─────────────────────────────────────────────────────────────────
def get_all_dates(records):
    return sorted({r["ds"] for r in records if str(r.get("ds", "")).isdigit()})

def get_window(records, n):
    all_dates = get_all_dates(records)
    if not all_dates: return []
    end_dt   = datetime.strptime(all_dates[-1], "%Y%m%d")
    start_dt = end_dt - timedelta(days=n - 1)
    return [d for d in all_dates if datetime.strptime(d, "%Y%m%d") >= start_dt]

# ─── 聚合 ────────────────────────────────────────────────────────────────────
def aggregate(records, dates):
    date_set    = set(dates)
    proj_agg    = defaultdict(lambda: defaultdict(lambda: {"spend":0,"gmv":0}))
    country_agg = defaultdict(lambda: defaultdict(lambda: {"spend":0,"gmv":0}))
    camp_agg    = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"spend":0,"ws":0,"wsr":0})))
    cid_to_short = {}

    for r in records:
        if r["ds"] not in date_set: continue
        kind, label = classify(r["name"], r.get("geo"))
        target = proj_agg[label] if kind == "project" else country_agg[label]
        b = target[r["ds"]]
        b["spend"] += r["spend"]
        b["gmv"]   += (r.get("gmv") or 0)          # 项目/国家聚合 ROI = Σ24h_gmv / Σ花费
        cid = str(r.get("cid") or "").strip()
        if cid:
            sn = cid_to_short.setdefault((kind, label, cid), short_name(r["name"]))
        else:
            sn = short_name(r["name"])
        cb = camp_agg[label][sn][r["ds"]]
        cb["spend"] += r["spend"]
        if r["roi"] and r["roi"] > 0 and r["spend"] > 0:   # 单 campaign 仍直接用 24h-gmvroi 字段
            cb["ws"]  += r["spend"]
            cb["wsr"] += r["roi"] * r["spend"]

    def to_series(agg_dict):
        result = {}
        for label, ds_map in agg_dict.items():
            spend_list, roi_list, gmv_list = [], [], []
            for d in dates:
                b = ds_map.get(d, {"spend":0,"gmv":0})
                sp = round(b["spend"], 2); gm = round(b.get("gmv", 0), 2)
                spend_list.append(sp); gmv_list.append(gm)
                # 当日聚合 ROI = 当日 Σ24h_gmv / 当日 Σ花费
                roi_list.append(round(gm / b["spend"], 2) if b["spend"] > 0 else None)
            if sum(s for s in spend_list if s) > 0:
                result[label] = {"spend": spend_list, "roi": roi_list, "gmv": gmv_list}
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

def build_platform_series(records, dates):
    date_set = set(dates)
    daily = defaultdict(lambda: {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0})

    for r in records:
        if r["ds"] not in date_set:
            continue
        b = daily[r["ds"]]
        b["spend"] += r["spend"]
        gmv = r.get("gmv")
        if gmv is not None:
            b["gmv"] += gmv
            b["gmv_rows"] += 1

    spend_list, gmv_list, roi_list = [], [], []
    for d in dates:
        b = daily.get(d, {"spend": 0.0, "gmv": 0.0, "gmv_rows": 0})
        spend = round(b["spend"], 2)
        gmv = round(b["gmv"], 2)
        spend_list.append(spend)
        gmv_list.append(gmv)
        roi_list.append(round(gmv / spend, 2) if spend > 0 and b["gmv_rows"] > 0 else None)

    total_spend = sum(spend_list)
    total_gmv = sum(gmv_list)
    total_gmv_rows = sum(b["gmv_rows"] for b in daily.values())
    return {
        "spend": spend_list,
        "gmv": gmv_list,
        "roi": roi_list,
        "summary": {
            "spend": round(total_spend, 0),
            "gmv": round(total_gmv, 0),
            "roi": round(total_gmv / total_spend, 2) if total_spend > 0 and total_gmv_rows > 0 else None,
        },
    }

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
        _, label = classify(r["name"], r.get("geo"))
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
        _, label = classify(r["name"], r.get("geo"))
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
        kind, label = classify(r["name"], r.get("geo"))
        (all_projects if kind == "project" else all_countries).add(label)

    country_list = sorted(all_countries)
    c_colors = {c: country_color(i) for i, c in enumerate(country_list)}
    p_colors = {p: (PROJECT_COLORS.get(p,{"bar":rgb(100,100,100,.75),"line":"#555"})["bar"],
                    PROJECT_COLORS.get(p,{"bar":rgb(100,100,100,.75),"line":"#555"})["line"])
                for p in all_projects}
    all_colors = {**{k:{"bar":bc,"line":lc} for k,(bc,lc) in c_colors.items()},
                  **{k:{"bar":bc,"line":lc} for k,(bc,lc) in p_colors.items()}}

    result = {}
    windows = [("all", get_all_dates(records))] + [(str(n), get_window(records, n)) for n in (7, 14, 30)]
    for key, dates in windows:
        labels = fmt_dates(dates)
        proj, country, campaigns = aggregate(records, dates)
        platform = build_platform_series(records, dates)

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
            out = {}
            for lbl, d in series.items():
                tot_spend = sum(s for s in d["spend"] if s)
                tot_gmv   = sum(g for g in d.get("gmv", []) if g)
                # 整体（窗口）ROI = Σ24h_gmv / Σ花费（花费加权，不是每日 ROI 简单平均）
                out[lbl] = {"spend": round(tot_spend, 0),
                            "roi":   round(tot_gmv / tot_spend, 2) if tot_spend > 0 else None}
            return out

        all_camp_ds = {g: make_camp_ds(gc) for g, gc in campaigns.items()}
        platform_ds = [
            {"label": "平台消耗(costdsp)", "data": platform["spend"],
             "backgroundColor": "rgba(14,165,233,0.42)", "borderColor": "#0ea5e9",
             "borderWidth": 1, "yAxisID": "ySpend", "type": "bar", "stack": "spend"},
            {"label": "平台 ROI", "data": platform["roi"],
             "borderColor": "#059669", "backgroundColor": "transparent",
             "borderWidth": 3, "pointRadius": 4, "pointHoverRadius": 6, "tension": 0.3,
             "yAxisID": "yROI", "type": "line", "spanGaps": True},
        ]

        result[key] = {
            "dates":           dates,
            "labels":          labels,
            "platform":        {**platform, "ds": platform_ds},
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
def generate_html(data_json, all_signals_json, dac_json, dau_cost_json, generated_at):
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
.date-range-bar{{display:flex;align-items:center;gap:6px;margin-left:6px;padding-left:10px;border-left:1px solid #e2e8f0}}
.date-input{{height:30px;border:1.5px solid #e2e8f0;border-radius:8px;padding:4px 8px;font-size:12px;color:#334155;background:#fff}}
.date-input:focus{{outline:none;border-color:#4f46e5;box-shadow:0 0 0 2px rgba(79,70,229,.12)}}
.range-sep{{font-size:12px;color:#94a3b8}}
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
.table-wrap{{overflow:auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.07);margin-bottom:22px}}
.metric-table{{width:100%;border-collapse:collapse;font-size:12px;min-width:760px}}
.metric-table th,.metric-table td{{padding:9px 10px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap}}
.metric-table th:first-child,.metric-table td:first-child{{text-align:left;position:sticky;left:0;background:#fff;z-index:1;max-width:360px;overflow:hidden;text-overflow:ellipsis}}
.metric-table thead th{{background:#f8fafc;color:#475569;font-weight:700}}
.metric-table thead th:first-child{{background:#f8fafc;z-index:2}}
.group-chips{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px}}
.chip{{padding:6px 14px;border-radius:99px;border:1.5px solid #e2e8f0;background:#fff;font-size:13px;font-weight:500;cursor:pointer;color:#475569;transition:all .15s}}
.chip:hover{{border-color:#a5b4fc;color:#4f46e5}}
.chip.selected{{border-color:#4f46e5;background:#4f46e5;color:#fff}}
.chip.proj{{border-color:#fecaca}}
.chip.proj.selected{{background:#dc2626;border-color:#dc2626;color:#fff}}
/* ── 视图控件：指标切换 + 全选/清空 ── */
.view-controls{{display:flex;gap:16px;align-items:center;flex-wrap:wrap;margin-bottom:14px}}
.ctrl-grp{{display:flex;align-items:center;gap:7px}}
.ctrl-lbl{{font-size:12px;color:#64748b;font-weight:600}}
.seg{{display:inline-flex;border:1.5px solid #e2e8f0;border-radius:8px;overflow:hidden}}
.seg button{{border:none;background:#fff;padding:6px 16px;font-size:13px;font-weight:600;color:#475569;cursor:pointer;transition:all .15s}}
.seg button + button{{border-left:1.5px solid #e2e8f0}}
.seg button:hover{{color:#4f46e5}}
.seg button.active{{background:#4f46e5;color:#fff}}
.mini-btn{{border:1.5px solid #e2e8f0;background:#fff;border-radius:8px;padding:6px 13px;font-size:13px;font-weight:600;color:#475569;cursor:pointer;transition:all .15s}}
.mini-btn:hover{{border-color:#a5b4fc;color:#4f46e5}}
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
  <div class="tab" onclick="switchTab('dau')">DAU成本</div>
  <div class="tab" onclick="switchTab('dac')">DAC成本</div>
  <div class="period-bar">
    <button class="pbtn" data-period="7" onclick="setPeriod(7)">过去 7 天</button>
    <button class="pbtn active" data-period="14" onclick="setPeriod(14)">过去 14 天</button>
    <button class="pbtn" data-period="30" onclick="setPeriod(30)">过去 30 天</button>
    <button class="pbtn" data-period="all" onclick="setPeriod('all')">全部</button>
    <div class="date-range-bar">
      <input id="date-start" class="date-input" type="date" onchange="setCustomDateRange()">
      <span class="range-sep">至</span>
      <input id="date-end" class="date-input" type="date" onchange="setCustomDateRange()">
    </div>
  </div>
</div>

<!-- ══ 首页 ══ -->
<div id="panel-home" class="panel active">
  <div id="kpi-bar" class="kpi-bar"></div>
  <div class="chart-box">
    <div class="chart-hd">整体平台消耗（柱）& ROI（线）趋势
      <span class="badge">左轴: costdsp USD · 右轴: ROI · Σ24h_gmv / Σcostdsp</span></div>
    <canvas id="platformRoiChart"></canvas>
    <div class="note">整体平台维度：按天汇总所有 Moloco campaign 的 costdsp 与 24h_gmv</div>
  </div>
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
  <div class="view-controls">
    <div class="ctrl-grp"><span class="ctrl-lbl">指标</span>
      <div class="seg" id="camp-metric">
        <button data-metric="both" onclick="setCampMetric('both')">消耗+ROI</button>
        <button data-metric="spend" onclick="setCampMetric('spend')">消耗</button>
        <button data-metric="roi" onclick="setCampMetric('roi')">ROI</button>
      </div>
    </div>
  </div>
  <div id="group-chips" class="group-chips"></div>
  <div id="camp-chart-area">
    <div class="empty-hint">👆 选择上方某个国家或项目，查看其全部 Campaign 趋势</div>
  </div>
</div>

<!-- ══ DAU成本 面板 ══ -->
<div id="panel-dau" class="panel">
  <div class="section-title">
    各国家 / 项目 DAU成本（当前日期范围）
    <span id="clear-dau-country" class="clear-btn" style="display:none" onclick="selectDauCountry(null)">✕ 清除</span>
  </div>
  <div class="note" style="margin:-8px 0 16px">DAU成本 = Σcost / Σsession_dau；跟随顶部日期筛选。</div>
  <div id="dau-country-cards" class="cards"></div>
  <div id="dau-chart-area">
    <div class="empty-hint">👆 点击国家 / 项目卡片，查看 DAU成本日趋势和 Campaign 分天明细</div>
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
    <div class="empty-hint">👆 点击「整体」或单个 campaign 卡片，查看 DAC数 / 花费 / DAC成本 / 24H GMV / ROI 分天趋势</div>
  </div>
</div>

<script>
const ALL         = {data_json};
const ALL_SIGNALS = {all_signals_json};
const DAC_DATA    = {dac_json};
const DAU_COST_DATA = {dau_cost_json};

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
let campMetric     = "both";
let selDauCountry  = null;
let selDacCountry  = null;
let selDacSpecial  = null;

let platformChart=null, projChart=null, countryChart=null, projRefChart=null, campChart=null;
let dauTrendChart=null, dauCampChart=null;
let dacTrendChart=null, dacCampChart=null;
let dacSpecSpendChart=null, dacSpecCostChart=null, dacSpecRoiChart=null;

if (typeof ChartDataLabels !== "undefined") Chart.register(ChartDataLabels);

const BASE_DATA = ALL["all"] || ALL["30"] || ALL["14"] || ALL["7"] || {{}};
const ALL_DATES = BASE_DATA.dates || [];
let dateStart = null;
let dateEnd = null;

function dsToInput(ds) {{
  return ds ? `${{ds.slice(0,4)}}-${{ds.slice(4,6)}}-${{ds.slice(6,8)}}` : "";
}}

function inputToDs(v) {{
  return (v || "").replaceAll("-", "");
}}

function setDateInputs() {{
  const s = document.getElementById("date-start");
  const e = document.getElementById("date-end");
  if (!s || !e || !ALL_DATES.length) return;
  const min = dsToInput(ALL_DATES[0]);
  const max = dsToInput(ALL_DATES[ALL_DATES.length - 1]);
  [s, e].forEach(el => {{ el.min = min; el.max = max; }});
  s.value = dsToInput(dateStart);
  e.value = dsToInput(dateEnd);
}}

function updatePeriodButtons() {{
  document.querySelectorAll(".pbtn").forEach(b => {{
    b.classList.toggle("active", String(b.dataset.period) === String(period));
  }});
  [3, 7, 14].forEach(d => {{
    const btn = document.getElementById("dac-btn-" + d);
    if (btn) btn.classList.toggle("active", String(period) === String(d));
  }});
}}

function setRangeToLast(n) {{
  if (!ALL_DATES.length) return;
  const endIdx = ALL_DATES.length - 1;
  const startIdx = Math.max(0, ALL_DATES.length - Number(n));
  dateStart = ALL_DATES[startIdx];
  dateEnd = ALL_DATES[endIdx];
}}

function setFullRange() {{
  if (!ALL_DATES.length) return;
  dateStart = ALL_DATES[0];
  dateEnd = ALL_DATES[ALL_DATES.length - 1];
}}

function setCustomDateRange() {{
  const s = inputToDs(document.getElementById("date-start").value);
  const e = inputToDs(document.getElementById("date-end").value);
  if (!s || !e) return;
  dateStart = s <= e ? s : e;
  dateEnd = s <= e ? e : s;
  period = "custom";
  setDateInputs();
  updatePeriodButtons();
  renderAll();
  if (activeTab === "dau") renderDauCost();
  if (activeTab === "dac") renderDac();
}}

function setPeriod(n) {{
  period = n;
  if (n === "all") setFullRange();
  else setRangeToLast(n);
  setDateInputs();
  updatePeriodButtons();
  renderAll();
  if (activeTab === "dau") renderDauCost();
  if (activeTab === "dac") renderDac();
}}

function getSliceRange(dates) {{
  if (!dates || !dates.length) return [0, -1];
  const start = dateStart || dates[0];
  const end = dateEnd || dates[dates.length - 1];
  let s = dates.findIndex(d => d >= start);
  let e = dates.length - 1;
  for (let i = dates.length - 1; i >= 0; i--) {{
    if (dates[i] <= end) {{ e = i; break; }}
  }}
  if (s < 0) s = 0;
  if (e < s) return [0, -1];
  return [s, e];
}}

function sliceArr(arr, s, e) {{
  return (arr || []).slice(s, e + 1);
}}

function sumVals(arr) {{
  return (arr || []).reduce((s, v) => s + (Number(v) || 0), 0);
}}

function fmtDauCost(v) {{
  if (v === null || v === undefined || Number.isNaN(Number(v))) return "-";
  const n = Number(v);
  const digits = Math.abs(n) < 1 ? 4 : 2;
  return "$" + n.toLocaleString(undefined, {{minimumFractionDigits: digits, maximumFractionDigits: digits}});
}}

function sliceDataset(ds, s, e) {{
  return {{...ds, data: sliceArr(ds.data, s, e)}};
}}

function sliceSeriesMap(seriesMap, s, e) {{
  const out = {{}};
  Object.entries(seriesMap || {{}}).forEach(([k, v]) => {{
    out[k] = {{
      ...v,
      spend: sliceArr(v.spend, s, e),
      roi:   sliceArr(v.roi,   s, e),
      gmv:   sliceArr(v.gmv,   s, e),
    }};
  }});
  return out;
}}

function summarizeSeriesMap(seriesMap) {{
  const out = {{}};
  Object.entries(seriesMap || {{}}).forEach(([k, v]) => {{
    const spend = sumVals(v.spend);
    const gmv = sumVals(v.gmv);
    if (spend > 0) out[k] = {{spend: Math.round(spend), roi: gmv > 0 ? Number((gmv / spend).toFixed(2)) : null}};
  }});
  return out;
}}

function getFilteredData() {{
  const base = BASE_DATA;
  const dates = base.dates || [];
  const [s, e] = getSliceRange(dates);
  if (e < s) return {{...base, dates: [], labels: [], proj_summary: {{}}, country_summary: {{}}, group_summary: {{}}}};
  const labels = sliceArr(base.labels, s, e);
  const fd = {{
    ...base,
    dates: sliceArr(dates, s, e),
    labels,
    proj_ds: (base.proj_ds || []).map(ds => sliceDataset(ds, s, e)),
    country_ds: (base.country_ds || []).map(ds => sliceDataset(ds, s, e)),
    proj_raw: sliceSeriesMap(base.proj_raw, s, e),
    country_raw: sliceSeriesMap(base.country_raw, s, e),
    camp_ds: {{}},
  }};
  fd.proj_summary = summarizeSeriesMap(fd.proj_raw);
  fd.country_summary = summarizeSeriesMap(fd.country_raw);
  fd.group_summary = {{...fd.proj_summary, ...fd.country_summary}};
  Object.entries(base.camp_ds || {{}}).forEach(([g, arr]) => {{
    fd.camp_ds[g] = (arr || []).map(ds => sliceDataset(ds, s, e));
  }});
  const p = base.platform || {{}};
  const spend = sliceArr(p.spend, s, e);
  const gmv = sliceArr(p.gmv, s, e);
  const roi = sliceArr(p.roi, s, e);
  const totalSpend = sumVals(spend);
  const totalGmv = sumVals(gmv);
  fd.platform = {{
    ...p,
    spend, gmv, roi,
    ds: (p.ds || []).map(ds => sliceDataset(ds, s, e)),
    summary: {{
      spend: Math.round(totalSpend),
      gmv: Math.round(totalGmv),
      roi: totalSpend > 0 && totalGmv > 0 ? Number((totalGmv / totalSpend).toFixed(2)) : null,
    }},
  }};
  return fd;
}}

function sliceDacEntity(entity, s, e, labelCount) {{
  const spendSeries = sliceArr(entity.spend_series, s, e);
  const dacSeries = sliceArr(entity.dac_series, s, e);
  const costSeries = sliceArr(entity.cost_series, s, e);
  const gmvSeries = sliceArr(entity.gmv_series, s, e);
  const roiSeries = sliceArr(entity.roi_series, s, e);
  const cost = sumVals(spendSeries);
  const dac = sumVals(dacSeries);
  const gmv = sumVals(gmvSeries);
  return {{
    ...entity,
    cost: Math.round(cost),
    dac: Math.round(dac),
    dac_cost: dac > 0 ? Number((cost / dac).toFixed(2)) : null,
    gmv: Math.round(gmv),
    roi: cost > 0 && gmv > 0 ? Number((gmv / cost).toFixed(2)) : null,
    daily_spend: labelCount > 0 ? Math.round(cost / labelCount) : 0,
    spend_series: spendSeries,
    dac_series: dacSeries,
    cost_series: costSeries,
    gmv_series: gmvSeries,
    roi_series: roiSeries,
  }};
}}

function getFilteredDacData() {{
  const base = DAC_DATA["all"] || DAC_DATA[String(dacPeriod)] || {{}};
  const dates = base.dates || [];
  const [s, e] = getSliceRange(dates);
  if (e < s) return {{labels: [], country_summary: {{}}, country_daily: {{}}, camp_daily: {{}}, dac_special: {{}}}};
  const labels = sliceArr(base.labels, s, e);
  const country_summary = {{}};
  const country_daily = {{}};
  Object.entries(base.country_metrics || {{}}).forEach(([c, m]) => {{
    const spend = sliceArr(m.spend_series, s, e);
    const dac = sliceArr(m.dac_series, s, e);
    const cost = sliceArr(m.cost_series, s, e);
    const tc = sumVals(spend);
    const td = sumVals(dac);
    if (td > 0) {{
      country_summary[c] = {{
        cost: Math.round(tc),
        dac: Math.round(td),
        dac_cost: Number((tc / td).toFixed(2)),
        daily_spend: labels.length ? Math.round(tc / labels.length) : 0,
      }};
      country_daily[c] = cost;
    }}
  }});
  const camp_daily = {{}};
  Object.entries(base.camp_daily || {{}}).forEach(([c, camps]) => {{
    const out = {{}};
    Object.entries(camps || {{}}).forEach(([camp, vals]) => {{
      out[camp] = sliceArr(vals, s, e);
    }});
    if (Object.keys(out).length) camp_daily[c] = out;
  }});
  const sp = base.dac_special || {{}};
  const dac_special = {{}};
  if (sp.overall) dac_special.overall = sliceDacEntity(sp.overall, s, e, labels.length);
  if (sp.campaigns) {{
    dac_special.campaigns = {{}};
    Object.entries(sp.campaigns).forEach(([camp, entity]) => {{
      dac_special.campaigns[camp] = sliceDacEntity(entity, s, e, labels.length);
    }});
  }}
  return {{...base, dates: sliceArr(dates, s, e), labels, country_summary, country_daily, camp_daily, dac_special}};
}}

function getFilteredDauCostData() {{
  const base = (DAU_COST_DATA || {{}})["all"] || {{}};
  const dates = base.dates || [];
  const [s, e] = getSliceRange(dates);
  if (e < s) return {{labels: [], country_summary: {{}}, country_daily: {{}}, camp_daily: {{}}}};
  const labels = sliceArr(base.labels, s, e);
  const country_summary = {{}};
  const country_daily = {{}};
  Object.entries(base.country_metrics || {{}}).forEach(([c, m]) => {{
    const spend = sliceArr(m.spend_series, s, e);
    const dau = sliceArr(m.dau_series, s, e);
    const cost = sliceArr(m.cost_series, s, e);
    const tc = sumVals(spend);
    const td = sumVals(dau);
    if (td > 0) {{
      country_summary[c] = {{
        cost: Math.round(tc),
        dau: Math.round(td),
        dau_cost: Number((tc / td).toFixed(4)),
        daily_spend: labels.length ? Math.round(tc / labels.length) : 0,
      }};
      country_daily[c] = cost;
    }}
  }});
  const camp_daily = {{}};
  Object.entries(base.camp_daily || {{}}).forEach(([c, camps]) => {{
    const out = {{}};
    Object.entries(camps || {{}}).forEach(([camp, vals]) => {{
      const sliced = sliceArr(vals, s, e);
      if (sliced.some(v => v !== null && v !== undefined)) out[camp] = sliced;
    }});
    if (Object.keys(out).length) camp_daily[c] = out;
  }});
  return {{...base, dates: sliceArr(dates, s, e), labels, country_summary, country_daily, camp_daily}};
}}

// ── Chart factory ────────────────────────────────────────────────────────────
function makeChart(id, datasets, labels, existing, opts) {{
  if (existing) existing.destroy();
  opts = opts || {{}};
  const SHOW_SPEND = opts.spendAxis  !== false;
  const SHOW_ROI   = opts.roiAxis    !== false;
  const SHOW_DL    = opts.dataLabels !== false;
  return new Chart(document.getElementById(id), {{
    type:"bar", data:{{labels,datasets}},
    options:{{
      responsive:true, interaction:{{mode:"index",intersect:false}},
      scales:{{
        x:{{grid:{{color:"rgba(0,0,0,.05)"}},ticks:{{font:{{size:12}}}}}},
        ySpend:{{type:"linear",position:"left",display:SHOW_SPEND,
          title:{{display:true,text:"消耗 (USD)",font:{{size:12}}}},
          grid:{{color:"rgba(0,0,0,.06)"}},
          ticks:{{callback:v=>"$"+v.toLocaleString(),font:{{size:11}}}}}},
        yROI:{{type:"linear",position:"right",display:SHOW_ROI,
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
          display: ctx => SHOW_DL && ctx.dataset.yAxisID === "yROI",
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
  const d = getFilteredData();
  const platformSummary = (d.platform || {{}}).summary || {{}};
  const fallbackSpend = Object.values(d.group_summary).reduce((s,x)=>s+(x.spend||0),0);
  const allSpend = platformSummary.spend != null ? platformSummary.spend : fallbackSpend;
  const dailyAvg = (allSpend / d.labels.length).toFixed(0);
  const platformRoi = platformSummary.roi != null ? platformSummary.roi + "x" : "-";
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
    <div class="kpi"><div class="k-label">Moloco平台 ROI</div>
      <div class="k-val">${{platformRoi}}</div>
      <div class="k-sub" style="color:#64748b">Σ24h_gmv / Σcostdsp</div></div>
    <div class="kpi red"><div class="k-label">需关注项</div>
      <div class="k-val">${{issues7.length}}</div>
      <div class="k-sub" style="color:#dc2626">ROI下滑 / 持续偏低</div></div>
    <div class="kpi green"><div class="k-label">可放量项</div>
      <div class="k-val">${{opps7.length}}</div>
      <div class="k-sub" style="color:#059669">ROI显著提升</div></div>
    <div class="kpi amber"><div class="k-label">异常波动</div>
      <div class="k-val">${{anomalies7.length}}</div>
      <div class="k-sub" style="color:#d97706">需核实数据</div></div>`;

  renderPlatformTrend(d);
  renderSignals();
}}

function renderPlatformTrend(d) {{
  if (!d.platform || !d.platform.ds) return;
  platformChart = makeChart("platformRoiChart", d.platform.ds, d.labels, platformChart);
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
  const d = getFilteredData();
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
  const d = getFilteredData();
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
  const d=getFilteredData();
  const chips=document.getElementById("group-chips");
  chips.innerHTML="";
  Object.keys(d.proj_summary).sort().forEach(lbl=>chips.appendChild(makeChip(lbl,d.proj_summary[lbl],true)));
  Object.entries(d.country_summary).sort((a,b)=>(b[1].spend||0)-(a[1].spend||0))
    .forEach(([lbl,s])=>chips.appendChild(makeChip(lbl,s,false)));

  document.querySelectorAll("#camp-metric button").forEach(b=>
    b.classList.toggle("active", b.dataset.metric===campMetric));

  const area=document.getElementById("camp-chart-area");
  if (!selGroup || !d.camp_ds[selGroup]) {{
    area.innerHTML=`<div class="empty-hint">👆 选择上方某个国家或项目，查看其全部 Campaign 趋势</div>`;
    campChart=null;
    document.getElementById("clear-camp").style.display="none";
    return;
  }}
  document.getElementById("clear-camp").style.display="inline";

  const showSpend = campMetric !== "roi";
  const showRoi   = campMetric !== "spend";
  const ds = d.camp_ds[selGroup].filter(x =>
    (x.yAxisID==="ySpend" && showSpend) || (x.yAxisID==="yROI" && showRoi));
  const metricTxt = campMetric==="both" ? "日消耗（柱）& ROI（线）"
                  : campMetric==="spend" ? "日消耗（柱）" : "日 ROI（线）";
  const axisTxt = campMetric==="both" ? "左轴: 消耗 USD · 右轴: ROI"
                : campMetric==="spend" ? "左轴: 消耗 USD" : "右轴: ROI";
  area.innerHTML=`<div class="chart-box">
    <div class="chart-hd">${{selGroup}} — Campaign ${{metricTxt}}
      <span class="badge">${{axisTxt}} · 图例可点击隐藏</span></div>
    <canvas id="campChart"></canvas>
    <div class="note">图例可点击隐藏/显示单个 Campaign；ROI 线带数值</div>
  </div>`;
  campChart=makeChart("campChart",ds,d.labels,null,
    {{spendAxis:showSpend, roiAxis:showRoi, dataLabels:showRoi}});
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

// ── DAU成本 面板 ─────────────────────────────────────────────────────────────
function selectDauCountry(c) {{
  selDauCountry = c;
  renderDauCost();
}}

function buildDauCostTable(campData, labels) {{
  const entries = Object.entries(campData || {{}})
    .sort((a, b) => {{
      const av = (a[1] || []).filter(v => v !== null && v !== undefined);
      const bv = (b[1] || []).filter(v => v !== null && v !== undefined);
      const aa = av.length ? av.reduce((s, v) => s + Number(v || 0), 0) / av.length : 999;
      const bb = bv.length ? bv.reduce((s, v) => s + Number(v || 0), 0) / bv.length : 999;
      return aa - bb;
    }});
  if (!entries.length) return '<div class="empty-hint">当前筛选下暂无 Campaign DAU成本明细</div>';
  const head = labels.map(l => `<th>${{l}}</th>`).join('');
  const rows = entries.map(([camp, vals]) => {{
    const cells = (vals || []).map(v => `<td>${{v != null ? fmtDauCost(v) : '-'}}</td>`).join('');
    return `<tr><td title="${{camp}}">${{camp}}</td>${{cells}}</tr>`;
  }}).join('');
  return `<div class="table-wrap"><table class="metric-table">
    <thead><tr><th>Campaign</th>${{head}}</tr></thead>
    <tbody>${{rows}}</tbody>
  </table></div>`;
}}

function renderDauCost() {{
  if (!DAU_COST_DATA) return;
  const d = getFilteredDauCostData();
  const summary = d.country_summary || {{}};
  const labels = d.labels || [];

  const cardsEl = document.getElementById('dau-country-cards');
  if (cardsEl) {{
    cardsEl.innerHTML = '';
    Object.entries(summary)
      .sort((a, b) => a[1].dau_cost - b[1].dau_cost)
      .forEach(([c, s]) => {{
        const sel = selDauCountry === c;
        const div = document.createElement('div');
        div.className = 'card' + (sel ? ' selected' : '');
        div.innerHTML = `<div class="name">${{c}}</div>
          <div class="spend">${{fmtDauCost(s.dau_cost)}}</div>
          <div class="roi">总花费: $${{(s.cost||0).toLocaleString()}} · 日均: $${{(s.daily_spend||0).toLocaleString()}}</div>
          <div class="roi">session_dau: ${{(s.dau||0).toLocaleString()}}</div>`;
        div.onclick = () => selectDauCountry(c === selDauCountry ? null : c);
        cardsEl.appendChild(div);
      }});
  }}
  const clearBtn = document.getElementById('clear-dau-country');
  if (clearBtn) clearBtn.style.display = selDauCountry ? 'inline' : 'none';

  const area = document.getElementById('dau-chart-area');
  if (!selDauCountry || !(d.country_daily || {{}})[selDauCountry]) {{
    area.innerHTML = '<div class="empty-hint">👆 点击国家 / 项目卡片，查看 DAU成本日趋势和 Campaign 分天明细</div>';
    if (dauTrendChart) {{ dauTrendChart.destroy(); dauTrendChart = null; }}
    if (dauCampChart)  {{ dauCampChart.destroy();  dauCampChart = null; }}
    return;
  }}

  area.innerHTML = `
    <div class="chart-box">
      <div class="chart-hd">${{selDauCountry}} — DAU成本日趋势
        <span class="badge">Σcost / Σsession_dau</span></div>
      <canvas id="dauTrendChart"></canvas>
      <div class="note">按天汇总该国家 / 项目下所有 campaign 的 cost 与 session_dau</div>
    </div>
    <div class="chart-box">
      <div class="chart-hd">${{selDauCountry}} — Campaign DAU成本
        <span class="badge">分 Campaign 分天</span></div>
      <canvas id="dauCampChart"></canvas>
      <div class="note">图例可点击隐藏 / 显示单个 Campaign</div>
    </div>
    <div class="chart-hd">${{selDauCountry}} — Campaign DAU成本明细表
      <span class="badge">每格 = 当日 cost / session_dau</span></div>
    ${{buildDauCostTable((d.camp_daily || {{}})[selDauCountry], labels)}}`;

  const cData = d.country_daily[selDauCountry];
  if (dauTrendChart) dauTrendChart.destroy();
  dauTrendChart = new Chart(document.getElementById('dauTrendChart'), {{
    type: 'line',
    data: {{ labels, datasets: [{{
      label: selDauCountry + ' DAU成本',
      data: cData,
      borderColor: '#0ea5e9',
      backgroundColor: 'rgba(14,165,233,0.08)',
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
          title: {{ display: true, text: 'DAU成本 (USD)', font: {{ size: 12 }} }},
          ticks: {{ callback: v => fmtDauCost(v), font: {{ size: 11 }} }}, min: 0,
        }},
      }},
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': ' + fmtDauCost(ctx.parsed.y) }} }},
        datalabels: {{
          display: true,
          formatter: v => v != null ? fmtDauCost(v) : null,
          color: '#0ea5e9', font: {{ size: 11, weight: '700' }},
          anchor: 'top', align: 'top', offset: 3,
          backgroundColor: 'rgba(255,255,255,0.85)', borderRadius: 3,
          padding: {{ top: 2, bottom: 2, left: 4, right: 4 }},
        }},
      }},
    }},
  }});

  const campData = ((d.camp_daily || {{}})[selDauCountry]) || {{}};
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
  if (dauCampChart) dauCampChart.destroy();
  dauCampChart = new Chart(document.getElementById('dauCampChart'), {{
    type: 'line',
    data: {{ labels, datasets: campDs }},
    options: {{
      responsive: true,
      interaction: {{ mode: 'index', intersect: false }},
      scales: {{
        x: {{ grid: {{ color: 'rgba(0,0,0,.05)' }}, ticks: {{ font: {{ size: 12 }} }} }},
        y: {{
          title: {{ display: true, text: 'DAU成本 (USD)', font: {{ size: 12 }} }},
          ticks: {{ callback: v => fmtDauCost(v), font: {{ size: 11 }} }}, min: 0,
        }},
      }},
      plugins: {{
        legend: {{ position: 'top', labels: {{ font: {{ size: 12 }}, usePointStyle: true, pointStyleWidth: 12 }} }},
        tooltip: {{ callbacks: {{ label: ctx => {{
          const v = ctx.parsed.y;
          return v != null ? ctx.dataset.label + ': ' + fmtDauCost(v) : null;
        }} }} }},
        datalabels: {{
          display: false,
        }},
      }},
    }},
  }});
}}

// ── DAC成本 面板 ─────────────────────────────────────────────────────────────
function setDacPeriod(n) {{
  dacPeriod = n;
  setPeriod(n);
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
  if (!DAC_DATA) return;
  renderDacSpecial();
  const d       = getFilteredDacData();
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
  if (!DAC_DATA) return;
  const d  = getFilteredDacData();
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
    if (dacSpecRoiChart)   {{ dacSpecRoiChart.destroy();   dacSpecRoiChart   = null; }}
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
    <div class="roi">DAC: ${{ov.dac}}</div>
    <div class="roi">24H GMV: $${{(ov.gmv||0).toLocaleString()}} · ROI: ${{ov.roi!=null?ov.roi+'x':'-'}}</div>`;
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
        <div class="roi">DAC: ${{s.dac}}</div>
        <div class="roi">24H GMV: $${{(s.gmv||0).toLocaleString()}} · ROI: ${{s.roi!=null?s.roi+'x':'-'}}</div>`;
      div.onclick = () => selectDacSpecial(sel ? null : camp);
      cardsEl.appendChild(div);
    }});

  const clearBtn = document.getElementById('clear-dac-special');
  if (clearBtn) clearBtn.style.display = selDacSpecial ? 'inline' : 'none';

  // 趋势图
  if (!selDacSpecial) {{
    area.innerHTML = '<div class="empty-hint">👆 点击「整体」或单个 campaign 卡片，查看 DAC数 / 花费 / DAC成本 / 24H GMV / ROI 分天趋势</div>';
    if (dacSpecSpendChart) {{ dacSpecSpendChart.destroy(); dacSpecSpendChart = null; }}
    if (dacSpecCostChart)  {{ dacSpecCostChart.destroy();  dacSpecCostChart  = null; }}
    if (dacSpecRoiChart)   {{ dacSpecRoiChart.destroy();   dacSpecRoiChart   = null; }}
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
      <div class="chart-hd">${{title}} — 24H GMV & ROI
        <span class="badge">左轴: 24H GMV USD · 右轴: ROI</span></div>
      <canvas id="dacSpecRoiChart"></canvas>
      <div class="note">ROI = 当日 Σ24h_gmv / 当日 Σ花费</div>
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

  if (dacSpecRoiChart) dacSpecRoiChart.destroy();
  dacSpecRoiChart = new Chart(document.getElementById('dacSpecRoiChart'), {{
    type: 'bar',
    data: {{ labels, datasets: [
      {{label:'24H GMV', data: target.gmv_series || [], backgroundColor:'rgba(14,165,233,0.58)',
        borderColor:'#0ea5e9', borderWidth:1, yAxisID:'yGmv', type:'bar'}},
      {{label:'ROI', data: target.roi_series || [], borderColor:'#059669',
        backgroundColor:'transparent', borderWidth:2.5, pointRadius:4, pointHoverRadius:6,
        tension:0.3, yAxisID:'yROI', type:'line', spanGaps:true}},
    ]}},
    options: {{
      responsive:true, interaction:{{mode:'index',intersect:false}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{font:{{size:12}}}}}},
        yGmv:{{type:'linear',position:'left',
          title:{{display:true,text:'24H GMV (USD)',font:{{size:12}}}},
          grid:{{color:'rgba(0,0,0,.06)'}},
          ticks:{{callback:v=>'$'+v.toLocaleString(),font:{{size:11}}}}, min:0}},
        yROI:{{type:'linear',position:'right',
          title:{{display:true,text:'ROI (x)',font:{{size:12}}}},
          grid:{{drawOnChartArea:false}},
          ticks:{{callback:v=>v+'x',font:{{size:11}}}}, min:0}},
      }},
      plugins:{{
        legend:{{position:'top',labels:{{font:{{size:12}},usePointStyle:true,pointStyleWidth:12}}}},
        tooltip:{{mode:'index',intersect:false,
          callbacks:{{label:ctx=>{{
            const v=ctx.parsed.y;
            if(v===null||v===undefined) return null;
            return ctx.dataset.yAxisID==='yGmv'
              ? ctx.dataset.label+': $'+v.toLocaleString()
              : ctx.dataset.label+': '+v+'x';
          }}}}}},
        datalabels:{{display:ctx=>ctx.dataset.yAxisID==='yROI',
          formatter:v=>v!=null?v+'x':null,
          color:'#059669', font:{{size:11,weight:'700'}},
          anchor:'top', align:'top', offset:3,
          backgroundColor:'rgba(255,255,255,0.85)', borderRadius:3,
          padding:{{top:2,bottom:2,left:4,right:4}}}},
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
function setCampMetric(m){{campMetric=m;renderCampaign();}}

function switchTab(name){{
  activeTab=name;
  const names=["home","project","country","campaign","dau","dac"];
  document.querySelectorAll(".tab").forEach((t,i)=>t.classList.toggle("active",names[i]===name));
  document.querySelectorAll(".panel").forEach(p=>p.classList.remove("active"));
  document.getElementById("panel-"+name).classList.add("active");
  // 切到某 tab 时重渲染：图表在隐藏面板里创建会塌成 0 宽，进入时需重画
  if(name==="home") renderHome();
  else if(name==="project") renderProject();
  else if(name==="country") renderCountry();
  else if(name==="campaign") renderCampaign();
  else if(name==="dau") renderDauCost();
  else if(name==="dac") renderDac();
}}

function renderAll(){{
  renderHome();
  renderProject();
  renderCountry();
  renderCampaign();
}}

// 默认进入过去 14 天；日期控件可扩展到最新 sheet 的最早日期
setRangeToLast(14);
setDateInputs();
updatePeriodButtons();
renderAll();
</script>
</body>
</html>"""

# ─── 主程序 ───────────────────────────────────────────────────────────────────
def main():
    no_push = "--no-push" in sys.argv
    records      = load_data()
    data         = build_data(records)
    dac_data     = build_dac_data(records)
    dau_cost_data = build_dau_cost_data(records)
    all_signals, overall_roi = compute_all_signals(records)

    data_json        = json.dumps(data,        ensure_ascii=False)
    all_signals_json = json.dumps(all_signals, ensure_ascii=False)
    dac_json         = json.dumps(dac_data,    ensure_ascii=False)
    dau_cost_json    = json.dumps(dau_cost_data, ensure_ascii=False)
    generated        = datetime.now().strftime("%Y-%m-%d %H:%M")
    html             = generate_html(data_json, all_signals_json, dac_json, dau_cost_json, generated)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    for n in (7, 14, 30):
        dates = get_window(records, n)
        print(f"  {n}天窗口: {dates[0] if dates else '?'} ~ {dates[-1] if dates else '?'} ({len(dates)}天)")
    m7 = all_signals["7"]["meta"]
    print(f"  信号(7天): {m7['n_issues']} 个问题, {m7['n_opps']} 个放量机会")
    print(f"\n✓ 看板已生成: {OUTPUT_FILE}")

    # ── 自动推送到 GitHub ──────────────────────────────────────────────────────
    if no_push:
        print("  (--no-push：仅本地生成 index.html，跳过 git 提交与推送)")
        return
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
