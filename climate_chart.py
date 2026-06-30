#!/usr/bin/env python3
"""
氣候風險圖表自動化工具
使用 Open-Meteo Archive API 抓取城市的月均溫和月均濕度（以當前月為終點、往回滾動 48 個月），填入 Excel 模板。

用法：
  python3 climate_chart.py "城市名稱, 國家" "輸出 Excel 路徑"

範例：
  python3 climate_chart.py "Sukabumi, Indonesia" "/path/to/案件/氣候圖/氣候風險圖表.xlsx"
"""

import io
import base64
import os
import re
import sys
import json
import shutil
import tempfile
import zipfile
import urllib.request
import urllib.parse
import urllib.error
from datetime import date, datetime
from pathlib import Path

_NAS_TEMPLATE = Path(
    "/Volumes/實驗室共用區MRC/#YCT資料區#/4.Smart EA/案件資料區"
    "/[COPY ONLY] 報告號 採樣日 國家 品牌 工廠代號/氣候圖/氣候風險圖表.xlsx"
)
_LOCAL_TEMPLATE = Path(__file__).parent / "氣候風險圖表_template.xlsx"
_ENCRYPTED_TEMPLATE = Path(__file__).parent / "氣候風險圖表_template.xlsx.fernet"
TEMPLATE_PATH = _NAS_TEMPLATE if _NAS_TEMPLATE.exists() else _LOCAL_TEMPLATE
_SECRET_TEMPLATE_PATH: Path | None = None

def get_template_path() -> Path:
    """Return the Excel template path from NAS, Streamlit secret, or local fallback."""
    global _SECRET_TEMPLATE_PATH

    if _NAS_TEMPLATE.exists():
        return _NAS_TEMPLATE

    template_b64 = os.environ.get("CLIMATE_TEMPLATE_XLSX_BASE64")
    if template_b64:
        if _SECRET_TEMPLATE_PATH is None or not _SECRET_TEMPLATE_PATH.exists():
            tmp_dir = Path(tempfile.mkdtemp(prefix="climate-chart-template-"))
            _SECRET_TEMPLATE_PATH = tmp_dir / "氣候風險圖表_template.xlsx"
            _SECRET_TEMPLATE_PATH.write_bytes(base64.b64decode(template_b64))
        return _SECRET_TEMPLATE_PATH

    template_key = os.environ.get("CLIMATE_TEMPLATE_FERNET_KEY")
    if template_key and _ENCRYPTED_TEMPLATE.exists():
        if _SECRET_TEMPLATE_PATH is None or not _SECRET_TEMPLATE_PATH.exists():
            from cryptography.fernet import Fernet

            tmp_dir = Path(tempfile.mkdtemp(prefix="climate-chart-template-"))
            _SECRET_TEMPLATE_PATH = tmp_dir / "氣候風險圖表_template.xlsx"
            encrypted = _ENCRYPTED_TEMPLATE.read_bytes()
            _SECRET_TEMPLATE_PATH.write_bytes(Fernet(template_key.encode()).decrypt(encrypted))
        return _SECRET_TEMPLATE_PATH

    if _LOCAL_TEMPLATE.exists():
        return _LOCAL_TEMPLATE

    raise FileNotFoundError("找不到 Excel 模板。")

def fetch_json(url: str) -> dict:
    with urllib.request.urlopen(url, timeout=30) as r:
        return json.loads(r.read().decode())

def geocode(city_query: str) -> tuple[float, float, str]:
    """回傳 (latitude, longitude, 地點全名)"""
    params = urllib.parse.urlencode({
        "name": city_query,
        "count": 1,
        "language": "en",
        "format": "json",
    })
    data = fetch_json(f"https://geocoding-api.open-meteo.com/v1/search?{params}")
    results = data.get("results")
    if not results:
        raise ValueError(f"找不到城市：{city_query}")
    r = results[0]
    name = r.get("name", city_query)
    country = r.get("country", "")
    return r["latitude"], r["longitude"], f"{name}, {country}"

WINDOW_MONTHS = 48  # 滾動視窗長度（48 = 每個月份都平均到 4 個年度）


def get_climate_data(lat: float, lon: float) -> dict:
    """
    回傳 {year: {month: {"temp": float, "humidity": float}}}
    資料範圍：以「當前月」為終點，往回剛好 WINDOW_MONTHS 個月的滾動視窗。
      例：當前 2026-06 → 2022-07 ~ 2026-06（48 個月）。
    用滾動視窗（而非行事曆整年）是為了讓每個月份都平均到相同年數，
    季節曲線不會被頭尾不完整的半年帶偏。
    """
    today = date.today()
    end_year, end_month = today.year, today.month
    end_idx = end_year * 12 + (end_month - 1)
    start_idx = end_idx - (WINDOW_MONTHS - 1)
    start_year, start_month = divmod(start_idx, 12)
    start_month += 1
    start_date = f"{start_year}-{start_month:02d}-01"
    end_date = today.strftime("%Y-%m-%d")

    def _archive_url(ed: str) -> str:
        params = urllib.parse.urlencode({
            "latitude": lat,
            "longitude": lon,
            "start_date": start_date,
            "end_date": ed,
            "daily": "temperature_2m_mean,relative_humidity_2m_mean,precipitation_sum",
            "timezone": "auto",
        })
        return f"https://archive-api.open-meteo.com/v1/archive?{params}"

    # archive API 有約 1 天延遲：請求「今天」會回 400，並在訊息裡告知可用的最新日期。
    # 抓到就退到那天重抓，讓視窗終點維持在最近一筆可用資料。
    try:
        data = fetch_json(_archive_url(end_date))
    except urllib.error.HTTPError as e:
        if e.code != 400:
            raise
        m = re.search(r"to (\d{4}-\d{2}-\d{2})", e.read().decode())
        if not m:
            raise
        end_date = m.group(1)
        data = fetch_json(_archive_url(end_date))

    daily = data["daily"]
    dates = daily["time"]
    temps = daily["temperature_2m_mean"]
    humids = daily["relative_humidity_2m_mean"]
    precips = daily.get("precipitation_sum") or [None] * len(dates)

    # 按年月分組：溫度/濕度取月均，降雨取月總量
    sums: dict = {}
    counts: dict = {}
    rain: dict = {}
    for d, t, h, p in zip(dates, temps, humids, precips):
        if t is None or h is None:
            continue
        dt = datetime.strptime(d, "%Y-%m-%d")
        y, m = dt.year, dt.month
        sums.setdefault(y, {}).setdefault(m, {"temp": 0.0, "humidity": 0.0})
        counts.setdefault(y, {}).setdefault(m, 0)
        sums[y][m]["temp"] += t
        sums[y][m]["humidity"] += h
        counts[y][m] += 1
        if p is not None:
            rain.setdefault(y, {}).setdefault(m, 0.0)
            rain[y][m] += p

    result: dict = {}
    for y in sums:
        result[y] = {}
        for m in sums[y]:
            n = counts[y][m]
            result[y][m] = {
                "temp": round(sums[y][m]["temp"] / n, 1),
                "humidity": round(sums[y][m]["humidity"] / n, 1),
                "precip": round(rain.get(y, {}).get(m, 0.0), 1),  # 該月降雨總量 (mm)
            }
    return result, start_year, end_year


def monthly_rainfall(climate: dict, start_year: int, end_year: int) -> dict:
    """每月平均降雨 (mm)：把各年同月的月總量平均。回傳 {month: mm}。"""
    rain = {}
    for m in range(1, 13):
        vals = [climate[y][m]["precip"]
                for y in range(start_year, end_year + 1)
                if climate.get(y, {}).get(m, {}).get("precip") is not None]
        rain[m] = round(sum(vals) / len(vals), 1) if vals else 0.0
    return rain


def classify_seasons(rain: dict) -> dict:
    """
    從每月平均降雨判定乾濕季，收斂成「單一濕季」的 3 區塊模型。

    回傳 {pattern, blocks, threshold, wet_months}
      pattern: "DWD"（乾-濕-乾，濕季在年中）或 "WDW"（濕-乾-濕，濕季跨年界）
      blocks : [(type, start_month, end_month), ...]，涵蓋 1-12 月、相鄰交替
      threshold: 判定門檻 = 全年月均降雨 (mm)

    判定規則（透明、可被使用者覆寫）：某月降雨 ≥ 全年月均 → 濕月；
    取最長的「連續濕月」（環狀，可跨 12→1）當作主濕季，其餘為乾季。
    """
    months = list(range(1, 13))
    threshold = round(sum(rain[m] for m in months) / 12, 1)
    wet = set(m for m in months if rain[m] >= threshold)

    best = (0, None, None)  # (length, start, end)
    for start in months:
        prev = 12 if start == 1 else start - 1
        if start not in wet or prev in wet:   # 只從 run 起點算，避免重複
            continue
        length, mm, seq = 0, start, []
        while mm in wet and length < 12:
            seq.append(mm); length += 1
            mm = 1 if mm == 12 else mm + 1
        if length > best[0]:
            best = (length, seq[0], seq[-1])

    length, ws, we = best
    if length == 0:
        return {"pattern": "DWD", "blocks": [("dry", 1, 12)],
                "threshold": threshold, "wet_months": []}
    if length >= 12:
        return {"pattern": "WDW", "blocks": [("wet", 1, 12)],
                "threshold": threshold, "wet_months": months}

    wet_months, mm = [], ws
    while True:
        wet_months.append(mm)
        if mm == we:
            break
        mm = 1 if mm == 12 else mm + 1

    if ws <= we:   # 濕季在年中 → 乾-濕-乾
        blocks = ([("dry", 1, ws - 1)] if ws > 1 else []) \
            + [("wet", ws, we)] \
            + ([("dry", we + 1, 12)] if we < 12 else [])
        pattern = "DWD"
    else:          # 濕季跨年界 → 濕-乾-濕
        blocks = [("wet", 1, we), ("dry", we + 1, ws - 1), ("wet", ws, 12)]
        pattern = "WDW"
    return {"pattern": pattern, "blocks": blocks,
            "threshold": threshold, "wet_months": wet_months}


_MONTH_ABBR = ["Jan.", "Feb.", "Mar.", "Apr.", "May", "Jun.",
               "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."]


def _patch_drawing_xml(xml: str, full_name: str, start_year: int, start_month: int, end_year: int, end_month: int) -> str:
    """替換 drawing XML 中的城市名、年份、月份佔位符。"""
    from lxml import etree

    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    start_month_abbr = _MONTH_ABBR[start_month - 1]
    end_month_abbr = _MONTH_ABBR[end_month - 1]

    root = etree.fromstring(xml.encode("utf-8"))

    for p_el in root.iter(f"{{{A}}}p"):
        t_els = list(p_el.iter(f"{{{A}}}t"))
        texts = [t.text or "" for t in t_els]
        joined = "".join(texts)

        # 英文標題段落：「The average climate in #城市名」+ 中間 # + 「 from ... to ...」
        if "The average climate in" in joined and "月至" not in joined:
            for t_el in t_els:
                text = t_el.text or ""
                if text.startswith("The average climate in #"):
                    t_el.text = f"The average climate in {full_name}"
                elif re.match(r"^ from [A-Za-z]+\.? \d{4} to [A-Za-z]+\.? \d{4}\.$", text):
                    t_el.text = f" from {start_month_abbr} {start_year} to {end_month_abbr} {end_year}."
                elif text == "#":
                    t_el.text = ""

        # 中文年份 + 城市名段落：含「月至」和「平均氣候條件」
        elif "月至" in texts and "平均氣候條件" in texts:
            yue_zhi = texts.index("月至")

            # start_year + start_month 在「月至」前（格式：YEAR 年 MONTH 月至）
            if yue_zhi >= 3 and texts[yue_zhi - 2] == "年":
                t_els[yue_zhi - 3].text = str(start_year)
                t_els[yue_zhi - 1].text = str(start_month)

            # end_year 在「月至」後 1 位（格式：月至 YEAR 年 MONTH 月）
            if yue_zhi + 2 < len(texts) and texts[yue_zhi + 2] == "年":
                t_els[yue_zhi + 1].text = str(end_year)

            # end_month 在「月至」後 3 位（後面跟「月」）
            if yue_zhi + 4 < len(texts) and texts[yue_zhi + 4] == "月":
                t_els[yue_zhi + 3].text = str(end_month)

            # 城市名：清空前後 # 佔位符，替換中間的城市名文字
            hash_indices = [i for i, t in enumerate(texts) if t == "#"]
            if len(hash_indices) >= 2:
                t_els[hash_indices[0]].text = ""
                t_els[hash_indices[1]].text = ""
                city_idx = hash_indices[0] + 1
                if city_idx < hash_indices[1]:
                    t_els[city_idx].text = full_name

    return etree.tostring(root, encoding="unicode")


# ── 乾濕季色塊自動化 ────────────────────────────────────────────────
_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# 乾濕季型態 → 模板版型 drawing / 分頁索引（0=AVG CAL,1=DWD,2=WDW）
_PATTERN_DRAWING = {"DWD": "xl/drawings/drawing1.xml", "WDW": "xl/drawings/drawing2.xml"}
_PATTERN_SHEET_INDEX = {"DWD": 1, "WDW": 2}

# 圖例月份縮寫（沿用模板字樣：5 月無點、9 月用 Sept.）
_LEG_ABBR = {1: "Jan.", 2: "Feb.", 3: "Mar.", 4: "Apr.", 5: "May", 6: "Jun.",
             7: "Jul.", 8: "Aug.", 9: "Sept.", 10: "Oct.", 11: "Nov.", 12: "Dec."}
_MONTH_NUM = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7,
              "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}


def _shape_text(sp):
    return "".join(t.text or "" for t in sp.iter(f"{{{_NS_A}}}t")).strip()


def _set_shape_text(sp, s):
    ts = list(sp.iter(f"{{{_NS_A}}}t"))
    if ts:
        ts[0].text = s
        for e in ts[1:]:
            e.text = ""


def _shape_off_ext(sp):
    xf = sp.find(f".//{{{_NS_A}}}xfrm")
    return xf.find(f"{{{_NS_A}}}off"), xf.find(f"{{{_NS_A}}}ext")


def _emu_chain(el, parent):
    """蒐集 el 到 root 的群組變換鏈（innermost first），用來合成絕對座標。"""
    chain = []
    p = parent.get(el)
    while p is not None:
        if p.tag == f"{{{_NS_XDR}}}grpSp":
            xf = p.find(f"{{{_NS_XDR}}}grpSpPr/{{{_NS_A}}}xfrm")
            if xf is not None:
                o = xf.find(f"{{{_NS_A}}}off"); e = xf.find(f"{{{_NS_A}}}ext")
                co = xf.find(f"{{{_NS_A}}}chOff"); ce = xf.find(f"{{{_NS_A}}}chExt")
                chain.append((float(o.get("x")), float(co.get("x")),
                              float(e.get("cx")), float(ce.get("cx"))))
        p = parent.get(p)
    return chain


def _to_abs(x, chain):
    for ox, cox, ecx, cecx in chain:
        x = ox + (x - cox) * (ecx / cecx)
    return x


def _set_abs(sp, chain, abs_left, abs_right):
    """把絕對左右緣換算回 sp 所在群組座標系，寫入 off.x / ext.cx。"""
    b0 = _to_abs(0, chain); a0 = _to_abs(1, chain) - b0
    off, ext = _shape_off_ext(sp)
    ll = (abs_left - b0) / a0; lr = (abs_right - b0) / a0
    off.set("x", str(int(round(ll))))
    ext.set("cx", str(int(round(lr - ll))))


def _reposition_self(sp, a_old, b_old, a_new, b_new):
    """單一座標系內自我校準位移（給六邊形圖例用）。"""
    off, ext = _shape_off_ext(sp)
    x = int(off.get("x")); w = int(ext.get("cx"))
    per = w / ((b_old - a_old) + 1)
    base = x - (a_old - 1) * per
    off.set("x", str(int(round(base + (a_new - 1) * per))))
    ext.set("cx", str(int(round(((b_new - a_new) + 1) * per))))


def _set_legend_text(sp, typ, a, b, cn):
    """改六邊形圖例文字，保留模板「類型 <br> 月份」的兩行結構。
    以 <a:br> 為界：換行前第一個 run 填類型、其餘清空；換行後填月份
    （EN 單一 run；CN 的純數字 run 依序填起訖月，保留「月」「- 」run）。"""
    for p in sp.iter(f"{{{_NS_A}}}p"):
        before, after, seen_br = [], [], False
        for el in list(p):
            tag = el.tag.split("}")[-1]
            if tag == "br":
                seen_br = True
            elif tag == "r":
                t = el.find(f"{{{_NS_A}}}t")
                if t is not None:
                    (after if seen_br else before).append(t)
        if not before:
            continue
        before[0].text = ("濕季" if typ == "wet" else "乾季") if cn \
            else ("Wet" if typ == "wet" else "Dry")
        for t in before[1:]:
            t.text = ""
        if cn:
            nums = [t for t in after if re.fullmatch(r"\d+", (t.text or "").strip())]
            if len(nums) >= 2:
                nums[0].text, nums[1].text = str(a), str(b)
            elif after:
                after[0].text = f"{a}月- {b}月"
                for t in after[1:]:
                    t.text = ""
        elif after:
            after[0].text = f"{_LEG_ABBR[a]} - {_LEG_ABBR[b]}"
            for t in after[1:]:
                t.text = ""
        return


def _parse_legend_months(t):
    m = re.match(r"(?:乾季|濕季)(\d+)月\s*-\s*(\d+)月", t)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"(?:Dry|Wet)([A-Za-z]+)\.?\s*-\s*([A-Za-z]+)\.?", t)
    if m:
        a = _MONTH_NUM.get(m.group(1).lower().rstrip(".").lower())
        b = _MONTH_NUM.get(m.group(2).lower().rstrip("."))
        if a and b:
            return a, b
    return None


def _band_color(sp):
    """背景帶矩形回傳填色 (FFC000 黃乾 / 2F5597 藍濕)，否則 None。"""
    geom = sp.find(f".//{{{_NS_A}}}prstGeom")
    if geom is None or geom.get("prst") != "rect":
        return None
    fill = sp.find(f"{{{_NS_XDR}}}spPr/{{{_NS_A}}}solidFill/{{{_NS_A}}}srgbClr")
    return fill.get("val") if fill is not None else None


def _read_chart_layouts(zin, drawing_name):
    """讀某 drawing 的 rels + 各 chart 的 manualLayout，回傳 {rId: (x, w)}。"""
    import posixpath
    from lxml import etree
    base = posixpath.basename(drawing_name)
    try:
        rels = etree.fromstring(zin.read(f"xl/drawings/_rels/{base}.rels"))
    except KeyError:
        return {}
    out = {}
    for r in rels:
        tgt = r.get("Target")
        if "chart" not in tgt:
            continue
        try:
            cx = zin.read("xl/charts/" + posixpath.basename(tgt)).decode("utf-8")
        except KeyError:
            continue
        mx = re.search(r'<c:x val="([\d.eE+-]+)"', cx)
        mw = re.search(r'<c:w val="([\d.eE+-]+)"', cx)
        if mx and mw:
            out[r.get("Id")] = (float(mx.group(1)), float(mw.group(1)))
    return out


def _ref_chart_axis(anc, parent, chart_layouts):
    """從 anchor 內參考圖表的 manualLayout 算出 (innerL, slot) 絕對座標。"""
    for gf in anc.iter(f"{{{_NS_XDR}}}graphicFrame"):
        cid = gf.find(f".//{{{_NS_C}}}chart")
        rid = cid.get(f"{{{_NS_R}}}id") if cid is not None else None
        if rid in chart_layouts:
            xf = gf.find(f"{{{_NS_XDR}}}xfrm")
            off = xf.find(f"{{{_NS_A}}}off"); ext = xf.find(f"{{{_NS_A}}}ext")
            chain = _emu_chain(gf, parent)
            fx = _to_abs(float(off.get("x")), chain)
            fr = _to_abs(float(off.get("x")) + float(ext.get("cx")), chain)
            ix, iw = chart_layouts[rid]
            return fx + ix * (fr - fx), iw * (fr - fx) / 12.0
    return None


def _place_legend(hexes, blocks):
    """六邊形圖例：自我校準位移 + 改字（依現有文字判語系）。"""
    n = min(len(hexes), len(blocks))
    for i in range(n):
        sp = hexes[i]; typ, a, b = blocks[i]
        cur = _shape_text(sp)
        cn = bool(re.match(r"(乾季|濕季)", cur))
        old = _parse_legend_months(cur)
        if old:
            _reposition_self(sp, old[0], old[1], a, b)
        _set_legend_text(sp, typ, a, b, cn)
    for i in range(n, len(hexes)):       # 多餘六邊形隱藏
        _, ext = _shape_off_ext(hexes[i]); ext.set("cx", "0")


def _place_bands(anc, parent, blocks, axis):
    """背景帶：對齊圖表真實月份刻度，外緣固定貼齊繪圖區。"""
    innerL, slot = axis
    rects = []
    for sp in anc.iter(f"{{{_NS_XDR}}}sp"):
        if _shape_text(sp):
            continue
        color = _band_color(sp)
        off, ext = _shape_off_ext(sp)
        if color not in ("FFC000", "2F5597") or ext is None:
            continue
        cy = ext.get("cy")
        if not cy or int(cy) < 2000000:
            continue
        chain = _emu_chain(sp, parent)
        rects.append({"sp": sp, "chain": chain,
                      "absL": _to_abs(int(off.get("x")), chain)})
    if len(rects) < len(blocks):
        return
    rects.sort(key=lambda r: r["absL"])
    plot_left = rects[0]["absL"]
    off_n, ext_n = _shape_off_ext(rects[-1]["sp"])
    plot_right = _to_abs(int(off_n.get("x")) + int(ext_n.get("cx")), rects[-1]["chain"])

    def boundary(k):
        return innerL + k * slot

    n = min(len(rects), len(blocks))
    for i in range(n):
        r = rects[i]; _typ, a, b = blocks[i]
        left = plot_left if a == 1 else boundary(a - 1)
        right = plot_right if b == 12 else boundary(b)
        _set_abs(r["sp"], r["chain"], left, right)
    for i in range(n, len(rects)):       # 多餘背景帶隱藏
        _, ext = _shape_off_ext(rects[i]["sp"]); ext.set("cx", "0")


def _apply_seasons_to_drawing(xml, blocks, chart_layouts):
    """在指定版型 drawing 上，重定位乾濕季色塊與圖例。"""
    from lxml import etree
    root = etree.fromstring(xml.encode("utf-8"))
    parent = {c: p for p in root.iter() for c in p}
    for anc in [a for a in root if a.tag == f"{{{_NS_XDR}}}twoCellAnchor"]:
        hexes = [sp for sp in anc.iter(f"{{{_NS_XDR}}}sp")
                 if re.match(r"(Dry|Wet)[A-Za-z]|(乾季|濕季)\d", _shape_text(sp))]
        for pred in (r"(Dry|Wet)", r"(乾季|濕季)"):
            grp = sorted([h for h in hexes if re.match(pred, _shape_text(h))],
                         key=lambda s: int(_shape_off_ext(s)[0].get("x")))
            _place_legend(grp, blocks)
        axis = _ref_chart_axis(anc, parent, chart_layouts)
        if axis:
            _place_bands(anc, parent, blocks, axis)
    return etree.tostring(root, encoding="unicode")


def _write_rainfall_block(sheet_xml, climate, start_year, end_year, rain, seasons):
    """在 AVG CAL. 表寫入完整降雨量區塊（跟溫度/濕度同款：各年資料 + 平均），
    並附上每月乾濕標記與判定摘要，讓同仁看得到原始數據自行核對。"""
    from lxml import etree
    root = etree.fromstring(sheet_xml.encode("utf-8"))
    sd = root.find(f"{{{_NS_S}}}sheetData")
    cols = list("BCDEFGHIJKLM")
    months_en = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    wet = set(seasons["wet_months"])

    def row(rn):
        r = etree.SubElement(sd, f"{{{_NS_S}}}row"); r.set("r", str(rn)); return r

    def num(r, ref, val):
        c = etree.SubElement(r, f"{{{_NS_S}}}c"); c.set("r", ref)
        etree.SubElement(c, f"{{{_NS_S}}}v").text = str(val)

    def txt(r, ref, s):
        c = etree.SubElement(r, f"{{{_NS_S}}}c"); c.set("r", ref); c.set("t", "inlineStr")
        is_ = etree.SubElement(c, f"{{{_NS_S}}}is")
        etree.SubElement(is_, f"{{{_NS_S}}}t").text = s

    rn = 24
    # 標題
    txt(row(rn), f"A{rn}", "降雨量 RAINFALL (mm)"); rn += 1
    # 月份表頭
    hdr = row(rn); txt(hdr, f"A{rn}", "年份")
    for i, col in enumerate(cols):
        txt(hdr, f"{col}{rn}", months_en[i])
    rn += 1
    # 各年資料列（缺月留空，跟溫度/濕度一致）
    for year in range(start_year, end_year + 1):
        r = row(rn); num(r, f"A{rn}", year)
        for i, col in enumerate(cols):
            md = climate.get(year, {}).get(i + 1)
            if md and md.get("precip") is not None:
                num(r, f"{col}{rn}", md["precip"])
        rn += 1
    # 平均列
    avg = row(rn); txt(avg, f"A{rn}", "平均 AVG")
    for i, col in enumerate(cols):
        num(avg, f"{col}{rn}", rain[i + 1])
    rn += 1
    # 乾濕季標記列
    wd = row(rn); txt(wd, f"A{rn}", "乾濕季 WET/DRY")
    for i, col in enumerate(cols):
        txt(wd, f"{col}{rn}", "濕" if (i + 1) in wet else "乾")
    rn += 1
    # 判定摘要
    desc = "乾-濕-乾 Dry-Wet-Dry" if seasons["pattern"] == "DWD" else "濕-乾-濕 Wet-Dry-Wet"
    wm = seasons["wet_months"]
    wet_txt = f"{wm[0]}-{wm[-1]} 月" if wm else "無明顯濕季"
    txt(row(rn), f"A{rn}",
        f"自動判定：{desc}｜濕季 {wet_txt}｜門檻 {seasons['threshold']:.0f}mm"
        "（門檻=全年月均；濕月=雨量≥門檻。如有出入請直接於圖表微調）")

    rows = sorted(sd.findall(f"{{{_NS_S}}}row"), key=lambda r: int(r.get("r")))
    for r in rows:
        sd.remove(r)
    for r in rows:
        sd.append(r)
    return etree.tostring(root, encoding="unicode")


def _set_active_sheet(workbook_xml, idx):
    """把開檔時的作用分頁設成 idx（讓乾/濕版型分頁自動跳到最前）。"""
    from lxml import etree
    root = etree.fromstring(workbook_xml.encode("utf-8"))
    wv = root.find(f"{{{_NS_S}}}bookViews/{{{_NS_S}}}workbookView")
    if wv is not None:
        wv.set("activeTab", str(idx))
    return etree.tostring(root, encoding="unicode")


def fill_excel(climate: dict, start_year: int, end_year: int, output_path: Path,
               full_name: str = "", seasons: dict = None, rain: dict = None):
    """
    直接操作 xlsx zip 內的 sheet XML，保留所有圖表檔案。
    不使用 openpyxl（會丟圖表）或 xlwings（需要 Excel 授權）。
    """
    col_letters = list("BCDEFGHIJKLM")  # B=Jan … M=Dec
    years = list(range(start_year, end_year + 1))

    # 建立 {cell_ref: value} 對照表
    data: dict[str, float] = {}
    for row_offset, year in enumerate(years):
        for month_idx, col in enumerate(col_letters):
            month = month_idx + 1
            md = climate.get(year, {}).get(month)
            if md:
                data[f"{col}{3 + row_offset}"] = md["temp"]      # 溫度 row 3-7
                data[f"{col}{12 + row_offset}"] = md["humidity"]  # 濕度 row 12-16

    start_month = min(climate.get(start_year, {}).keys(), default=1)
    end_month = max(climate.get(end_year, {}).keys(), default=12)

    # 讀模板 zip，改 sheet1.xml 數據 + drawing XMLs 文字佔位符
    template_path = get_template_path()
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:

        # 乾濕季：決定版型 drawing、預讀該版型的圖表佈局（算月份刻度用）
        target_drawing = None
        chart_layouts = {}
        if seasons:
            target_drawing = _PATTERN_DRAWING.get(seasons["pattern"])
            chart_layouts = _read_chart_layouts(zin, target_drawing)

        for item in zin.infolist():
            raw = zin.read(item.filename)
            name = item.filename

            if name == "xl/worksheets/sheet1.xml":
                xml = _patch_sheet_xml(raw.decode("utf-8"), data)
                if rain is not None and seasons is not None:
                    xml = _write_rainfall_block(xml, climate, start_year, end_year, rain, seasons)
                raw = xml.encode("utf-8")
            elif name == "xl/styles.xml":
                raw = _clear_mask_fill(raw.decode("utf-8")).encode("utf-8")
            elif name == "xl/workbook.xml":
                xml = _set_full_calc_on_load(raw.decode("utf-8"))
                if seasons is not None:
                    xml = _set_active_sheet(xml, _PATTERN_SHEET_INDEX[seasons["pattern"]])
                raw = xml.encode("utf-8")
            elif name == "xl/calcChain.xml":
                continue  # 讓 Excel 重建，避免過期的 calc chain 導致問題
            elif name in ("xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml"):
                xml = raw.decode("utf-8")
                if full_name:
                    xml = _patch_drawing_xml(xml, full_name, start_year, start_month, end_year, end_month)
                if seasons is not None and name == target_drawing:
                    xml = _apply_seasons_to_drawing(xml, seasons["blocks"], chart_layouts)
                raw = xml.encode("utf-8")

            zout.writestr(item, raw)


def _col_to_num(col: str) -> int:
    """'A'→1, 'B'→2, …, 'M'→13"""
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n

def _patch_sheet_xml(sheet_xml: str, data: dict[str, float]) -> str:
    """用 lxml 解析 sheet1.xml，注入資料格，保留所有 namespace。"""
    from lxml import etree

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    root = etree.fromstring(sheet_xml.encode("utf-8"))
    sheet_data = root.find(f"{{{NS}}}sheetData")

    # 以 row 號為 key 建立快速查詢
    rows_by_num: dict[int, etree._Element] = {
        int(r.get("r")): r for r in sheet_data.findall(f"{{{NS}}}row")
    }

    # 按 row 分組
    by_row: dict[int, dict[str, float]] = {}
    for ref, val in data.items():
        col = re.match(r"([A-Z]+)", ref).group(1)
        row_num = int(re.search(r"(\d+)", ref).group(1))
        by_row.setdefault(row_num, {})[col] = val

    for row_num in sorted(by_row):
        row_el = rows_by_num.get(row_num)
        if row_el is None:
            row_el = etree.SubElement(sheet_data, f"{{{NS}}}row")
            row_el.set("r", str(row_num))
            rows_by_num[row_num] = row_el

        # 現有格以 r 屬性為 key
        existing: dict[str, etree._Element] = {
            c.get("r"): c for c in row_el.findall(f"{{{NS}}}c")
        }

        for col, val in by_row[row_num].items():
            ref = f"{col}{row_num}"
            cell_el = existing.get(ref)

            if cell_el is None:
                cell_el = etree.SubElement(row_el, f"{{{NS}}}c")
                cell_el.set("r", ref)
                existing[ref] = cell_el

            # 移除 string type，數字格不需要
            if "t" in cell_el.attrib:
                del cell_el.attrib["t"]

            # 移除舊 value 與 formula
            for tag in ("v", "f"):
                old = cell_el.find(f"{{{NS}}}{tag}")
                if old is not None:
                    cell_el.remove(old)

            v_el = etree.SubElement(cell_el, f"{{{NS}}}v")
            v_el.text = str(val)

        # 確保同一 row 內的格依欄位排序（xlsx 規範要求）
        cells = row_el.findall(f"{{{NS}}}c")
        cells.sort(key=lambda c: _col_to_num(re.match(r"([A-Z]+)", c.get("r")).group(1)))
        for c in cells:
            row_el.remove(c)
        for c in cells:
            row_el.append(c)

    # 確保 sheetData 內的 row 也依序排列
    all_rows = sheet_data.findall(f"{{{NS}}}row")
    all_rows.sort(key=lambda r: int(r.get("r")))
    for r in all_rows:
        sheet_data.remove(r)
    for r in all_rows:
        sheet_data.append(r)

    return etree.tostring(root, encoding="unicode")


def _set_full_calc_on_load(workbook_xml: str) -> str:
    """在 workbook.xml 的 calcPr 加上 fullCalcOnLoad='1'，讓 Excel 開檔時完整重算。"""
    if "fullCalcOnLoad" in workbook_xml:
        return workbook_xml
    return workbook_xml.replace(
        "<calcPr", '<calcPr fullCalcOnLoad="1"', 1
    )


def _clear_mask_fill(styles_xml: str) -> str:
    """把模板用來「遮住無資料月份」的黑色填滿（solid + fgColor theme=1）改成無填滿。

    舊版行事曆年設計用黑底遮住當年未來月份；改成滾動視窗後，這塊黑底會蓋住
    視窗頭尾真正有值的月份（如 2026 的 4–6 月），所以一律清成無填滿。
    只動黑色 solid 填滿，黃色高亮（fgColor=FFFFFF00）等其他填滿不受影響。
    """
    from lxml import etree

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    root = etree.fromstring(styles_xml.encode("utf-8"))
    for pf in root.iter(f"{{{NS}}}patternFill"):
        if pf.get("patternType") != "solid":
            continue
        fg = pf.find(f"{{{NS}}}fgColor")
        if fg is not None and fg.get("theme") == "1":
            pf.set("patternType", "none")
            for child in list(pf):
                pf.remove(child)
    return etree.tostring(root, encoding="unicode")

def main():
    if len(sys.argv) < 2:
        print("用法：python3 climate_chart.py \"城市, 國家\" [輸出路徑]")
        sys.exit(1)

    city_query = sys.argv[1]
    print(f"🔍 查詢城市座標：{city_query}")
    lat, lon, full_name = geocode(city_query)
    print(f"   → {full_name}（{lat}, {lon}）")

    print("📡 抓取 Open-Meteo 氣候歷史資料...")
    climate, start_year, end_year = get_climate_data(lat, lon)
    print(f"   → 完成（{start_year}–{end_year}）")

    # 印出預覽
    print("\n── 月均溫 (°C) ──")
    header = "年份  " + "  ".join(f"{m:>5}" for m in ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
    print(header)
    for year in range(start_year, end_year + 1):
        row = f"{year}  "
        for m in range(1, 13):
            v = climate.get(year, {}).get(m, {}).get("temp")
            row += f"  {v:>5}" if v is not None else "    —"
        print(row)

    print("\n── 月均濕度 (%) ──")
    print(header)
    for year in range(start_year, end_year + 1):
        row = f"{year}  "
        for m in range(1, 13):
            v = climate.get(year, {}).get(m, {}).get("humidity")
            row += f"  {v:>5}" if v is not None else "    —"
        print(row)

    # 乾濕季判定
    rain = monthly_rainfall(climate, start_year, end_year)
    seasons = classify_seasons(rain)
    print("\n── 月均雨量 (mm) ──")
    print(header)
    print("雨量  " + "  ".join(f"{rain[m]:>5.0f}" for m in range(1, 13)))
    pat = "乾-濕-乾 (DWD)" if seasons["pattern"] == "DWD" else "濕-乾-濕 (WDW)"
    wm = seasons["wet_months"]
    wet_txt = f"{wm[0]}-{wm[-1]} 月" if wm else "無明顯濕季"
    print(f"   → 判定：{pat}｜濕季 {wet_txt}｜門檻 {seasons['threshold']:.0f}mm")
    print(f"   → 區塊：{seasons['blocks']}")

    # 決定輸出路徑
    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        safe_name = city_query.replace(",", "").replace(" ", "_")
        output_path = Path(f"/tmp/氣候風險圖表_{safe_name}.xlsx")

    print(f"\n📊 填入 Excel 模板 → {output_path}")
    fill_excel(climate, start_year, end_year, output_path,
               full_name=full_name, seasons=seasons, rain=rain)
    print(f"✅ 完成：{output_path}（已自動標乾濕季色塊，使用 {seasons['pattern']} 版型分頁）")

if __name__ == "__main__":
    main()
