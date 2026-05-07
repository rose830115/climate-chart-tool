#!/usr/bin/env python3
"""
氣候風險圖表自動化工具
使用 Open-Meteo Archive API 抓取城市近 5 年月均溫和月均濕度，填入 Excel 模板。

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

def get_climate_data(lat: float, lon: float) -> dict:
    """
    回傳 {year: {month: {"temp": float, "humidity": float}}}
    年份範圍：前 5 年到今年
    """
    today = date.today()
    end_year = today.year
    start_year = end_year - 4
    start_date = f"{start_year}-01-01"
    end_date = today.strftime("%Y-%m-%d")

    params = urllib.parse.urlencode({
        "latitude": lat,
        "longitude": lon,
        "start_date": start_date,
        "end_date": end_date,
        "daily": "temperature_2m_mean,relative_humidity_2m_mean",
        "timezone": "auto",
    })
    data = fetch_json(f"https://archive-api.open-meteo.com/v1/archive?{params}")

    daily = data["daily"]
    dates = daily["time"]
    temps = daily["temperature_2m_mean"]
    humids = daily["relative_humidity_2m_mean"]

    # 按年月分組加總
    sums: dict = {}
    counts: dict = {}
    for d, t, h in zip(dates, temps, humids):
        if t is None or h is None:
            continue
        dt = datetime.strptime(d, "%Y-%m-%d")
        y, m = dt.year, dt.month
        sums.setdefault(y, {}).setdefault(m, {"temp": 0.0, "humidity": 0.0})
        counts.setdefault(y, {}).setdefault(m, 0)
        sums[y][m]["temp"] += t
        sums[y][m]["humidity"] += h
        counts[y][m] += 1

    result: dict = {}
    for y in sums:
        result[y] = {}
        for m in sums[y]:
            n = counts[y][m]
            result[y][m] = {
                "temp": round(sums[y][m]["temp"] / n, 1),
                "humidity": round(sums[y][m]["humidity"] / n, 1),
            }
    return result, start_year, end_year

def fill_excel(climate: dict, start_year: int, end_year: int, output_path: Path):
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

    # 讀模板 zip，只改 sheet1.xml（AVG CAL.），其他全數保留
    template_path = get_template_path()
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            raw = zin.read(item.filename)

            if item.filename == "xl/worksheets/sheet1.xml":
                raw = _patch_sheet_xml(raw.decode("utf-8"), data).encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                raw = _set_full_calc_on_load(raw.decode("utf-8")).encode("utf-8")
            elif item.filename == "xl/calcChain.xml":
                continue  # 讓 Excel 重建，避免過期的 calc chain 導致問題

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

    # 決定輸出路徑
    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        safe_name = city_query.replace(",", "").replace(" ", "_")
        output_path = Path(f"/tmp/氣候風險圖表_{safe_name}.xlsx")

    print(f"\n📊 填入 Excel 模板 → {output_path}")
    fill_excel(climate, start_year, end_year, output_path)
    print(f"✅ 完成：{output_path}")

if __name__ == "__main__":
    main()
