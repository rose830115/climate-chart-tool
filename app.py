"""
氣候風險圖表 Web 工具
輸入工廠城市名稱，自動產生填好的氣候風險圖表 Excel。
"""

import io
import hmac
import os
import shutil
import tempfile
from pathlib import Path

import streamlit as st

# ── 路徑設定 ──────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
CLIMATE_SCRIPT = SCRIPT_DIR / "climate_chart.py"

# 動態 import 主腳本的函式
import sys
sys.path.insert(0, str(SCRIPT_DIR))
from climate_chart import geocode, get_climate_data, fill_excel, get_template_path


# ── UI ────────────────────────────────────────────────────
st.set_page_config(page_title="氣候風險圖表產生器", layout="centered")
st.title("🌡️ 氣候風險圖表產生器")
st.caption("輸入工廠所在城市，自動抓取近 5 年月均溫與月均濕度，填入 Excel 並產生圖表。")

def get_secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default

template_b64 = get_secret("TEMPLATE_XLSX_BASE64")
if template_b64:
    os.environ["CLIMATE_TEMPLATE_XLSX_BASE64"] = str(template_b64).strip()

template_key = get_secret("TEMPLATE_FERNET_KEY")
if template_key:
    os.environ["CLIMATE_TEMPLATE_FERNET_KEY"] = str(template_key).strip()

app_password = get_secret("APP_PASSWORD")
if app_password and not st.session_state.get("authenticated"):
    with st.form("login_form"):
        password_input = st.text_input("使用密碼", type="password")
        login_submitted = st.form_submit_button("進入工具", use_container_width=True)

    if login_submitted and hmac.compare_digest(password_input, str(app_password)):
        st.session_state["authenticated"] = True
        st.rerun()

    if login_submitted:
        st.error("密碼不正確。")
    st.stop()

with st.form("climate_form"):
    city_input = st.text_input(
        "工廠城市（英文）",
        placeholder="例如：Sukabumi, Indonesia / Ho Chi Minh City, Vietnam",
    )
    submitted = st.form_submit_button("產生氣候風險圖表", type="primary", use_container_width=True)

if submitted:
    city = city_input.strip()
    if not city:
        st.warning("請輸入城市名稱。")
        st.stop()

    try:
        template_path = get_template_path()
    except FileNotFoundError:
        st.error("找不到 Excel 模板，請確認 Streamlit Secrets 已設定或 NAS 已掛載。")
        st.stop()

    if not template_path.exists():
        st.error(f"找不到 Excel 模板：\n`{template_path}`")
        st.stop()

    with st.spinner(f"查詢 {city} 座標中…"):
        try:
            lat, lon, full_name = geocode(city)
        except ValueError:
            st.error(f"找不到「{city}」的氣候資料。請更換鄰近城市或較大城市後再試。")
            st.stop()

    st.info(f"**{full_name}**（{lat:.4f}, {lon:.4f}）")

    with st.spinner("從 Open-Meteo 抓取 5 年氣候資料…"):
        climate, start_year, end_year = get_climate_data(lat, lon)

    # 顯示資料預覽
    with st.expander("資料預覽", expanded=False):
        import pandas as pd
        months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        temp_rows, humid_rows = [], []
        for y in range(start_year, end_year + 1):
            temp_rows.append({"年份": y, **{m: climate.get(y, {}).get(i+1, {}).get("temp") for i, m in enumerate(months)}})
            humid_rows.append({"年份": y, **{m: climate.get(y, {}).get(i+1, {}).get("humidity") for i, m in enumerate(months)}})
        st.markdown("**月均溫 (°C)**")
        st.dataframe(pd.DataFrame(temp_rows).set_index("年份"), use_container_width=True)
        st.markdown("**月均濕度 (%)**")
        st.dataframe(pd.DataFrame(humid_rows).set_index("年份"), use_container_width=True)

    with st.spinner("填入 Excel 模板（Excel 在背景開啟中）…"):
        safe_name = city.replace(",", "").replace(" ", "_")
        tmp_path = Path(tempfile.mkdtemp()) / f"氣候風險圖表_{safe_name}.xlsx"
        fill_excel(climate, start_year, end_year, tmp_path)

    st.success("圖表產生完成！")
    st.download_button(
        label=f"📥 下載 {safe_name} 氣候風險圖表.xlsx",
        data=tmp_path.read_bytes(),
        file_name=f"氣候風險圖表_{safe_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
