"""
氣候風險圖表 Web 工具
輸入工廠城市名稱，自動產生填好的氣候風險圖表 Excel。
"""

import io
import hashlib
import hmac
import os
import shutil
import tempfile
from pathlib import Path

import streamlit as st

# ── 路徑設定 ──────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
DEMO_PASSWORD_SHA256 = "30c02e9a58710296111c685ed1ceadf48e28b0240190c8338e2f39a35dcf4de4"
from climate_chart import (
    geocode, get_climate_data, fill_excel, get_template_path,
    monthly_rainfall, classify_seasons,
)


# ── UI ────────────────────────────────────────────────────
st.set_page_config(page_title="氣候風險圖表產生器", layout="centered")
st.title("🌡️ 氣候風險圖表產生器")
st.caption("輸入工廠所在城市，自動抓取近 4 年月均溫、月均濕度與雨量，判定乾濕季並填入 Excel 產生圖表。")

def get_secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default

def password_matches(password_input: str, app_password: str) -> bool:
    # 用 bytes 比對：hmac.compare_digest 的字串版不接受非 ASCII 字元（會 TypeError）
    if app_password and hmac.compare_digest(
        password_input.encode("utf-8"), app_password.encode("utf-8")
    ):
        return True

    password_hash = hashlib.sha256(password_input.encode("utf-8")).hexdigest()
    return hmac.compare_digest(password_hash, DEMO_PASSWORD_SHA256)

template_b64 = get_secret("TEMPLATE_XLSX_BASE64")
if template_b64:
    os.environ["CLIMATE_TEMPLATE_XLSX_BASE64"] = str(template_b64).strip()

template_key = get_secret("TEMPLATE_FERNET_KEY")
if template_key:
    os.environ["CLIMATE_TEMPLATE_FERNET_KEY"] = str(template_key).strip()

app_password = str(get_secret("APP_PASSWORD")).strip()
if not st.session_state.get("authenticated"):
    with st.form("login_form"):
        password_input = st.text_input("使用密碼", type="password")
        login_submitted = st.form_submit_button("進入工具", use_container_width=True)

    if login_submitted and password_matches(password_input, app_password):
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

    with st.spinner("從 Open-Meteo 抓取近年氣候與雨量資料…"):
        climate, start_year, end_year = get_climate_data(lat, lon)
        rain = monthly_rainfall(climate, start_year, end_year)
        seasons = classify_seasons(rain)

    # 乾濕季判定（醒目顯示）
    pat = "乾-濕-乾 (Dry-Wet-Dry)" if seasons["pattern"] == "DWD" else "濕-乾-濕 (Wet-Dry-Wet)"
    wm = seasons["wet_months"]
    wet_txt = f"{wm[0]}–{wm[-1]} 月" if wm else "無明顯濕季"
    st.success(f"**乾濕季判定：{pat}**　|　濕季 {wet_txt}　|　門檻 {seasons['threshold']:.0f} mm（全年月均）")
    st.caption("乾濕季依雨量自動推算（濕月＝月雨量 ≥ 全年月均）。如與當地實況不符，下載後可直接在圖表上微調色塊與圖例。")

    # 資料預覽（溫度 / 濕度 / 雨量）— 讓同仁核對原始數據
    import pandas as pd
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    with st.expander("資料預覽（近年每月數據）", expanded=True):
        temp_rows, humid_rows, rain_rows = [], [], []
        for y in range(start_year, end_year + 1):
            temp_rows.append({"年份": y, **{m: climate.get(y, {}).get(i+1, {}).get("temp") for i, m in enumerate(months)}})
            humid_rows.append({"年份": y, **{m: climate.get(y, {}).get(i+1, {}).get("humidity") for i, m in enumerate(months)}})
            rain_rows.append({"年份": y, **{m: climate.get(y, {}).get(i+1, {}).get("precip") for i, m in enumerate(months)}})
        wet_set = set(seasons["wet_months"])
        st.markdown("**月均溫 (°C)**")
        st.dataframe(pd.DataFrame(temp_rows).set_index("年份"), use_container_width=True)
        st.markdown("**月均濕度 (%)**")
        st.dataframe(pd.DataFrame(humid_rows).set_index("年份"), use_container_width=True)
        st.markdown("**月雨量 (mm)** — 乾濕季判定依據")
        df_rain = pd.DataFrame(rain_rows).set_index("年份")
        df_rain.loc["平均 AVG"] = [rain[i + 1] for i in range(12)]
        df_rain.loc["乾濕季"] = ["濕" if (i + 1) in wet_set else "乾" for i in range(12)]
        st.dataframe(df_rain, use_container_width=True)

    with st.spinner("填入 Excel 模板…"):
        safe_name = city.replace(",", "").replace(" ", "_")
        tmp_path = Path(tempfile.mkdtemp()) / f"氣候風險圖表_{safe_name}.xlsx"
        fill_excel(climate, start_year, end_year, tmp_path,
                   full_name=full_name, seasons=seasons, rain=rain)

    st.success("圖表產生完成！")
    st.download_button(
        label=f"📥 下載 {safe_name} 氣候風險圖表.xlsx",
        data=tmp_path.read_bytes(),
        file_name=f"氣候風險圖表_{safe_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
