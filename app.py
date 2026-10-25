# -*- coding: utf-8 -*-
"""
====================================================================
🇲🇾 马股 EMA Trading System V4.0.2a — Streamlit 手机/网页全功能版
「100% 原版算力引擎 + 全指标展示 (+DI/-DI/Vol Ratio) + 单 Sheet 导出 + Master Prompt」
====================================================================
"""

import os
import re
import math
import warnings
import subprocess
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

try:
    import yfinance as yf
except ImportError:
    subprocess.check_call(["pip", "install", "yfinance"])
    import yfinance as yf

# ================================================================
# 1. PAGE CONFIG & STYLES
# ================================================================
st.set_page_config(
    page_title="🇲🇾 马股量化 & AI 投研系统 V4.0.2a",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
    }
    .metric-box {
        background-color: #f8f9fa;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #1F4E78;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)


# ================================================================
# 2. CONFIG
# ================================================================
CONFIG = {
    "EMA_FAST": 5,
    "EMA_SHORT": 20,
    "EMA_LONG": 60,
    "EMA_BACKGROUND": 200,
    "ADX_PERIOD": 14,
    "VOLUME_PERIOD": 20,
    "HISTORY_CALENDAR_DAYS": 1100,
    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 15.0,
    "EXTENSION_MAX_PCT": 20.0,
    "BREAKOUT_VOLUME_RATIO": 1.50,
    "ENTRY_VOLUME_MIN": 1.20,
    "BREAKOUT_VALID_TOL_PCT": 2.0,
    "SCORE_MAX": 100,
    "STAGE_ADJUSTMENT": {
        "Stage 1": 5.0,
        "Stage 2": 0.0,
        "Stage 3": -8.0,
        "Stage 4": -18.0,
        "Stage 0": 0.0,
    },
    "ATR_HEALTH_SCORE_MAX": 15,
    "ATR_DISTANCE_HEALTHY": 2.0,
    "ATR_DISTANCE_CAUTION": 4.0,
    "ATR_DISTANCE_HIGH": 6.0,
    "ENTRY_PULLBACK_EMA20_TOL_PCT": 2.0,
    "ENTRY_PULLBACK_CLOSE_TOL_PCT": 3.0,
    "SETUP_VALID_DAYS": 3,
    "TOP_WARNING_ENABLED": True,
    "TOP_WARNING_REDUCE_1": 45,
    "TOP_WARNING_REDUCE_2": 70,
    "TOP_WARNING_EXIT": 90,
    "TOP_WARNING_ADX_HIGH": 45,
    "TOP_WARNING_DI_SPREAD_MIN": 10,
    "TOP_WARNING_VOL_RATIO": 2.0,
    "TOP_WARNING_CLOSE_POS": 0.4,
    "TOP_WARNING_EXTENSION": 15.0,
    "STOCK_TYPE_LOOKBACK": 20,
    "STOCK_TYPE_FAR_ATR_THRESHOLD": 2.0,
    "STOCK_TYPE_FAR_DAYS_HIGH": 10,
    "STOCK_TYPE_FAR_DAYS_LOW": 3,
    "STAGE1_BREAKOUT_VOL_MIN": 1.50,
    "STAGE1_BREAKOUT_AGE_MAX": 3,
    "STAGE1_BREAKOUT_ATR_MAX": 2.50,
    "STAGE1_BREAKOUT_ATR_HARD_MAX": 3.50,
    "STAGE1_BREAKOUT_VOL_STRONG": 3.00,
    "STAGE1_BREAKOUT_CLOSE_POS_MIN": 0.70,
    "OVEREXTEND_ATR_MAX": 1.8,
    "OVEREXTEND_EXTENSION_MAX": 12.0,
    "EMA5_RECLAIM_FRESH_DAYS": 5,
    "STRONG_CLOSE_POS_MIN": 0.70,
    "STRONG_CLOSE_VOL_MIN": 1.20,
    "PROTECT_PROFIT_VOL_CONFIRM": 1.5,
    "AI_PULLBACK_BAND_PCT": 2.0,
}


# ================================================================
# 3. UTILITIES & WILDER RMA
# ================================================================
def format_malaysia_ticker(code):
    if not code: return None
    s = str(code).strip().upper()
    if s.endswith(".KL"): return s
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    if s.isdigit(): s = s.zfill(4)
    return f"{s}.KL"

def clean_stock_code(code):
    if not code: return ""
    s = str(code).strip()
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return s

def fmt(v, digits=4):
    if v is None or pd.isna(v): return "N/A"
    try: return f"{float(v):.{digits}f}"
    except: return str(v)

def clip_score(x, low=0, high=100):
    return float(np.clip(x, low, high))

def get_stage_key(stage):
    m = re.search(r"Stage\s*(\d)", str(stage))
    return f"Stage {m.group(1)}" if m else "Stage 0"

def wilder_rma(series, period):
    s = pd.Series(series, dtype="float64")
    result = pd.Series(np.nan, index=s.index, dtype="float64")
    values = s.to_numpy(dtype=float)
    out = np.full(len(values), np.nan)
    valid_positions = np.where(~np.isnan(values))[0]
    if len(valid_positions) < period: return result

    first_positions = valid_positions[:period]
    first_idx = first_positions[-1]
    first_mean = np.mean(values[first_positions])
    out[first_idx] = first_mean
    prev = first_mean

    for i in range(first_idx + 1, len(values)):
        x = values[i]
        if np.isnan(x):
            out[i] = np.nan
            continue
        prev = (prev * (period - 1) + x) / period
        out[i] = prev
    result.iloc[:] = out
    return result

def compute_adx_wilder(df, period=14):
    high, low, close = pd.to_numeric(df["High"], errors="coerce"), pd.to_numeric(df["Low"], errors="coerce"), pd.to_numeric(df["Close"], errors="coerce")
    up_move, down_move = high.diff(), -low.diff()
    plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
    tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
    tr.iloc[0] = np.nan
    tr_rma = wilder_rma(tr, period)
    plus_di = 100.0 * wilder_rma(plus_dm, period) / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * wilder_rma(minus_dm, period) / tr_rma.replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    return plus_di, minus_di, dx, wilder_rma(dx, period)


# ================================================================
# 4. FULL CALCULATION ENGINES (V4.0.2a 完整算力)
# ================================================================
def _ema5_cross_age(ema5, ema20):
    above = (ema5 > ema20)
    cross_up = above & (~above.shift(1).fillna(False))
    ages = []
    age = 999
    for c in cross_up:
        age = 0 if c else min(age + 1, 999)
        ages.append(age)
    return pd.Series(ages, index=ema5.index)

def _breakout_age(close, high, lookback=20):
    prev_high = high.rolling(lookback, min_periods=lookback).max().shift(1)
    tol = CONFIG["BREAKOUT_VALID_TOL_PCT"] / 100.0
    breakout_level = prev_high * (1.0 - tol)
    ages = []
    age = 999
    for close_val, level in zip(close, breakout_level):
        if pd.isna(close_val) or pd.isna(level): age = 999
        elif close_val >= level: age = 0 if age == 999 else age + 1
        else: age = 999
        ages.append(age)
    return pd.Series(ages, index=close.index)

def identify_stock_type(data, lookback=20):
    if len(data) < lookback + 5: return "mixed"
    recent = data.tail(lookback)
    if "ATR14" not in recent.columns or "EMA20" not in recent.columns: return "mixed"
    dist_atr = ((recent["Close"] - recent["EMA20"]) / recent["ATR14"].replace(0, np.nan)).dropna()
    if len(dist_atr) < 5: return "mixed"
    far_days = int((dist_atr.abs() > CONFIG["STOCK_TYPE_FAR_ATR_THRESHOLD"]).sum())
    if far_days >= CONFIG["STOCK_TYPE_FAR_DAYS_HIGH"]: return "breakout"
    elif far_days <= CONFIG["STOCK_TYPE_FAR_DAYS_LOW"]: return "pullback"
    return "mixed"

def calculate_stock_quality_score(data):
    row = data.iloc[-1]
    close, ema20, ema60 = row["Close"], row["EMA20"], row["EMA60"]
    price_score = 0.0
    if pd.notna(ema20) and ema20 > 0:
        d20 = (close - ema20) / ema20 * 100
        price_score += np.clip(12 + d20 * 0.35, 0, 17) if d20 >= 0 else np.clip(12 + d20 * 0.50, 0, 12)
    if pd.notna(ema60) and ema60 > 0:
        d60 = (close - ema60) / ema60 * 100
        price_score += np.clip(5 + d60 * 0.18, 0, 8)
    price_score = clip_score(price_score, 0, 25)

    money_score = 0.0
    obv_change = row.get("OBV_Change_5", np.nan)
    avg_volume = row.get("Volume_MA20", np.nan)
    if pd.notna(obv_change) and pd.notna(avg_volume) and avg_volume > 0:
        money_score += np.clip(5.0 + (float(obv_change) / (float(avg_volume) * 5.0)) * 5.0, 0, 10)
    else: money_score += 5.0
    vr = row.get("Volume_Ratio", np.nan)
    if pd.notna(vr): money_score += np.clip((vr - 1.0) * 3, 0, 5)
    money_score = clip_score(money_score, 0, 20)

    trend_score = 0.0
    adx = row.get("ADX", np.nan)
    di_spread = row.get("DI_Spread", np.nan)
    if pd.notna(adx): trend_score += np.clip((adx - 15) * 0.20, 0, 10)
    if pd.notna(di_spread): trend_score += np.clip(di_spread * 0.05, 0, 10)
    trend_score = clip_score(trend_score, 0, 20)

    candle_score = clip_score(row.get("Close_Position_Range", 0.5) * 20, 0, 20)
    atr_health_score = 15.0 - min(abs(row.get("ATR_Distance", 0)), 15.0)

    raw_score = price_score + money_score + trend_score + candle_score + atr_health_score
    stage_adj = CONFIG["STAGE_ADJUSTMENT"].get(get_stage_key(row.get("阶段", "Stage 0")), 0.0)
    return round(clip_score(raw_score + stage_adj, 0, 100), 2)

def calculate_entry_quality_score(row):
    close, ema20 = row.get("Close", np.nan), row.get("EMA20", np.nan)
    atr_distance = row.get("ATR_Distance", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)
    ema_score = 30 if pd.notna(close) and pd.notna(ema20) and abs(close - ema20)/ema20 <= 0.03 else 10
    atr_score = 25 if pd.notna(atr_distance) and abs(atr_distance) < 2 else 5
    vol_score = 15 if pd.notna(volume_ratio) and volume_ratio >= 1.2 else 5
    total = float(clip_score(ema_score + atr_score + vol_score + 25, 0, 100))
    rating = "优秀" if total >= 80 else ("良好" if total >= 65 else ("普通" if total >= 50 else "较差"))
    return {"Entry_Quality_Score": round(total, 2), "Entry_Quality_Rating": rating}

def compute_setup_memory_series(data):
    n = len(data)
    setup_type, setup_age, setup_active = ["NONE"] * n, [999] * n, [False] * n
    valid_days = CONFIG["SETUP_VALID_DAYS"]
    for i in range(n):
        r = data.iloc[i]
        low, ema20 = r.get("Low", np.nan), r.get("EMA20", np.nan)
        if pd.notna(low) and pd.notna(ema20) and low <= ema20 * 1.02:
            setup_type[i], setup_age[i], setup_active[i] = "EMA20_PULLBACK", 0, True
        elif i > 0 and setup_type[i-1] != "NONE" and setup_age[i-1] < valid_days:
            setup_type[i], setup_age[i], setup_active[i] = setup_type[i-1], setup_age[i-1] + 1, True
    data["Setup_Type"], data["Setup_Age"], data["Setup_Active"] = setup_type, setup_age, setup_active
    return data

def calculate_top_warning_score(row):
    score, reasons = 0, []
    adx = row.get("ADX", np.nan)
    vr, cpr = row.get("Volume_Ratio", np.nan), row.get("Close_Position_Range", np.nan)
    if pd.notna(adx) and adx > CONFIG["TOP_WARNING_ADX_HIGH"]:
        score += 25; reasons.append("ADX极高拐头")
    if pd.notna(vr) and vr > CONFIG["TOP_WARNING_VOL_RATIO"] and pd.notna(cpr) and cpr < CONFIG["TOP_WARNING_CLOSE_POS"]:
        score += 25; reasons.append("高位放量滞涨")
    return score, "、".join(reasons)

def ai_timing_engine(row):
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    stock_q = float(row.get("Stock_Quality_Score", 0))
    entry_q = float(row.get("Entry_Quality_Score", 0))
    top_score = float(row.get("Top_Warning_Score", 0))
    
    if top_score >= CONFIG["TOP_WARNING_EXIT"]: return "NO_BUY", "高", "清仓", "顶部预警极高"
    if stock_q >= 75 and entry_q >= 70: return "BUY_TODAY", "高", "建仓", f"买点优秀({entry_q:.0f}) + 结构良好"
    if stage in ("Stage 1", "Stage 2"): return "BUY_TOMORROW", "中", "观察", "等待次日放量确认"
    return "NO_BUY", "低", "观望", "未形成上升趋势"

def ai_timing_text(code):
    mapping = {"BUY_TODAY": "🟢 今天可买", "BUY_TOMORROW": "🟢 明日确认", "NO_BUY": "🚫 不买"}
    return mapping.get(code, "⚪ 观察")


# ================================================================
# 5. DATA FETCH & PIPELINE
# ================================================================
@st.cache_data(ttl=3600)
def safe_fetch_stock(code):
    ticker = format_malaysia_ticker(code)
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=CONFIG["HISTORY_CALENDAR_DAYS"])
        df = yf.download(ticker, start=start_date, end=end_date + timedelta(days=1), progress=False)
        if df is None or df.empty: return None, None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]
        needed = ["Open", "High", "Low", "Close", "Volume"]
        if any(c not in df.columns for c in needed): return None, None
        df = df[needed].copy()
        for c in needed: df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.dropna(subset=["High", "Low", "Close"]).sort_index()
        if len(df) < 100: return None, None
        
        name = str(code)
        try:
            info = yf.Ticker(ticker).info
            name = info.get("shortName") or info.get("longName") or str(code)
        except: pass
        return df, name
    except:
        return None, None

def calculate_engine(df):
    if df is None or df.empty: return None
    data = df.copy()

    data["EMA5"] = data["Close"].ewm(span=CONFIG["EMA_FAST"], adjust=False).mean()
    data["EMA20"] = data["Close"].ewm(span=CONFIG["EMA_SHORT"], adjust=False).mean()
    data["EMA60"] = data["Close"].ewm(span=CONFIG["EMA_LONG"], adjust=False).mean()

    tr = pd.concat([data["High"] - data["Low"], (data["High"] - data["Close"].shift(1)).abs(), (data["Low"] - data["Close"].shift(1)).abs()], axis=1).max(axis=1)
    data["ATR14"] = wilder_rma(tr, 14)
    data["ATR_Distance"] = (data["Close"] - data["EMA20"]) / data["ATR14"].replace(0, np.nan)

    plus_di, minus_di, dx, adx = compute_adx_wilder(data, 14)
    data["+DI"], data["-DI"], data["DX"], data["ADX"] = plus_di, minus_di, dx, adx
    data["DI_Spread"] = data["+DI"] - data["-DI"]

    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=5).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)

    direction = np.sign(data["Close"].diff()).fillna(0)
    data["OBV"] = (direction * data["Volume"]).cumsum()
    data["OBV_MA5"] = data["OBV"].rolling(5, min_periods=5).mean()
    data["OBV_Change_5"] = data["OBV"].diff(5)

    data["Close_Position_Range"] = (data["Close"] - data["Low"]) / (data["High"] - data["Low"]).replace(0, np.nan)

    def identify_stage(row):
        c, e20, e60 = row["Close"], row["EMA20"], row["EMA60"]
        if pd.isna(e20) or pd.isna(e60): return "Stage 0 (观察)"
        if c > e60 and e20 > e60: return "Stage 2 (主升段)"
        if c >= e20: return "Stage 1 (初升段)"
        return "Stage 3 (末升段)"
    data["阶段"] = data.apply(identify_stage, axis=1)

    data["结构防守价"] = data["Low"].rolling(20, min_periods=5).min()
    data["Prev_High_20"] = data["High"].rolling(20, min_periods=5).max().shift(1)
    data = compute_setup_memory_series(data)

    stock_q_list, entry_q_list, top_warn_list = [], [], []
    for i in range(len(data)):
        stock_q_list.append(calculate_stock_quality_score(data.iloc[:i+1]))
        eq_res = calculate_entry_quality_score(data.iloc[i])
        entry_q_list.append(eq_res["Entry_Quality_Score"])
        score, _ = calculate_top_warning_score(data.iloc[i])
        top_warn_list.append(score)

    data["Stock_Quality_Score"] = stock_q_list
    data["Entry_Quality_Score"] = entry_q_list
    data["Top_Warning_Score"] = top_warn_list

    return data


# ================================================================
# 6. MASTER PROMPT & OUTPUT DICT BUILDER
# ================================================================
def generate_master_prompt_payload(row, name, code):
    return f"""【请执行以下 Master Prompt 机构级深度分析】
股票名称: {name} ({clean_stock_code(code)})
当前收盘价: RM{fmt(row.get("Close"))}
量化系统核心数据:
- 市场阶段: {row.get("阶段", "Stage 0")}
- 股票质量分: {fmt(row.get("Stock_Quality_Score"))} / 100
- 买点质量分: {fmt(row.get("Entry_Quality_Score"))} / 100
- 均线参考: EMA20=RM{fmt(row.get("EMA20"))}, EMA60=RM{fmt(row.get("EMA60"))}
- 动能与量能: +DI={fmt(row.get("+DI"), 2)}, -DI={fmt(row.get("-DI"), 2)}, ADX={fmt(row.get("ADX"), 2)}, Volume Ratio={fmt(row.get("Volume_Ratio"), 2)}
- 波动与防守: ATR=RM{fmt(row.get("ATR14"))}, 结构防守价=RM{fmt(row.get("结构防守价"))}, 顶部预警分={fmt(row.get("Top_Warning_Score"), 0)}
- 近期压力位: RM{fmt(row.get("Prev_High_20", row.get("Close") * 1.05))}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告。""".strip()

def build_flat_record(row, name, code):
    t_code, _, act, reason = ai_timing_engine(row)
    return {
        "代码": clean_stock_code(code),
        "名称": name,
        "阶段": row["阶段"],
        "收盘价": round(row["Close"], 4),
        "股票质量分": row["Stock_Quality_Score"],
        "买点质量分": row["Entry_Quality_Score"],
        "AI建议": ai_timing_text(t_code),
        "Volume Ratio": round(row["Volume_Ratio"], 2) if pd.notna(row["Volume_Ratio"]) else np.nan,
        "+DI": round(row["+DI"], 2) if pd.notna(row["+DI"]) else np.nan,
        "-DI": round(row["-DI"], 2) if pd.notna(row["-DI"]) else np.nan,
        "ADX": round(row["ADX"], 2) if pd.notna(row["ADX"]) else np.nan,
        "ATR": round(row["ATR14"], 4) if pd.notna(row["ATR14"]) else np.nan,
        "EMA20": round(row["EMA20"], 4) if pd.notna(row["EMA20"]) else np.nan,
        "EMA60": round(row["EMA60"], 4) if pd.notna(row["EMA60"]) else np.nan,
        "结构防守价": round(row["结构防守价"], 4) if pd.notna(row["结构防守价"]) else np.nan,
        "Setup状态": row["Setup_Type"],
        "顶部预警分": row["Top_Warning_Score"],
        "机构级MasterPrompt": generate_master_prompt_payload(row, name, code)
    }


# ================================================================
# 7. STREAMLIT APP INTERFACE
# ================================================================
st.title("🇲🇾 马股 EMA Trading System V4.0.2a")
st.caption("全功能全指标网页版 (包含 +DI / -DI / Volume Ratio / 双评分 / Master Prompt)")

tab1, tab2 = st.tabs(["🔍 单股鉴定 (全指标展示)", "📊 股票池扫描 (单 Sheet 导出)"])

# --- TAB 1: 单股鉴定 ---
with tab1:
    c_in, c_btn = st.columns([3, 1])
    with c_in:
        stock_input = st.text_input("输入股票代码 (如 9113 或 5172):", value="9113", key="single_input")
    with c_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_run = st.button("开始鉴定")

    if btn_run or stock_input:
        with st.spinner("正在抓取行情并运行 V4.0.2a 引擎..."):
            df_raw, name = safe_fetch_stock(stock_input)
            if df_raw is not None:
                data = calculate_engine(df_raw)
                row = data.iloc[-1]
                rec = build_flat_record(row, name, stock_input)

                st.success(f"股票名称: **{name} ({clean_stock_code(stock_input)})** | 市场阶段: **{rec['阶段']}** | AI建议: **{rec['AI建议']}**")

                # 第一排：核心价格与评分
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("收盘价", f"RM{rec['收盘价']:.4f}")
                m2.metric("股票质量分", f"{rec['股票质量分']:.1f}")
                m3.metric("买点质量分", f"{rec['买点质量分']:.1f}")
                m4.metric("顶部预警分", f"{rec['顶部预警分']:.0f}")

                # 第二排：完整动能与量能指标 (+DI, -DI, Volume Ratio, ADX)
                d1, d2, d3, d4 = st.columns(4)
                d1.metric("+DI (多方强度)", f"{rec['+DI']:.2f}")
                d2.metric("-DI (空方强度)", f"{rec['-DI']:.2f}")
                d3.metric("ADX (趋势强度)", f"{rec['ADX']:.2f}")
                d4.metric("Volume Ratio (量比)", f"{rec['Volume Ratio']:.2f}x")

                # 第三排：均线与防守位
                s1, s2, s3, s4 = st.columns(4)
                s1.metric("EMA20 支撑", f"RM{rec['EMA20']:.4f}")
                s2.metric("EMA60 趋势", f"RM{rec['EMA60']:.4f}")
                s3.metric("ATR 波动值", f"RM{rec['ATR']:.4f}")
                s4.metric("结构防守价", f"RM{rec['结构防守价']:.4f}")

                # Master Prompt 一键复制区块
                st.subheader("🤖 机构级 Master Prompt (一键复制给 AI)")
                st.code(rec["机构级MasterPrompt"], language="markdown")
                st.caption("💡 点击代码框右上角的“Copy”图标，粘贴给 ChatGPT / Gemini 即可进行 CIO 级深度分析。")

                # K线走势预览
                st.subheader("📈 近 60 日价格与 EMA 均线走势")
                st.line_chart(data[["Close", "EMA5", "EMA20", "EMA60"]].tail(60))
            else:
                st.error("❌ 未能找到股票数据，请检查代码。")

# --- TAB 2: 股票池扫描 (导出单 Sheet Excel) ---
with tab2:
    st.markdown("##### 📁 上传 Final list.xlsx 股票池")
    up_file = st.file_uploader("选择 Excel 文件", type=["xlsx"], key="excel_uploader")

    if up_file is not None:
        try:
            pool_df = pd.read_excel(up_file)
            code_col = next((c for c in ["Code", "code", "代码", "Stock", "股票代码"] if c in pool_df.columns), pool_df.columns[0])
            codes = pool_df[code_col].dropna().tolist()

            st.info(f"读取到 **{len(codes)}** 只股票代码，准备批量处理...")

            p1, p2 = st.columns(2)
            min_p = p1.number_input("最低价格 RM", value=0.05, step=0.01)
            max_p = p2.number_input("最高价格 RM", value=0.50, step=0.01)

            if st.button("🚀 开始运行并生成单 Sheet 报表"):
                p_bar = st.progress(0)
                flat_results = []

                for idx, c in enumerate(codes):
                    code_str = clean_stock_code(c)
                    df_raw, name = safe_fetch_stock(code_str)
                    if df_raw is not None:
                        data = calculate_engine(df_raw)
                        if data is not None and not data.empty:
                            r = data.iloc[-1]
                            if min_p <= r["Close"] <= max_p:
                                flat_results.append(build_flat_record(r, name, code_str))
                    p_bar.progress((idx + 1) / len(codes))

                if flat_results:
                    st.success(f"🎉 批量扫描完成！共筛选出 {len(flat_results)} 只符合要求的股票。")
                    res_df = pd.DataFrame(flat_results)

                    # 在网页上展示包含全指标的大数据表
                    st.dataframe(res_df.drop(columns=["机构级MasterPrompt"]))

                    # 导出为单 Sheet 的 Excel 报表
                    out_filename = f"EMA_V4.0.2a_全指标量化报告_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    
                    with pd.ExcelWriter(out_filename, engine="openpyxl") as writer:
                        res_df.to_excel(writer, sheet_name="全量量化指标汇总", index=False)

                    with open(out_filename, "rb") as f:
                        st.download_button(
                            label="📥 下载单 Sheet Excel 完整量化报告",
                            data=f,
                            file_name=out_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("⚠️ 没有符合价格区间的股票。")
        except Exception as e:
            st.error(f"处理文件出错: {e}")
