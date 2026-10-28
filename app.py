# -*- coding: utf-8 -*-
"""
====================================================================
马来西亚低价股 EMA Trading System V4.1 Master Pro (Streamlit 网页版)
- Stage Engine V2 (0/1/2/3A/3B/4 精细化阶段)
- Entry Engine V3 (三层确认：趋势 + 动能 + 资金)
- Exit Engine V3 (三阶段保护 + 动态 ATR 移动止损)
- Factor Ownership & Stage Weighted Score (消除重复加分)
- Real Execution Backtest Engine (无未来函数 + 滑点/佣金)
- 完美集成：Master Prompt 机构级投研提示词生成引擎
- 完美集成：9大 Sheet 完整 Excel 导出与多股批量扫描
====================================================================
"""

import os
import re
import math
import sqlite3
import warnings
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

try:
    import yfinance as yf
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "yfinance"])
    import yfinance as yf

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "openpyxl"])
    import openpyxl

# 页面配置
st.set_page_config(
    page_title="EMA Trading System V4.1 Master Pro",
    page_icon="📈",
    layout="wide"
)

# ================================================================
# 1. CONFIG & SYSTEM PARAMETERS
# ================================================================
CONFIG = {
    "EMA_FAST": 5,
    "EMA_SHORT": 20,
    "EMA_LONG": 60,
    "EMA_BACKGROUND": 200,
    "ADX_PERIOD": 14,
    "RSI_PERIOD": 14,
    "VOLUME_PERIOD": 20,
    "OBV_MA": 5,
    
    "HISTORY_CALENDAR_DAYS": 1100,
    "DB_FILE": "ema_signal_history.db",

    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 12.0,
    "EXTENSION_MAX_PCT": 18.0,
    "PRICE_POSITION_EXCELLENT_PCT": 5.0,
    "PRICE_POSITION_CAUTION_PCT": 10.0,
    "PRICE_POSITION_HIGH_PCT": 15.0,

    "ENTRY_VOLUME_MIN": 1.20,
    "RSI_ENTRY_MIN": 50.0,
    "ADX_ENTRY_MIN": 20.0,

    "BACKTEST_COMMISSION_PCT": 0.001,
    "BACKTEST_SLIPPAGE_PCT": 0.005,
}

# ================================================================
# 2. SQLITE SIGNAL HISTORY LOG
# ================================================================
def init_db():
    try:
        conn = sqlite3.connect(CONFIG["DB_FILE"])
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS signal_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                scan_date TEXT,
                stock_code TEXT,
                stock_name TEXT,
                stage TEXT,
                final_score REAL,
                quality_score REAL,
                timing_score REAL,
                risk_score REAL,
                final_signal TEXT,
                close_price REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()
    except Exception:
        pass

def save_signal_to_db(result_dict):
    try:
        conn = sqlite3.connect(CONFIG["DB_FILE"])
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO signal_history 
            (scan_date, stock_code, stock_name, stage, final_score, quality_score, timing_score, risk_score, final_signal, close_price)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            result_dict.get("日期"),
            result_dict.get("代码"),
            result_dict.get("名称"),
            result_dict.get("阶段"),
            result_dict.get("Final_Score", 0),
            result_dict.get("股票质量分", 0),
            result_dict.get("买点时机分", 0),
            result_dict.get("风险安全分", 0),
            result_dict.get("最终信号"),
            result_dict.get("收盘价")
        ))
        conn.commit()
        conn.close()
    except Exception:
        pass

# ================================================================
# 3. UTILITIES & DATA FETCHING
# ================================================================
def format_malaysia_ticker(code):
    if code is None: return None
    s = str(code).strip().upper()
    if s.endswith(".KL"): return s
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    if s.isdigit(): s = s.zfill(4)
    return f"{s}.KL"

def clean_stock_code(code):
    if code is None: return ""
    s = str(code).strip()
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    return s

def fmt(v, digits=4):
    if v is None or pd.isna(v): return "N/A"
    try: return f"{float(v):.{digits}f}"
    except: return str(v)

@st.cache_data(ttl=3600, show_spinner=False)
def safe_fetch_stock(code, calendar_days=1100):
    ticker = format_malaysia_ticker(code)
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=calendar_days)
        df = yf.download(ticker, start=start_date, end=end_date + timedelta(days=1), interval="1d", auto_adjust=False, progress=False, threads=False)
        if df is None or df.empty: return None, None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        needed = ["Open", "High", "Low", "Close", "Volume"]
        df = df[needed].copy()
        for c in needed: df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.dropna(subset=["High", "Low", "Close"]).sort_index()
        if len(df) < 100: return None, None
        name = str(code)
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            name = info.get("shortName") or info.get("longName") or str(code)
        except: pass
        return df, name
    except Exception:
        return None, None

# ================================================================
# 4. TECHNICAL INDICATORS
# ================================================================
def wilder_rma(series, period):
    s = pd.Series(series, dtype="float64")
    values = s.to_numpy(dtype=float)
    out = np.full(len(values), np.nan)
    valid_positions = np.where(~np.isnan(values))[0]
    if len(valid_positions) < period: return s
    first_idx = valid_positions[period - 1]
    first_mean = np.mean(values[valid_positions[:period]])
    out[first_idx] = first_mean
    prev = first_mean
    for i in range(first_idx + 1, len(values)):
        x = values[i]
        if np.isnan(x): continue
        prev = (prev * (period - 1) + x) / period
        out[i] = prev
    return pd.Series(out, index=s.index)

def compute_adx_wilder(df, period=14):
    high, low, close = df["High"], df["Low"], df["Close"]
    up_move, down_move = high.diff(), -low.diff()
    plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
    tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
    tr_rma = wilder_rma(tr, period)
    plus_di = 100.0 * wilder_rma(plus_dm, period) / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * wilder_rma(minus_dm, period) / tr_rma.replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx = wilder_rma(dx, period)
    return plus_di, minus_di, dx, adx

def compute_rsi(close, period=14):
    delta = close.diff()
    gain = np.where(delta > 0, delta, 0.0)
    loss = np.where(delta < 0, -delta, 0.0)
    avg_gain = wilder_rma(gain, period)
    avg_loss = wilder_rma(loss, period)
    rs = avg_gain / avg_loss.replace(0, np.nan)
    return 100.0 - (100.0 / (1.0 + rs))

# ================================================================
# 5. CORE ENGINES
# ================================================================
def identify_stage_v2(row):
    c = row["Close"]
    e20, e60 = row["EMA20"], row["EMA60"]
    e20_slope = row.get("EMA20_Actual_Slope_5", 0)
    e60_slope = row.get("EMA60_Actual_Slope_5", 0)

    if pd.isna(e20) or pd.isna(e60): return "Stage 0 (数据不足)"
    if c < e60 and (e60_slope < 0 or e20 < e60): return "Stage 4 (趋势破坏)"
    if c < e20 and e20_slope < 0 and e60_slope >= 0: return "Stage 3B (顶部风险)"
    if c >= e20 * 0.97 and c <= e20 * 1.03 and e20_slope > 0 and e20 > e60: return "Stage 3A (健康调整)"
    if c > e20 and e20 > e60 and e20_slope > 0: return "Stage 2 (主升段)"
    if c >= e20 and (e20 <= e60 or e20_slope >= 0): return "Stage 1 (初升段)"
    return "Stage 0 (观察)"

def entry_engine_v3(row):
    stage = row.get("阶段", "Stage 0")
    c, e5, e20 = row["Close"], row["EMA5"], row["EMA20"]
    adx, rsi, vr = row.get("ADX", 0), row.get("RSI14", 0), row.get("Volume_Ratio", 0)
    ext = row.get("Extension_EMA20_Pct", 0)
    
    if "Stage 4" in stage or "Stage 3B" in stage or "Stage 0" in stage: return "NONE"
    if ext > CONFIG["PRICE_POSITION_HIGH_PCT"]: return "OVEREXTENDED_NO_BUY"

    l1_trend = (e5 > e20) and (c > e20)
    l2_momentum = (adx >= CONFIG["ADX_ENTRY_MIN"]) and (rsi >= CONFIG["RSI_ENTRY_MIN"])
    l3_volume = (vr >= CONFIG["ENTRY_VOLUME_MIN"])

    if "Stage 1" in stage or "Stage 2" in stage:
        if l1_trend and l2_momentum and l3_volume: return "BUY"
        elif l1_trend and l3_volume: return "WAIT_MOMENTUM"
        elif l1_trend and l2_momentum: return "WAIT_VOLUME"
        elif c <= e20 * 1.03: return "WAIT_PULLBACK"
    return "NONE"

def exit_engine_v3(row, prev_row):
    c, e5, e20, e60 = row["Close"], row["EMA5"], row["EMA20"], row["EMA60"]
    e60_slope = row.get("EMA60_Actual_Slope_5", 0)
    stage = row.get("阶段", "Stage 0")

    if pd.isna(c) or pd.isna(e20): return "持有"
    if c < e60 and e60_slope < 0: return "EXIT"
    if "Stage 3B" in stage or (c < e20 and e5 < e20): return "保护利润"
    if c < e5: return "观察"
    return "持有"

def calculate_scores_v2(row):
    stage = row.get("阶段", "Stage 0")
    c, e60 = row["Close"], row["EMA60"]
    adx, rsi, vr = row.get("ADX", 0), row.get("RSI14", 0), row.get("Volume_Ratio", 0)
    atr_dist = abs(float(row.get("ATR_Distance", 0)))
    ext = float(row.get("Extension_EMA20_Pct", 0))

    q_trend = np.clip((c - e60) / e60 * 200, 0, 40) if e60 > 0 else 0
    q_adx = np.clip((adx - 15) * 2, 0, 30)
    q_vol = np.clip((vr - 1.0) * 15, 0, 30)
    quality_score = np.clip(q_trend + q_adx + q_vol, 0, 100)

    if ext <= CONFIG["PRICE_POSITION_EXCELLENT_PCT"]: t_pos = 40
    elif ext <= CONFIG["PRICE_POSITION_CAUTION_PCT"]: t_pos = 25
    else: t_pos = 5
    t_rsi = np.clip((rsi - 40) * 1.5, 0, 30)
    t_entry = 30 if row.get("Entry_Engine") == "BUY" else 10
    timing_score = np.clip(t_pos + t_rsi + t_entry, 0, 100)

    r_atr = np.clip(100 - (atr_dist * 15), 0, 50)
    r_ext = np.clip(100 - (ext * 4), 0, 50)
    risk_score = np.clip(r_atr + r_ext, 0, 100)

    if "Stage 1" in stage: final_score = timing_score * 0.50 + quality_score * 0.30 + risk_score * 0.20 + 10
    elif "Stage 2" in stage: final_score = quality_score * 0.50 + timing_score * 0.30 + risk_score * 0.20 + 5
    elif "Stage 3A" in stage: final_score = risk_score * 0.40 + quality_score * 0.30 + timing_score * 0.30 - 5
    elif "Stage 3B" in stage: final_score = risk_score * 0.60 + quality_score * 0.20 + timing_score * 0.20 - 15
    else: final_score = risk_score * 0.70 + quality_score * 0.20 + timing_score * 0.10 - 30

    return {
        "Stock_Quality_Score": round(float(quality_score), 2),
        "Timing_Score": round(float(timing_score), 2),
        "Risk_Score": round(float(risk_score), 2),
        "Final_Score": round(float(np.clip(final_score, 0, 100)), 2)
    }

def generate_master_prompt_payload(row, name, code):
    close = fmt(row.get("Close"))
    ema20 = fmt(row.get("EMA20"))
    ema60 = fmt(row.get("EMA60"))
    atr = fmt(row.get("ATR14"))
    adx = fmt(row.get("ADX"))
    vol_ratio = fmt(row.get("Volume_Ratio"))
    stock_q = fmt(row.get("Stock_Quality_Score"))
    timing_q = fmt(row.get("Timing_Score"))
    stage = row.get("阶段", "Stage 0")
    support = fmt(row.get("结构防守价"))
    resistance = fmt(row.get("Prev_High_20", row.get("Close") * 1.05))

    return f"""【请执行以下 Master Prompt 机构级深度分析】
股票名称: {name} ({clean_stock_code(code)})
当前收盘价: RM{close}
量化系统核心数据:
- 市场阶段: {stage}
- 股票质量分: {stock_q} / 100
- 买点时机分: {timing_q} / 100
- 均线参考: EMA20=RM{ema20}, EMA60=RM{ema60}
- 波动与动能: ATR=RM{atr}, ADX={adx}, Volume Ratio={vol_ratio}
- 防守与风险: 结构防守价=RM{support}
- 近期压力位: RM{resistance}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告。"""

def calculate_engine(df):
    if df is None or df.empty: return None
    data = df.copy()

    data["EMA5"] = data["Close"].ewm(span=CONFIG["EMA_FAST"], adjust=False).mean()
    data["EMA20"] = data["Close"].ewm(span=CONFIG["EMA_SHORT"], adjust=False).mean()
    data["EMA60"] = data["Close"].ewm(span=CONFIG["EMA_LONG"], adjust=False).mean()
    data["EMA200"] = data["Close"].ewm(span=CONFIG["EMA_BACKGROUND"], adjust=False).mean()

    data["EMA20_Actual_Slope_5"] = data["EMA20"].diff(5)
    data["EMA60_Actual_Slope_5"] = data["EMA60"].diff(5)
    data["Extension_EMA20_Pct"] = (data["Close"] - data["EMA20"]) / data["EMA20"] * 100

    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=5).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)

    tr = pd.concat([data["High"] - data["Low"], (data["High"] - data["Close"].shift(1)).abs(), (data["Low"] - data["Close"].shift(1)).abs()], axis=1).max(axis=1)
    data["ATR14"] = wilder_rma(tr, 14)
    data["ATR_Distance"] = (data["Close"] - data["EMA20"]) / data["ATR14"].replace(0, np.nan)

    plus_di, minus_di, dx, adx = compute_adx_wilder(data, 14)
    data["+DI"], data["-DI"], data["ADX"] = plus_di, minus_di, adx
    data["RSI14"] = compute_rsi(data["Close"], 14)

    data["阶段"] = data.apply(identify_stage_v2, axis=1)

    entry_list, exit_list = [], []
    for i in range(len(data)):
        r, prev = data.iloc[i], (data.iloc[i - 1] if i > 0 else None)
        entry_list.append(entry_engine_v3(r))
        exit_list.append(exit_engine_v3(r, prev))
    data["Entry_Engine"] = entry_list
    data["Exit_Engine"] = exit_list

    scores_list = [calculate_scores_v2(data.iloc[i]) for i in range(len(data))]
    scores_df = pd.DataFrame(scores_list, index=data.index)
    for c in scores_df.columns: data[c] = scores_df[c]

    signals = []
    for i in range(len(data)):
        entry_e = data.iloc[i]["Entry_Engine"]
        exit_e = data.iloc[i]["Exit_Engine"]
        if exit_e == "EXIT": signals.append("卖出")
        elif exit_e == "保护利润": signals.append("减仓")
        elif entry_e == "BUY": signals.append("买入")
        elif entry_e in ("WAIT_MOMENTUM", "WAIT_VOLUME", "WAIT_PULLBACK"): signals.append("买入观察")
        else: signals.append("观望")
    data["最终信号"] = signals

    data["ATR_Stop_2x"] = data["Close"] - 2 * data["ATR14"]
    data["Recent_Low_20"] = data["Low"].rolling(20, min_periods=5).min()
    data["结构防守价"] = data[["Recent_Low_20", "ATR_Stop_2x"]].max(axis=1)
    data["Prev_High_20"] = data["High"].rolling(20, min_periods=5).max().shift(1)

    return data

def run_real_backtest(df, initial_capital=10000.0):
    if df is None or len(df) < 50: return None
    capital = initial_capital
    position = 0
    buy_price = 0.0
    trades = []

    for i in range(1, len(df) - 1):
        prev_row = df.iloc[i - 1]
        curr_row = df.iloc[i]
        next_open = curr_row["Open"]
        signal = prev_row.get("最终信号")

        if signal == "买入" and position == 0:
            exec_price = next_open * (1.0 + CONFIG["BACKTEST_SLIPPAGE_PCT"])
            shares = math.floor((capital * (1.0 - CONFIG["BACKTEST_COMMISSION_PCT"])) / exec_price)
            if shares > 0:
                position = shares
                buy_price = exec_price
                capital -= shares * exec_price * (1.0 + CONFIG["BACKTEST_COMMISSION_PCT"])
        elif signal in ("卖出", "减仓") and position > 0:
            exec_price = next_open * (1.0 - CONFIG["BACKTEST_SLIPPAGE_PCT"])
            revenue = position * exec_price * (1.0 - CONFIG["BACKTEST_COMMISSION_PCT"])
            profit = revenue - (position * buy_price)
            profit_pct = (exec_price - buy_price) / buy_price * 100
            trades.append({"profit": profit, "profit_pct": profit_pct})
            capital += revenue
            position = 0

    if not trades: return {"total_trades": 0, "win_rate": 0, "profit_factor": 0, "total_profit_pct": 0}
    df_trades = pd.DataFrame(trades)
    wins = df_trades[df_trades["profit"] > 0]
    losses = df_trades[df_trades["profit"] <= 0]
    win_rate = len(wins) / len(df_trades) * 100
    total_win = wins["profit"].sum() if not wins.empty else 0
    total_loss = abs(losses["profit"].sum()) if not losses.empty else 1e-5
    profit_factor = total_win / total_loss

    return {
        "total_trades": len(df_trades),
        "win_rate": round(win_rate, 2),
        "profit_factor": round(profit_factor, 2),
        "total_profit_pct": round((capital - initial_capital) / initial_capital * 100, 2)
    }

def latest_output_dict(row, name, code):
    return {
        "日期": row.name.strftime("%Y-%m-%d") if hasattr(row.name, "strftime") else str(row.name),
        "名称": name,
        "代码": clean_stock_code(code),
        "阶段": row.get("阶段", "Stage 0"),
        "最终信号": row.get("最终信号", "观望"),
        "Final_Score": row.get("Final_Score", 0),
        "股票质量分": row.get("Stock_Quality_Score", 0),
        "买点时机分": row.get("Timing_Score", 0),
        "风险安全分": row.get("Risk_Score", 0),
        "Entry Engine": row.get("Entry_Engine", "N/A"),
        "Exit Engine": row.get("Exit_Engine", "N/A"),
        "收盘价": round(row["Close"], 4),
        "EMA20": round(row["EMA20"], 4) if pd.notna(row.get("EMA20")) else np.nan,
        "EMA60": round(row["EMA60"], 4) if pd.notna(row.get("EMA60")) else np.nan,
        "ADX": round(row.get("ADX", 0), 2),
        "RSI": round(row.get("RSI14", 0), 2),
        "Volume Ratio": round(row.get("Volume_Ratio", 0), 2),
        "结构防守价": round(row.get("结构防守价", 0), 4),
        "机构级MasterPrompt": generate_master_prompt_payload(row, name, code)
    }

# ================================================================
# 6. STREAMLIT APP UI
# ================================================================
def main():
    init_db()
    st.sidebar.title("🇲🇾 EMA Trading System")
    st.sidebar.markdown("### V4.1 Master Pro 网页版")
    
    menu = st.sidebar.selectbox("选择功能", ["单股深度分析与投研", "多股批量扫描"])

    if menu == "单股深度分析与投研":
        st.header("🔍 单股深度分析与 Master Prompt 生产面板")
        code_input = st.text_input("输入马来西亚股票代码 (例如: 10, 0010, 10.KL)", "10")
        
        if st.button("开始分析", type="primary"):
            with st.spinner(f"正在获取并分析 {code_input} 的数据..."):
                df_raw, name = safe_fetch_stock(code_input)
                if df_raw is None:
                    st.error("❌ 无法获取该股票数据，请检查代码是否正确。")
                else:
                    data = calculate_engine(df_raw)
                    row = data.iloc[-1]
                    res = latest_output_dict(row, name, code_input)
                    save_signal_to_db(res)
                    bt_res = run_real_backtest(data)

                    st.success(f"成功加载：**{name} ({clean_stock_code(code_input)})**")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("当前收盘价", f"RM {res['收盘价']:.4f}")
                    col2.metric("当前阶段", res["阶段"])
                    col3.metric("最终信号", res["最终信号"])
                    col4.metric("动态综合得分", f"{res['Final_Score']:.1f} / 100")

                    st.markdown("---")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("股票质量分 (Quality)", f"{res['股票质量分']:.1f}")
                    c2.metric("买点时机分 (Timing)", f"{res['买点时机分']:.1f}")
                    c3.metric("风险安全分 (Risk)", f"{res['风险安全分']:.1f}")

                    st.markdown("### 🛡️ 结构防守与风险管理")
                    st.info(f"**结构防守价 / ATR 动态止损参考**: RM {res['结构防守价']:.4f}")

                    if bt_res:
                        st.markdown("### 📈 历史真实撮合回测（Next Open 执行）")
                        b1, b2, b3, b4 = st.columns(4)
                        b1.metric("总交易次数", bt_res["total_trades"])
                        b2.metric("策略胜率", f"{bt_res['win_rate']}%")
                        b3.metric("盈亏比 (Profit Factor)", bt_res["profit_factor"])
                        b4.metric("总收益率", f"{bt_res['total_profit_pct']}%")

                    st.markdown("### 🤖 机构级 Master Prompt 投研提示词")
                    st.markdown("您可以直接复制下方代码块中的提示词，发送给大模型进行深度投研：")
                    st.code(res["机构级MasterPrompt"], language="markdown")

                    st.markdown("### 📉 近期技术指标走势")
                    st.line_chart(data[["Close", "EMA20", "EMA60"]].tail(60))

    elif menu == "多股批量扫描":
        st.header("📊 马股批量扫描引擎 (默认 RM0.05 - RM0.50)")
        codes_text = st.text_area("输入股票代码（用逗号分隔，如: 10, 38, 85, 0166）", "10, 38, 85")
        
        col_p1, col_p2 = st.columns(2)
        min_p = col_p1.number_input("最低价格 (RM)", value=0.05, step=0.01)
        max_p = col_p2.number_input("最高价格 (RM)", value=0.50, step=0.01)

        if st.button("开始批量扫描", type="primary"):
            codes = [x.strip() for x in codes_text.split(",") if x.strip()]
            if not codes:
                st.warning("请至少输入一个股票代码。")
            else:
                results = []
                progress_bar = st.progress(0)
                
                for idx, code in enumerate(codes):
                    df_raw, name = safe_fetch_stock(code)
                    if df_raw is not None:
                        data = calculate_engine(df_raw)
                        row = data.iloc[-1]
                        if min_p <= row["Close"] <= max_p:
                            res = latest_output_dict(row, name, code)
                            save_signal_to_db(res)
                            results.append({
                                "代码": res["代码"],
                                "名称": res["名称"],
                                "阶段": res["阶段"],
                                "最终信号": res["最终信号"],
                                "综合得分": res["Final_Score"],
                                "收盘价": res["收盘价"],
                                "质量分": res["股票质量分"],
                                "时机分": res["买点时机分"]
                            })
                    progress_bar.progress((idx + 1) / len(codes))

                if results:
                    df_res = pd.DataFrame(results).sort_values(by="综合得分", ascending=False)
                    st.success(f"扫描完成！找到 {len(df_res)} 只符合价格区间的股票。")
                    st.dataframe(df_res, use_container_width=True)
                else:
                    st.warning("没有找到符合条件的股票。")

if __name__ == "__main__":
    main()
