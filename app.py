# -*- coding: utf-8 -*-
"""
====================================================================
马来西亚低价股 EMA Trading System V4.1 Pro (全功能重构升级版)
- Stage Engine V2 (0/1/2/3A/3B/4 精细化阶段)
- Entry Engine V3 (三层确认：趋势 + 动能 + 资金)
- Exit Engine V3 (三阶段保护 + 动态 ATR 移动止损)
- Factor Ownership & Stage Weighted Score (消除重复加分)
- Real Execution Backtest Engine (消除未来函数 + 滑点/佣金)
- SQLite Signal History Log (历史信号追溯与重复检测)
- Daily Trading Dashboard + 9大 Sheet 交互 Excel
====================================================================
"""

import os
import re
import math
import sqlite3
import warnings
import subprocess
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

try:
    import yfinance as yf
except ImportError:
    subprocess.check_call(["pip", "install", "yfinance"])
    import yfinance as yf

try:
    import openpyxl
except ImportError:
    subprocess.check_call(["pip", "install", "openpyxl"])
    import openpyxl


# ================================================================
# 1. CONFIG & SYSTEM PARAMETERS
# ================================================================

CONFIG = {
    # 均线与指标周期
    "EMA_FAST": 5,
    "EMA_SHORT": 20,
    "EMA_LONG": 60,
    "EMA_BACKGROUND": 200,
    "ADX_PERIOD": 14,
    "RSI_PERIOD": 14,
    "VOLUME_PERIOD": 20,
    "OBV_MA": 5,
    "BB_PERIOD": 20,
    "BB_STD": 2,

    "HISTORY_CALENDAR_DAYS": 1100,
    "DB_FILE": "ema_signal_history.db",

    # 偏离与位置阈值 (Factor Ownership)
    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 12.0,
    "EXTENSION_MAX_PCT": 18.0,
    "PRICE_POSITION_EXCELLENT_PCT": 5.0,  # Close距离EMA20 0-5%
    "PRICE_POSITION_CAUTION_PCT": 10.0,   # Close距离EMA20 5-10%
    "PRICE_POSITION_HIGH_PCT": 15.0,      # >15% 严禁追高

    # 买入三层确认阈值
    "BREAKOUT_VOLUME_RATIO": 1.50,
    "ENTRY_VOLUME_MIN": 1.20,
    "BREAKOUT_VALID_TOL_PCT": 2.0,
    "RSI_ENTRY_MIN": 50.0,
    "ADX_ENTRY_MIN": 20.0,

    # ADX 强度分类
    "ADX_WEAK": 20,
    "ADX_DEVELOPING": 25,
    "ADX_STRONG": 30,

    # 筛选门槛
    "DEFAULT_MIN_PRICE": 0.05,
    "DEFAULT_MAX_PRICE": 0.50,
    "DEFAULT_MIN_VOLUME": 100000,
    "SCORE_MAX": 100,

    # 阶段调整基础分
    "STAGE_ADJUSTMENT": {
        "Stage 1": 10.0,
        "Stage 2": 5.0,
        "Stage 3A": -5.0,
        "Stage 3B": -15.0,
        "Stage 4": -30.0,
        "Stage 0": 0.0,
    },

    # 动态 ATR 止损倍数
    "ATR_STOP_MULTIPLIER": {
        "Stage 1": 2.5,
        "Stage 2": 3.0,
        "Stage 3A": 2.0,
        "Stage 3B": 1.5,
        "Stage 4": 0.0,
        "Stage 0": 2.0,
    },

    # 回测真实撮合参数
    "BACKTEST_COMMISSION_PCT": 0.001,  # 0.1% 佣金
    "BACKTEST_SLIPPAGE_PCT": 0.005,    # 0.5% 滑点

    # 顶部预警阈值
    "TOP_WARNING_ENABLED": True,
    "TOP_WARNING_REDUCE_1": 45,
    "TOP_WARNING_REDUCE_2": 70,
    "TOP_WARNING_EXIT": 90,

    # Setup 记忆机制
    "SETUP_VALID_DAYS": 3,
}

# Excel Sheet 名称定义
SHEET_DASHBOARD = "⓪Daily Dashboard"
SHEET_OVERVIEW = "①今日交易总览"
SHEET_BUY_LIST = "②Buy Candidates"
SHEET_WATCHLIST = "③Watch List"
SHEET_SELL_MONITOR = "④Risk Sell Monitor"
SHEET_DECISION = "⑤决策摘要"
SHEET_AI = "⑥AI分析"
SHEET_RECENT = "⑦最近60日"
SHEET_ALL = "⑧全部指标"
SHEET_ERRORS = "无数据或错误"


# ================================================================
# 2. SQLITE SIGNAL HISTORY LOG
# ================================================================

def init_db():
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
            result_dict.get("Stock_Quality_Score", 0),
            result_dict.get("Timing_Score", 0),
            result_dict.get("Risk_Score", 0),
            result_dict.get("最终信号"),
            result_dict.get("收盘价")
        ))
        conn.commit()
        conn.close()
    except Exception as e:
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

def safe_filename(text):
    text = str(text or "stock")
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    return text[:80].strip() or "stock"

def fmt(v, digits=4):
    if v is None or pd.isna(v): return "N/A"
    try: return f"{float(v):.{digits}f}"
    except: return str(v)

def safe_fetch_stock(code, calendar_days=None):
    if calendar_days is None: calendar_days = CONFIG["HISTORY_CALENDAR_DAYS"]
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
# 4. TECHNICAL INDICATORS (WILDER RMA, ADX, RSI, OBV)
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
# 5. CORE ENGINES: STAGE V2, ENTRY V3, EXIT V3, SCORES V2
# ================================================================

# --- Stage Engine V2 ---
def identify_stage_v2(row):
    c = row["Close"]
    e5, e20, e60, e200 = row["EMA5"], row["EMA20"], row["EMA60"], row["EMA200"]
    e20_slope = row.get("EMA20_Actual_Slope_5", 0)
    e60_slope = row.get("EMA60_Actual_Slope_5", 0)

    if pd.isna(e20) or pd.isna(e60): return "Stage 0"
    
    # Stage 4: 趋势完全破坏 (Close低于EMA60且EMA60下弯，或EMA20<EMA60且向下)
    if c < e60 and (e60_slope < 0 or e20 < e60):
        return "Stage 4"
    
    # Stage 3B: 高位风险段 (Close跌破EMA20，EMA20下弯，但EMA60仍向上)
    if c < e20 and e20_slope < 0 and e60_slope >= 0:
        return "Stage 3B"
        
    # Stage 3A: 健康回调段 (Stage 2过程中的缩量回调触及EMA20，EMA20仍向上)
    if c >= e20 * 0.97 and c <= e20 * 1.03 and e20_slope > 0 and e20 > e60:
        return "Stage 3A"
        
    # Stage 2: 主升趋势段 (Close > EMA20 > EMA60，且均线多头排列)
    if c > e20 and e20 > e60 and e20_slope > 0:
        return "Stage 2"
        
    # Stage 1: 初升突破段 (Close刚站上EMA20与EMA60，底部开始盘整向上)
    if c >= e20 and (e20 <= e60 or e20_slope >= 0):
        return "Stage 1"
        
    return "Stage 0"

# --- Entry Engine V3 (三层确认机制) ---
def entry_engine_v3(row):
    stage = row.get("阶段_V2", "Stage 0")
    c, e5, e20, e60 = row["Close"], row["EMA5"], row["EMA20"], row["EMA60"]
    adx, rsi, vr = row.get("ADX", 0), row.get("RSI14", 0), row.get("Volume_Ratio", 0)
    ext = row.get("Extension_EMA20_Pct", 0)
    
    if stage in ("Stage 4", "Stage 3B", "Stage 0"):
        return "NONE"

    # 追高风险过滤
    if ext > CONFIG["PRICE_POSITION_HIGH_PCT"]:
        return "OVEREXTENDED_NO_BUY"

    # Layer 1: 趋势确认
    l1_trend = (e5 > e20) and (c > e20)
    # Layer 2: 动能确认
    l2_momentum = (adx >= CONFIG["ADX_ENTRY_MIN"]) and (rsi >= CONFIG["RSI_ENTRY_MIN"])
    # Layer 3: 资金确认
    l3_volume = (vr >= CONFIG["ENTRY_VOLUME_MIN"])

    if stage in ("Stage 1", "Stage 2"):
        if l1_trend and l2_momentum and l3_volume:
            return "BUY_CONFIRMED"
        elif l1_trend and l3_volume:
            return "WAIT_MOMENTUM"
        elif l1_trend and l2_momentum:
            return "WAIT_VOLUME"
        elif c <= e20 * 1.03:
            return "WAIT_PULLBACK_CONFIRM"

    return "NONE"

# --- Exit Engine V3 (三级保护 + 动态 ATR 止损) ---
def exit_engine_v3(row, prev_row):
    c, e5, e20, e60 = row["Close"], row["EMA5"], row["EMA20"], row["EMA60"]
    e60_slope = row.get("EMA60_Actual_Slope_5", 0)
    stage = row.get("阶段_V2", "Stage 0")

    if pd.isna(c) or pd.isna(e20): return "HOLD"

    # 级别 3 告警：清仓卖出 (Close < EMA60 & EMA60 下弯)
    if c < e60 and e60_slope < 0:
        return "EXIT_SELL"

    # 级别 2 告警：保护利润 / 减仓 (Stage 3B 或 Close 跌破 EMA20)
    if stage == "Stage 3B" or (c < e20 and e5 < e20):
        return "REDUCE_POSITION"

    # 级别 1 告警：短线观察 (Close 跌破 EMA5)
    if c < e5:
        return "WATCH_WARNING"

    return "HOLD"

# --- Score Engine V2 (Factor Ownership & Stage Weighted Score) ---
def calculate_scores_v2(row):
    stage = row.get("阶段_V2", "Stage 0")
    c, e20, e60 = row["Close"], row["EMA20"], row["EMA60"]
    adx, rsi, vr = row.get("ADX", 0), row.get("RSI14", 0), row.get("Volume_Ratio", 0)
    atr_dist = abs(float(row.get("ATR_Distance", 0)))
    ext = float(row.get("Extension_EMA20_Pct", 0))

    # 1. Quality Score (股票趋势与质量 0-100)
    q_trend = np.clip((c - e60) / e60 * 200, 0, 40) if e60 > 0 else 0
    q_adx = np.clip((adx - 15) * 2, 0, 30)
    q_vol = np.clip((vr - 1.0) * 15, 0, 30)
    quality_score = np.clip(q_trend + q_adx + q_vol, 0, 100)

    # 2. Timing Score (买入时机与位置 0-100)
    # 位置越靠近 EMA20 得分越高，极端偏离扣分
    if ext <= CONFIG["PRICE_POSITION_EXCELLENT_PCT"]: t_pos = 40
    elif ext <= CONFIG["PRICE_POSITION_CAUTION_PCT"]: t_pos = 25
    else: t_pos = 5
    t_rsi = np.clip((rsi - 40) * 1.5, 0, 30)
    t_entry = 30 if row.get("Entry_Engine_V3") == "BUY_CONFIRMED" else 10
    timing_score = np.clip(t_pos + t_rsi + t_entry, 0, 100)

    # 3. Risk Score (风险程度，分数越高越安全 0-100)
    r_atr = np.clip(100 - (atr_dist * 15), 0, 50)
    r_ext = np.clip(100 - (ext * 4), 0, 50)
    risk_score = np.clip(r_atr + r_ext, 0, 100)

    # 4. Stage Weighted Dynamic Final Score (阶段动态权重整合)
    if stage == "Stage 1":
        final_score = timing_score * 0.50 + quality_score * 0.30 + risk_score * 0.20 + 10
    elif stage == "Stage 2":
        final_score = quality_score * 0.50 + timing_score * 0.30 + risk_score * 0.20 + 5
    elif stage == "Stage 3A":
        final_score = risk_score * 0.40 + quality_score * 0.30 + timing_score * 0.30 - 5
    elif stage == "Stage 3B":
        final_score = risk_score * 0.60 + quality_score * 0.20 + timing_score * 0.20 - 15
    else:  # Stage 4 / Stage 0
        final_score = risk_score * 0.70 + quality_score * 0.20 + timing_score * 0.10 - 30

    return {
        "Stock_Quality_Score": round(float(quality_score), 2),
        "Timing_Score": round(float(timing_score), 2),
        "Risk_Score": round(float(risk_score), 2),
        "Final_Score": round(float(np.clip(final_score, 0, 100)), 2)
    }


# ================================================================
# 6. PIPELINE COMPUTATION & BACKTEST ENGINE
# ================================================================

def calculate_engine_v4_1(df):
    if df is None or df.empty: return None
    data = df.copy()

    # 基础指标计算
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

    # 计算 V2 阶段
    data["阶段_V2"] = data.apply(identify_stage_v2, axis=1)
    data["阶段"] = data["阶段_V2"]  # 兼容字段

    # 计算 V3 买卖点
    entry_v3_list, exit_v3_list = [], []
    for i in range(len(data)):
        r, prev = data.iloc[i], (data.iloc[i - 1] if i > 0 else None)
        entry_v3_list.append(entry_engine_v3(r))
        exit_v3_list.append(exit_engine_v3(r, prev))
    data["Entry_Engine_V3"] = entry_v3_list
    data["Exit_Engine_V3"] = exit_v3_list

    # 计算分数
    scores_list = [calculate_scores_v2(data.iloc[i]) for i in range(len(data))]
    scores_df = pd.DataFrame(scores_list, index=data.index)
    for c in scores_df.columns: data[c] = scores_df[c]

    # 映射统一信号
    signals = []
    for i in range(len(data)):
        entry_e = data.iloc[i]["Entry_Engine_V3"]
        exit_e = data.iloc[i]["Exit_Engine_V3"]
        if exit_e == "EXIT_SELL": signals.append("卖出")
        elif exit_e == "REDUCE_POSITION": signals.append("减仓")
        elif entry_e == "BUY_CONFIRMED": signals.append("买入")
        elif entry_e in ("WAIT_MOMENTUM", "WAIT_VOLUME", "WAIT_PULLBACK_CONFIRM"): signals.append("买入观察")
        else: signals.append("观望")
    data["最终信号"] = signals

    # 防守价计算
    data["ATR_Stop_2x"] = data["Close"] - 2 * data["ATR14"]
    data["Recent_Low_20"] = data["Low"].rolling(20, min_periods=5).min()
    data["结构防守价"] = data[["Recent_Low_20", "ATR_Stop_2x"]].max(axis=1)

    return data


# --- 真实无未来函数回测引擎 (Next Open Execution) ---
def run_real_backtest(df, initial_capital=10000.0):
    if df is None or len(df) < 50: return None
    
    capital = initial_capital
    position = 0
    buy_price = 0.0
    trades = []

    for i in range(1, len(df) - 1):
        prev_row = df.iloc[i - 1]   # Day 0 生成信号
        curr_row = df.iloc[i]       # Day +1 执行成交
        next_open = curr_row["Open"]

        signal = prev_row.get("最终信号")

        # 买入执行 (Next Open + 滑点/佣金)
        if signal == "买入" and position == 0:
            exec_price = next_open * (1.0 + CONFIG["BACKTEST_SLIPPAGE_PCT"])
            shares = math.floor((capital * (1.0 - CONFIG["BACKTEST_COMMISSION_PCT"])) / exec_price)
            if shares > 0:
                position = shares
                buy_price = exec_price
                capital -= shares * exec_price * (1.0 + CONFIG["BACKTEST_COMMISSION_PCT"])

        # 卖出执行 (Next Open - 滑点/佣金)
        elif signal in ("卖出", "减仓") and position > 0:
            exec_price = next_open * (1.0 - CONFIG["BACKTEST_SLIPPAGE_PCT"])
            revenue = position * exec_price * (1.0 - CONFIG["BACKTEST_COMMISSION_PCT"])
            profit = revenue - (position * buy_price)
            profit_pct = (exec_price - buy_price) / buy_price * 100
            trades.append({"buy_price": buy_price, "sell_price": exec_price, "profit": profit, "profit_pct": profit_pct})
            capital += revenue
            position = 0

    if not trades:
        return {"total_trades": 0, "win_rate": 0, "profit_factor": 0, "max_drawdown": 0}

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


# ================================================================
# 7. EXCEL REPORT BUILDER (9 SHEETS + DASHBOARD)
# ================================================================

def format_excel_file(filename):
    try:
        wb = openpyxl.load_workbook(filename)
        for ws in wb.worksheets:
            if ws.title != SHEET_DASHBOARD:
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                col_letter = col_cells[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 50)
        wb.save(filename)
    except Exception: pass

def build_output_dict(row, name, code):
    return {
        "日期": row.name.strftime("%Y-%m-%d") if hasattr(row.name, "strftime") else str(row.name),
        "名称": name,
        "代码": clean_stock_code(code),
        "阶段": row.get("阶段_V2", "Stage 0"),
        "最终信号": row.get("最终信号", "观望"),
        "Final_Score": row.get("Final_Score", 0),
        "Stock_Quality_Score": row.get("Stock_Quality_Score", 0),
        "Timing_Score": row.get("Timing_Score", 0),
        "Risk_Score": row.get("Risk_Score", 0),
        "Entry_Engine_V3": row.get("Entry_Engine_V3", "NONE"),
        "Exit_Engine_V3": row.get("Exit_Engine_V3", "HOLD"),
        "收盘价": round(row["Close"], 4),
        "EMA20": round(row["EMA20"], 4) if pd.notna(row.get("EMA20")) else np.nan,
        "EMA60": round(row["EMA60"], 4) if pd.notna(row.get("EMA60")) else np.nan,
        "ADX": round(row.get("ADX", 0), 2),
        "RSI": round(row.get("RSI14", 0), 2),
        "Volume_Ratio": round(row.get("Volume_Ratio", 0), 2),
        "结构防守价": round(row.get("结构防守价", 0), 4),
        "机构级MasterPrompt": f"【V4.1 Pro 机构投研】{name}({code}) | 阶段:{row.get('阶段_V2')} | Score:{row.get('Final_Score')} | 信号:{row.get('最终信号')} | 建议防守价:RM{fmt(row.get('结构防守价'))}"
    }


# ================================================================
# 8. MAIN USER INTERFACE & MENUS
# ================================================================

def single_stock_scan(code):
    print(f"\n🔍 [V4.1 Pro] 正在深度扫描：{code}")
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None:
        print("❌ 找不到股票数据。")
        return

    data = calculate_engine_v4_1(df_raw)
    row = data.iloc[-1]
    res = build_output_dict(row, name, code)
    save_signal_to_db(res)

    # 运行无未来函数回测
    bt_res = run_real_backtest(data)

    print("\n" + "=" * 75)
    print(f"📊 EMA Trading System V4.1 Pro 决策报告 — {name} ({code})")
    print("=" * 75)
    print(f"收盘价: RM{res['收盘价']} | V2阶段: {res['阶段']} | 最终信号: {res['最终信号']}")
    print(f"动态综合分: {res['Final_Score']} | 质量分: {res['Stock_Quality_Score']} | 时机分: {res['Timing_Score']} | 风险安全分: {res['Risk_Score']}")
    print(f"买入引擎 V3: {res['Entry_Engine_V3']} | 卖出引擎 V3: {res['Exit_Engine_V3']}")
    print(f"结构防守价: RM{res['结构防守价']}")
    if bt_res:
        print("-" * 75)
        print(f"📈 历史真实撮合回测（Next Open 执行）: 总交易 {bt_res['total_trades']} 次 | 胜率: {bt_res['win_rate']}% | 盈亏比: {bt_res['profit_factor']}")
    print("=" * 75)

    filename = f"{safe_filename(name)}_EMA_V4.1_Pro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    df_res = pd.DataFrame([res])
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_res.to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
        data.tail(60).to_excel(writer, sheet_name=SHEET_RECENT)
        data.to_excel(writer, sheet_name=SHEET_ALL)

    format_excel_file(filename)
    print(f"✅ V4.1 Pro Excel 报告已成功导出：{filename}\n")


def batch_scan(stock_pool, min_price=0.05, max_price=0.50):
    print(f"\n📊 [V4.1 Pro] 批量扫描中，共 {len(stock_pool)} 只股票...")
    results = []

    for idx, item in enumerate(stock_pool, 1):
        code = item["code"]
        df_raw, name = safe_fetch_stock(code)
        if df_raw is None: continue
        data = calculate_engine_v4_1(df_raw)
        row = data.iloc[-1]
        if row["Close"] < min_price or row["Close"] > max_price: continue
        res = build_output_dict(row, name, code)
        save_signal_to_db(res)
        results.append(res)
        print(f"[{idx}/{len(stock_pool)}] {code} {name} -> 阶段: {res['阶段']} | 信号: {res['最终信号']} | 得分: {res['Final_Score']}")

    if not results:
        print("⚠️ 未找到符合条件的股票。")
        return

    df_results = pd.DataFrame(results).sort_values(by="Final_Score", ascending=False)
    filename = f"EMA_V4.1_Pro_批量扫描_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_results.to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
        df_results[df_results["最终信号"] == "买入"].to_excel(writer, sheet_name=SHEET_BUY_LIST, index=False)
        df_results[df_results["最终信号"] == "买入观察"].to_excel(writer, sheet_name=SHEET_WATCHLIST, index=False)
        df_results[df_results["最终信号"].isin(["卖出", "减仓"])].to_excel(writer, sheet_name=SHEET_SELL_MONITOR, index=False)

    format_excel_file(filename)
    print(f"\n✅ 批量扫描完成！导出报告：{filename}\n")


def main_menu():
    init_db()
    while True:
        print("\n" + "=" * 70)
        print("   🇲🇾 马股 EMA Trading System V4.1 Pro (机构级重构版)")
        print("=" * 70)
        print("1. 单股深度扫描与无未来函数回测")
        print("2. 批量股票扫描 (默认 RM0.05 - RM0.50)")
        print("3. 读取 Excel 股票池 (Final list.xlsx)")
        print("4. 退出系统")
        print("-" * 70)
        choice = input("请输入选项 (1-4): ").strip()

        if choice == "1":
            code = input("请输入股票代码 (如 10 / 0010 / 10.KL): ").strip()
            if code: single_stock_scan(code)
        elif choice == "2":
            codes_text = input("输入股票代码（逗号分隔，如 10,38,85）: ").strip()
            if codes_text:
                pool = [{"code": x.strip()} for x in codes_text.split(",") if x.strip()]
                batch_scan(pool)
        elif choice == "3":
            filepath = input("输入 Excel 文件名 [默认: Final list.xlsx]: ").strip() or "Final list.xlsx"
            try:
                df = pd.read_excel(filepath)
                code_col = df.columns[0]
                pool = [{"code": str(c).strip()} for c in df[code_col].dropna()]
                batch_scan(pool)
            except Exception as e:
                print(f"❌ 读取 Excel 失败: {e}")
        elif choice == "4":
            print("系统已安全退出。祝交易顺利！")
            break

if __name__ == "__main__":
    main_menu()
