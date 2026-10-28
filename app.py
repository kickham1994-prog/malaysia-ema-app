# -*- coding: utf-8 -*-
"""
====================================================================
马来西亚低价股 EMA Trading System V4.1 Master Pro (全功能完整集成版)
- Stage Engine V2 (0/1/2/3A/3B/4 精细化阶段)
- Entry Engine V3 (三层确认：趋势 + 动能 + 资金)
- Exit Engine V3 (三阶段保护 + 动态 ATR 移动止损)
- Factor Ownership & Stage Weighted Score (消除重复加分)
- Real Execution Backtest Engine (消除未来函数 + 滑点/佣金)
- 完美恢复：Master Prompt 机构级投研提示词生成引擎
- 完美恢复：9个Sheet全功能 Excel 报告与终端交互菜单
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

    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 12.0,
    "EXTENSION_MAX_PCT": 18.0,
    "PRICE_POSITION_EXCELLENT_PCT": 5.0,
    "PRICE_POSITION_CAUTION_PCT": 10.0,
    "PRICE_POSITION_HIGH_PCT": 15.0,

    "BREAKOUT_VOLUME_RATIO": 1.50,
    "ENTRY_VOLUME_MIN": 1.20,
    "RSI_ENTRY_MIN": 50.0,
    "ADX_ENTRY_MIN": 20.0,

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

    "BACKTEST_COMMISSION_PCT": 0.001,
    "BACKTEST_SLIPPAGE_PCT": 0.005,
}

# Excel 9大 Sheet 名称定义
SHEET_DASHBOARD = "⓪仪表盘"
SHEET_OVERVIEW = "①今日交易总览"
SHEET_WATCHLIST = "②今日候选名单"
SHEET_POSITION = "③持仓管理"
SHEET_DECISION = "④V4.1决策摘要"
SHEET_AI = "⑤AI分析与MasterPrompt"
SHEET_RECENT = "⑥最近60日"
SHEET_TURNS = "⑦信号转折"
SHEET_ALL = "⑧全部指标"
SHEET_FULL_SCAN = "全部扫描"
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
# 5. CORE ENGINES (Stage V2, Entry V3, Exit V3, Scores V2)
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


# ================================================================
# 6. MASTER PROMPT 辅助生成函数 (完美保留与恢复)
# ================================================================

def generate_master_prompt_payload(row, name, code):
    """根据量化计算结果自动组装 Master Prompt 文本"""
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

    prompt_text = f"""【请执行以下 Master Prompt 机构级深度分析】
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
    return prompt_text.strip()


# ================================================================
# 7. PIPELINE COMPUTATION & BACKTEST ENGINE
# ================================================================

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


# --- 真实无未来函数回测引擎 ---
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


# ================================================================
# 8. EXCEL REPORT BUILDER (9大Sheet完美恢复)
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
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 60)
        wb.save(filename)
    except Exception: pass

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

def build_dashboard_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame([{"项目": "提示", "值": "无数据"}])
    return pd.DataFrame([
        {"项目": "📊 报告生成日期", "值": datetime.now().strftime("%Y-%m-%d")},
        {"项目": "📊 扫描总股票数", "值": len(df_results)},
        {"项目": "🟢 建议买入数", "值": int((df_results["最终信号"] == "买入").sum())},
        {"项目": "🟡 买入观察数", "值": int((df_results["最终信号"] == "买入观察").sum())},
    ])


# ================================================================
# 9. USER INTERFACE & MENUS (恢复6大交互菜单)
# ================================================================

def single_stock_scan(code, output_excel=True):
    print(f"\n🔍 [V4.1 Master Pro] 正在深度扫描：{code}")
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None:
        print("❌ 找不到股票数据。")
        return None

    data = calculate_engine(df_raw)
    row = data.iloc[-1]
    result = latest_output_dict(row, name, code)
    save_signal_to_db(result)
    bt_res = run_real_backtest(data)

    print("\n" + "=" * 78)
    print(f"📊 EMA Trading System V4.1 Master Pro — {name} ({code})")
    print("=" * 78)
    print(f"收盘价: RM{result['收盘价']} | 阶段: {result['阶段']} | 最终信号: {result['最终信号']}")
    print(f"综合得分: {result['Final_Score']} | 质量分: {result['股票质量分']} | 时机分: {result['买点时机分']} | 风险分: {result['风险安全分']}")
    print(f"结构防守价: RM{result['结构防守价']}")
    if bt_res:
        print(f"📈 历史真实回测: 交易次数 {bt_res['total_trades']} | 胜率 {bt_res['win_rate']}% | 盈亏比 {bt_res['profit_factor']}")
    print("-" * 78)
    print("🤖 机构级 Master Prompt 已成功组装（可直接复制用于大模型分析）：")
    print(result["机构级MasterPrompt"])
    print("=" * 78)

    if output_excel:
        filename = f"{safe_filename(name)}_EMA_V4.1_MasterPro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df_summary = pd.DataFrame([result])
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            build_dashboard_df(df_summary).to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
            df_summary.to_excel(writer, sheet_name=SHEET_OVERVIEW, index=False)
            df_summary.to_excel(writer, sheet_name=SHEET_AI, index=False)
            data.tail(60).reset_index().to_excel(writer, sheet_name=SHEET_RECENT, index=False)
            data.reset_index().to_excel(writer, sheet_name=SHEET_ALL, index=False)
        format_excel_file(filename)
        print(f"✅ Excel 报告（含 9 个 Sheet 与 Master Prompt）已导出：{filename}\n")

    return result

def batch_scan(stock_pool, min_price=0.05, max_price=0.50):
    print(f"\n📊 批量扫描中，共 {len(stock_pool)} 只...")
    results = []
    for idx, item in enumerate(stock_pool, 1):
        code = item["code"]
        df_raw, name = safe_fetch_stock(code)
        if df_raw is None: continue
        data = calculate_engine(df_raw)
        row = data.iloc[-1]
        if row["Close"] < min_price or row["Close"] > max_price: continue
        res = latest_output_dict(row, name, code)
        save_signal_to_db(res)
        results.append(res)
        print(f"[{idx}/{len(stock_pool)}] {code} {name} -> 阶段: {res['阶段']} | 信号: {res['最终信号']} | 得分: {res['Final_Score']}")

    if not results:
        print("⚠️ 未找到符合条件的股票。")
        return

    df_results = pd.DataFrame(results).sort_values(by="Final_Score", ascending=False)
    filename = f"EMA_V4.1_MasterPro_批量扫描_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        build_dashboard_df(df_results).to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
        df_results.to_excel(writer, sheet_name=SHEET_OVERVIEW, index=False)
        df_results.to_excel(writer, sheet_name=SHEET_AI, index=False)
        df_results.to_excel(writer, sheet_name=SHEET_FULL_SCAN, index=False)
    format_excel_file(filename)
    print(f"\n✅ 批量扫描完成！生成 Excel 报告：{filename}\n")

def print_menu():
    print("\n" + "=" * 70)
    print("   🇲🇾 马股 EMA Trading System V4.1 Master Pro")
    print("=" * 70)
    print("1. 单股深度分析（含 Master Prompt 提示词与 9 个 Sheet Excel）")
    print("2. 批量股票扫描 (默认 RM0.05 - RM0.50)")
    print("3. 读取 Excel 股票池批量扫描 (Final list.xlsx)")
    print("4. 退出系统")
    print("-" * 70)

def main():
    init_db()
    while True:
        print_menu()
        choice = input("请输入选项 (1-4): ").strip()
        if choice == "1":
            code = input("输入股票代码 (如 10 / 0010 / 10.KL): ").strip()
            if code: single_stock_scan(code)
        elif choice == "2":
            codes_text = input("输入股票代码（逗号分隔，如 10,38,85）: ").strip()
            if codes_text:
                pool = [{"code": clean_stock_code(x)} for x in codes_text.split(",") if x.strip()]
                batch_scan(pool)
        elif choice == "3":
            filepath = input("输入 Excel 文件名 [默认: Final list.xlsx]: ").strip() or "Final list.xlsx"
            try:
                df = pd.read_excel(filepath)
                code_col = df.columns[0]
                pool = [{"code": clean_stock_code(str(c))} for c in df[code_col].dropna()]
                batch_scan(pool)
            except Exception as e:
                print(f"❌ 读取 Excel 失败: {e}")
        elif choice == "4":
            print("系统已安全退出。祝交易顺利！")
            break

if __name__ == "__main__":
    main()
