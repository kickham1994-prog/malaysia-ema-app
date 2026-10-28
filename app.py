# -*- coding: utf-8 -*-
"""
====================================================================
马来西亚低价股 EMA Trading System V4.0.2a (全功能 + Master Prompt 集成版)
「EMA5 + EMA20 + EMA60 + EMA扣底 + ADX + ATR
 + Stock Quality + Entry Quality + Entry/Exit V2 + AI Timing
 + Stage 1 突破追买 + 股票类型识别 + 保护利润分批
 + 顶部预警分 + Setup 记忆机制 + 6大菜单 + 9个Sheet + Master Prompt」
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
# 1. CONFIG
# ================================================================

CONFIG = {
    "EMA_FAST": 5,
    "EMA_SHORT": 20,
    "EMA_LONG": 60,
    "EMA_BACKGROUND": 200,
    "ADX_PERIOD": 14,
    "VOLUME_PERIOD": 20,
    "OBV_MA": 5,
    "BB_PERIOD": 20,
    "BB_STD": 2,

    "HISTORY_CALENDAR_DAYS": 1100,

    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 15.0,
    "EXTENSION_MAX_PCT": 20.0,

    "BREAKOUT_VOLUME_RATIO": 1.50,
    "ENTRY_VOLUME_MIN": 1.20,
    "BREAKOUT_VALID_TOL_PCT": 2.0,

    "ADX_WEAK": 20,
    "ADX_DEVELOPING": 25,
    "ADX_STRONG": 30,

    "DEFAULT_MIN_PRICE": 0.05,
    "DEFAULT_MAX_PRICE": 0.50,
    "DEFAULT_MIN_VOLUME": 100000,

    "SCORE_MAX": 100,

    "STAGE_CONFIRMATION_DAYS": 2,

    "STAGE_ADJUSTMENT": {
        "Stage 1": 5.0,
        "Stage 2": 0.0,
        "Stage 3": -8.0,
        "Stage 4": -18.0,
        "Stage 0": 0.0,
    },
    "STAGE1_BUILD_SCORE": 85,
    "STAGE2_HOLD_SCORE": 80,
    "STAGE3_PROTECT_SCORE": 82,
    "STAGE4_RISK_SCORE": 60,
    "ATR_DISTANCE_HEALTHY": 2.0,
    "ATR_DISTANCE_CAUTION": 4.0,
    "ATR_DISTANCE_HIGH": 6.0,
    "ATR_HEALTH_SCORE_MAX": 15,

    "ENTRY_PULLBACK_EMA20_TOL_PCT": 2.0,
    "ENTRY_PULLBACK_CLOSE_TOL_PCT": 3.0,

    "ENTRY_QUALITY_BUY_MIN": 80,
    "ENTRY_QUALITY_WAIT_MIN": 55,

    "ENTRY_V2_CLOSE_DIST_MAX_PCT": 5.0,
    "ENTRY_V2_ATR_DIST_MAX": 3.0,
    "ENTRY_V2_CROSS_AGE_MAX": 5,

    "AI_BUY_TODAY_ATR_MAX": 2.0,
    "AI_BUY_TODAY_EXT_MAX": 6.0,
    "AI_BUY_TOMORROW_ATR_MAX": 4.0,
    "AI_BUY_TOMORROW_EXT_MAX": 10.0,
    "AI_PULLBACK_BAND_PCT": 2.0,

    "STAGE1_BREAKOUT_AGE_MAX": 3,
    "STAGE1_BREAKOUT_VOL_MIN": 1.50,
    "STAGE1_BREAKOUT_ATR_MAX": 2.50,

    "STAGE1_BREAKOUT_ATR_HARD_MAX": 3.50,
    "STAGE1_BREAKOUT_VOL_STRONG": 3.00,
    "STAGE1_BREAKOUT_CLOSE_POS_MIN": 0.70,

    "STOCK_TYPE_LOOKBACK": 20,
    "STOCK_TYPE_FAR_ATR_THRESHOLD": 2.0,
    "STOCK_TYPE_FAR_DAYS_HIGH": 10,
    "STOCK_TYPE_FAR_DAYS_LOW": 3,

    "PULLBACK_CLOSE_DIST_MAX_PCT": 8.0,

    "BREAKOUT_AGE_MAX_EXTENDED": 5,
    "BREAKOUT_VOL_MIN_EXTENDED": 1.20,

    "ADX_CHOPPY_THRESHOLD": 22,
    "PULLBACK_ATR_NEAR_MAX": 1.8,
    "PULLBACK_EMA5_TOL_PCT": 1.5,
    "STAGE2_ATR_NEAR_MAX": 2.5,
    "RE_ENTRY_ENABLED": True,
    "PROTECT_PROFIT_VOL_CONFIRM": 1.5,
    "STRONG_BUY_STOCK_Q_MIN": 80,
    "STRONG_BUY_ENTRY_Q_MIN": 70,
    "STRONG_BUY_ATR_MAX": 2.5,
    "STAGE2_ADX_ADD_MIN": 20,

    "INITIAL_POSITION": 0.30,
    "ADD_POSITION": 0.30,
    "PROTECT_PROFIT_REDUCE_PCT": 0.33,
    "REDUCE_POSITION_PCT": 0.33,
    "ATR_TRAIL_MULTIPLIER": 1.5,
    "TRAIL_STOP_ENABLED": True,
    "TRAIL_STOP_USE_STRUCTURE": True,
    "TRAIL_STOP_USE_EMA5": True,
    "RE_ENTRY_POSITION": 0.15,
    "RE_ENTRY_REQUIRE_VOL": True,
    "RE_ENTRY_REQUIRE_STAGE": True,
    "RE_ENTRY_VOL_MIN": 1.20,

    "TOP_WARNING_ENABLED": True,
    "TOP_WARNING_REDUCE_1": 45,
    "TOP_WARNING_REDUCE_2": 70,
    "TOP_WARNING_EXIT": 90,
    "TOP_WARNING_ADX_HIGH": 45,
    "TOP_WARNING_DI_SPREAD_MIN": 10,
    "TOP_WARNING_VOL_RATIO": 2.0,
    "TOP_WARNING_CLOSE_POS": 0.4,
    "TOP_WARNING_EXTENSION": 15.0,
    "TOP_WARNING_ATR_DIST": 4.0,
    "TOP_WARNING_ENTRY_Q_LOW": 35,

    "STAGE3_REDUCE_MIN_CONDITIONS": 2,

    "STAGE1_ADX_MIN": 20,
    "BLOCK_BUY_WHEN_EMA60_DOWN": True,

    "WATCH_VOL_TRIGGER": 1.50,
    "WATCH_VOL_NEAR": 1.20,
    "WATCH_EMA20_NEAR_PCT": 0.02,
    "WATCH_EMA20_FAR_PCT": 0.05,

    "SETUP_VALID_DAYS": 3,
    "BREAKOUT_MEMORY_VALID_DAYS": 3,
    "ATR_LEVEL_EXCELLENT": 0.8,
    "ATR_LEVEL_ACCEPTABLE": 1.2,
    "OVEREXTEND_ATR_MAX": 1.8,
    "OVEREXTEND_EXTENSION_MAX": 12.0,
    "EMA5_RECLAIM_FRESH_DAYS": 5,
    "STRONG_CLOSE_POS_MIN": 0.70,
    "STRONG_CLOSE_VOL_MIN": 1.20,
}


# ================================================================
# SHEET 名称
# ================================================================

SHEET_DASHBOARD = "⓪仪表盘"
SHEET_OVERVIEW = "①今日交易总览"
SHEET_WATCHLIST = "②今日候选名单"
SHEET_POSITION = "③持仓管理"
SHEET_DECISION = "④V3.8决策摘要"
SHEET_AI = "⑤AI分析"
SHEET_RECENT = "⑥最近60日"
SHEET_TURNS = "⑦信号转折"
SHEET_ALL = "⑧全部指标"
SHEET_FULL_SCAN = "全部扫描"
SHEET_FOCUS = "重点关注"
SHEET_OBSERVE = "观察候选"
SHEET_CLEAR = "清除"
SHEET_ERRORS = "无数据或错误"


TIMING_PRIORITY_ORDER = {
    "BUY_TODAY": 1,
    "BUY_TOMORROW": 2,
    "BUY_ON_VOLUME": 3,
    "BUY_ON_EMA5_RECLAIM": 4,
    "BUY_ON_PULLBACK_EMA20": 5,
    "BUY_ON_BREAKOUT": 6,
    "HOLD": 7,
    "NO_BUY": 8,
}


# ================================================================
# 2. BASIC UTILITIES
# ================================================================

def format_malaysia_ticker(code):
    if code is None:
        return None
    s = str(code).strip().upper()
    if s.endswith(".KL"):
        return s
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if s.isdigit():
        s = s.zfill(4)
    return f"{s}.KL"


def clean_stock_code(code):
    if code is None:
        return ""
    s = str(code).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def safe_filename(text):
    text = str(text or "stock")
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    return text[:80].strip() or "stock"


def fmt(v, digits=4):
    if v is None or pd.isna(v):
        return "N/A"
    try:
        return f"{float(v):.{digits}f}"
    except Exception:
        return str(v)


def pct(v, digits=2):
    if v is None or pd.isna(v):
        return "N/A"
    return f"{float(v):.{digits}f}%"


def extract_price_range(zone_str):
    if zone_str is None or (isinstance(zone_str, float) and pd.isna(zone_str)):
        return np.nan, np.nan
    s = str(zone_str)
    nums = re.findall(r"RM(\d+\.?\d*)", s)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    if len(nums) == 1:
        v = float(nums[0])
        return v, v
    return np.nan, np.nan


def classify_action(comprehensive_advice):
    s = str(comprehensive_advice or "")
    if "今天可买" in s: return "今日可买"
    if "明日确认" in s: return "明日确认"
    if "等放量" in s: return "等放量"
    if "等回" in s or "回 EMA20" in s: return "等回EMA20"
    if "等 EMA5" in s: return "等EMA5"
    if "等突破" in s: return "等突破"
    if "持有不加仓" in s: return "持有不加仓"
    if "不买" in s: return "不买"
    return "其他"


# ================================================================
# 3. MASTER PROMPT 辅助生成函数
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
    entry_q = fmt(row.get("Entry_Quality_Score"))
    stage = row.get("阶段", "Stage 0")
    support = fmt(row.get("结构防守价"))
    resistance = fmt(row.get("Prev_High_20", row.get("Close") * 1.05))
    top_warn = fmt(row.get("Top_Warning_Score", 0))

    prompt_text = f"""【请执行以下 Master Prompt 机构级深度分析】
股票名称: {name} ({clean_stock_code(code)})
当前收盘价: RM{close}
量化系统核心数据:
- 市场阶段: {stage}
- 股票质量分: {stock_q} / 100
- 买点质量分: {entry_q} / 100
- 均线参考: EMA20=RM{ema20}, EMA60=RM{ema60}
- 波动与动能: ATR=RM{atr}, ADX={adx}, Volume Ratio={vol_ratio}
- 防守与风险: 结构防守价=RM{support}, 顶部预警分={top_warn}
- 近期压力位: RM{resistance}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告。"""
    return prompt_text.strip()


# ================================================================
# 4. DATA LOADER
# ================================================================

def safe_fetch_stock(code, calendar_days=None):
    if calendar_days is None:
        calendar_days = CONFIG["HISTORY_CALENDAR_DAYS"]
    ticker = format_malaysia_ticker(code)

    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=calendar_days)

        df = yf.download(
            ticker,
            start=start_date,
            end=end_date + timedelta(days=1),
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=False,
        )

        if df is None or df.empty:
            return None, None

        if isinstance(df.columns, pd.MultiIndex):
            try:
                df.columns = df.columns.get_level_values(0)
            except Exception:
                df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]

        needed = ["Open", "High", "Low", "Close", "Volume"]
        missing = [c for c in needed if c not in df.columns]
        if missing:
            return None, None

        df = df[needed].copy()
        for c in needed:
            df[c] = pd.to_numeric(df[c], errors="coerce")

        df = df.dropna(subset=["High", "Low", "Close"]).copy()
        df = df[df["High"] != df["Low"]].copy()
        df = df[df["Volume"].fillna(0) > 0].copy()
        df = df[~df.index.duplicated(keep="last")].sort_index()
        df.index.name = "Date"

        if len(df) < 100:
            return None, None

        name = str(code)
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            name = (
                info.get("shortName")
                or info.get("longName")
                or info.get("displayName")
                or str(code)
            )
        except Exception:
            pass

        return df, name

    except Exception as e:
        return None, None


# ================================================================
# 5. WILDER RMA / ADX / GAP / DEDUCT
# ================================================================

def wilder_rma(series, period):
    s = pd.Series(series, dtype="float64")
    result = pd.Series(np.nan, index=s.index, dtype="float64")
    values = s.to_numpy(dtype=float)
    out = np.full(len(values), np.nan)
    valid_positions = np.where(~np.isnan(values))[0]

    if len(valid_positions) < period:
        return result

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
    high = pd.to_numeric(df["High"], errors="coerce")
    low = pd.to_numeric(df["Low"], errors="coerce")
    close = pd.to_numeric(df["Close"], errors="coerce")

    up_move = high.diff()
    down_move = -low.diff()

    plus_dm = pd.Series(np.nan, index=df.index, dtype="float64")
    minus_dm = pd.Series(np.nan, index=df.index, dtype="float64")
    valid = up_move.notna() & down_move.notna()

    plus_dm.loc[valid] = np.where((up_move.loc[valid] > down_move.loc[valid]) & (up_move.loc[valid] > 0), up_move.loc[valid], 0.0)
    minus_dm.loc[valid] = np.where((down_move.loc[valid] > up_move.loc[valid]) & (down_move.loc[valid] > 0), down_move.loc[valid], 0.0)

    tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
    tr.iloc[0] = np.nan

    tr_rma = wilder_rma(tr, period)
    plus_dm_rma = wilder_rma(plus_dm, period)
    minus_dm_rma = wilder_rma(minus_dm, period)

    plus_di = 100.0 * plus_dm_rma / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * minus_dm_rma / tr_rma.replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx = wilder_rma(dx, period)

    return plus_di, minus_di, dx, adx


def classify_distance(distance, distance_change, price):
    if pd.isna(distance) or pd.isna(price) or price == 0: return "数据不足"
    threshold = abs(price) * CONFIG["GAP_EPSILON_PCT"] / 100.0
    if distance > threshold:
        if pd.notna(distance_change) and distance_change > 0: return "价格上偏离扩大"
        elif pd.notna(distance_change) and distance_change < 0: return "价格上偏离收窄"
        return "价格上偏离"
    if distance < -threshold:
        if pd.notna(distance_change) and distance_change < 0: return "价格下偏离扩大"
        elif pd.notna(distance_change) and distance_change > 0: return "价格下偏离收窄"
        return "价格下偏离"
    if pd.notna(distance_change):
        if distance_change > 0: return "接近上穿"
        elif distance_change < 0: return "接近下穿"
    return "贴合均线"


def classify_deduct(deduct, deduct_change, price):
    if pd.isna(deduct) or pd.isna(price) or price == 0: return "数据不足"
    threshold = abs(price) * CONFIG["GAP_EPSILON_PCT"] / 100.0
    if deduct > threshold:
        if pd.notna(deduct_change) and deduct_change > 0: return "正向扣底扩大(动能增强)"
        elif pd.notna(deduct_change) and deduct_change < 0: return "正向扣底收窄(动能放缓)"
        return "正向扣底"
    if deduct < -threshold:
        if pd.notna(deduct_change) and deduct_change < 0: return "负向扣底扩大(下行增强)"
        elif pd.notna(deduct_change) and deduct_change > 0: return "负向扣底收窄(下行放缓)"
        return "负向扣底"
    if pd.notna(deduct_change):
        if deduct_change > 0: return "扣底转正"
        elif deduct_change < 0: return "扣底转负"
    return "扣底中性"


def arrow_trend(series, max_days=5):
    s = pd.Series(series).dropna().tail(max_days)
    if len(s) < 2: return "N/A"
    arrows = []
    for i in range(1, len(s)):
        if s.iloc[i] > s.iloc[i - 1]: arrows.append("↑")
        elif s.iloc[i] < s.iloc[i - 1]: arrows.append("↓")
        else: arrows.append("→")
    return "".join(arrows)


# ================================================================
# 6. SCORES & ENGINES
# ================================================================

def clip_score(x, low=0, high=100):
    return float(np.clip(x, low, high))


def get_stage_key(stage):
    m = re.search(r"Stage\s*(\d)", str(stage))
    return f"Stage {m.group(1)}" if m else "Stage 0"


def calculate_stage_adjustment(stage):
    return float(CONFIG["STAGE_ADJUSTMENT"].get(get_stage_key(stage), 0.0))


def calculate_stock_quality_score(data):
    row = data.iloc[-1]
    close = row["Close"]
    ema20 = row["EMA20"]
    ema60 = row["EMA60"]

    price_score = 0.0
    if pd.notna(ema20) and ema20 > 0:
        distance20 = (close - ema20) / ema20 * 100
        price_score += np.clip(12 + distance20 * 0.35, 0, 17) if distance20 >= 0 else np.clip(12 + distance20 * 0.50, 0, 12)
    if pd.notna(ema60) and ema60 > 0:
        distance60 = (close - ema60) / ema60 * 100
        price_score += np.clip(5 + distance60 * 0.18, 0, 8)
    price_score = clip_score(price_score, 0, 25)

    money_score = 0.0
    obv_change = row.get("OBV_Change_5", np.nan)
    obv_ma_distance = row.get("OBV_MA_Distance_Pct", np.nan)
    avg_volume = row.get("Volume_MA20", np.nan)

    if pd.notna(obv_change):
        money_score += np.clip(5.0 + (float(obv_change) / (float(avg_volume) * 5.0)) * 5.0, 0, 10) if pd.notna(avg_volume) and avg_volume > 0 else 5.0
    if pd.notna(obv_ma_distance):
        money_score += np.clip(5 + obv_ma_distance * 0.5, 0, 10)

    vr = row.get("Volume_Ratio", np.nan)
    if pd.notna(vr):
        money_score += np.clip((vr - 1.0) * 3, 0, 5)

    money_score = clip_score(money_score, 0, 20)

    trend_score = 0.0
    gap20_pct = row["Gap20_Pct"]
    gap60_pct = row["Gap60_Pct"]
    adx = row["ADX"]
    adx_change = row["ADX_Change_5"]
    di_spread = row["DI_Spread"]
    if pd.notna(gap20_pct): trend_score += np.clip(5 + gap20_pct * 0.8, 0, 8)
    if pd.notna(gap60_pct): trend_score += np.clip(5 + gap60_pct * 0.4, 0, 7)
    if pd.notna(adx): trend_score += np.clip((adx - 15) * 0.20, 0, 4)
    if pd.notna(adx_change): trend_score += np.clip(adx_change * 0.10, 0, 1)
    if pd.notna(di_spread): trend_score += np.clip(di_spread * 0.03, 0, 1)
    trend_score = clip_score(trend_score, 0, 20)

    candle_score = 0.0
    body = row["Candle_Body_Pct"]
    close_position = row["Close_Position_Range"]
    return_5 = row["Return_5"]
    if pd.notna(body): candle_score += np.clip(7 + body * 0.15, 0, 10)
    if pd.notna(close_position): candle_score += np.clip(close_position * 8, 0, 8)
    if pd.notna(return_5): candle_score += np.clip(2 + return_5 * 0.15, 0, 5)
    candle_score = clip_score(candle_score, 0, 20)

    atr_distance = row.get("ATR_Distance", np.nan)
    atr_pct = row.get("ATR_Pct", np.nan)
    extension = row.get("Extension_EMA20_Pct", np.nan)
    permission = row.get("Trading_Permission", "YELLOW")
    atr_health_score = float(CONFIG["ATR_HEALTH_SCORE_MAX"])

    if pd.notna(atr_distance):
        ad = abs(float(atr_distance))
        if ad <= CONFIG["ATR_DISTANCE_HEALTHY"]: atr_health_score -= ad * 0.5
        elif ad <= CONFIG["ATR_DISTANCE_CAUTION"]: atr_health_score -= 1.0 + (ad - 2.0) * 1.5
        elif ad <= CONFIG["ATR_DISTANCE_HIGH"]: atr_health_score -= 4.0 + (ad - 4.0) * 1.5
        else: atr_health_score -= 7.0 + min((ad - 6.0) * 1.0, 5.0)

    if pd.notna(atr_pct): atr_health_score -= np.clip((float(atr_pct) - 4.0) * 0.35, 0, 3)
    if pd.notna(extension):
        if extension > CONFIG["EXTENSION_MAX_PCT"]: atr_health_score -= 3
        elif extension > CONFIG["EXTENSION_CAUTION_PCT"]: atr_health_score -= 1
    if permission == "YELLOW": atr_health_score -= 1
    elif permission == "RED": atr_health_score -= 5
    atr_health_score = clip_score(atr_health_score, 0, CONFIG["ATR_HEALTH_SCORE_MAX"])

    atr_risk_score = float(CONFIG["ATR_HEALTH_SCORE_MAX"]) - atr_health_score

    raw_score = price_score + money_score + trend_score + candle_score + atr_health_score
    stage = row.get("阶段", "Stage 0 (观察)")
    stage_adjustment = calculate_stage_adjustment(stage)
    final_score = clip_score(raw_score + stage_adjustment, 0, CONFIG["SCORE_MAX"])

    return {
        "Price_Structure_Score": round(price_score, 2),
        "Money_Flow_Score": round(money_score, 2),
        "Trend_Momentum_Score": round(trend_score, 2),
        "Candlestick_Momentum_Score": round(candle_score, 2),
        "ATR_Health_Score": round(atr_health_score, 2),
        "ATR_Risk_Score": round(atr_risk_score, 2),
        "Risk_Control_Score": round(atr_health_score, 2),
        "Raw_Score": round(raw_score, 2),
        "Stage_Adjustment": round(stage_adjustment, 2),
        "Stock_Quality_Score": round(final_score, 2),
        "Final_Score": round(final_score, 2),
        "Dynamic_Total_Score": round(final_score, 2),
    }


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


def calculate_entry_quality_score(row):
    close = row.get("Close", np.nan)
    ema20 = row.get("EMA20", np.nan)
    atr_distance = row.get("ATR_Distance", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)
    cross_age = row.get("EMA5_Cross_Age", np.nan)
    breakout_age = row.get("Breakout_Age", np.nan)

    ema_score = 0
    if pd.notna(close) and pd.notna(ema20) and ema20 > 0:
        dist_pct = (close - ema20) / ema20 * 100
        if dist_pct < -3: ema_score = 8
        elif dist_pct < 0: ema_score = 15
        elif dist_pct <= 3: ema_score = 30
        elif dist_pct <= 6: ema_score = 22
        elif dist_pct <= 10: ema_score = 10
        else: ema_score = -10

    cross_score = 0
    if pd.notna(cross_age):
        if cross_age <= 5: cross_score = 20
        elif cross_age <= 10: cross_score = 12
        elif cross_age <= 15: cross_score = 5

    atr_score = 0
    if pd.notna(atr_distance):
        ad = abs(float(atr_distance))
        if ad < 2: atr_score = 25
        elif ad < 4: atr_score = 10
        elif ad < 6: atr_score = -10
        else: atr_score = -20

    brk_score = 0
    if pd.notna(breakout_age) and breakout_age != 999:
        if breakout_age <= 5: brk_score = 15
        elif breakout_age <= 10: brk_score = 8

    vol_score = 0
    if pd.notna(volume_ratio):
        if volume_ratio >= 1.5: vol_score = 10
        elif volume_ratio >= 1.2: vol_score = 7
        elif volume_ratio >= 1.0: vol_score = 4

    total = float(np.clip(ema_score + cross_score + atr_score + brk_score + vol_score, 0, 100))
    rating = "优秀" if total >= 80 else ("良好" if total >= 65 else ("普通" if total >= 50 else ("偏差" if total >= 35 else "差(追高)")))

    return {
        "Entry_Quality_Score": round(total, 2),
        "Entry_Quality_Rating": rating,
        "EQ_EMA20_Position": round(float(ema_score), 2),
        "EQ_EMA5_Cross": round(float(cross_score), 2),
        "EQ_ATR_Distance": round(float(atr_score), 2),
        "EQ_Breakout_Age": round(float(brk_score), 2),
        "EQ_Volume": round(float(vol_score), 2),
    }


def stock_quality_rating(score):
    if pd.isna(score): return "N/A"
    s = float(score)
    if s >= 85: return "优秀"
    if s >= 70: return "良好"
    if s >= 55: return "普通"
    if s >= 40: return "偏差"
    return "差"


def combined_buy_rating(stock_q, entry_q, timing_code):
    if timing_code == "NO_BUY": return "🚫 不买"
    if timing_code == "HOLD": return "🟡 持有不加仓"
    if timing_code == "BUY_TODAY": return "🟢 今天可买"
    if timing_code == "BUY_TOMORROW": return "🟢 明日确认后买"
    if timing_code == "BUY_ON_PULLBACK_EMA20": return "🔵 等回 EMA20"
    if timing_code == "BUY_ON_VOLUME": return "🔵 等放量"
    if timing_code == "BUY_ON_EMA5_RECLAIM": return "🔵 等 EMA5 站上"
    if timing_code == "BUY_ON_BREAKOUT": return "🔵 等突破"
    return "⚪ 观察"


def classify_atr_level(atr_distance):
    if pd.isna(atr_distance): return "UNKNOWN"
    ad = abs(float(atr_distance))
    if ad <= CONFIG["ATR_LEVEL_EXCELLENT"]: return "EXCELLENT"
    if ad <= CONFIG["ATR_LEVEL_ACCEPTABLE"]: return "ACCEPTABLE"
    return "TOO_FAR"


def check_overextended(row):
    atr_distance = row.get("ATR_Distance", np.nan)
    extension = row.get("Extension_EMA20_Pct", np.nan)
    ad = abs(float(atr_distance)) if pd.notna(atr_distance) else 0.0
    ext = abs(float(extension)) if pd.notna(extension) else 0.0
    over = (ad >= CONFIG["OVEREXTEND_ATR_MAX"] and ext >= CONFIG["OVEREXTEND_EXTENSION_MAX"])
    return over, ad, ext


def compute_setup_memory_series(data):
    n = len(data)
    setup_type_arr = ["NONE"] * n
    setup_age_arr = [999] * n
    setup_active_arr = [False] * n
    setup_brk_level_arr = [np.nan] * n
    brk_confirmed_arr = [False] * n

    valid_days = CONFIG["SETUP_VALID_DAYS"]
    prev_type, prev_age, prev_level, prev_confirmed = "NONE", 999, np.nan, False

    for i in range(n):
        r = data.iloc[i]
        stage_key = get_stage_key(r.get("阶段", "Stage 0"))
        stage_ok = stage_key in ("Stage 1", "Stage 2")

        close, low = r.get("Close", np.nan), r.get("Low", np.nan)
        ema5, ema20 = r.get("EMA5", np.nan), r.get("EMA20", np.nan)
        atr_distance, volume_ratio = r.get("ATR_Distance", np.nan), r.get("Volume_Ratio", np.nan)

        cur_type, cur_age, cur_level, cur_confirmed = "NONE", 999, np.nan, False

        if prev_type != "NONE" and prev_age < valid_days:
            if prev_type == "EMA20_PULLBACK":
                if pd.notna(close) and pd.notna(ema20) and close > ema20 * 0.97:
                    cur_type, cur_age = prev_type, prev_age + 1
            elif prev_type == "BREAKOUT":
                if pd.notna(close) and pd.notna(prev_level) and close >= prev_level * 0.98:
                    cur_type, cur_age, cur_level, cur_confirmed = prev_type, prev_age + 1, prev_level, prev_confirmed

        if cur_type == "NONE" and stage_ok:
            touched_ema20 = False
            if pd.notna(low) and pd.notna(ema20) and low <= ema20 * (1 + CONFIG["ENTRY_PULLBACK_EMA20_TOL_PCT"] / 100.0): touched_ema20 = True
            if not touched_ema20 and pd.notna(close) and pd.notna(ema20) and close <= ema20 * (1 + CONFIG["ENTRY_PULLBACK_CLOSE_TOL_PCT"] / 100.0): touched_ema20 = True
            if not touched_ema20 and pd.notna(atr_distance) and abs(float(atr_distance)) <= CONFIG["PULLBACK_ATR_NEAR_MAX"]: touched_ema20 = True
            if not touched_ema20 and pd.notna(low) and pd.notna(ema5) and low <= ema5 * (1 + CONFIG["PULLBACK_EMA5_TOL_PCT"] / 100.0): touched_ema20 = True
            if touched_ema20: cur_type, cur_age = "EMA20_PULLBACK", 0

        if cur_type == "NONE" and stage_ok:
            prev_high = r.get("Prev_High_20", np.nan)
            if pd.notna(prev_high) and pd.notna(close):
                breakout_level = prev_high * (1.0 - CONFIG["BREAKOUT_VALID_TOL_PCT"] / 100.0)
                if close >= breakout_level:
                    cur_type, cur_age, cur_level = "BREAKOUT", 0, breakout_level

        if cur_type == "BREAKOUT":
            if pd.notna(volume_ratio) and volume_ratio >= CONFIG["STAGE1_BREAKOUT_VOL_MIN"]:
                cur_confirmed = True

        setup_type_arr[i] = cur_type
        setup_age_arr[i] = cur_age
        setup_active_arr[i] = (cur_type != "NONE" and cur_age <= valid_days)
        setup_brk_level_arr[i] = cur_level
        brk_confirmed_arr[i] = cur_confirmed

        prev_type, prev_age, prev_level, prev_confirmed = cur_type, cur_age, cur_level, cur_confirmed

    data["Setup_Type"] = setup_type_arr
    data["Setup_Age"] = setup_age_arr
    data["Setup_Active"] = setup_active_arr
    data["Setup_Breakout_Level"] = setup_brk_level_arr
    data["Breakout_Confirmed"] = brk_confirmed_arr
    return data


def entry_engine_v2(row, prev_row):
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    stock_type = row.get("Stock_Type", "mixed")
    close, ema5, ema20 = row.get("Close", np.nan), row.get("EMA5", np.nan), row.get("EMA20", np.nan)
    gap20, volume_ratio = row.get("Gap20", np.nan), row.get("Volume_Ratio", np.nan)
    atr_distance, cross_age = row.get("ATR_Distance", np.nan), row.get("EMA5_Cross_Age", np.nan)
    breakout_age, close_position = row.get("Breakout_Age", np.nan), row.get("Close_Position_Range", np.nan)

    setup_type, setup_active = row.get("Setup_Type", "NONE"), bool(row.get("Setup_Active", False))
    breakout_confirmed = bool(row.get("Breakout_Confirmed", False))

    if pd.isna(ema5) or pd.isna(ema20) or pd.isna(close): return "NONE"

    atr_level = classify_atr_level(atr_distance)
    overextended, _, _ = check_overextended(row)

    prev_close = prev_row.get("Close", np.nan) if prev_row is not None else np.nan
    prev_ema5 = prev_row.get("EMA5", np.nan) if prev_row is not None else np.nan
    prev_ema20 = prev_row.get("EMA20", np.nan) if prev_row is not None else np.nan

    ema5_reclaim = False
    if pd.notna(prev_close) and pd.notna(prev_ema5) and prev_close < prev_ema5 and close > ema5: ema5_reclaim = True
    if pd.notna(prev_ema5) and pd.notna(prev_ema20) and prev_ema5 <= prev_ema20 and ema5 > ema20: ema5_reclaim = True
    if pd.notna(cross_age) and cross_age <= CONFIG["EMA5_RECLAIM_FRESH_DAYS"] and ema5 > ema20 and close >= ema5: ema5_reclaim = True

    strong_close = (pd.notna(close_position) and close_position >= CONFIG["STRONG_CLOSE_POS_MIN"] and pd.notna(volume_ratio) and volume_ratio >= CONFIG["STRONG_CLOSE_VOL_MIN"])

    if stage == "Stage 1":
        above_ema20, gap20_ok, ema5_above_ema20 = close >= ema20, pd.notna(gap20) and gap20 >= 0, ema5 > ema20
        breakout_age_max = CONFIG["BREAKOUT_AGE_MAX_EXTENDED"] if stock_type == "breakout" else CONFIG["STAGE1_BREAKOUT_AGE_MAX"]
        breakout_vol_min = CONFIG["BREAKOUT_VOL_MIN_EXTENDED"] if stock_type == "breakout" else CONFIG["STAGE1_BREAKOUT_VOL_MIN"]

        breakout_ok = (pd.notna(breakout_age) and breakout_age != 999 and breakout_age <= breakout_age_max)
        volume_strong = (pd.notna(volume_ratio) and volume_ratio >= breakout_vol_min)
        atr_soft_ok = (pd.notna(atr_distance) and abs(float(atr_distance)) < CONFIG["STAGE1_BREAKOUT_ATR_MAX"])
        atr_hard_ok = (pd.notna(atr_distance) and abs(float(atr_distance)) <= CONFIG["STAGE1_BREAKOUT_ATR_HARD_MAX"])

        if above_ema20 and gap20_ok and ema5_above_ema20 and breakout_ok and volume_strong and atr_soft_ok and not overextended: return "BUY"
        if above_ema20 and gap20_ok and ema5_above_ema20 and breakout_ok and volume_strong and atr_hard_ok:
            if (pd.notna(volume_ratio) and volume_ratio >= CONFIG["STAGE1_BREAKOUT_VOL_STRONG"]) and (pd.notna(close_position) and close_position >= CONFIG["STAGE1_BREAKOUT_CLOSE_POS_MIN"]) and not overextended:
                return "BUY_STRONG_BREAKOUT"

        if setup_active and (ema5_reclaim or strong_close):
            return "WAIT_PULLBACK" if overextended or atr_level == "TOO_FAR" else "BUY"

        if above_ema20 and gap20_ok and ema5_above_ema20:
            if overextended or atr_level == "TOO_FAR": return "WAIT_PULLBACK"
            return "BUY" if ema5_reclaim else "WAIT_EMA5"
        if above_ema20 and gap20_ok: return "WAIT_EMA5"
        return "NONE"

    if stage == "Stage 2":
        if setup_active:
            if ema5_reclaim or breakout_confirmed or strong_close:
                return "WAIT_PULLBACK" if overextended or atr_level == "TOO_FAR" else "BUY"
            return "WAIT_EMA5" if setup_type == "EMA20_PULLBACK" else ("WAIT_VOL" if setup_type == "BREAKOUT" else "WAIT_EMA5")
        return "WAIT_PULLBACK"

    return "NONE"


def exit_engine_v2(row, prev_row):
    close, ema5, ema20 = row.get("Close", np.nan), row.get("EMA5", np.nan), row.get("EMA20", np.nan)
    ema20_slope, ema60_slope = row.get("EMA20_Actual_Slope_5", np.nan), row.get("EMA60_Actual_Slope_5", np.nan)
    adx, volume_ratio = row.get("ADX", np.nan), row.get("Volume_Ratio", np.nan)

    if pd.isna(close) or pd.isna(ema5) or pd.isna(ema20): return "数据不足"
    if pd.notna(ema60_slope) and ema60_slope < 0: return "EXIT"

    if prev_row is not None:
        prev_close, prev_ema20 = prev_row.get("Close", np.nan), prev_row.get("EMA20", np.nan)
        if pd.notna(prev_close) and pd.notna(prev_ema20) and close < ema20 and prev_close < prev_ema20:
            return "减仓"

    if pd.notna(adx) and adx < CONFIG["ADX_CHOPPY_THRESHOLD"]:
        return "观察" if close < ema5 else "持有"

    if pd.notna(ema20_slope) and ema20_slope < 0:
        if close < ema20 and ema5 < ema20: return "保护利润"
        if close < ema5 and pd.notna(volume_ratio) and volume_ratio > CONFIG["PROTECT_PROFIT_VOL_CONFIRM"]: return "保护利润"
        return "观察"

    return "观察" if close < ema5 else "持有"


def calculate_top_warning_score(row, prev_rows):
    score = 0
    reasons = []
    adx, adx_chg = row.get("ADX", np.nan), row.get("ADX_Change_5", np.nan)
    if pd.notna(adx) and adx > CONFIG["TOP_WARNING_ADX_HIGH"] and pd.notna(adx_chg) and adx_chg < 0:
        score += 25; reasons.append("ADX见顶回落")

    di_spread = row.get("DI_Spread", np.nan)
    if pd.notna(di_spread) and di_spread < CONFIG["TOP_WARNING_DI_SPREAD_MIN"]:
        score += 15; reasons.append("DI收窄")

    vr, cpr = row.get("Volume_Ratio", np.nan), row.get("Close_Position_Range", np.nan)
    if pd.notna(vr) and vr > CONFIG["TOP_WARNING_VOL_RATIO"] and pd.notna(cpr) and cpr < CONFIG["TOP_WARNING_CLOSE_POS"]:
        score += 15; reasons.append("放量滞涨")

    ext = row.get("Extension_EMA20_Pct", np.nan)
    if pd.notna(ext) and ext > CONFIG["TOP_WARNING_EXTENSION"]:
        score += 15; reasons.append("Extension过大")

    return score, reasons


def ai_timing_engine(row):
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    entry = row.get("Entry_Engine", "NONE")
    exit_e = row.get("Exit_Engine", "持有")
    permission = row.get("Trading_Permission", "YELLOW")
    stock_q = float(row.get("Stock_Quality_Score", row.get("Final_Score", 0)) or 0)
    entry_q = float(row.get("Entry_Quality_Score", 50) or 0)
    atr_distance = row.get("ATR_Distance", np.nan)

    ad = abs(float(atr_distance)) if pd.notna(atr_distance) else np.nan

    if CONFIG["TOP_WARNING_ENABLED"]:
        top_score = float(row.get("Top_Warning_Score", 0) or 0)
        top_reasons = row.get("Top_Warning_Reasons", "")
        if top_score >= CONFIG["TOP_WARNING_EXIT"]: return "NO_BUY", "高", "清仓", f"顶部预警分{top_score:.0f}：{top_reasons}"
        if top_score >= CONFIG["TOP_WARNING_REDUCE_2"]: return "NO_BUY", "高", "减仓1/3", f"顶部预警分{top_score:.0f}：{top_reasons}"

    if exit_e == "EXIT": return "NO_BUY", "高", "退出", "EMA60 下弯，趋势破坏"
    if exit_e == "保护利润": return "NO_BUY", "高", "减仓1/3", "EMA20 下弯，减仓1/3锁利"

    if entry in ("BUY", "BUY_STRONG_BREAKOUT"):
        if stock_q >= 80 and entry_q >= 70:
            return "BUY_TODAY", "高", "建仓", f"买点优秀({entry_q:.0f}) + 结构良好"
        return "BUY_TOMORROW", "中", "观察", f"买点一般({entry_q:.0f})，等待次日确认"

    if entry == "WAIT_VOL": return "BUY_ON_VOLUME", "中", "观察", "等量能放大"
    if entry == "WAIT_EMA5": return "BUY_ON_EMA5_RECLAIM", "中", "观察", "等 EMA5 重新站上"
    if entry == "WAIT_PULLBACK": return "BUY_ON_PULLBACK_EMA20", "低", "等回调", "等回调至 EMA20 附近"

    if stage == "Stage 1": return "BUY_ON_BREAKOUT", "低", "观察", "Stage 1 结构待突破"
    if stage == "Stage 2": return "HOLD", "低", "持有", "Stage 2 持有，无加仓触发"

    return "NO_BUY", "低", "观望", "未形成明确趋势"


def ai_buy_zone(row, timing_code):
    close, ema20 = row.get("Close", np.nan), row.get("EMA20", np.nan)
    if pd.isna(close): return "N/A"
    band = CONFIG["AI_PULLBACK_BAND_PCT"] / 100.0

    if timing_code in ["BUY_TODAY", "BUY_TOMORROW", "BUY_ON_VOLUME", "BUY_ON_EMA5_RECLAIM"]:
        return f"RM{close*(1-band):.4f} – RM{close*(1+band):.4f}"
    if timing_code == "BUY_ON_PULLBACK_EMA20" and pd.notna(ema20):
        return f"RM{ema20*(1-band):.4f} – RM{ema20*(1+band):.4f}（EMA20 附近）"
    return "N/A"


def ai_position_suggestion(row, timing_code, action_level):
    if timing_code == "NO_BUY": return "0%（不新增）"
    if action_level == "建仓": return "首仓 30%，确认后加 30%"
    if action_level == "持有": return "维持现有仓位"
    return "0%（等确认后建仓）"


def ai_invalidation(row, timing_code):
    ema20 = row.get("EMA20", np.nan)
    if timing_code == "NO_BUY": return "等 Stage 回到 Stage 1/2 再重新评估。"
    return f"若 Close < EMA20（≈{ema20:.4f}）连续2天，取消买入计划。"


def ai_timing_text(timing_code):
    mapping = {
        "BUY_TODAY": "今天可买", "BUY_TOMORROW": "明日确认后买",
        "BUY_ON_PULLBACK_EMA20": "等回 EMA20 再买", "BUY_ON_VOLUME": "等放量再买",
        "BUY_ON_EMA5_RECLAIM": "等 EMA5 重新站上再买", "BUY_ON_BREAKOUT": "等突破再买",
        "HOLD": "持有不加仓", "NO_BUY": "不买"
    }
    return mapping.get(timing_code, timing_code)


def ai_reduce_ratio(exit_e, timing_code):
    if exit_e == "EXIT": return "100% 清仓"
    if exit_e == "保护利润": return "减 1/3，剩余用 ATR 移动止损"
    return "0% 持有"


def map_ai_to_final_signal(timing_code, exit_e):
    if timing_code in ("BUY_TODAY", "BUY_TOMORROW"): return "买入"
    if timing_code in ("BUY_ON_PULLBACK_EMA20", "BUY_ON_VOLUME", "BUY_ON_EMA5_RECLAIM", "BUY_ON_BREAKOUT"): return "买入观察"
    if timing_code == "HOLD": return "持有"
    if exit_e == "EXIT": return "卖出"
    if exit_e in ("保护利润", "减仓"): return "减仓"
    return "观望"


def generate_ai_summary(row, code, name):
    stage = row["阶段"]
    stock_q = row.get("Stock_Quality_Score", np.nan)
    entry_q = row.get("Entry_Quality_Score", np.nan)
    close = row["Close"]
    timing_code, priority, action_level, reason = ai_timing_engine(row)

    return (
        f"{name} ({code}) 当前属于 {stage}。"
        f"收盘价 {close:.4f}，股票质量分 {stock_q:.1f}/100，买点质量分 {entry_q:.1f}/100。"
        f"【AI 时机建议】：{ai_timing_text(timing_code)}（{action_level}）。原因：{reason}。"
    )


def print_terminal_report(row, code, name):
    print("\n" + "=" * 78)
    print(f"📊 EMA Trading System V4.0.2a — {name} ({code})")
    print("=" * 78)
    print(f"收盘价: {fmt(row['Close'])} | 阶段: {row['阶段']} | 最终信号: {row['最终信号']}")
    print(f"股票质量分: {fmt(row.get('Stock_Quality_Score'))} | 买点质量分: {fmt(row.get('Entry_Quality_Score'))}")
    print("-" * 78)
    print("AI 摘要：")
    print(generate_ai_summary(row, code, name))
    print("=" * 78)


def calculate_trading_priority(row, timing_code=None):
    if timing_code is None: timing_code = ai_timing_engine(row)[0]
    if timing_code == "BUY_TODAY": return "★★★★★"
    if timing_code == "BUY_TOMORROW": return "★★★★☆"
    if timing_code in ("BUY_ON_VOLUME", "BUY_ON_EMA5_RECLAIM", "BUY_ON_PULLBACK_EMA20", "BUY_ON_BREAKOUT"): return "★★★☆☆"
    if timing_code == "HOLD": return "★★☆☆☆"
    return "★☆☆☆☆"


# ================================================================
# 7. EXCEL FORMATTING & OUTPUT DICT
# ================================================================

def format_excel_file(filename):
    try:
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.load_workbook(filename)
        for ws in wb.worksheets:
            is_dashboard = (ws.title == SHEET_DASHBOARD)
            if not is_dashboard:
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                col_letter = col_cells[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 60)
        wb.save(filename)
    except Exception as e:
        pass


def latest_output_dict(row, name, code):
    timing_code, priority, action_level, reason = ai_timing_engine(row)
    timing_text = ai_timing_text(timing_code)
    buy_zone = ai_buy_zone(row, timing_code)
    position = ai_position_suggestion(row, timing_code, action_level)
    invalidation = ai_invalidation(row, timing_code)
    trading_priority = calculate_trading_priority(row, timing_code)
    exit_e = row.get("Exit_Engine", "持有")
    reduce_ratio = ai_reduce_ratio(exit_e, timing_code)

    stock_q = float(row.get("Stock_Quality_Score", row.get("Final_Score", 0)) or 0)
    entry_q = float(row.get("Entry_Quality_Score", 0) or 0)
    stock_type = row.get("Stock_Type", "mixed")
    type_label = {"pullback": "回调型", "breakout": "突破型", "mixed": "混合型"}.get(stock_type, stock_type)

    return {
        "日期": row.name.strftime("%Y-%m-%d") if hasattr(row.name, "strftime") else str(row.name),
        "名称": name,
        "代码": clean_stock_code(code),
        "股票类型": type_label,
        "交易优先级": trading_priority,

        "股票质量分": round(stock_q, 2),
        "股票质量评级": stock_quality_rating(stock_q),
        "买点质量分": round(entry_q, 2),
        "买点质量评级": row.get("Entry_Quality_Rating", "N/A"),
        "综合建议": combined_buy_rating(stock_q, entry_q, timing_code),

        "Top_Warning_Score": row.get("Top_Warning_Score", 0),
        "Top_Warning_Reasons": row.get("Top_Warning_Reasons", ""),

        "Trading Permission": row["Trading_Permission"],
        "阶段": row["阶段"],
        "最终信号": row["最终信号"],

        "Entry Engine": row.get("Entry_Engine", "N/A"),
        "Exit Engine": row.get("Exit_Engine", "N/A"),

        "Setup_Type": row.get("Setup_Type", "NONE"),
        "Setup_Age": row.get("Setup_Age", 999),
        "Setup_Active": bool(row.get("Setup_Active", False)),

        "AI_结论": ai_timing_text(timing_code),
        "AI_时机": timing_code,
        "AI_动作": action_level,
        "AI_原因": reason,
        "AI_参考买入区": buy_zone,
        "AI_仓位建议": position,
        "结构防守价": round(row["结构防守价"], 4) if pd.notna(row.get("结构防守价")) else np.nan,
        "ATR防守参考": round(row["ATR_Stop_2x"], 4) if pd.notna(row.get("ATR_Stop_2x")) else np.nan,
        "AI_失效条件": invalidation,
        "AI 自然语言摘要": generate_ai_summary(row, code, name),

        # 核心 Master Prompt 字段
        "机构级MasterPrompt": generate_master_prompt_payload(row, name, code),

        "收盘价": round(row["Close"], 4),
        "EMA5": round(row["EMA5"], 4) if pd.notna(row.get("EMA5")) else np.nan,
        "EMA20": round(row["EMA20"], 4) if pd.notna(row["EMA20"]) else np.nan,
        "EMA60": round(row["EMA60"], 4) if pd.notna(row["EMA60"]) else np.nan,
        "ADX": round(row["ADX"], 2) if pd.notna(row["ADX"]) else np.nan,
        "Volume Ratio": round(row["Volume_Ratio"], 2) if pd.notna(row["Volume_Ratio"]) else np.nan,
        "ATR Distance": round(row.get("ATR_Distance", np.nan), 2) if pd.notna(row.get("ATR_Distance")) else np.nan,
    }


# ================================================================
# 8. BUILDERS FOR 9 SHEETS
# ================================================================

def build_dashboard_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame([{"项目": "提示", "值": "无数据"}])
    df = df_results.copy()

    today_buy = int((df["综合建议"].apply(classify_action) == "今日可买").sum())
    tomorrow = int((df["综合建议"].apply(classify_action) == "明日确认").sum())
    setup_active_cnt = int(df["Setup_Active"].astype(bool).sum()) if "Setup_Active" in df.columns else 0

    rows = [
        {"项目": "📊 报告日期", "值": datetime.now().strftime("%Y-%m-%d")},
        {"项目": "📊 扫描总数", "值": len(df)},
        {"项目": "=== 行动卡片 ===", "值": ""},
        {"项目": "🟢 今日可买", "值": today_buy},
        {"项目": "🟢 明日确认", "值": tomorrow},
        {"项目": "=== Setup 记忆（V4.0.2a）===", "值": ""},
        {"项目": "活跃 Setup 数", "值": setup_active_cnt},
    ]
    return pd.DataFrame(rows)


def build_trading_overview_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame()
    df = df_results.copy()
    df.insert(0, "排名", range(1, len(df) + 1))
    return df


def build_watchlist_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame()
    return df_results.copy()


def build_position_management_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame()
    return df_results.copy()


def build_decision_summary_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame()
    return df_results.copy()


def build_ai_analysis_df(df_results):
    if df_results is None or df_results.empty: return pd.DataFrame()
    cols = ["名称", "代码", "股票类型", "综合建议", "股票质量分", "买点质量分", "阶段", "最终信号", "机构级MasterPrompt"]
    return df_results[[c for c in cols if c in df_results.columns]].copy()


# ================================================================
# 9. ENGINE CALCULATION PIPELINE
# ================================================================

def calculate_engine(df):
    if df is None or df.empty: return None
    data = df.copy()

    data["EMA5"] = data["Close"].ewm(span=CONFIG["EMA_FAST"], adjust=False).mean()
    data["EMA20"] = data["Close"].ewm(span=CONFIG["EMA_SHORT"], adjust=False).mean()
    data["EMA60"] = data["Close"].ewm(span=CONFIG["EMA_LONG"], adjust=False).mean()
    data["EMA200"] = data["Close"].ewm(span=CONFIG["EMA_BACKGROUND"], adjust=False).mean()

    data["Gap5"] = data["Close"] - data["EMA5"]
    data["Gap20"] = data["Close"] - data["EMA20"]
    data["Gap60"] = data["Close"] - data["EMA60"]

    data["Gap20_Pct"] = data["Gap20"] / data["Close"] * 100
    data["Gap60_Pct"] = data["Gap60"] / data["Close"] * 100

    data["EMA20_Actual_Slope_5"] = data["EMA20"].diff(5)
    data["EMA60_Actual_Slope_5"] = data["EMA60"].diff(5)
    data["Extension_EMA20_Pct"] = (data["Close"] - data["EMA20"]) / data["EMA20"] * 100

    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=5).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)

    tr = pd.concat([data["High"] - data["Low"], (data["High"] - data["Close"].shift(1)).abs(), (data["Low"] - data["Close"].shift(1)).abs()], axis=1).max(axis=1)
    data["ATR14"] = wilder_rma(tr, 14)
    data["ATR_Pct"] = data["ATR14"] / data["Close"] * 100
    data["ATR_Distance"] = (data["Close"] - data["EMA20"]) / data["ATR14"].replace(0, np.nan)

    plus_di, minus_di, dx, adx = compute_adx_wilder(data, 14)
    data["+DI"], data["-DI"], data["DX"], data["ADX"] = plus_di, minus_di, dx, adx
    data["ADX_Change_5"] = data["ADX"].diff(5)
    data["DI_Spread"] = data["+DI"] - data["-DI"]

    direction = np.sign(data["Close"].diff()).fillna(0)
    data["OBV"] = (direction * data["Volume"]).cumsum()
    data["OBV_MA5"] = data["OBV"].rolling(5, min_periods=5).mean()
    data["OBV_Change_5"] = data["OBV"].diff(5)
    data["OBV_MA_Distance_Pct"] = ((data["OBV"] - data["OBV_MA5"]) / data["OBV_MA5"].abs().replace(0, np.nan) * 100)

    data["Candle_Range"] = (data["High"] - data["Low"]).replace(0, np.nan)
    data["Candle_Body"] = (data["Close"] - data["Open"]).abs()
    data["Candle_Body_Pct"] = data["Candle_Body"] / data["Candle_Range"] * 100
    data["Close_Position_Range"] = (data["Close"] - data["Low"]) / data["Candle_Range"]
    data["Return_5"] = data["Close"].pct_change(5) * 100

    data["EMA5_Cross_Age"] = _ema5_cross_age(data["EMA5"], data["EMA20"])
    data["Breakout_Age"] = _breakout_age(data["Close"], data["High"], lookback=20)

    stock_types = [identify_stock_type(data.iloc[:i + 1], CONFIG["STOCK_TYPE_LOOKBACK"]) for i in range(len(data))]
    data["Stock_Type"] = stock_types

    data["Trading_Permission"] = "GREEN"

    def identify_stage(row):
        c, e20, e60 = row["Close"], row["EMA20"], row["EMA60"]
        if pd.isna(e20) or pd.isna(e60): return "Stage 0 (数据不足)"
        if c > e60 and e20 > e60: return "Stage 2 (主升段)"
        if c >= e20: return "Stage 1 (初升段)"
        return "Stage 3 (末升段)"
    data["阶段"] = data.apply(identify_stage, axis=1)

    data["Prev_High_20"] = data["High"].rolling(20, min_periods=5).max().shift(1)
    data = compute_setup_memory_series(data)

    entry_list, exit_list = [], []
    for i in range(len(data)):
        r, prev = data.iloc[i], (data.iloc[i - 1] if i > 0 else None)
        entry_list.append(entry_engine_v2(r, prev))
        exit_list.append(exit_engine_v2(r, prev))
    data["Entry_Engine"] = entry_list
    data["Exit_Engine"] = exit_list

    data["最终信号"] = "买入"
    data["ATR_Stop_2x"] = data["Close"] - 2 * data["ATR14"]
    data["Recent_Low_20"] = data["Low"].rolling(20, min_periods=5).min()
    data["结构防守价"] = data[["Recent_Low_20", "ATR_Stop_2x"]].max(axis=1)

    stock_q_rows = [calculate_stock_quality_score(data.iloc[:i + 1]) for i in range(len(data))]
    entry_q_rows = [calculate_entry_quality_score(data.iloc[i]) for i in range(len(data))]

    stock_q_df = pd.DataFrame(stock_q_rows, index=data.index)
    entry_q_df = pd.DataFrame(entry_q_rows, index=data.index)

    for c in stock_q_df.columns: data[c] = stock_q_df[c]
    for c in entry_q_df.columns: data[c] = entry_q_df[c]

    top_scores, top_reasons_list = [], []
    for i in range(len(data)):
        s, r = calculate_top_warning_score(data.iloc[i], data.iloc[max(0, i - 5):i])
        top_scores.append(s); top_reasons_list.append("、".join(r))
    data["Top_Warning_Score"] = top_scores
    data["Top_Warning_Reasons"] = top_reasons_list

    final_signals_unified = [map_ai_to_final_signal(ai_timing_engine(data.iloc[i])[0], data.iloc[i].get("Exit_Engine", "持有")) for i in range(len(data))]
    data["最终信号"] = final_signals_unified

    return data


# ================================================================
# 10. SIX MAIN FUNCTIONS
# ================================================================

# --- 菜单 1：单股分析 ---
def single_stock_scan(code, output_excel=True):
    print(f"\n🔍 正在扫描：{code}")
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None:
        print("❌ 找不到股票数据或获取失败。")
        return None

    data = calculate_engine(df_raw)
    if data is None or data.empty: return None

    row = data.iloc[-1]
    print_terminal_report(row, code, name)
    result = latest_output_dict(row, name, code)

    if output_excel:
        filename = f"{safe_filename(name)}_EMA_V4.0.2a_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df_summary = pd.DataFrame([result])

        dashboard_df = build_dashboard_df(df_summary)
        overview_df = build_trading_overview_df(df_summary)
        watchlist_df = build_watchlist_df(df_summary)
        position_df = build_position_management_df(df_summary)
        decision_df = build_decision_summary_df(df_summary)
        ai_df = build_ai_analysis_df(df_summary)

        tail_output = data.tail(60).copy().reset_index()
        signal_log = tail_output.copy()
        all_indicators = data.copy().reset_index()

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            dashboard_df.to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
            overview_df.to_excel(writer, sheet_name=SHEET_OVERVIEW, index=False)
            watchlist_df.to_excel(writer, sheet_name=SHEET_WATCHLIST, index=False)
            position_df.to_excel(writer, sheet_name=SHEET_POSITION, index=False)
            decision_df.to_excel(writer, sheet_name=SHEET_DECISION, index=False)
            ai_df.to_excel(writer, sheet_name=SHEET_AI, index=False)
            tail_output.to_excel(writer, sheet_name=SHEET_RECENT, index=False)
            signal_log.to_excel(writer, sheet_name=SHEET_TURNS, index=False)
            all_indicators.to_excel(writer, sheet_name=SHEET_ALL, index=False)

        format_excel_file(filename)
        print(f"\n✅ Excel 已生成（包含 9 个完整 Sheet + Master Prompt）：{filename}")

    return result


# --- 菜单 2 & 3：批量扫描与 Excel 股票池扫描 ---
def batch_scan(stock_pool, output_excel=True, min_price=0.05, max_price=0.50, min_volume=100000):
    print(f"\n📊 批量扫描：共 {len(stock_pool)} 只")
    results, errors = [], []

    for idx, item in enumerate(stock_pool, 1):
        code = item["code"]
        print(f"[{idx}/{len(stock_pool)}] {code} ...", end=" ")

        try:
            df_raw, name = safe_fetch_stock(code)
            if df_raw is None:
                errors.append(code); print("无数据")
                continue

            data = calculate_engine(df_raw)
            if data is None or data.empty:
                errors.append(code); print("计算失败")
                continue

            row = data.iloc[-1]
            if pd.isna(row["Close"]) or row["Close"] < min_price or row["Close"] > max_price:
                print("跳过(价格范围外)")
                continue

            result = latest_output_dict(row, name, code)
            results.append(result)
            print("成功")

        except Exception as e:
            errors.append(code)
            print(f"错误：{e}")

    if not results:
        print("\n⚠️ 没有符合条件的股票。")
        return None

    df_results = pd.DataFrame(results)

    if output_excel:
        filename = f"EMA_V4.0.2a_批量扫描_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        dashboard_df = build_dashboard_df(df_results)
        overview_df = build_trading_overview_df(df_results)
        watchlist_df = build_watchlist_df(df_results)
        position_df = build_position_management_df(df_results)
        decision_df = build_decision_summary_df(df_results)
        ai_df = build_ai_analysis_df(df_results)

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            dashboard_df.to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
            overview_df.to_excel(writer, sheet_name=SHEET_OVERVIEW, index=False)
            watchlist_df.to_excel(writer, sheet_name=SHEET_WATCHLIST, index=False)
            position_df.to_excel(writer, sheet_name=SHEET_POSITION, index=False)
            decision_df.to_excel(writer, sheet_name=SHEET_DECISION, index=False)
            ai_df.to_excel(writer, sheet_name=SHEET_AI, index=False)
            df_results.to_excel(writer, sheet_name=SHEET_FULL_SCAN, index=False)
            if errors:
                pd.DataFrame({"错误代码": errors}).to_excel(writer, sheet_name=SHEET_ERRORS, index=False)

        format_excel_file(filename)
        print(f"\n✅ 批量报告已生成（含 Master Prompt）：{filename}")

    return df_results


def load_stock_pool_from_excel(filepath):
    try:
        df = pd.read_excel(filepath)
        code_col = next((c for c in ["Code", "code", "代码", "Stock", "股票代码"] if c in df.columns), df.columns[0])
        pool = [{"code": clean_stock_code(c), "name": ""} for c in df[code_col].dropna()]
        print(f"✅ 成功从 Excel 加载 {len(pool)} 只股票")
        return pool
    except Exception as e:
        print(f"❌ 读取 Excel 股票池失败：{e}")
        return None


# --- 菜单 4：单股历史回溯 ---
def single_stock_backtrack(code, days=60, output_excel=True):
    print(f"\n⏪ 正在回溯 {code} 最近 {days} 个交易日...")
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None: return None

    data = calculate_engine(df_raw)
    if data is None: return None

    back = data.tail(days).copy()
    records = [latest_output_dict(row, name, code) for _, row in back.iterrows()]
    df_back = pd.DataFrame(records)

    filename = f"{safe_filename(name)}_EMA_V4.0.2a_回溯{days}天_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_back.to_excel(writer, sheet_name="逐日日志", index=False)
    format_excel_file(filename)
    print(f"\n✅ 回溯报告已导出：{filename}")
    return df_back


# --- 菜单 5：ADX 对照测试 ---
def compare_adx_recent(code, rows=30):
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None: return None
    data = calculate_engine(df_raw)
    cols = ["Close", "High", "Low", "+DI", "-DI", "DI_Spread", "DX", "ADX"]
    res = data[cols].tail(rows).copy()
    print("\n" + "=" * 80)
    print(f"ADX 对照表 (Wilder RMA): {name} ({code})")
    print("=" * 80)
    print(res.round(4).to_string())
    return res


# ================================================================
# 11. MAIN MENU (完美恢复交互界面)
# ================================================================

def print_menu():
    print("\n" + "=" * 70)
    print("   🇲🇾 马股 EMA Trading System V4.0.2a (集成 Master Prompt)")
    print("=" * 70)
    print("1. 单股分析（输出 9 个 Sheet 双层报告 + Master Prompt）")
    print("2. 批量扫描（默认价格 RM0.05–RM0.50）")
    print("3. Excel 股票池批量扫描 (自动读取 Final list.xlsx)")
    print("4. 单股历史回溯（含顶部预警 + 持仓跟踪）")
    print("5. ADX 对照测试（与 TradingView 对齐）")
    print("6. 退出")
    print("-" * 70)


def main():
    while True:
        print_menu()
        choice = input("请输入选项 (1-6): ").strip()

        if choice == "1":
            code = input("输入股票代码，例如 10 / 0010 / 10.KL：").strip()
            if code: single_stock_scan(code)

        elif choice == "2":
            min_p = input("最低价格 RM [0.05]：").strip()
            max_p = input("最高价格 RM [0.50]：").strip()
            min_p = float(min_p) if min_p else 0.05
            max_p = float(max_p) if max_p else 0.50

            codes_text = input("输入股票代码（逗号分隔，如 10,38,85）：").strip()
            if codes_text:
                codes = [{"code": clean_stock_code(x), "name": ""} for x in codes_text.split(",") if x.strip()]
                batch_scan(codes, min_price=min_p, max_price=max_p)

        elif choice == "3":
            filepath = input("输入 Excel 文件名 [默认: Final list.xlsx]：").strip()
            if not filepath: filepath = "Final list.xlsx"
            pool = load_stock_pool_from_excel(filepath)
            if pool: batch_scan(pool)

        elif choice == "4":
            code = input("输入股票代码：").strip()
            days = input("回溯多少交易日 [60]：").strip()
            days = int(days) if days.isdigit() else 60
            if code: single_stock_backtrack(code, days)

        elif choice == "5":
            code = input("输入股票代码：").strip()
            if code: compare_adx_recent(code)

        elif choice == "6":
            print("系统已安全退出，祝交易顺利！")
            break
        else:
            print("⚠️ 无效选项，请重新输入。")


if __name__ == "__main__":
    main()
