# -*- coding: utf-8 -*-
"""
====================================================================
马来西亚低价股 EMA Trading System V4.2
「EMA5 + EMA20 + EMA60 + EMA200 + EMA扣底 + ADX(+DI/-DI) + ATR + OBV
 + Stock Quality + Entry Quality + Entry/Exit V2 + AI Timing」
====================================================================

V4.2 = V4.0.2（B 主干）+ V4.1（A）移植：
  ★ 主干保留：Setup+Trigger+Memory、真实 Trading Permission(GREEN/YELLOW/RED)、
              完整顶部预警(7 项, 满分 105)、Stage 3 持有逻辑(转弱条件计数)、
              再入场、强买点覆盖、仓位建议、持仓跟踪回测
  ★ 新增 EMA200 长趋势过滤：Close<=EMA200 时 Permission 由 GREEN 降为 YELLOW，
              Stage 2 标注「未确认长期」
  ★ 新增 Stage 3A（末升-健康调整）/ 3B（末升-趋势恶化）拆分；3B 只需 1 个转弱条件即减仓
  ★ 新增动态 ATR 倍数（按阶段）→ ATR 动态防守价 / 结构防守价 / 移动止损
  ★ 新增 Entry 确认层（5 因子）与 AI 信心分（低信心 BUY → 次日确认）
  ★ 新增完整 Master Prompt：含 +DI/-DI、ADX、OBV、价量变化、背离判断、防守/压力位
  ★ 新增 Streamlit 界面：streamlit run EMA_trading_system_V4_2.py

依赖：
pip install yfinance pandas numpy openpyxl streamlit(可选)
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
        "Stage 3A": -3.0,
        "Stage 3B": -10.0,
        "Stage 4": -18.0,
        "Stage 0": 0.0,
    },
    # ============ V4.2 新增 ============
    "LONG_TREND_FILTER_ENABLED": True,      # Close<=EMA200 → Permission 降为 YELLOW
    "STAGE3B_REDUCE_MIN_CONDITIONS": 1,     # Stage 3B 只需 1 个转弱条件即减仓
    "ATR_MULTIPLIER_BY_STAGE": {
        "Stage 1": 2.5, "Stage 2": 3.0,
        "Stage 3": 2.0, "Stage 3A": 2.0, "Stage 3B": 1.5,
        "Stage 4": 1.0, "Stage 0": 2.0,
    },
    "USE_DYNAMIC_ATR_TRAIL": True,          # 移动止损使用按阶段的 ATR 倍数
    "ENTRY_CONFIRM_ENABLED": True,
    "ENTRY_CONFIRM_MIN_FACTORS": 3,         # V4.1 原值为 4/3，过严会砍掉约 70% 买入信号
    "ENTRY_CONFIRM_SOFT_FACTORS": 2,
    "AI_CONFIDENCE_MIN_BUY": 60,
    "AI_CONFIDENCE_STRONG": 80,

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

    # 报告优化用
    "WATCH_VOL_TRIGGER": 1.50,
    "WATCH_VOL_NEAR": 1.20,
    "WATCH_EMA20_NEAR_PCT": 0.02,
    "WATCH_EMA20_FAR_PCT": 0.05,

    # ============ V4.2 Setup + Trigger 记忆机制 ============
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
    """从 'RM0.3724 – RM0.3876' 或 'RM0.1629 – RM0.1696（EMA20 附近）' 提取范围"""
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
    """将综合建议归类为 6 大行动类别"""
    s = str(comprehensive_advice or "")
    if "今天可买" in s:
        return "今日可买"
    if "明日确认" in s:
        return "明日确认"
    if "等放量" in s:
        return "等放量"
    if "等回" in s or "回 EMA20" in s:
        return "等回EMA20"
    if "等 EMA5" in s:
        return "等EMA5"
    if "等突破" in s:
        return "等突破"
    if "持有不加仓" in s:
        return "持有不加仓"
    if "不买" in s:
        return "不买"
    return "其他"


# ================================================================
# 3. DATA LOADER
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
            print(f"⚠️ {ticker} 无数据，可能已退市或代码错误")
            return None, None

        if isinstance(df.columns, pd.MultiIndex):
            try:
                df.columns = df.columns.get_level_values(0)
            except Exception:
                df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]

        needed = ["Open", "High", "Low", "Close", "Volume"]
        missing = [c for c in needed if c not in df.columns]
        if missing:
            print(f"❌ {ticker} 缺少栏位: {missing}")
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
            print(f"⚠️ {ticker} 有效交易日不足: {len(df)}")
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
        print(f"❌ {ticker} 抓取失败: {e}")
        return None, None


# ================================================================
# 4. WILDER RMA / ADX
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

    plus_dm.loc[valid] = np.where(
        (up_move.loc[valid] > down_move.loc[valid]) &
        (up_move.loc[valid] > 0),
        up_move.loc[valid],
        0.0
    )

    minus_dm.loc[valid] = np.where(
        (down_move.loc[valid] > up_move.loc[valid]) &
        (down_move.loc[valid] > 0),
        down_move.loc[valid],
        0.0
    )

    tr = pd.concat(
        [
            high - low,
            (high - close.shift(1)).abs(),
            (low - close.shift(1)).abs(),
        ],
        axis=1
    ).max(axis=1)

    tr.iloc[0] = np.nan

    tr_rma = wilder_rma(tr, period)
    plus_dm_rma = wilder_rma(plus_dm, period)
    minus_dm_rma = wilder_rma(minus_dm, period)

    plus_di = 100.0 * plus_dm_rma / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * minus_dm_rma / tr_rma.replace(0, np.nan)

    denominator = (plus_di + minus_di).replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / denominator

    adx = wilder_rma(dx, period)

    return plus_di, minus_di, dx, adx


# ================================================================
# 5. TREND / GAP / DEDUCT UTILITIES
# ================================================================

def classify_gap(gap, price, epsilon_pct=0.15):
    if pd.isna(gap) or pd.isna(price) or price == 0:
        return "数据不足"
    threshold = abs(price) * epsilon_pct / 100.0
    if gap > threshold:
        return "上扬"
    elif gap < -threshold:
        return "下弯"
    return "走平"


def classify_distance(distance, distance_change, price):
    if pd.isna(distance) or pd.isna(price) or price == 0:
        return "数据不足"
    threshold = abs(price) * CONFIG["GAP_EPSILON_PCT"] / 100.0

    if distance > threshold:
        if pd.notna(distance_change) and distance_change > 0:
            return "价格上偏离扩大"
        elif pd.notna(distance_change) and distance_change < 0:
            return "价格上偏离收窄"
        return "价格上偏离"
    if distance < -threshold:
        if pd.notna(distance_change) and distance_change < 0:
            return "价格下偏离扩大"
        elif pd.notna(distance_change) and distance_change > 0:
            return "价格下偏离收窄"
        return "价格下偏离"

    if pd.notna(distance_change):
        if distance_change > 0:
            return "接近上穿"
        elif distance_change < 0:
            return "接近下穿"
    return "贴合均线"


def classify_deduct(deduct, deduct_change, price):
    if pd.isna(deduct) or pd.isna(price) or price == 0:
        return "数据不足"
    threshold = abs(price) * CONFIG["GAP_EPSILON_PCT"] / 100.0

    if deduct > threshold:
        if pd.notna(deduct_change) and deduct_change > 0:
            return "正向扣底扩大(动能增强)"
        elif pd.notna(deduct_change) and deduct_change < 0:
            return "正向扣底收窄(动能放缓)"
        return "正向扣底"

    if deduct < -threshold:
        if pd.notna(deduct_change) and deduct_change < 0:
            return "负向扣底扩大(下行增强)"
        elif pd.notna(deduct_change) and deduct_change > 0:
            return "负向扣底收窄(下行放缓)"
        return "负向扣底"

    if pd.notna(deduct_change):
        if deduct_change > 0:
            return "扣底转正"
        elif deduct_change < 0:
            return "扣底转负"
    return "扣底中性"


def arrow_trend(series, max_days=5):
    s = pd.Series(series).dropna().tail(max_days)
    if len(s) < 2:
        return "N/A"
    arrows = []
    for i in range(1, len(s)):
        if s.iloc[i] > s.iloc[i - 1]:
            arrows.append("↑")
        elif s.iloc[i] < s.iloc[i - 1]:
            arrows.append("↓")
        else:
            arrows.append("→")
    return "".join(arrows)


# ================================================================
# 6. STOCK QUALITY SCORE
# ================================================================

def clip_score(x, low=0, high=100):
    return float(np.clip(x, low, high))


def get_stage_key(stage):
    m = re.search(r"Stage\s*(3A|3B|\d)", str(stage), re.IGNORECASE)
    return f"Stage {m.group(1).upper()}" if m else "Stage 0"


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
        if distance20 >= 0:
            price_score += np.clip(12 + distance20 * 0.35, 0, 17)
        else:
            price_score += np.clip(12 + distance20 * 0.50, 0, 12)
    if pd.notna(ema60) and ema60 > 0:
        distance60 = (close - ema60) / ema60 * 100
        price_score += np.clip(5 + distance60 * 0.18, 0, 8)
    price_score = clip_score(price_score, 0, 25)

    money_score = 0.0
    obv_change = row.get("OBV_Change_5", np.nan)
    obv_ma_distance = row.get("OBV_MA_Distance_Pct", np.nan)
    avg_volume = row.get("Volume_MA20", np.nan)

    if pd.notna(obv_change):
        if pd.notna(avg_volume) and avg_volume > 0:
            obv_change_norm = float(obv_change) / (float(avg_volume) * 5.0)
            money_score += np.clip(5.0 + obv_change_norm * 5.0, 0, 10)
        else:
            money_score += 5.0

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
    if pd.notna(gap20_pct):
        trend_score += np.clip(5 + gap20_pct * 0.8, 0, 8)
    if pd.notna(gap60_pct):
        trend_score += np.clip(5 + gap60_pct * 0.4, 0, 7)
    if pd.notna(adx):
        trend_score += np.clip((adx - 15) * 0.20, 0, 4)
    if pd.notna(adx_change):
        trend_score += np.clip(adx_change * 0.10, 0, 1)
    if pd.notna(di_spread):
        trend_score += np.clip(di_spread * 0.03, 0, 1)
    trend_score = clip_score(trend_score, 0, 20)

    candle_score = 0.0
    body = row["Candle_Body_Pct"]
    close_position = row["Close_Position_Range"]
    return_5 = row["Return_5"]
    upper_shadow_pct = row.get("Upper_Shadow_Pct", np.nan)
    lower_shadow_pct = row.get("Lower_Shadow_Pct", np.nan)

    if pd.notna(body):
        candle_score += np.clip(7 + body * 0.15, 0, 10)
    if pd.notna(close_position):
        candle_score += np.clip(close_position * 8, 0, 8)
    if pd.notna(return_5):
        candle_score += np.clip(2 + return_5 * 0.15, 0, 5)

    if pd.notna(upper_shadow_pct) and upper_shadow_pct > 40:
        candle_score -= float(np.clip((upper_shadow_pct - 40) * 0.05, 0, 2))
    if pd.notna(lower_shadow_pct) and lower_shadow_pct > 40:
        candle_score += float(np.clip((lower_shadow_pct - 40) * 0.03, 0, 1))

    candle_score = clip_score(candle_score, 0, 20)

    atr_distance = row.get("ATR_Distance", np.nan)
    atr_pct = row.get("ATR_Pct", np.nan)
    extension = row.get("Extension_EMA20_Pct", np.nan)
    permission = row.get("Trading_Permission", "YELLOW")
    atr_health_score = float(CONFIG["ATR_HEALTH_SCORE_MAX"])

    if pd.notna(atr_distance):
        ad = abs(float(atr_distance))
        if ad <= CONFIG["ATR_DISTANCE_HEALTHY"]:
            atr_health_score -= ad * 0.5
        elif ad <= CONFIG["ATR_DISTANCE_CAUTION"]:
            atr_health_score -= 1.0 + (ad - 2.0) * 1.5
        elif ad <= CONFIG["ATR_DISTANCE_HIGH"]:
            atr_health_score -= 4.0 + (ad - 4.0) * 1.5
        else:
            atr_health_score -= 7.0 + min((ad - 6.0) * 1.0, 5.0)

    if pd.notna(atr_pct):
        atr_health_score -= np.clip((float(atr_pct) - 4.0) * 0.35, 0, 3)
    if pd.notna(extension):
        if extension > CONFIG["EXTENSION_MAX_PCT"]:
            atr_health_score -= 3
        elif extension > CONFIG["EXTENSION_CAUTION_PCT"]:
            atr_health_score -= 1
    if permission == "YELLOW":
        atr_health_score -= 1
    elif permission == "RED":
        atr_health_score -= 5
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


# ================================================================
# 6b. ENTRY QUALITY SCORE
# ================================================================

def _ema5_cross_age(ema5, ema20):
    above = (ema5 > ema20)
    cross_up = above & (~above.shift(1).fillna(False))
    ages = []
    age = 999
    for c in cross_up:
        if c:
            age = 0
        else:
            age = min(age + 1, 999)
        ages.append(age)
    return pd.Series(ages, index=ema5.index)


def _breakout_age(close, high, lookback=20):
    prev_high = high.rolling(lookback, min_periods=lookback).max().shift(1)
    tol = CONFIG["BREAKOUT_VALID_TOL_PCT"] / 100.0
    breakout_level = prev_high * (1.0 - tol)

    ages = []
    age = 999

    for close_val, level in zip(close, breakout_level):
        if pd.isna(close_val) or pd.isna(level):
            age = 999
        elif close_val >= level:
            if age == 999:
                age = 0
            else:
                age += 1
        else:
            age = 999
        ages.append(age)

    return pd.Series(ages, index=close.index)


def identify_stock_type(data, lookback=20):
    if len(data) < lookback + 5:
        return "mixed"

    recent = data.tail(lookback)

    if "ATR14" not in recent.columns or "EMA20" not in recent.columns:
        return "mixed"

    dist_atr = (recent["Close"] - recent["EMA20"]) / recent["ATR14"].replace(0, np.nan)
    dist_atr = dist_atr.dropna()

    if len(dist_atr) < 5:
        return "mixed"

    far_days = int((dist_atr.abs() > CONFIG["STOCK_TYPE_FAR_ATR_THRESHOLD"]).sum())

    if far_days >= CONFIG["STOCK_TYPE_FAR_DAYS_HIGH"]:
        return "breakout"
    elif far_days <= CONFIG["STOCK_TYPE_FAR_DAYS_LOW"]:
        return "pullback"
    else:
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
        if dist_pct < -3:
            ema_score = 8
        elif dist_pct < 0:
            ema_score = 15
        elif dist_pct <= 3:
            ema_score = 30
        elif dist_pct <= 6:
            ema_score = 22
        elif dist_pct <= 10:
            ema_score = 10
        else:
            ema_score = -10

    cross_score = 0
    if pd.notna(cross_age):
        if cross_age <= 5:
            cross_score = 20
        elif cross_age <= 10:
            cross_score = 12
        elif cross_age <= 15:
            cross_score = 5
        else:
            cross_score = 0

    atr_score = 0
    if pd.notna(atr_distance):
        ad = abs(float(atr_distance))
        if ad < 2:
            atr_score = 25
        elif ad < 4:
            atr_score = 10
        elif ad < 6:
            atr_score = -10
        else:
            atr_score = -20

    brk_score = 0
    if pd.notna(breakout_age) and breakout_age != 999:
        if breakout_age <= 5:
            brk_score = 15
        elif breakout_age <= 10:
            brk_score = 8
        else:
            brk_score = 0

    vol_score = 0
    if pd.notna(volume_ratio):
        if volume_ratio >= 1.5:
            vol_score = 10
        elif volume_ratio >= 1.2:
            vol_score = 7
        elif volume_ratio >= 1.0:
            vol_score = 4

    total = float(np.clip(ema_score + cross_score + atr_score + brk_score + vol_score, 0, 100))

    if total >= 80:
        rating = "优秀"
    elif total >= 65:
        rating = "良好"
    elif total >= 50:
        rating = "普通"
    elif total >= 35:
        rating = "偏差"
    else:
        rating = "差(追高)"

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
    if pd.isna(score):
        return "N/A"
    s = float(score)
    if s >= 85:
        return "优秀"
    if s >= 70:
        return "良好"
    if s >= 55:
        return "普通"
    if s >= 40:
        return "偏差"
    return "差"


def combined_buy_rating(stock_q, entry_q, timing_code):
    if timing_code == "NO_BUY":
        return "🚫 不买"
    if timing_code == "HOLD":
        return "🟡 持有不加仓"
    if timing_code == "BUY_TODAY":
        return "🟢 今天可买"
    if timing_code == "BUY_TOMORROW":
        return "🟢 明日确认后买"
    if timing_code == "BUY_ON_PULLBACK_EMA20":
        return "🔵 等回 EMA20"
    if timing_code == "BUY_ON_VOLUME":
        return "🔵 等放量"
    if timing_code == "BUY_ON_EMA5_RECLAIM":
        return "🔵 等 EMA5 站上"
    if timing_code == "BUY_ON_BREAKOUT":
        return "🔵 等突破"
    return "⚪ 观察"


# ================================================================
# 6c-0. V4.2 新增：ATR 分级 / 防追高 / Setup 记忆
# ================================================================

def classify_atr_level(atr_distance):
    """ATR Distance 三级：EXCELLENT / ACCEPTABLE / TOO_FAR"""
    if pd.isna(atr_distance):
        return "UNKNOWN"
    ad = abs(float(atr_distance))
    if ad <= CONFIG["ATR_LEVEL_EXCELLENT"]:
        return "EXCELLENT"
    if ad <= CONFIG["ATR_LEVEL_ACCEPTABLE"]:
        return "ACCEPTABLE"
    return "TOO_FAR"


def check_overextended(row):
    """防止追高：ATR 与 Extension 同时过高 => overextended"""
    atr_distance = row.get("ATR_Distance", np.nan)
    extension = row.get("Extension_EMA20_Pct", np.nan)
    ad = abs(float(atr_distance)) if pd.notna(atr_distance) else 0.0
    ext = abs(float(extension)) if pd.notna(extension) else 0.0

    over = (ad >= CONFIG["OVEREXTEND_ATR_MAX"]
            and ext >= CONFIG["OVEREXTEND_EXTENSION_MAX"])
    return over, ad, ext


def compute_setup_memory_series(data):
    """
    V4.2：为每一行生成 Setup 记忆。

    输出列：
      Setup_Type          : NONE / EMA20_PULLBACK / BREAKOUT
      Setup_Age           : 0,1,2 ...（超过有效期后失效）
      Setup_Active        : True/False
      Setup_Breakout_Level: 突破型形态的突破位
      Breakout_Confirmed  : 突破型形态是否已被量能确认
    """
    n = len(data)
    setup_type_arr = ["NONE"] * n
    setup_age_arr = [999] * n
    setup_active_arr = [False] * n
    setup_brk_level_arr = [np.nan] * n
    brk_confirmed_arr = [False] * n

    valid_days = CONFIG["SETUP_VALID_DAYS"]

    prev_type = "NONE"
    prev_age = 999
    prev_level = np.nan
    prev_confirmed = False

    for i in range(n):
        r = data.iloc[i]
        stage_key = get_stage_key(r.get("阶段", "Stage 0"))
        stage_ok = stage_key in ("Stage 1", "Stage 2")

        close = r.get("Close", np.nan)
        low = r.get("Low", np.nan)
        ema5 = r.get("EMA5", np.nan)
        ema20 = r.get("EMA20", np.nan)
        atr_distance = r.get("ATR_Distance", np.nan)
        volume_ratio = r.get("Volume_Ratio", np.nan)

        cur_type = "NONE"
        cur_age = 999
        cur_level = np.nan
        cur_confirmed = False

        # ---------- (A) 尝试延续旧 Setup ----------
        if prev_type != "NONE" and prev_age < valid_days:
            if prev_type == "EMA20_PULLBACK":
                # 只要 Close 没有明显跌破 EMA20（> -3%），Setup 有效
                if (pd.notna(close) and pd.notna(ema20)
                        and close > ema20 * 0.97):
                    cur_type = prev_type
                    cur_age = prev_age + 1
            elif prev_type == "BREAKOUT":
                if pd.notna(close) and pd.notna(prev_level):
                    if close >= prev_level * 0.98:
                        cur_type = prev_type
                        cur_age = prev_age + 1
                        cur_level = prev_level
                        cur_confirmed = prev_confirmed
                # 跌破突破位 -> 形态失效，不延续

        # ---------- (B) 尝试建立 EMA20 回调 Setup ----------
        if cur_type == "NONE" and stage_ok:
            touched_ema20 = False
            if pd.notna(low) and pd.notna(ema20):
                tol = ema20 * (1 + CONFIG["ENTRY_PULLBACK_EMA20_TOL_PCT"] / 100.0)
                if low <= tol:
                    touched_ema20 = True
            if not touched_ema20 and pd.notna(close) and pd.notna(ema20):
                tol = ema20 * (1 + CONFIG["ENTRY_PULLBACK_CLOSE_TOL_PCT"] / 100.0)
                if close <= tol:
                    touched_ema20 = True
            if not touched_ema20 and pd.notna(atr_distance):
                if abs(float(atr_distance)) <= CONFIG["PULLBACK_ATR_NEAR_MAX"]:
                    touched_ema20 = True
            if not touched_ema20 and pd.notna(low) and pd.notna(ema5):
                tol = ema5 * (1 + CONFIG["PULLBACK_EMA5_TOL_PCT"] / 100.0)
                if low <= tol:
                    touched_ema20 = True

            if touched_ema20:
                cur_type = "EMA20_PULLBACK"
                cur_age = 0

        # ---------- (C) 尝试建立 Breakout Setup ----------
        if cur_type == "NONE" and stage_ok:
            prev_high = r.get("Prev_High_20", np.nan)
            if pd.notna(prev_high) and pd.notna(close):
                tol = CONFIG["BREAKOUT_VALID_TOL_PCT"] / 100.0
                breakout_level = prev_high * (1.0 - tol)
                if close >= breakout_level:
                    cur_type = "BREAKOUT"
                    cur_age = 0
                    cur_level = breakout_level

        # ---------- (D) 突破量能确认 ----------
        if cur_type == "BREAKOUT":
            if pd.notna(volume_ratio) and volume_ratio >= CONFIG["STAGE1_BREAKOUT_VOL_MIN"]:
                cur_confirmed = True

        # ---------- 写回 ----------
        setup_type_arr[i] = cur_type
        setup_age_arr[i] = cur_age
        setup_active_arr[i] = (cur_type != "NONE" and cur_age <= valid_days)
        setup_brk_level_arr[i] = cur_level
        brk_confirmed_arr[i] = cur_confirmed

        prev_type = cur_type
        prev_age = cur_age
        prev_level = cur_level
        prev_confirmed = cur_confirmed

    data["Setup_Type"] = setup_type_arr
    data["Setup_Age"] = setup_age_arr
    data["Setup_Active"] = setup_active_arr
    data["Setup_Breakout_Level"] = setup_brk_level_arr
    data["Breakout_Confirmed"] = brk_confirmed_arr
    return data


# ================================================================
# 6c. ENTRY / EXIT ENGINE V2（V4.2 改写）
# ================================================================

def entry_engine_v2(row, prev_row):
    """
    V4.2：
      - 不再要求 touched_ema20 + ema5_reclaim + vol + atr 同日同时满足
      - 改为：Setup（EMA20回调 / Breakout）+ 1~3 天记忆 + Trigger
      - 成交量由 Hard Filter → Position Sizing 因子
      - ATR 改为三级：EXCELLENT / ACCEPTABLE / TOO_FAR
      - BUY 前最后做防止追高检查
    """
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    stock_type = row.get("Stock_Type", "mixed")

    close = row.get("Close", np.nan)
    low = row.get("Low", np.nan)
    ema5 = row.get("EMA5", np.nan)
    ema20 = row.get("EMA20", np.nan)
    gap20 = row.get("Gap20", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)
    atr_distance = row.get("ATR_Distance", np.nan)
    cross_age = row.get("EMA5_Cross_Age", np.nan)
    breakout_age = row.get("Breakout_Age", np.nan)
    close_position = row.get("Close_Position_Range", np.nan)

    setup_type = row.get("Setup_Type", "NONE")
    setup_age = row.get("Setup_Age", 999)
    setup_active = bool(row.get("Setup_Active", False))
    breakout_confirmed = bool(row.get("Breakout_Confirmed", False))

    if pd.isna(ema5) or pd.isna(ema20) or pd.isna(close):
        return "NONE"

    atr_level = classify_atr_level(atr_distance)
    overextended, _, _ = check_overextended(row)

    # ---------- Trigger 判定 ----------
    prev_close = prev_row.get("Close", np.nan) if prev_row is not None else np.nan
    prev_ema5 = prev_row.get("EMA5", np.nan) if prev_row is not None else np.nan
    prev_ema20 = prev_row.get("EMA20", np.nan) if prev_row is not None else np.nan

    # EMA5 Reclaim（主要 Trigger）——不再要求 Volume
    ema5_reclaim = False
    if (pd.notna(prev_close) and pd.notna(prev_ema5)
            and prev_close < prev_ema5 and close > ema5):
        ema5_reclaim = True
    if (pd.notna(prev_ema5) and pd.notna(prev_ema20)
            and prev_ema5 <= prev_ema20 and ema5 > ema20):
        ema5_reclaim = True
    if (pd.notna(cross_age)
            and cross_age <= CONFIG["EMA5_RECLAIM_FRESH_DAYS"]
            and ema5 > ema20 and close >= ema5):
        ema5_reclaim = True

    # Strong Close Trigger（次强 Trigger）
    strong_close = (
        pd.notna(close_position)
        and close_position >= CONFIG["STRONG_CLOSE_POS_MIN"]
        and pd.notna(volume_ratio)
        and volume_ratio >= CONFIG["STRONG_CLOSE_VOL_MIN"]
    )

    # ==================================================
    # Stage 1
    # ==================================================
    if stage == "Stage 1":
        above_ema20 = close >= ema20
        gap20_ok = pd.notna(gap20) and gap20 >= 0
        ema5_above_ema20 = ema5 > ema20

        if stock_type == "pullback":
            breakout_age_max = CONFIG["STAGE1_BREAKOUT_AGE_MAX"]
            breakout_vol_min = CONFIG["STAGE1_BREAKOUT_VOL_MIN"]
        elif stock_type == "breakout":
            breakout_age_max = CONFIG["BREAKOUT_AGE_MAX_EXTENDED"]
            breakout_vol_min = CONFIG["BREAKOUT_VOL_MIN_EXTENDED"]
        else:
            breakout_age_max = CONFIG["STAGE1_BREAKOUT_AGE_MAX"]
            breakout_vol_min = CONFIG["STAGE1_BREAKOUT_VOL_MIN"]

        breakout_ok = (pd.notna(breakout_age) and breakout_age != 999
                       and breakout_age <= breakout_age_max)
        volume_strong = (pd.notna(volume_ratio) and volume_ratio >= breakout_vol_min)
        atr_soft_ok = (pd.notna(atr_distance)
                       and abs(float(atr_distance)) < CONFIG["STAGE1_BREAKOUT_ATR_MAX"])
        atr_hard_ok = (pd.notna(atr_distance)
                       and abs(float(atr_distance)) <= CONFIG["STAGE1_BREAKOUT_ATR_HARD_MAX"])

        # 原始强硬突破逻辑（保留）
        if (above_ema20 and gap20_ok and ema5_above_ema20
                and breakout_ok and volume_strong and atr_soft_ok
                and not overextended):
            return "BUY"

        if (above_ema20 and gap20_ok and ema5_above_ema20
                and breakout_ok and volume_strong and atr_hard_ok):
            vol_strong_extra = (pd.notna(volume_ratio)
                                and volume_ratio >= CONFIG["STAGE1_BREAKOUT_VOL_STRONG"])
            close_strong = (pd.notna(close_position)
                            and close_position >= CONFIG["STAGE1_BREAKOUT_CLOSE_POS_MIN"])
            if vol_strong_extra and close_strong and not overextended:
                return "BUY_STRONG_BREAKOUT"

        # 新增：Setup + Trigger（放宽成交量）
        if setup_active and (ema5_reclaim or strong_close):
            if overextended or atr_level == "TOO_FAR":
                return "WAIT_PULLBACK"
            return "BUY"

        # 后备：站上 EMA20 且非追高
        if above_ema20 and gap20_ok and ema5_above_ema20:
            if overextended or atr_level == "TOO_FAR":
                return "WAIT_PULLBACK"
            if ema5_reclaim:
                return "BUY"
            return "WAIT_EMA5"
        if above_ema20 and gap20_ok:
            return "WAIT_EMA5"
        return "NONE"

    # ==================================================
    # Stage 2（核心改写）
    # ==================================================
    if stage == "Stage 2":
        if setup_active:
            trigger_fired = ema5_reclaim or breakout_confirmed or strong_close

            if trigger_fired:
                # 防止追高检查
                if overextended or atr_level == "TOO_FAR":
                    return "WAIT_PULLBACK"
                return "BUY"

            # Setup 已建立但 Trigger 未到
            if setup_type == "EMA20_PULLBACK":
                return "WAIT_EMA5"
            if setup_type == "BREAKOUT":
                return "WAIT_VOL"
            return "WAIT_EMA5"

        # 没有 Setup → 等回调形成 Setup
        return "WAIT_PULLBACK"

    return "NONE"


def exit_engine_v2(row, prev_row):
    close = row.get("Close", np.nan)
    ema5 = row.get("EMA5", np.nan)
    ema20 = row.get("EMA20", np.nan)
    ema20_slope = row.get("EMA20_Actual_Slope_5", np.nan)
    ema60_slope = row.get("EMA60_Actual_Slope_5", np.nan)
    adx = row.get("ADX", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)

    if pd.isna(close) or pd.isna(ema5) or pd.isna(ema20):
        return "数据不足"

    if pd.notna(ema60_slope) and ema60_slope < 0:
        return "EXIT"

    if prev_row is not None:
        prev_close = prev_row.get("Close", np.nan)
        prev_ema20 = prev_row.get("EMA20", np.nan)
        if (pd.notna(prev_close) and pd.notna(prev_ema20)
                and close < ema20 and prev_close < prev_ema20):
            return "减仓"

    if pd.notna(adx) and adx < CONFIG["ADX_CHOPPY_THRESHOLD"]:
        if close < ema5:
            return "观察"
        return "持有"

    if pd.notna(ema20_slope) and ema20_slope < 0:
        if close < ema20 and ema5 < ema20:
            return "保护利润"
        if (close < ema5
                and pd.notna(volume_ratio)
                and volume_ratio > CONFIG["PROTECT_PROFIT_VOL_CONFIRM"]):
            return "保护利润"
        return "观察"

    if close < ema5:
        return "观察"

    return "持有"


# ================================================================
# 6d. 顶部预警分
# ================================================================

def calculate_top_warning_score(row, prev_rows):
    score = 0
    reasons = []

    adx = row.get("ADX", np.nan)
    adx_chg = row.get("ADX_Change_5", np.nan)
    if (pd.notna(adx) and adx > CONFIG["TOP_WARNING_ADX_HIGH"]
            and pd.notna(adx_chg) and adx_chg < 0):
        score += 25
        reasons.append("ADX见顶回落")

    di_spread = row.get("DI_Spread", np.nan)
    if pd.notna(di_spread) and di_spread < CONFIG["TOP_WARNING_DI_SPREAD_MIN"]:
        score += 15
        reasons.append("DI收窄")

    vr = row.get("Volume_Ratio", np.nan)
    cpr = row.get("Close_Position_Range", np.nan)
    if (pd.notna(vr) and vr > CONFIG["TOP_WARNING_VOL_RATIO"]
            and pd.notna(cpr) and cpr < CONFIG["TOP_WARNING_CLOSE_POS"]):
        score += 15
        reasons.append("放量滞涨")

    if len(prev_rows) >= 2 and pd.notna(cpr):
        last2 = prev_rows.tail(2)["Close_Position_Range"].tolist()
        if all(pd.notna(x) and x < 0.35 for x in last2) and cpr < 0.35:
            score += 15
            reasons.append("连续低位收盘")

    ext = row.get("Extension_EMA20_Pct", np.nan)
    if pd.notna(ext) and ext > CONFIG["TOP_WARNING_EXTENSION"]:
        score += 15
        reasons.append("Extension过大")

    atr_d = row.get("ATR_Distance", np.nan)
    if pd.notna(atr_d) and abs(atr_d) > CONFIG["TOP_WARNING_ATR_DIST"]:
        score += 10
        reasons.append("ATR Distance过大")

    if len(prev_rows) >= 1:
        prev_eq = prev_rows.iloc[-1].get("Entry_Quality_Score", np.nan)
        cur_eq = row.get("Entry_Quality_Score", np.nan)
        if (pd.notna(prev_eq) and pd.notna(cur_eq)
                and prev_eq < CONFIG["TOP_WARNING_ENTRY_Q_LOW"]
                and cur_eq < CONFIG["TOP_WARNING_ENTRY_Q_LOW"]):
            score += 10
            reasons.append("连续买点质量差")

    return score, reasons


# ================================================================
# 6e. AI TIMING ENGINE
# ================================================================

# ================================================================
# 6d-2. V4.2 新增模块：安全工具 / 动态 ATR / 价量与 OBV 分析 /
#        Entry 确认层 / AI 信心分
# ================================================================

def _safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        f = float(v)
        if math.isnan(f) or math.isinf(f):
            return default
        return f
    except Exception:
        return default


def _safe_str(v, default=""):
    if v is None:
        return default
    try:
        if isinstance(v, float) and math.isnan(v):
            return default
    except Exception:
        pass
    s = str(v)
    return s if s.strip() else default


def calculate_dynamic_atr_multiplier(stage):
    """V4.2：按阶段取 ATR 倍数（Stage 2 宽松，Stage 3B/4 收紧）"""
    key = get_stage_key(stage)
    return float(CONFIG["ATR_MULTIPLIER_BY_STAGE"].get(key, 2.0))


def _trail_mult(stage):
    """移动止损所用的 ATR 倍数（可在 CONFIG 中切回固定值）"""
    if CONFIG.get("USE_DYNAMIC_ATR_TRAIL", False):
        return calculate_dynamic_atr_multiplier(stage)
    return float(CONFIG["ATR_TRAIL_MULTIPLIER"])


def classify_price_volume(price_chg, vol_chg, vol_thr=5.0):
    """价量关系：价涨量增 / 价涨量缩 / 价跌量增 / 价跌量缩 ..."""
    if pd.isna(price_chg) or pd.isna(vol_chg):
        return "数据不足"
    if vol_chg > vol_thr:
        v = "量增"
    elif vol_chg < -vol_thr:
        v = "量缩"
    else:
        v = "量平"
    if price_chg > 0:
        return "价涨" + v
    if price_chg < 0:
        return "价跌" + v
    return "价平" + v


def classify_obv_divergence(row):
    """价格 5 日变化 vs OBV 5 日变化，判断同步 / 背离"""
    r5 = row.get("Return_5", np.nan)
    o5 = row.get("OBV_Change_5", np.nan)
    if pd.isna(r5) or pd.isna(o5):
        return "数据不足"
    if r5 > 0 and o5 < 0:
        return "⚠️ 顶背离（价涨 OBV 降，资金不认同）"
    if r5 < 0 and o5 > 0:
        return "底背离迹象（价跌 OBV 升，有资金承接）"
    if r5 > 0 and o5 > 0:
        return "价量同步上行（健康）"
    if r5 < 0 and o5 < 0:
        return "价量同步下行（偏弱）"
    return "中性"


def entry_confirmation_layer(row, base_signal):
    """
    V4.2：BUY 确认层（移植自 V4.1）
    5 个因子：EMA5>EMA20、EMA20 斜率>0、ADX>=20、+DI>=-DI、量比>=1.2
      - >=4 个：维持原信号
      - 3 个：BUY_STRONG_BREAKOUT 降级为 BUY，BUY 降级为 WAIT_VOL
      - <3 个：降级为 WAIT_VOL
    返回 (signal, confirm_dict)
    """
    ema5 = row.get("EMA5", np.nan)
    ema20 = row.get("EMA20", np.nan)
    ema20_slope = row.get("EMA20_Actual_Slope_5", np.nan)
    adx = row.get("ADX", np.nan)
    di_spread = row.get("DI_Spread", np.nan)
    vol_ratio = row.get("Volume_Ratio", np.nan)

    ema5_ok = bool(pd.notna(ema5) and pd.notna(ema20) and ema5 > ema20)
    slope_ok = bool(pd.notna(ema20_slope) and ema20_slope > 0)
    adx_ok = bool(pd.notna(adx) and adx >= CONFIG["ADX_WEAK"])
    di_ok = bool(pd.notna(di_spread) and di_spread >= 0)
    vol_ok = bool(pd.notna(vol_ratio) and vol_ratio >= CONFIG["ENTRY_VOLUME_MIN"])
    confirmations = int(sum([ema5_ok, slope_ok, adx_ok, di_ok, vol_ok]))

    confirm_dict = {
        "BUY_Confirmations": confirmations,
        "BUY_Confirm_EMA5": ema5_ok,
        "BUY_Confirm_EMA20Slope": slope_ok,
        "BUY_Confirm_ADX": adx_ok,
        "BUY_Confirm_DI": di_ok,
        "BUY_Confirm_Vol": vol_ok,
    }

    if not CONFIG.get("ENTRY_CONFIRM_ENABLED", True):
        return base_signal, confirm_dict
    if base_signal not in ("BUY", "BUY_STRONG_BREAKOUT"):
        return base_signal, confirm_dict
    if confirmations >= CONFIG["ENTRY_CONFIRM_MIN_FACTORS"]:
        return base_signal, confirm_dict
    if confirmations >= CONFIG["ENTRY_CONFIRM_SOFT_FACTORS"]:
        if base_signal == "BUY_STRONG_BREAKOUT":
            return "BUY", confirm_dict
        return "WAIT_VOL", confirm_dict
    return "WAIT_VOL", confirm_dict


def calculate_ai_confidence_score(row):
    """V4.2：AI 信心分（趋势 / 股票质量 / 买点质量 / 量能 / 风险，按阶段加权）"""
    stage_key = get_stage_key(row.get("阶段", "Stage 0"))
    stock_q = _safe_float(row.get("Stock_Quality_Score", 0), 0.0)
    entry_q = _safe_float(row.get("Entry_Quality_Score", 0), 0.0)
    adx = _safe_float(row.get("ADX", 0), 0.0)
    vol_ratio = _safe_float(row.get("Volume_Ratio", 1), 1.0)
    top_warn = _safe_float(row.get("Top_Warning_Score", 0), 0.0)

    trend_comp = float(np.clip((adx / 40.0) * 100, 0, 100))
    volume_comp = float(np.clip((vol_ratio / 2.0) * 100, 0, 100))
    risk_comp = float(np.clip(100 - top_warn * 2.5, 0, 100))

    weights = {
        "Stage 1":  {"trend": 0.20, "stock_q": 0.30, "entry_q": 0.30, "volume": 0.10, "risk": 0.10},
        "Stage 2":  {"trend": 0.30, "stock_q": 0.30, "entry_q": 0.20, "volume": 0.10, "risk": 0.10},
        "Stage 3":  {"trend": 0.20, "stock_q": 0.20, "entry_q": 0.15, "volume": 0.15, "risk": 0.30},
        "Stage 3A": {"trend": 0.20, "stock_q": 0.20, "entry_q": 0.15, "volume": 0.15, "risk": 0.30},
        "Stage 3B": {"trend": 0.15, "stock_q": 0.15, "entry_q": 0.10, "volume": 0.10, "risk": 0.50},
        "Stage 4":  {"trend": 0.10, "stock_q": 0.10, "entry_q": 0.05, "volume": 0.05, "risk": 0.70},
        "Stage 0":  {"trend": 0.20, "stock_q": 0.30, "entry_q": 0.30, "volume": 0.10, "risk": 0.10},
    }
    w = weights.get(stage_key, weights["Stage 1"])
    total = (trend_comp * w["trend"] + stock_q * w["stock_q"] + entry_q * w["entry_q"]
             + volume_comp * w["volume"] + risk_comp * w["risk"])
    return round(float(np.clip(total, 0, 100)), 2)


def _get_ai_confidence(row):
    v = row.get("AI_Confidence_Score", np.nan)
    if v is None or pd.isna(v):
        v = calculate_ai_confidence_score(row)
    return float(v)


def generate_master_prompt_payload(row, name, code):
    """
    V4.2：机构级 Master Prompt（完整版）
    含：均线 / 扣底 / ADX +DI -DI / 价量变化 / OBV / ATR 与防守位 /
        评分 / 引擎信号 / 顶部预警
    """
    def f(key, d=4):
        return fmt(row.get(key, np.nan), d)

    def n(key):
        v = row.get(key, np.nan)
        try:
            return f"{float(v):,.0f}" if pd.notna(v) else "N/A"
        except Exception:
            return "N/A"

    def p(key, d=2):
        v = row.get(key, np.nan)
        try:
            return f"{float(v):+.{d}f}%" if pd.notna(v) else "N/A"
        except Exception:
            return "N/A"

    timing_code, priority, action_level, reason = ai_timing_engine(row)
    date_str = row.name.strftime("%Y-%m-%d") if hasattr(row.name, "strftime") else str(row.name)
    stage = _safe_str(row.get("阶段", "Stage 0"), "Stage 0")
    stock_type = row.get("Stock_Type", "mixed")
    type_label = {"pullback": "回调型", "breakout": "突破型", "mixed": "混合型"}.get(stock_type, stock_type)
    long_ok = "是" if bool(row.get("Long_Trend_OK", True)) else "否（Close<EMA200 或 EMA200 下行）"
    pv1 = classify_price_volume(row.get("Return_1", np.nan), row.get("Volume_Change_1", np.nan))
    pv5 = classify_price_volume(row.get("Return_5", np.nan), row.get("Volume_Change_5", np.nan))
    obv_div = classify_obv_divergence(row)
    resistance = row.get("Prev_High_20", np.nan)
    atr_stop = row.get("ATR_Stop_Dynamic", np.nan)
    setup_txt = (f"{row.get('Setup_Type', 'NONE')}（年龄 {row.get('Setup_Age', 'N/A')} 天，"
                 f"活跃={bool(row.get('Setup_Active', False))}，突破已确认={bool(row.get('Breakout_Confirmed', False))}）")

    prompt_text = f"""【请执行以下 Master Prompt 机构级深度分析】
股票名称: {name} ({clean_stock_code(code)})　日期: {date_str}
当前收盘价: RM{f('Close')}　股票类型: {type_label}

■ 市场阶段与权限
- 阶段: {stage}
- Trading Permission: {row.get('Trading_Permission', 'N/A')}　EMA200 长趋势确认: {long_ok}
- 最终信号: {row.get('最终信号', 'N/A')}　AI 时机: {timing_code}（动作: {action_level}；原因: {reason}）

■ 均线结构
- EMA5=RM{f('EMA5')}, EMA20=RM{f('EMA20')}, EMA60=RM{f('EMA60')}, EMA200=RM{f('EMA200')}
- 价格距离: EMA20 乖离 {p('Extension_EMA20_Pct')}, EMA60 乖离 {p('Extension_EMA60_Pct')}, 距 EMA200 {p('Price_vs_EMA200_Pct')}
- 斜率: EMA20(5日)={f('EMA20_Actual_Slope_5')}, EMA60(5日)={f('EMA60_Actual_Slope_5')}, EMA200(20日)={f('EMA200_Actual_Slope_20')}
- 扣底状态: EMA5={row.get('EMA5_扣底状态', 'N/A')}, EMA20={row.get('EMA20_扣底状态', 'N/A')}, EMA60={row.get('EMA60_扣底状态', 'N/A')}

■ 趋势强度（ADX / +DI / -DI）
- ADX={f('ADX', 2)}（{row.get('ADX_强度', 'N/A')}），5日变化={f('ADX_Change_5', 2)}，5日趋势={row.get('ADX_5日趋势', 'N/A')}
- +DI={f('+DI', 2)}, -DI={f('-DI', 2)}, DI Spread={f('DI_Spread', 2)}（{row.get('DI_方向', 'N/A')}）

■ 价量变化
- 成交量={n('Volume')}, 20日均量={n('Volume_MA20')}, 量比(Volume Ratio)={f('Volume_Ratio', 2)}
- 量能变化: 1日 {p('Volume_Change_1', 1)}, 5日均量变化 {p('Volume_Change_5', 1)}, MA5 对 MA20 {p('Volume_MA5_vs_MA20_Pct', 1)}
- 涨跌幅: 1日 {p('Return_1')}, 5日 {p('Return_5')}, 20日 {p('Return_20')}
- 价量关系: 1日={pv1}；5日={pv5}
- K线: 收盘位置(Close Position)={f('Close_Position_Range', 2)}, 实体占比={f('Candle_Body_Pct', 1)}%, 上影线={f('Upper_Shadow_Pct', 1)}%, 下影线={f('Lower_Shadow_Pct', 1)}%
- 突破量确认: {'是' if bool(row.get('Breakout_Volume', False)) else '否'}

■ OBV 资金流
- OBV={n('OBV')}, OBV_MA5={n('OBV_MA5')}, OBV 5日变化={n('OBV_Change_5')}, 距 MA5 {p('OBV_MA_Distance_Pct', 1)}
- OBV 趋势={row.get('OBV_趋势', 'N/A')}（5日趋势 {row.get('OBV_5日趋势', 'N/A')}）
- 价量/OBV 判断: {obv_div}
- 布林带位置: {row.get('BB_位置', 'N/A')}

■ 波动与防守
- ATR14=RM{f('ATR14')}, ATR%={f('ATR_Pct', 2)}%, ATR Distance={f('ATR_Distance', 2)}（{row.get('ATR_Risk', 'N/A')}）
- 动态 ATR 倍数={f('ATR_Multiplier', 1)}x（按阶段），ATR 动态防守价=RM{fmt(atr_stop)}
- 结构防守价（支撑）=RM{f('结构防守价')}, 近20日低点=RM{f('Recent_Low_20')}
- 近期压力位（20日高点）=RM{fmt(resistance)}

■ 量化评分与引擎
- 股票质量分={f('Stock_Quality_Score', 1)}/100, 买点质量分={f('Entry_Quality_Score', 1)}/100（{row.get('Entry_Quality_Rating', 'N/A')}）, AI 信心={_get_ai_confidence(row):.0f}/100
- 顶部预警分={f('Top_Warning_Score', 0)}（{_safe_str(row.get('Top_Warning_Reasons', ''), '无')}）
- Entry Engine={row.get('Entry_Engine', 'N/A')}（原始={row.get('Entry_Engine_Raw', 'N/A')}，确认因子 {row.get('BUY_Confirmations', 'N/A')}/5）, Exit Engine={row.get('Exit_Engine', 'N/A')}
- Setup: {setup_txt}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告，
并特别回答：(1) +DI/-DI 与 ADX 是否支持当前阶段；(2) 价量关系与 OBV 是否确认或背离；(3) 防守价与压力位之间的盈亏比。"""
    return prompt_text.strip()


def ai_timing_engine(row):
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    stock_type = row.get("Stock_Type", "mixed")
    entry = row.get("Entry_Engine", "NONE")
    exit_e = row.get("Exit_Engine", "持有")
    permission = row.get("Trading_Permission", "YELLOW")
    atr_distance = row.get("ATR_Distance", np.nan)
    extension = row.get("Extension_EMA20_Pct", np.nan)
    stock_q = float(row.get("Stock_Quality_Score", row.get("Final_Score", 0)) or 0)
    entry_q = float(row.get("Entry_Quality_Score", 50) or 0)
    breakout_age = row.get("Breakout_Age", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)
    adx = row.get("ADX", np.nan)
    di_spread = row.get("DI_Spread", np.nan)
    close = row.get("Close", np.nan)
    ema5 = row.get("EMA5", np.nan)
    ema20 = row.get("EMA20", np.nan)
    ema60_slope = row.get("EMA60_Actual_Slope_5", np.nan)

    ad = abs(float(atr_distance)) if pd.notna(atr_distance) else np.nan
    ext = abs(float(extension)) if pd.notna(extension) else np.nan

    strong_buy = (
        stage in ("Stage 1", "Stage 2")
        and entry in ("BUY", "BUY_STRONG_BREAKOUT")
        and permission == "GREEN"
        and stock_q >= CONFIG["STRONG_BUY_STOCK_Q_MIN"]
    )

    if CONFIG["TOP_WARNING_ENABLED"]:
        top_score = float(row.get("Top_Warning_Score", 0) or 0)
        top_reasons = row.get("Top_Warning_Reasons", "")
        if top_score >= CONFIG["TOP_WARNING_EXIT"]:
            return "NO_BUY", "高", "清仓", f"顶部预警分{top_score:.0f}：{top_reasons}"
        if top_score >= CONFIG["TOP_WARNING_REDUCE_2"]:
            return "NO_BUY", "高", "减仓1/3", f"顶部预警分{top_score:.0f}：{top_reasons}"
        if top_score >= CONFIG["TOP_WARNING_REDUCE_1"]:
            return "NO_BUY", "中", "减仓1/3", f"顶部预警分{top_score:.0f}：{top_reasons}"

    if exit_e == "EXIT":
        return "NO_BUY", "高", "退出", "EMA60 下弯，趋势破坏"

    if strong_buy and exit_e in ("保护利润", "观察"):
        if (entry_q >= CONFIG["STRONG_BUY_ENTRY_Q_MIN"]
                and pd.notna(ad) and ad <= CONFIG["STRONG_BUY_ATR_MAX"]):
            return ("BUY_TODAY", "中", "建仓-半仓",
                    f"强买点覆盖保护利润(StockQ={stock_q:.0f}, EntryQ={entry_q:.0f})，仓位减半")
        return ("BUY_ON_PULLBACK_EMA20", "中", "等回调",
                f"强买点但需等回踩(StockQ={stock_q:.0f}, EntryQ={entry_q:.0f})")

    if exit_e == "保护利润":
        return "NO_BUY", "高", "减仓1/3", "EMA20 下弯，减仓1/3锁利，剩余仓位用ATR移动止损"
    if exit_e == "减仓":
        return "NO_BUY", "高", "减仓1/3", "连续2天 < EMA20，再减1/3，剩余仓位用ATR移动止损"

    if permission == "RED":
        return "NO_BUY", "高", "减码", "EMA60 转弱，停止新增多仓"

    if stage == "Stage 4":
        return "NO_BUY", "高", "退出", "Stage 4 趋势破坏"

    if stage in ("Stage 3", "Stage 3A", "Stage 3B"):
        if stage != "Stage 3B" and pd.notna(adx) and adx < CONFIG["ADX_CHOPPY_THRESHOLD"]:
            return "HOLD", "低", "持有", f"{stage} 但 ADX 弱，暂不保护利润"

        cond = 0
        cond_reasons = []

        adx_chg = row.get("ADX_Change_5", np.nan)
        if pd.notna(adx) and adx > 40 and pd.notna(adx_chg) and adx_chg < -3:
            cond += 1
            cond_reasons.append("ADX见顶回落")

        if pd.notna(di_spread) and di_spread < 0:
            cond += 1
            cond_reasons.append("DI转空")

        prev_close = row.get("Prev_Close", np.nan)
        prev_ema20 = row.get("Prev_EMA20", np.nan)
        ema20_slope = row.get("EMA20_Actual_Slope_5", np.nan)
        if (pd.notna(close) and pd.notna(ema20)
                and close < ema20
                and pd.notna(prev_close) and pd.notna(prev_ema20)
                and prev_close < prev_ema20
                and pd.notna(ema20_slope) and ema20_slope < 0):
            cond += 1
            cond_reasons.append("连续2天<EMA20且EMA20下弯")

        top_score = float(row.get("Top_Warning_Score", 0) or 0)
        if top_score >= 45:
            cond += 1
            cond_reasons.append(f"顶部预警{top_score:.0f}")

        cpr = row.get("Close_Position_Range", np.nan)
        if (pd.notna(volume_ratio) and volume_ratio > 1.5
                and pd.notna(cpr) and cpr < 0.3):
            cond += 1
            cond_reasons.append("放量下跌")

        structure_stop = row.get("结构防守价", np.nan)
        if pd.notna(structure_stop) and pd.notna(close) and close < structure_stop:
            cond += 1
            cond_reasons.append("跌破结构防守价")

        min_cond = (CONFIG["STAGE3B_REDUCE_MIN_CONDITIONS"] if stage == "Stage 3B"
                    else CONFIG["STAGE3_REDUCE_MIN_CONDITIONS"])
        if cond >= min_cond:
            return ("NO_BUY", "高", "减仓1/3",
                    f"{stage} 满足{cond}个转弱条件(阈值{min_cond})：{'、'.join(cond_reasons)}")

        if (exit_e == "持有"
                and stock_q >= CONFIG["STAGE3_PROTECT_SCORE"]
                and permission == "GREEN"):
            return "HOLD", "中", "保护利润", f"{stage} 未确认转弱，继续持有"
        return "HOLD", "中", "持有", f"{stage} 未确认转弱，继续持有"

    if (CONFIG["RE_ENTRY_ENABLED"]
            and entry == "NONE"
            and permission == "GREEN"
            and stage in ("Stage 1", "Stage 2")
            and pd.notna(close) and pd.notna(ema20) and pd.notna(ema5)):
        if close > ema20 and ema5 > ema20:
            re_ok = True
            if CONFIG["RE_ENTRY_REQUIRE_VOL"]:
                vr = row.get("Volume_Ratio", np.nan)
                if pd.isna(vr) or vr < CONFIG["RE_ENTRY_VOL_MIN"]:
                    re_ok = False
            if re_ok:
                return ("BUY_ON_EMA5_RECLAIM", "中", "再入场",
                        "重新站上 EMA20，可再入场")

    if entry in ("BUY", "BUY_STRONG_BREAKOUT"):
        if stock_q < 60:
            return "NO_BUY", "高", "观望", f"股票质量不足({stock_q:.0f})"

        if CONFIG["BLOCK_BUY_WHEN_EMA60_DOWN"]:
            if pd.notna(ema60_slope) and ema60_slope < 0:
                return "NO_BUY", "高", "观望", "EMA60 下弯，禁止买入"

        # ---- V4.2：AI 信心门槛 ----
        confidence = _get_ai_confidence(row)
        if confidence < CONFIG["AI_CONFIDENCE_MIN_BUY"]:
            return ("BUY_TOMORROW", "中", "观察",
                    f"AI信心偏低({confidence:.0f})，等次日确认")

        if entry == "BUY_STRONG_BREAKOUT":
            return ("BUY_TODAY", "中", "建仓-半仓",
                    f"Stage 1 强势突破({entry_q:.0f}) + 巨量收盘强，仓位减半")

        if stock_type == "breakout":
            breakout_age_max = CONFIG["BREAKOUT_AGE_MAX_EXTENDED"]
            breakout_vol_min = CONFIG["BREAKOUT_VOL_MIN_EXTENDED"]
        else:
            breakout_age_max = CONFIG["STAGE1_BREAKOUT_AGE_MAX"]
            breakout_vol_min = CONFIG["STAGE1_BREAKOUT_VOL_MIN"]

        stage1_breakout_soft = (
            stage == "Stage 1"
            and pd.notna(breakout_age) and breakout_age != 999
            and breakout_age <= breakout_age_max
            and pd.notna(volume_ratio)
            and volume_ratio >= breakout_vol_min
            and pd.notna(ad)
            and ad < CONFIG["STAGE1_BREAKOUT_ATR_MAX"]
        )
        if stage1_breakout_soft:
            adx_ok = (pd.notna(adx) and adx > CONFIG["STAGE1_ADX_MIN"]
                      and pd.notna(di_spread) and di_spread > 0)
            if not adx_ok:
                return ("BUY_ON_PULLBACK_EMA20", "中", "等回调",
                        f"Stage 1 突破但 ADX 不足({adx:.0f})，等回 EMA20")
            return ("BUY_TODAY", "高", "建仓",
                    f"Stage 1 突破追买({entry_q:.0f}) + 放量确认")

        if entry_q < CONFIG["ENTRY_QUALITY_WAIT_MIN"]:
            return ("BUY_ON_PULLBACK_EMA20", "中", "等回调",
                    f"买点质量差({entry_q:.0f})，等回 EMA20")

        if (entry_q >= CONFIG["ENTRY_QUALITY_BUY_MIN"]
                and pd.notna(ad) and ad <= CONFIG["AI_BUY_TODAY_ATR_MAX"]
                and pd.notna(ext) and ext <= CONFIG["AI_BUY_TODAY_EXT_MAX"]):
            action = "建仓" if stage == "Stage 1" else "加仓"
            return ("BUY_TODAY", "高", action,
                    f"买点优秀({entry_q:.0f}) + ATR 健康")

        if (pd.notna(ad) and ad <= CONFIG["AI_BUY_TOMORROW_ATR_MAX"]
                and pd.notna(ext) and ext <= CONFIG["AI_BUY_TOMORROW_EXT_MAX"]):
            action = "建仓" if stage == "Stage 1" else "加仓"
            return ("BUY_TOMORROW", "中", action,
                    f"买点一般({entry_q:.0f})，次日确认")

        return ("BUY_ON_PULLBACK_EMA20", "中", "等回调",
                f"买点质量不足({entry_q:.0f})，等回 EMA20")

    if entry == "WAIT_VOL":
        return "BUY_ON_VOLUME", "中", "观察", "等量能放大"

    if entry == "WAIT_EMA5":
        return "BUY_ON_EMA5_RECLAIM", "中", "观察", "等 EMA5 重新站上"

    if entry == "WAIT_PULLBACK":
        return "BUY_ON_PULLBACK_EMA20", "低", "等回调", "等回调至 EMA20 附近"

    if stage == "Stage 1":
        return "BUY_ON_BREAKOUT", "低", "观察", "Stage 1 结构待突破"

    if stage == "Stage 2":
        return "HOLD", "低", "持有", "Stage 2 持有，无加仓触发"

    return "NO_BUY", "低", "观望", "未形成明确趋势"


def ai_buy_zone(row, timing_code):
    close = row.get("Close", np.nan)
    ema20 = row.get("EMA20", np.nan)
    if pd.isna(close):
        return "N/A"

    band = CONFIG["AI_PULLBACK_BAND_PCT"] / 100.0

    if timing_code in ["BUY_TODAY", "BUY_TOMORROW", "BUY_ON_VOLUME", "BUY_ON_EMA5_RECLAIM"]:
        low = close * (1 - band)
        high = close * (1 + band)
        return f"RM{low:.4f} – RM{high:.4f}"

    if timing_code == "BUY_ON_PULLBACK_EMA20":
        if pd.notna(ema20):
            low = ema20 * (1 - band)
            high = ema20 * (1 + band)
            return f"RM{low:.4f} – RM{high:.4f}（EMA20 附近）"
        low = close * (1 - band)
        high = close * (1 + band)
        return f"RM{low:.4f} – RM{high:.4f}"

    if timing_code == "BUY_ON_BREAKOUT":
        high_n = row.get("High", np.nan)
        if pd.notna(high_n):
            low = high_n * (1 - band)
            high = high_n * (1 + band)
            return f"突破 RM{high_n:.4f}（区间 RM{low:.4f} – RM{high:.4f}）"
        return "突破前高后确认"

    return "N/A"


def ai_position_suggestion(row, timing_code, action_level):
    stage = get_stage_key(row.get("阶段", "Stage 0"))
    permission = row.get("Trading_Permission", "YELLOW")
    atr_pct = row.get("ATR_Pct", np.nan)
    entry_q = float(row.get("Entry_Quality_Score", 50) or 50)
    breakout_age = row.get("Breakout_Age", np.nan)
    volume_ratio = row.get("Volume_Ratio", np.nan)

    if timing_code == "NO_BUY":
        return "0%（不新增）"

    if action_level == "建仓-半仓":
        base = "首仓 15%，确认后加 15%"
    elif action_level == "建仓":
        base = "首仓 30%，确认后加 30%"
    elif action_level == "加仓":
        base = "加 20–30%"
    elif action_level == "观察":
        base = "0%（等确认后首仓 20–30%）"
    elif action_level == "等回调":
        base = "0%（回调到位后首仓 20–30%）"
    elif action_level == "再入场":
        base = "首仓 15%"
    elif action_level == "持有":
        base = "维持现有仓位"
    elif action_level == "减仓1/3":
        base = "减 1/3，剩余用 ATR 移动止损"
    elif action_level == "清仓":
        base = "清仓"
    else:
        base = "谨慎"

    if stage == "Stage 1":
        base = base.replace("首仓 30%", "首仓 20–30%")

    stage1_breakout = (
        stage == "Stage 1"
        and pd.notna(breakout_age) and breakout_age != 999
        and breakout_age <= CONFIG["STAGE1_BREAKOUT_AGE_MAX"]
        and pd.notna(volume_ratio)
        and volume_ratio >= CONFIG["STAGE1_BREAKOUT_VOL_MIN"]
    )

    if permission == "YELLOW":
        if stage1_breakout:
            base += "；Stage 1 突破追买，YELLOW 不减半"
        else:
            base += "；YELLOW 权限，减半"

    if pd.notna(atr_pct) and float(atr_pct) > 6:
        base += "；ATR% 偏高，进一步减半"

    if entry_q < 55:
        base += "；买点质量偏低，仓位减半"

    # ---- V4.2：成交量 / ATR 分级决定仓位 ----
    vr = row.get("Volume_Ratio", np.nan)
    atr_level = classify_atr_level(row.get("ATR_Distance", np.nan))

    if timing_code in ("BUY_TODAY", "BUY_TOMORROW"):
        if pd.notna(vr) and vr >= 1.5:
            base += "；成交量强确认（正常仓位）"
        elif pd.notna(vr) and vr >= 1.2:
            base += "；成交量良好（正常仓位）"
        elif pd.notna(vr) and vr >= 1.0:
            base += "；成交量一般（仓位 ×0.7）"
        elif pd.notna(vr) and vr < 1.0:
            base += "；成交量偏弱（仓位 ×0.5）"

        if atr_level == "EXCELLENT":
            base += "；ATR 优秀（可正常仓位）"
        elif atr_level == "ACCEPTABLE":
            base += "；ATR 可接受（仓位 ×0.8）"
        elif atr_level == "TOO_FAR":
            base += "；ATR 偏远（建议减半或等回调）"

    return base


def ai_invalidation(row, timing_code):
    ema20 = row.get("EMA20", np.nan)
    ema60 = row.get("EMA60", np.nan)
    ema20_slope = row.get("EMA20_Actual_Slope_5", np.nan)
    ema60_slope = row.get("EMA60_Actual_Slope_5", np.nan)

    if timing_code == "NO_BUY":
        if pd.notna(ema60) and pd.notna(ema60_slope) and ema60_slope > 0:
            return f"若 EMA60 重新上弯且 Stage 回到 1/2（参考 EMA60 ≈ {ema60:.4f}），可重新评估。"
        return "等 Stage 回到 Stage 1 或 Stage 2，且 Permission 回到 GREEN，再重新评估。"

    conds = []
    if pd.notna(ema20):
        conds.append(f"Close < EMA20（≈{ema20:.4f}）连续2天")
    if pd.notna(ema20_slope):
        conds.append("EMA20 下弯")
    conds.append("Trading Permission 变 RED")
    return "；若出现以下任一情况，取消买入计划：" + "、".join(conds)


def ai_timing_text(timing_code):
    mapping = {
        "BUY_TODAY": "今天可买",
        "BUY_TOMORROW": "明日确认后买",
        "BUY_ON_PULLBACK_EMA20": "等回 EMA20 再买",
        "BUY_ON_VOLUME": "等放量再买",
        "BUY_ON_EMA5_RECLAIM": "等 EMA5 重新站上再买",
        "BUY_ON_BREAKOUT": "等突破再买",
        "HOLD": "持有不加仓",
        "NO_BUY": "不买",
    }
    return mapping.get(timing_code, timing_code)


def ai_reduce_ratio(exit_e, timing_code):
    if exit_e == "EXIT":
        return "100% 清仓"
    if exit_e == "减仓":
        return "再减 1/3，剩余用 ATR 移动止损"
    if exit_e == "保护利润":
        return "减 1/3，剩余用 ATR 移动止损"
    if exit_e == "观察":
        return "0% 仅观察"
    if timing_code == "NO_BUY":
        return "0% 不新增"
    return "0% 持有"


# ================================================================
# AI 时机 → 最终信号 统一映射
# ================================================================

def map_ai_to_final_signal(timing_code, exit_e):
    if timing_code in ("BUY_TODAY", "BUY_TOMORROW"):
        return "买入"
    if timing_code in ("BUY_ON_PULLBACK_EMA20",
                       "BUY_ON_VOLUME",
                       "BUY_ON_EMA5_RECLAIM",
                       "BUY_ON_BREAKOUT"):
        return "买入观察"
    if timing_code == "HOLD":
        return "持有"

    if exit_e == "EXIT":
        return "卖出"
    if exit_e in ("保护利润", "减仓"):
        return "减仓"
    if exit_e == "观察":
        return "观察"
    return "观望"


def generate_ai_recommendation(row):
    timing_code, priority, action_level, reason = ai_timing_engine(row)
    timing_text = ai_timing_text(timing_code)
    buy_zone = ai_buy_zone(row, timing_code)
    position = ai_position_suggestion(row, timing_code, action_level)
    invalidation = ai_invalidation(row, timing_code)
    exit_e = row.get("Exit_Engine", "持有")
    reduce_ratio = ai_reduce_ratio(exit_e, timing_code)

    stage = row.get("阶段", "Stage 0 (观察)")
    entry = row.get("Entry_Engine", "NONE")
    permission = row.get("Trading_Permission", "YELLOW")
    stock_q = float(row.get("Stock_Quality_Score", row.get("Final_Score", 0)) or 0)
    entry_q = float(row.get("Entry_Quality_Score", 0) or 0)
    entry_rating = row.get("Entry_Quality_Rating", "N/A")
    top_score = float(row.get("Top_Warning_Score", 0) or 0)
    top_reasons = row.get("Top_Warning_Reasons", "")

    conclusion_map = {
        "BUY_TODAY": "可以买，今天可执行。",
        "BUY_TOMORROW": "可以买，但不是今天，明日确认后执行。",
        "BUY_ON_PULLBACK_EMA20": "可以买，但不是今天，等回 EMA20。",
        "BUY_ON_VOLUME": "可以买，但今天不行，等放量确认。",
        "BUY_ON_EMA5_RECLAIM": "可以买，但今天不行，等 EMA5 重新站上。",
        "BUY_ON_BREAKOUT": "可以留意，等突破后买。",
        "HOLD": "不加仓，继续持有。",
        "NO_BUY": "不建议买入。",
    }
    conclusion = conclusion_map.get(timing_code, "观察")

    return (
        f"结论：{conclusion}"
        f" | 时机：{timing_text}（{timing_code}）"
        f" | 动作：{action_level}"
        f" | 原因：{reason}"
        f" | 股票质量：{stock_q:.0f}/100"
        f" | 买点质量：{entry_q:.0f}/100（{entry_rating}）"
        f" | AI信心：{_get_ai_confidence(row):.0f}/100"
        f" | 顶部预警分：{top_score:.0f}（{top_reasons}）"
        f" | 触发条件：Stage={stage}；Entry={entry}；Exit={exit_e}；Permission={permission}"
        f" | 参考买入区：{buy_zone}"
        f" | 仓位建议：{position}"
        f" | 减仓比例：{reduce_ratio}"
        f" | 优先级：{priority}"
        f" | 失效条件：{invalidation}"
    )


def generate_ai_summary(row, code, name):
    stage = row["阶段"]
    stage_key = get_stage_key(stage)
    signal = row["最终信号"]
    permission = row["Trading_Permission"]
    stock_type = row.get("Stock_Type", "mixed")
    close = row["Close"]
    ema5 = row.get("EMA5", np.nan)
    ema20 = row["EMA20"]
    ema60 = row["EMA60"]
    ema200 = row["EMA200"]
    adx = row["ADX"]
    plus_di = row["+DI"]
    minus_di = row["-DI"]
    raw_score = row.get("Raw_Score", np.nan)
    adjustment = row.get("Stage_Adjustment", np.nan)
    stock_q = row.get("Stock_Quality_Score", row.get("Final_Score", np.nan))
    entry_q = row.get("Entry_Quality_Score", np.nan)
    entry_rating = row.get("Entry_Quality_Rating", "N/A")
    atr_distance = row.get("ATR_Distance", np.nan)
    atr_pct = row.get("ATR_Pct", np.nan)
    atr_label = row.get("ATR_Risk", "数据不足")
    entry = row.get("Entry_Engine", "NONE")
    exit_e = row.get("Exit_Engine", "持有")
    top_score = float(row.get("Top_Warning_Score", 0) or 0)
    top_reasons = row.get("Top_Warning_Reasons", "")

    timing_code, priority, action_level, reason = ai_timing_engine(row)
    timing_text = ai_timing_text(timing_code)
    buy_zone = ai_buy_zone(row, timing_code)
    position = ai_position_suggestion(row, timing_code, action_level)
    invalidation = ai_invalidation(row, timing_code)
    reduce_ratio = ai_reduce_ratio(exit_e, timing_code)

    type_label = {"pullback": "回调型", "breakout": "突破型", "mixed": "混合型"}.get(stock_type, stock_type)

    reasons = [f"股票类型：{type_label}"]
    if permission == "GREEN":
        reasons.append("三层趋势绿色（EMA60↑+EMA200↑+Close>EMA60）")
    elif permission == "YELLOW":
        reasons.append("三层趋势黄色（EMA60↑但 Close<=EMA200 长趋势未确认，或趋势未完全转强）")
    else:
        reasons.append("三层趋势红色（EMA60↓），优先控制风险")

    if pd.notna(atr_distance):
        reasons.append(f"ATR Distance = {atr_distance:.2f} ATR，{atr_label}")
    if pd.notna(atr_pct):
        reasons.append(f"ATR% = {atr_pct:.2f}%")
    if pd.notna(plus_di) and pd.notna(minus_di) and pd.notna(adx):
        direction = "多方占优" if plus_di > minus_di else "空方占优"
        reasons.append(f"ADX {adx:.1f}，+DI {plus_di:.1f} / -DI {minus_di:.1f}，{direction}")
    reasons.append(f"Entry Engine V2 = {entry}")
    reasons.append(f"Exit Engine V2 = {exit_e}")
    reasons.append(f"AI信心 {_get_ai_confidence(row):.0f}/100")
    if top_score > 0:
        reasons.append(f"顶部预警分 {top_score:.0f}（{top_reasons}）")

    if stage_key == "Stage 1":
        action = "Stage 1 初升：EMA5 站上 EMA20 且买点质量良好时，可分批建仓。"
    elif stage_key == "Stage 2":
        action = "Stage 2 主升：回调至 EMA20 或 ATR 接近 EMA20，且 EMA5 重新站上、量能放大时，可考虑回调加仓。"
    elif stage_key == "Stage 3A":
        action = "Stage 3A 末升-健康调整：默认持有，满足至少2个转弱条件才减1/3，剩余仓位用 EMA5+ATR 动态移动止损。"
    elif stage_key == "Stage 3B":
        action = "Stage 3B 末升-趋势恶化：不加仓，满足1个转弱条件即减1/3，收紧 ATR 倍数移动止损。"
    elif stage_key == "Stage 3":
        action = "Stage 3 末升：默认持有，满足至少2个转弱条件才减1/3，剩余仓位用 EMA5+ATR 移动止损。"
    elif stage_key == "Stage 4":
        action = "Stage 4 趋势破坏：EMA60 下弯视为 EXIT。"
    else:
        action = "Stage 0 观察：等待 Stage 明确转强。"

    ai_block = (
        f"\n【AI 时机建议】\n"
        f"  结论：{ai_timing_text(timing_code)}\n"
        f"  时机：{timing_code}（{timing_text}）\n"
        f"  动作：{action_level}\n"
        f"  原因：{reason}\n"
        f"  股票类型：{type_label}\n"
        f"  股票质量分：{float(stock_q) if pd.notna(stock_q) else 0:.1f}/100\n"
        f"  买点质量分：{float(entry_q) if pd.notna(entry_q) else 0:.1f}/100（{entry_rating}）\n"
        f"  顶部预警分：{top_score:.0f}（{top_reasons}）\n"
        f"  参考买入区：{buy_zone}\n"
        f"  仓位建议：{position}\n"
        f"  减仓比例：{reduce_ratio}\n"
        f"  优先级：{priority}\n"
        f"  失效条件：{invalidation}"
    )

    return (
        f"{name} ({code}) 当前属于 {stage}。"
        f"股票类型={type_label}。"
        f"Trading Permission={permission}，最终信号={signal}。"
        f"股票质量 {float(stock_q) if pd.notna(stock_q) else 0:.1f}/100，"
        f"买点质量 {float(entry_q) if pd.notna(entry_q) else 0:.1f}/100（{entry_rating}）。"
        f"Raw Score={raw_score:.1f}，Stage Adjustment={adjustment:+.1f}。"
        f"收盘价 {close:.4f}，EMA5 {fmt(ema5)}，EMA20 {fmt(ema20)}，EMA60 {fmt(ema60)}，EMA200 {fmt(ema200)}。"
        + "；".join(reasons) + "。" + action
        + ai_block
    )


def print_terminal_report(row, code, name):
    print("\n" + "=" * 78)
    print(f"📊 EMA Trading System V4.2 — {name} ({code})")
    print("=" * 78)

    stock_type = row.get("Stock_Type", "mixed")
    type_label = {"pullback": "回调型", "breakout": "突破型", "mixed": "混合型"}.get(stock_type, stock_type)

    fields = [
        ("日期", row.name.strftime("%Y-%m-%d") if hasattr(row.name, "strftime") else row.name),
        ("股票类型", type_label),
        ("收盘价", fmt(row["Close"])),
        ("EMA5", fmt(row.get("EMA5", np.nan))),
        ("EMA5扣底(真实)", fmt(row.get("Deduct5", np.nan))),
        ("EMA5扣底状态", row.get("EMA5_扣底状态", "N/A")),

        ("EMA20", fmt(row["EMA20"])),
        ("EMA20扣底(真实)", fmt(row.get("Deduct20", np.nan))),
        ("EMA20扣底状态", row.get("EMA20_扣底状态", "N/A")),

        ("EMA60", fmt(row["EMA60"])),
        ("EMA60扣底(真实)", fmt(row.get("Deduct60", np.nan))),
        ("EMA60扣底状态", row.get("EMA60_扣底状态", "N/A")),

        ("EMA200", fmt(row["EMA200"])),
        ("EMA200 20日斜率", fmt(row.get("EMA200_Actual_Slope_20", np.nan))),

        ("--- 分数分离 ---", ""),
        ("股票质量分", fmt(row.get("Stock_Quality_Score", np.nan), 2)),
        ("股票质量评级", stock_quality_rating(row.get("Stock_Quality_Score", np.nan))),
        ("买点质量分", fmt(row.get("Entry_Quality_Score", np.nan), 2)),
        ("买点质量评级", row.get("Entry_Quality_Rating", "N/A")),
        ("EQ_EMA20位置(分)", fmt(row.get("EQ_EMA20_Position", np.nan), 1)),
        ("EQ_EMA5金叉(分)", fmt(row.get("EQ_EMA5_Cross", np.nan), 1)),
        ("EQ_ATR距离(分)", fmt(row.get("EQ_ATR_Distance", np.nan), 1)),
        ("EQ_突破年龄(分)", fmt(row.get("EQ_Breakout_Age", np.nan), 1)),
        ("EQ_量能(分)", fmt(row.get("EQ_Volume", np.nan), 1)),

        ("--- V4.2 Setup 记忆 ---", ""),
        ("Setup 类型", row.get("Setup_Type", "NONE")),
        ("Setup 年龄(天)", fmt(row.get("Setup_Age", np.nan), 0)),
        ("Setup 活跃", "✅" if bool(row.get("Setup_Active", False)) else "—"),
        ("突破位", fmt(row.get("Setup_Breakout_Level", np.nan))),
        ("突破已确认", "✅" if bool(row.get("Breakout_Confirmed", False)) else "—"),

        ("--- 顶部预警 ---", ""),
        ("顶部预警分", fmt(row.get("Top_Warning_Score", 0), 0)),
        ("顶部预警原因", row.get("Top_Warning_Reasons", "")),

        ("--- 关键参考 ---", ""),
        ("EMA5金叉年龄(天)", fmt(row.get("EMA5_Cross_Age", np.nan), 1)),
        ("突破年龄(天)", fmt(row.get("Breakout_Age", np.nan), 1)),
        ("结构防守价", fmt(row.get("结构防守价", np.nan))),
        ("ATR防守参考(2x)", fmt(row.get("ATR_Stop_2x", np.nan))),

        ("--- 趋势与风险 ---", ""),
        ("Trading Permission", row["Trading_Permission"]),
        ("阶段", row["阶段"]),
        ("ADX", fmt(row["ADX"], 2)),
        ("+DI", fmt(row["+DI"], 2)),
        ("-DI", fmt(row["-DI"], 2)),
        ("Volume Ratio", fmt(row["Volume_Ratio"], 2)),
        ("EMA20乖离率", pct(row["Extension_EMA20_Pct"])),
        ("ATR Distance", fmt(row.get("ATR_Distance", np.nan), 2)),
        ("ATR Health Score", fmt(row.get("ATR_Health_Score", np.nan), 2)),
        ("ATR Risk Score", fmt(row.get("ATR_Risk_Score", np.nan), 2)),
        ("ATR Risk", row.get("ATR_Risk", "数据不足")),
        ("OBV趋势", row["OBV_趋势"]),
        ("BB位置", row["BB_位置"]),

        ("--- V4.2 AI 信心 / 确认层 ---", ""),
        ("AI 信心", fmt(_get_ai_confidence(row), 0)),
        ("EMA200 长趋势确认", "✅" if bool(row.get("Long_Trend_OK", True)) else "❌"),
        ("BUY 确认因子", f"{row.get('BUY_Confirmations', 'N/A')}/5"),
        ("ATR动态倍数(阶段)", fmt(row.get("ATR_Multiplier", np.nan), 1)),
        ("ATR动态防守价", fmt(row.get("ATR_Stop_Dynamic", np.nan))),

        ("--- 价量 / OBV ---", ""),
        ("涨跌幅 1日/5日", f"{pct(row.get('Return_1', np.nan))} / {pct(row.get('Return_5', np.nan))}"),
        ("量能变化 1日/5日", f"{pct(row.get('Volume_Change_1', np.nan), 1)} / {pct(row.get('Volume_Change_5', np.nan), 1)}"),
        ("价量关系 1日", row.get("PV_关系_1日", "N/A")),
        ("价量关系 5日", row.get("PV_关系_5日", "N/A")),
        ("OBV 5日变化", fmt(row.get("OBV_Change_5", np.nan), 0)),
        ("OBV背离判断", classify_obv_divergence(row)),

        ("--- 引擎 ---", ""),
        ("Entry Engine V2", row.get("Entry_Engine", "N/A")),
        ("Exit Engine V2", row.get("Exit_Engine", "N/A")),
        ("最终信号", row["最终信号"]),
    ]

    for k, v in fields:
        print(f"{k:<22}: {v}")

    print("-" * 78)
    print("AI 摘要：")
    print(generate_ai_summary(row, code, name))
    print("=" * 78)
    print("机构级 Master Prompt（可直接复制）：")
    print("-" * 78)
    print(generate_master_prompt_payload(row, name, code))
    print("=" * 78)


# ================================================================
# 7. TRADING PRIORITY
# ================================================================

def calculate_trading_priority(row, timing_code=None):
    if timing_code is None:
        timing_code = ai_timing_engine(row)[0]

    stage = get_stage_key(row.get("阶段", "Stage 0"))
    permission = row.get("Trading_Permission", "YELLOW")
    stock_q = float(row.get("Stock_Quality_Score", row.get("Final_Score", 0)) or 0)
    entry_q = float(row.get("Entry_Quality_Score", 0) or 0)
    action_level = ai_timing_engine(row)[2]

    if timing_code == "BUY_TODAY":
        if action_level == "建仓-半仓":
            return "★★★☆☆"
        if (stage == "Stage 1" and permission == "GREEN"
                and stock_q >= 90 and entry_q >= 80):
            return "★★★★★"
        return "★★★★☆"

    if timing_code == "BUY_TOMORROW":
        return "★★★★☆"

    if timing_code in (
        "BUY_ON_VOLUME",
        "BUY_ON_EMA5_RECLAIM",
        "BUY_ON_PULLBACK_EMA20",
        "BUY_ON_BREAKOUT",
    ):
        return "★★★☆☆"

    if timing_code == "HOLD":
        return "★★☆☆☆"

    return "★☆☆☆☆"


# ================================================================
# 8. EXCEL FORMATTING
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

            # 自动列宽
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                if is_dashboard:
                    ws.column_dimensions[col_letter].width = min(max(max_len + 2, 14), 60)
                else:
                    ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)

            # 表头加粗
            if not is_dashboard:
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # 仪表盘：A1 加大加粗
        if SHEET_DASHBOARD in wb.sheetnames:
            ws = wb[SHEET_DASHBOARD]
            ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
            ws["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            ws["A2"].font = Font(bold=True, size=11, color="1F4E78")

        wb.save(filename)
    except Exception as e:
        print(f"⚠️ Excel格式化失败，但文件已生成：{e}")


# ================================================================
# 9. OUTPUT ROW
# ================================================================

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
        "EQ_EMA20位置(分)": row.get("EQ_EMA20_Position", np.nan),
        "EQ_EMA5金叉(分)": row.get("EQ_EMA5_Cross", np.nan),
        "EQ_ATR距离(分)": row.get("EQ_ATR_Distance", np.nan),
        "EQ_突破年龄(分)": row.get("EQ_Breakout_Age", np.nan),
        "EQ_量能(分)": row.get("EQ_Volume", np.nan),
        "EMA5金叉年龄(天)": row.get("EMA5_Cross_Age", np.nan),
        "突破年龄(天)": row.get("Breakout_Age", np.nan),

        "Top_Warning_Score": row.get("Top_Warning_Score", 0),
        "Top_Warning_Reasons": row.get("Top_Warning_Reasons", ""),

        "Trading Permission": row["Trading_Permission"],
        "阶段": row["阶段"],
        "最终信号": row["最终信号"],

        "Entry Engine": row.get("Entry_Engine", "N/A"),
        "Exit Engine": row.get("Exit_Engine", "N/A"),

        # ---- V4.2 Setup 字段 ----
        "Setup_Type": row.get("Setup_Type", "NONE"),
        "Setup_Age": row.get("Setup_Age", 999),
        "Setup_Active": bool(row.get("Setup_Active", False)),
        "Breakout_Confirmed": bool(row.get("Breakout_Confirmed", False)),
        "Setup_Breakout_Level": (
            round(row["Setup_Breakout_Level"], 4)
            if pd.notna(row.get("Setup_Breakout_Level", np.nan)) else np.nan
        ),

        "AI_结论": ai_timing_text(timing_code),
        "AI_时机": timing_code,
        "AI_动作": action_level,
        "AI_原因": reason,
        "AI_优先级": priority,
        "AI_减仓比例": reduce_ratio,

        "AI_参考买入区": buy_zone,
        "AI_仓位建议": position,
        "结构防守价": round(row["结构防守价"], 4) if pd.notna(row["结构防守价"]) else np.nan,
        "ATR防守参考": round(row["ATR_Stop_2x"], 4) if pd.notna(row["ATR_Stop_2x"]) else np.nan,
        "AI_失效条件": invalidation,
        "AI Recommendation": generate_ai_recommendation(row),
        "AI 自然语言摘要": generate_ai_summary(row, code, name),
        "持有原则": row["持有原则"],

        # ---- V4.2 新增字段 ----
        "AI_Confidence": round(_get_ai_confidence(row), 2),
        "Long_Trend_OK": bool(row.get("Long_Trend_OK", True)),
        "Entry Engine(原始)": row.get("Entry_Engine_Raw", "N/A"),
        "确认因子数": row.get("BUY_Confirmations", np.nan),
        "ATR动态倍数": round(row.get("ATR_Multiplier", np.nan), 2) if pd.notna(row.get("ATR_Multiplier", np.nan)) else np.nan,
        "ATR动态防守价": round(row.get("ATR_Stop_Dynamic", np.nan), 4) if pd.notna(row.get("ATR_Stop_Dynamic", np.nan)) else np.nan,
        "涨跌幅1日%": round(row.get("Return_1", np.nan), 2) if pd.notna(row.get("Return_1", np.nan)) else np.nan,
        "涨跌幅5日%": round(row.get("Return_5", np.nan), 2) if pd.notna(row.get("Return_5", np.nan)) else np.nan,
        "成交量变化1日%": round(row.get("Volume_Change_1", np.nan), 1) if pd.notna(row.get("Volume_Change_1", np.nan)) else np.nan,
        "成交量变化5日%": round(row.get("Volume_Change_5", np.nan), 1) if pd.notna(row.get("Volume_Change_5", np.nan)) else np.nan,
        "价量关系_1日": row.get("PV_关系_1日", "N/A"),
        "价量关系_5日": row.get("PV_关系_5日", "N/A"),
        "OBV": round(row.get("OBV", np.nan), 0) if pd.notna(row.get("OBV", np.nan)) else np.nan,
        "OBV变化5日": round(row.get("OBV_Change_5", np.nan), 0) if pd.notna(row.get("OBV_Change_5", np.nan)) else np.nan,
        "OBV背离判断": classify_obv_divergence(row),

        "收盘价": round(row["Close"], 4),
        "EMA5": round(row["EMA5"], 4) if pd.notna(row.get("EMA5", np.nan)) else np.nan,
        "EMA5扣底": round(row["Deduct5"], 4) if pd.notna(row.get("Deduct5", np.nan)) else np.nan,
        "EMA5扣底状态": row.get("EMA5_扣底状态", "N/A"),
        "EMA5价格距离": round(row["Gap5"], 4) if pd.notna(row.get("Gap5", np.nan)) else np.nan,

        "EMA20": round(row["EMA20"], 4) if pd.notna(row["EMA20"]) else np.nan,
        "EMA20扣底": round(row["Deduct20"], 4) if pd.notna(row.get("Deduct20", np.nan)) else np.nan,
        "EMA20扣底状态": row.get("EMA20_扣底状态", "N/A"),
        "EMA20价格距离": round(row["Gap20"], 4) if pd.notna(row.get("Gap20", np.nan)) else np.nan,

        "EMA60": round(row["EMA60"], 4) if pd.notna(row["EMA60"]) else np.nan,
        "EMA60扣底": round(row["Deduct60"], 4) if pd.notna(row.get("Deduct60", np.nan)) else np.nan,
        "EMA60扣底状态": row.get("EMA60_扣底状态", "N/A"),
        "EMA60价格距离": round(row["Gap60"], 4) if pd.notna(row.get("Gap60", np.nan)) else np.nan,

        "EMA200": round(row["EMA200"], 4) if pd.notna(row["EMA200"]) else np.nan,
        "EMA200 20日斜率": round(row.get("EMA200_Actual_Slope_20", np.nan), 4) if pd.notna(row.get("EMA200_Actual_Slope_20", np.nan)) else np.nan,
        "EMA20斜率5日": round(row["EMA20_Actual_Slope_5"], 4) if pd.notna(row.get("EMA20_Actual_Slope_5", np.nan)) else np.nan,
        "EMA60斜率5日": round(row["EMA60_Actual_Slope_5"], 4) if pd.notna(row.get("EMA60_Actual_Slope_5", np.nan)) else np.nan,

        "ADX": round(row["ADX"], 2) if pd.notna(row["ADX"]) else np.nan,
        "+DI": round(row["+DI"], 2) if pd.notna(row["+DI"]) else np.nan,
        "-DI": round(row["-DI"], 2) if pd.notna(row["-DI"]) else np.nan,
        "DI Spread": round(row["DI_Spread"], 2) if pd.notna(row["DI_Spread"]) else np.nan,
        "ADX变化5日": round(row["ADX_Change_5"], 2) if pd.notna(row["ADX_Change_5"]) else np.nan,
        "ADX 5日趋势": row["ADX_5日趋势"],
        "ADX强度": row["ADX_强度"],
        "Volume": int(row["Volume"]) if pd.notna(row["Volume"]) else np.nan,
        "Volume Ratio": round(row["Volume_Ratio"], 2) if pd.notna(row["Volume_Ratio"]) else np.nan,
        "突破量确认": "✅" if bool(row["Breakout_Volume"]) else "—",
        "Close_Position_Range": round(row["Close_Position_Range"], 4) if pd.notna(row.get("Close_Position_Range", np.nan)) else np.nan,
        "EMA20乖离率": round(row["Extension_EMA20_Pct"], 2) if pd.notna(row["Extension_EMA20_Pct"]) else np.nan,
        "ATR Distance": round(row.get("ATR_Distance", np.nan), 2) if pd.notna(row.get("ATR_Distance", np.nan)) else np.nan,
        "ATR Health Score": round(row.get("ATR_Health_Score", np.nan), 2) if pd.notna(row.get("ATR_Health_Score", np.nan)) else np.nan,
        "ATR Risk Score": round(row.get("ATR_Risk_Score", np.nan), 2) if pd.notna(row.get("ATR_Risk_Score", np.nan)) else np.nan,
        "ATR Risk": row.get("ATR_Risk", "数据不足"),
        "ATR%": round(row["ATR_Pct"], 2) if pd.notna(row["ATR_Pct"]) else np.nan,
        "OBV趋势": row["OBV_趋势"],
        "OBV 5日趋势": row["OBV_5日趋势"],
        "BB位置": row["BB_位置"],

        "Price Structure": row["Price_Structure_Score"],
        "Money Flow": row["Money_Flow_Score"],
        "Trend Momentum": row["Trend_Momentum_Score"],
        "Candlestick Momentum": row["Candlestick_Momentum_Score"],
        "Risk Control": row["Risk_Control_Score"],
        "Raw Score": row.get("Raw_Score", np.nan),
        "Stage Adjustment": row.get("Stage_Adjustment", np.nan),
        "Final Score": row.get("Final_Score", np.nan),
        "机构级MasterPrompt": generate_master_prompt_payload(row, name, code),
    }


# ================================================================
# 10. TWO-LAYER REPORT BUILDERS
# ================================================================

def build_dashboard_df(df_results):
    """构建仪表盘数据框（两列：项目 / 值）"""
    if df_results is None or df_results.empty:
        return pd.DataFrame([{"项目": "提示", "值": "无数据"}])

    df = df_results.copy()

    def _count_action(action_label):
        return int((df["综合建议"].apply(classify_action) == action_label).sum())

    today_buy = _count_action("今日可买")
    tomorrow = _count_action("明日确认")
    wait_vol = _count_action("等放量")
    wait_pull = _count_action("等回EMA20")
    wait_ema5 = _count_action("等EMA5")
    wait_break = _count_action("等突破")
    hold_no_add = _count_action("持有不加仓")
    no_buy = _count_action("不买")

    wait_total = wait_vol + wait_pull + wait_ema5 + wait_break

    pos_risk = int(df["Exit Engine"].isin(["EXIT", "保护利润", "减仓"]).sum()) if "Exit Engine" in df.columns else 0

    stage_series = df.get("阶段", pd.Series([], dtype=str)).astype(str)
    stage_0 = int(stage_series.str.contains("Stage 0", na=False).sum())
    stage_1 = int(stage_series.str.contains("Stage 1", na=False).sum())
    stage_2 = int(stage_series.str.contains("Stage 2", na=False).sum())
    stage_3 = int(stage_series.str.contains("Stage 3", na=False).sum())
    stage_3a = int(stage_series.str.contains("Stage 3A", na=False).sum())
    stage_3b = int(stage_series.str.contains("Stage 3B", na=False).sum())
    stage_4 = int(stage_series.str.contains("Stage 4", na=False).sum())

    perm_series = df.get("Trading Permission", pd.Series([], dtype=str)).astype(str)
    perm_green = int((perm_series == "GREEN").sum())
    perm_yellow = int((perm_series == "YELLOW").sum())
    perm_red = int((perm_series == "RED").sum())

    top_series = pd.to_numeric(df.get("Top_Warning_Score", 0), errors="coerce").fillna(0)
    warn_high = int((top_series >= CONFIG["TOP_WARNING_REDUCE_2"]).sum())
    warn_mid = int(((top_series >= CONFIG["TOP_WARNING_REDUCE_1"]) & (top_series < CONFIG["TOP_WARNING_REDUCE_2"])).sum())

    # Setup 统计
    setup_active_cnt = 0
    setup_pullback_cnt = 0
    setup_breakout_cnt = 0
    if "Setup_Active" in df.columns:
        setup_active_cnt = int(df["Setup_Active"].astype(bool).sum())
    if "Setup_Type" in df.columns:
        setup_pullback_cnt = int((df["Setup_Type"] == "EMA20_PULLBACK").sum())
        setup_breakout_cnt = int((df["Setup_Type"] == "BREAKOUT").sum())

    rows = [
        {"项目": "📊 报告日期", "值": datetime.now().strftime("%Y-%m-%d")},
        {"项目": "📊 扫描总数", "值": len(df)},
        {"项目": "", "值": ""},
        {"项目": "=== 行动卡片 ===", "值": ""},
        {"项目": "🟢 今日可买", "值": today_buy},
        {"项目": "🟢 明日确认", "值": tomorrow},
        {"项目": "🔵 等触发（合计）", "值": wait_total},
        {"项目": "   └─ 等放量", "值": wait_vol},
        {"项目": "   └─ 等回EMA20", "值": wait_pull},
        {"项目": "   └─ 等EMA5", "值": wait_ema5},
        {"项目": "   └─ 等突破", "值": wait_break},
        {"项目": "🟡 持有不加仓", "值": hold_no_add},
        {"项目": "⚠️ 持仓风险（EXIT/保护利润/减仓）", "值": pos_risk},
        {"项目": "🚫 不买", "值": no_buy},
        {"项目": "", "值": ""},
        {"项目": "=== Setup 记忆（V4.2）===", "值": ""},
        {"项目": "活跃 Setup 数", "值": setup_active_cnt},
        {"项目": "   └─ EMA20 回调 Setup", "值": setup_pullback_cnt},
        {"项目": "   └─ Breakout Setup", "值": setup_breakout_cnt},
        {"项目": "", "值": ""},
        {"项目": "=== Stage 分布 ===", "值": ""},
        {"项目": "Stage 0 (观察)", "值": stage_0},
        {"项目": "Stage 1 (初升)", "值": stage_1},
        {"项目": "Stage 2 (主升)", "值": stage_2},
        {"项目": "Stage 3 (末升，合计)", "值": stage_3},
        {"项目": "   └─ 3A 健康调整", "值": stage_3a},
        {"项目": "   └─ 3B 趋势恶化", "值": stage_3b},
        {"项目": "Stage 4 (趋势破坏)", "值": stage_4},
        {"项目": "", "值": ""},
        {"项目": "=== Permission 分布 ===", "值": ""},
        {"项目": "GREEN", "值": perm_green},
        {"项目": "YELLOW", "值": perm_yellow},
        {"项目": "RED", "值": perm_red},
        {"项目": "", "值": ""},
        {"项目": "=== 顶部预警分布 ===", "值": ""},
        {"项目": "预警 ≥70（再减1/3）", "值": warn_high},
        {"项目": "预警 45–69（减1/3）", "值": warn_mid},
    ]

    # 今日执行清单
    exec_df = df[df["综合建议"].apply(classify_action).isin(["今日可买", "明日确认"])]
    if not exec_df.empty:
        rows.append({"项目": "", "值": ""})
        rows.append({"项目": "=== 今日执行清单 ===", "值": ""})
        for _, r in exec_df.iterrows():
            label = f"{r.get('名称','')} ({r.get('代码','')})"
            detail = f"{r.get('综合建议','')} | {r.get('AI_参考买入区','')} | {r.get('AI_仓位建议','')} | 防守 {r.get('结构防守价','')}"
            rows.append({"项目": label, "值": detail})

    # 持仓风险清单
    if "Exit Engine" in df.columns:
        risk_df = df[df["Exit Engine"].isin(["EXIT", "保护利润", "减仓"])]
        if not risk_df.empty:
            rows.append({"项目": "", "值": ""})
            rows.append({"项目": "=== 持仓风险清单 ===", "值": ""})
            for _, r in risk_df.iterrows():
                label = f"{r.get('名称','')} ({r.get('代码','')})"
                detail = f"{r.get('Exit Engine','')} | {r.get('AI_减仓比例','')} | 收盘 {r.get('收盘价','')} | 防守 {r.get('结构防守价','')}"
                rows.append({"项目": label, "值": detail})

    return pd.DataFrame(rows)


def build_trading_overview_df(df_results):
    """①今日交易总览：瘦身行动字段 + 分类 + 买入区拆分"""
    if df_results is None or df_results.empty:
        return pd.DataFrame()

    df = df_results.copy()
    df["_order"] = df["AI_时机"].map(TIMING_PRIORITY_ORDER).fillna(99)
    df["_score"] = pd.to_numeric(df.get("股票质量分", 0), errors="coerce").fillna(0)

    df = df.sort_values(["_order", "_score"], ascending=[True, False]).reset_index(drop=True)
    df.insert(0, "排名", range(1, len(df) + 1))

    # 分类
    df["分类"] = df["综合建议"].apply(classify_action)

    # 拆分买入区
    zone_low_high = df["AI_参考买入区"].apply(lambda s: pd.Series(extract_price_range(s)))
    df["买入区下沿"] = zone_low_high[0]
    df["买入区上沿"] = zone_low_high[1]

    close_num = pd.to_numeric(df.get("收盘价", np.nan), errors="coerce")
    up = pd.to_numeric(df["买入区上沿"], errors="coerce")
    down = pd.to_numeric(df["买入区下沿"], errors="coerce")
    df["距上沿%"] = ((close_num - up) / up * 100).round(2)
    df["距下沿%"] = ((close_num - down) / down * 100).round(2)

    cols = [
        "排名", "分类", "名称", "代码", "股票类型",
        "综合建议",
        "阶段", "Trading Permission",
        "最终信号", "AI_时机", "AI_动作", "AI_优先级",
        "Setup_Type", "Setup_Age", "Setup_Active",
        "收盘价",
        "买入区下沿", "买入区上沿", "距上沿%", "距下沿%",
        "AI_仓位建议",
        "结构防守价", "ATR防守参考",
        "AI_失效条件",
        "交易优先级",
        "股票质量分", "股票质量评级",
        "买点质量分", "买点质量评级",
        "Top_Warning_Score", "Top_Warning_Reasons",
    ]
    cols = [c for c in cols if c in df.columns]
    return df[cols].copy()


def build_watchlist_df(df_results):
    """②今日候选名单：加入触发条件 / 当前状态 / 距触发%"""
    if df_results is None or df_results.empty:
        return pd.DataFrame()

    wait_map = {
        "BUY_ON_VOLUME": "等量放大",
        "WAIT_VOL": "等量放大",
        "BUY_ON_EMA5_RECLAIM": "EMA5 重新站上",
        "WAIT_EMA5": "EMA5 重新站上",
        "BUY_ON_PULLBACK_EMA20": "回 EMA20",
        "WAIT_PULLBACK": "回 EMA20",
        "BUY_ON_BREAKOUT": "突破买",
    }

    records = []
    for _, r in df_results.iterrows():
        timing = r.get("AI_时机", "NO_BUY")
        entry = r.get("Entry Engine", "NONE")

        reason = None
        if timing in wait_map:
            reason = wait_map[timing]
        elif entry in wait_map:
            reason = wait_map[entry]

        if reason is None:
            continue

        close = r.get("收盘价", np.nan)
        ema5 = r.get("EMA5", np.nan)
        ema20 = r.get("EMA20", np.nan)
        vr = r.get("Volume Ratio", np.nan)
        breakout_ok = r.get("突破量确认", "—")

        # 触发条件
        if reason == "等量放大":
            trigger = f"Volume Ratio ≥ {CONFIG['WATCH_VOL_TRIGGER']:.2f} 且 突破量确认=✅"
        elif reason == "回 EMA20":
            trigger = "收盘价回到 EMA20 ±2% 内"
        elif reason == "EMA5 重新站上":
            trigger = "收盘价 ≥ EMA5 且 EMA5 ≥ EMA20"
        elif reason == "突破买":
            trigger = "突破前高且放量"
        else:
            trigger = "—"

        # 当前状态
        status = "等待"
        dist_trigger = np.nan
        try:
            if reason == "等量放大" and pd.notna(vr):
                if float(vr) >= CONFIG["WATCH_VOL_TRIGGER"] and breakout_ok == "✅":
                    status = "已触发"
                elif float(vr) >= CONFIG["WATCH_VOL_NEAR"]:
                    status = "接近"
                dist_trigger = (CONFIG["WATCH_VOL_TRIGGER"] / float(vr) - 1.0) * 100
            elif reason == "回 EMA20" and pd.notna(close) and pd.notna(ema20) and ema20 > 0:
                d = (float(close) - float(ema20)) / float(ema20)
                dist_trigger = d * 100
                if abs(d) <= CONFIG["WATCH_EMA20_NEAR_PCT"]:
                    status = "已触发"
                elif abs(d) <= CONFIG["WATCH_EMA20_FAR_PCT"]:
                    status = "接近"
            elif reason == "EMA5 重新站上" and pd.notna(close) and pd.notna(ema5):
                if float(close) >= float(ema5):
                    status = "已触发"
                else:
                    status = "等待"
            elif reason == "突破买":
                status = "等待"
        except Exception:
            pass

        # 拆分买入区
        zone = r.get("AI_参考买入区", "")
        low, high = extract_price_range(zone)

        records.append({
            "日期": r.get("日期", ""),
            "名称": r.get("名称", ""),
            "代码": r.get("代码", ""),
            "股票类型": r.get("股票类型", ""),
            "综合建议": r.get("综合建议", ""),
            "等什么": reason,
            "触发条件": trigger,
            "当前状态": status,
            "距触发%": round(dist_trigger, 2) if pd.notna(dist_trigger) else np.nan,
            "收盘价": close,
            "EMA5": ema5,
            "EMA20": ema20,
            "Volume Ratio": vr,
            "突破量确认": breakout_ok,
            "阶段": r.get("阶段", ""),
            "最终信号": r.get("最终信号", ""),
            "Setup_Type": r.get("Setup_Type", "NONE"),
            "Setup_Age": r.get("Setup_Age", 999),
            "Setup_Active": bool(r.get("Setup_Active", False)),
            "股票质量分": r.get("股票质量分", np.nan),
            "买点质量分": r.get("买点质量分", np.nan),
            "买点质量评级": r.get("买点质量评级", ""),
            "Top_Warning_Score": r.get("Top_Warning_Score", 0),
            "Top_Warning_Reasons": r.get("Top_Warning_Reasons", ""),
            "AI时机": timing,
            "Entry Engine": entry,
            "交易优先级": r.get("交易优先级", ""),
            "买入区下沿": low,
            "买入区上沿": high,
            "AI_参考买入区": zone,
            "AI_仓位建议": r.get("AI_仓位建议", ""),
            "结构防守价": r.get("结构防守价", np.nan),
        })

    df = pd.DataFrame(records)
    if not df.empty:
        status_order = {"已触发": 1, "接近": 2, "等待": 3}
        df["_s"] = df["当前状态"].map(status_order).fillna(9)
        df["_order"] = df["AI时机"].map(TIMING_PRIORITY_ORDER).fillna(99)
        df["_score"] = pd.to_numeric(df["买点质量分"], errors="coerce").fillna(0)
        df = df.sort_values(["_s", "_order", "_score"], ascending=[True, True, False])
        df = df.drop(columns=["_s", "_order", "_score"]).reset_index(drop=True)
    return df


def build_position_management_df(df_results):
    """③持仓管理：加入盈亏/距止损/仓位占比（默认空仓）"""
    if df_results is None or df_results.empty:
        return pd.DataFrame()

    def ab(price, ema):
        if pd.isna(price) or pd.isna(ema):
            return "N/A"
        return "Above" if float(price) >= float(ema) else "Below"

    def action_of(exit_e, top_score):
        if top_score >= CONFIG["TOP_WARNING_EXIT"]:
            return "顶部预警清仓"
        if top_score >= CONFIG["TOP_WARNING_REDUCE_2"]:
            return "顶部预警减1/3"
        if top_score >= CONFIG["TOP_WARNING_REDUCE_1"]:
            return "顶部预警减1/3"
        if exit_e == "EXIT":
            return "全部退出"
        if exit_e == "保护利润":
            return "减仓锁利(1/3) + EMA5/ATR移动止损"
        if exit_e == "减仓":
            return "减仓(再1/3) + EMA5/ATR移动止损"
        if exit_e == "观察":
            return "观察"
        if exit_e == "持有":
            return "继续持有"
        return "—"

    records = []
    for _, r in df_results.iterrows():
        close = r.get("收盘价", np.nan)
        ema5 = r.get("EMA5", np.nan)
        ema20 = r.get("EMA20", np.nan)
        ema60 = r.get("EMA60", np.nan)
        exit_e = r.get("Exit Engine", "持有")
        top_score = float(r.get("Top_Warning_Score", 0) or 0)
        structure_stop = r.get("结构防守价", np.nan)
        atr_stop = r.get("ATR防守参考", np.nan)
        moving_stop = np.nan
        try:
            candidates = [x for x in [structure_stop, atr_stop] if pd.notna(x) and x > 0]
            if candidates:
                moving_stop = max(candidates)
        except Exception:
            pass

        # 距止损%
        dist_stop_pct = np.nan
        if pd.notna(close) and pd.notna(moving_stop) and float(close) > 0:
            dist_stop_pct = (float(close) - float(moving_stop)) / float(close) * 100

        records.append({
            "日期": r.get("日期", ""),
            "名称": r.get("名称", ""),
            "代码": r.get("代码", ""),
            "股票类型": r.get("股票类型", ""),
            "综合建议": r.get("综合建议", ""),
            "阶段": r.get("阶段", ""),
            "Trading Permission": r.get("Trading Permission", ""),
            "收盘价": close,
            "EMA5": ema5,
            "EMA20": ema20,
            "EMA60": ema60,
            "EMA5关系": ab(close, ema5),
            "EMA20关系": ab(close, ema20),
            "EMA60关系": ab(close, ema60),
            "Exit Engine": exit_e,
            "Top_Warning_Score": top_score,
            "Top_Warning_Reasons": r.get("Top_Warning_Reasons", ""),
            "今日动作": action_of(exit_e, top_score),
            "AI_减仓比例": r.get("AI_减仓比例", ""),
            "结构防守价": structure_stop,
            "ATR防守参考": atr_stop,
            "移动止损参考": moving_stop,
            "距止损%": round(dist_stop_pct, 2) if pd.notna(dist_stop_pct) else np.nan,
            # 用户持仓区（可手动填写）
            "持仓数量": 0,
            "成本价": np.nan,
            "市值": np.nan,
            "盈亏": np.nan,
            "盈亏%": np.nan,
            "仓位占比": np.nan,
            "持有原则": r.get("持有原则", ""),
        })
    return pd.DataFrame(records)


def build_decision_summary_df(df_results):
    if df_results is None or df_results.empty:
        return pd.DataFrame()

    part0 = ["交易优先级", "综合建议"]
    part1 = ["名称", "代码", "日期", "股票类型"]
    part2 = ["股票质量分", "股票质量评级", "买点质量分", "买点质量评级"]
    part3 = ["Trading Permission", "阶段", "Final Score", "最终信号"]
    part4 = ["Top_Warning_Score", "Top_Warning_Reasons"]
    part5 = ["Entry Engine", "Exit Engine",
             "Setup_Type", "Setup_Age", "Setup_Active", "Breakout_Confirmed",
             "AI_时机", "AI_动作", "AI_原因", "AI_优先级", "AI_减仓比例"]
    part6 = ["AI_参考买入区", "AI_仓位建议", "结构防守价", "ATR防守参考", "AI_失效条件"]
    part7 = ["持有原则"]
    part8 = [
        "EMA5", "EMA20", "EMA60", "EMA200",
        "ADX", "+DI", "-DI",
        "Volume Ratio", "ATR%", "ATR Distance",
        "ATR Health Score", "ATR Risk Score",
        "EMA5金叉年龄(天)", "突破年龄(天)",
        "OBV趋势", "BB位置",
    ]
    part9 = [
        "EQ_EMA20位置(分)", "EQ_EMA5金叉(分)", "EQ_ATR距离(分)",
        "EQ_突破年龄(分)", "EQ_量能(分)",
    ]
    part10 = [
        "Price Structure", "Money Flow", "Trend Momentum",
        "Risk Control", "Stage Adjustment", "Raw Score",
    ]

    ordered = part0 + part1 + part2 + part3 + part4 + part5 + part6 + part7 + part8 + part9 + part10
    existing = [c for c in ordered if c in df_results.columns]
    df = df_results[existing].copy()

    if "AI_时机" in df_results.columns:
        order = df_results["AI_时机"].map(TIMING_PRIORITY_ORDER).fillna(99)
        df = df.assign(_order=order).sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)
    return df


def build_ai_analysis_df(df_results):
    if df_results is None or df_results.empty:
        return pd.DataFrame()

    records = []
    for _, r in df_results.iterrows():
        records.append({
            "股票": f"{r.get('名称','')} ({r.get('代码','')})",
            "股票类型": r.get("股票类型", ""),
            "综合建议": r.get("综合建议", ""),
            "股票质量分": r.get("股票质量分", np.nan),
            "股票质量评级": r.get("股票质量评级", ""),
            "买点质量分": r.get("买点质量分", np.nan),
            "买点质量评级": r.get("买点质量评级", ""),
            "Stage": r.get("阶段", ""),
            "最终信号": r.get("最终信号", ""),
            "Top_Warning_Score": r.get("Top_Warning_Score", 0),
            "Top_Warning_Reasons": r.get("Top_Warning_Reasons", ""),
            "Entry Engine": r.get("Entry Engine", ""),
            "Exit Engine": r.get("Exit Engine", ""),
            "Setup_Type": r.get("Setup_Type", "NONE"),
            "Setup_Age": r.get("Setup_Age", 999),
            "Setup_Active": r.get("Setup_Active", False),
            "Breakout_Confirmed": r.get("Breakout_Confirmed", False),
            "AI 结论": r.get("AI_结论", ""),
            "AI 时机": r.get("AI_时机", ""),
            "AI 动作": r.get("AI_动作", ""),
            "AI 原因": r.get("AI_原因", ""),
            "AI 参考买入区": r.get("AI_参考买入区", ""),
            "AI 仓位建议": r.get("AI_仓位建议", ""),
            "AI 减仓比例": r.get("AI_减仓比例", ""),
            "AI 优先级": r.get("AI_优先级", ""),
            "AI 失效条件": r.get("AI_失效条件", ""),
            "交易优先级": r.get("交易优先级", ""),
            "Raw Score": r.get("Raw Score", np.nan),
            "Stage Adjustment": r.get("Stage Adjustment", np.nan),
            "Final Score": r.get("Final Score", np.nan),
            "ATR Distance": r.get("ATR Distance", np.nan),
            "ATR Health Score": r.get("ATR Health Score", np.nan),
            "ATR Risk Score": r.get("ATR Risk Score", np.nan),
            "ATR Risk": r.get("ATR Risk", ""),
            "AI Recommendation": r.get("AI Recommendation", ""),
            "AI 自然语言摘要": r.get("AI 自然语言摘要", ""),
            "AI 信心": r.get("AI_Confidence", np.nan),
            "确认因子数": r.get("确认因子数", np.nan),
            "+DI": r.get("+DI", np.nan),
            "-DI": r.get("-DI", np.nan),
            "价量关系_1日": r.get("价量关系_1日", ""),
            "价量关系_5日": r.get("价量关系_5日", ""),
            "OBV背离判断": r.get("OBV背离判断", ""),
            "机构级MasterPrompt": r.get("机构级MasterPrompt", ""),
        })
    return pd.DataFrame(records)


# ================================================================
# 11. SINGLE STOCK
# ================================================================

def single_stock_scan(code, output_excel=True):
    print(f"\n🔍 正在扫描：{code}")

    df_raw, name = safe_fetch_stock(code)
    if df_raw is None:
        return None

    data = calculate_engine(df_raw)
    if data is None or data.empty:
        return None

    row = data.iloc[-1]
    print_terminal_report(row, code, name)

    result = latest_output_dict(row, name, code)

    if output_excel:
        filename = (
            f"{safe_filename(name)}_EMA_V4.2_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        df_summary = pd.DataFrame([result])

        dashboard_df = build_dashboard_df(df_summary)
        overview_df = build_trading_overview_df(df_summary)
        watchlist_df = build_watchlist_df(df_summary)
        position_df = build_position_management_df(df_summary)
        decision_df = build_decision_summary_df(df_summary)
        ai_df = build_ai_analysis_df(df_summary)

        tail = data.tail(60).copy()
        tail_output = tail.reset_index().copy()
        tail_output["日期"] = pd.to_datetime(tail_output["Date"]).dt.strftime("%Y-%m-%d")
        keep_cols = [
            "日期", "Close", "Stock_Type",
            "EMA5", "Deduct5", "Deduct5_Change", "EMA5_扣底状态",
            "Gap5", "Gap5_Change",
            "EMA20", "Deduct20", "Deduct20_Change", "EMA20_扣底状态",
            "Gap20", "Gap20_Change",
            "EMA60", "Deduct60", "Deduct60_Change", "EMA60_扣底状态",
            "Gap60", "Gap60_Change",
            "EMA200", "EMA200_Actual_Slope_20",
            "Trading_Permission", "阶段", "最终信号",
            "Entry_Engine", "Exit_Engine",
            "Setup_Type", "Setup_Age", "Setup_Active", "Breakout_Confirmed",
            "Stock_Quality_Score", "Entry_Quality_Score", "Entry_Quality_Rating",
            "EQ_EMA20_Position", "EQ_EMA5_Cross", "EQ_ATR_Distance",
            "EQ_Breakout_Age", "EQ_Volume",
            "EMA5_Cross_Age", "Breakout_Age",
            "Top_Warning_Score", "Top_Warning_Reasons",
            "AI_Confidence_Score", "Long_Trend_OK", "Entry_Engine_Raw", "BUY_Confirmations",
            "ATR_Multiplier", "ATR_Stop_Dynamic",
            "ADX", "+DI", "-DI", "DI_Spread", "ADX_Change_5",
            "Volume_Ratio", "Breakout_Volume",
            "Return_1", "Return_5", "Volume_Change_1", "Volume_Change_5", "PV_关系_1日", "PV_关系_5日",
            "OBV", "OBV_Change_5", "OBV_MA_Distance_Pct",
            "Extension_EMA20_Pct",
            "ATR_Distance", "ATR_Distance_Abs", "ATR_Pct",
            "ATR_Health_Score", "ATR_Risk_Score",
            "ATR_Risk",
            "OBV_趋势", "BB_位置",
            "Raw_Score", "Stage_Adjustment", "Final_Score",
        ]
        tail_output = tail_output[[c for c in keep_cols if c in tail_output.columns]]

        signal_log = tail_output.copy()
        signal_log["上一日信号"] = signal_log["最终信号"].shift(1)
        signal_log["是否变化"] = signal_log["最终信号"] != signal_log["上一日信号"]
        signal_log = signal_log[signal_log["是否变化"] == True]

        all_indicators = data.copy().reset_index()
        if "Date" in all_indicators.columns:
            all_indicators["日期"] = pd.to_datetime(all_indicators["Date"]).dt.strftime("%Y-%m-%d")
            all_indicators = all_indicators.drop(columns=["Date"])
        if "日期" in all_indicators.columns:
            cols = ["日期"] + [c for c in all_indicators.columns if c != "日期"]
            all_indicators = all_indicators[cols]

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
        print(f"\n✅ Excel 已输出：{filename}")

    return result


# ================================================================
# 12. LOAD STOCK POOL
# ================================================================

def load_stock_pool_from_excel(filepath):
    try:
        df = pd.read_excel(filepath)

        possible_code_cols = [
            "Code", "code", "CODE",
            "Stock", "stock", "STOCK",
            "代码", "股票代码", "股号",
            "Ticker", "ticker", "Yahoo Ticker",
            "Yahoo Finance"
        ]
        code_col = next((c for c in possible_code_cols if c in df.columns), None)
        if code_col is None:
            code_col = df.columns[0]
            print(f"⚠️ 未找到标准代码栏位，使用第一列：{code_col}")

        name_col = None
        for c in ["Name", "name", "名称", "股票简称", "Short Name"]:
            if c in df.columns:
                name_col = c
                break

        pool = []
        for _, r in df.iterrows():
            c = r[code_col]
            if pd.isna(c):
                continue
            c = clean_stock_code(c)
            item = {"code": c}
            if name_col and pd.notna(r[name_col]):
                item["name"] = str(r[name_col]).strip()
            else:
                item["name"] = ""
            pool.append(item)

        print(f"✅ 从 Excel 加载 {len(pool)} 只股票")
        return pool
    except Exception as e:
        print(f"❌ Excel读取失败：{e}")
        return None


# ================================================================
# 13. BATCH SCAN
# ================================================================

def batch_scan(stock_pool, output_excel=True,
               min_price=None, max_price=None, min_volume=None):
    if min_price is None:
        min_price = CONFIG["DEFAULT_MIN_PRICE"]
    if max_price is None:
        max_price = CONFIG["DEFAULT_MAX_PRICE"]
    if min_volume is None:
        min_volume = CONFIG["DEFAULT_MIN_VOLUME"]

    print(f"\n📊 V4.2 批量扫描：{len(stock_pool)} 只")
    print(f"价格范围：RM{min_price:.2f} - RM{max_price:.2f}")
    print(f"最低成交量：{min_volume:,.0f}")

    results = []
    errors = []

    for idx, item in enumerate(stock_pool, 1):
        code = item["code"]
        print(f"[{idx}/{len(stock_pool)}] {code} ...", end=" ")

        try:
            df_raw, name = safe_fetch_stock(code)
            if df_raw is None:
                errors.append(code)
                print("无数据")
                continue

            data = calculate_engine(df_raw)
            if data is None or data.empty:
                errors.append(code)
                print("计算失败")
                continue

            row = data.iloc[-1]

            if pd.isna(row["Close"]):
                errors.append(code)
                print("无价格")
                continue

            if row["Close"] < min_price or row["Close"] > max_price:
                print("跳过")
                continue

            if pd.notna(row["Volume"]) and row["Volume"] < min_volume:
                print("低流动性")
                continue

            result = latest_output_dict(row, name, code)

            signal = row["最终信号"]
            stage = row["阶段"]
            permission = row["Trading_Permission"]
            stock_q = row.get("Stock_Quality_Score", 0)
            entry_q = row.get("Entry_Quality_Score", 0)
            entry = row.get("Entry_Engine", "NONE")
            exit_e = row.get("Exit_Engine", "持有")
            timing = result.get("AI_时机", "NO_BUY")
            stype = result.get("股票类型", "混合型")
            top_score = result.get("Top_Warning_Score", 0)

            if signal in ["买入", "加码"]:
                classification = "重点关注"
            elif signal in ["买入观察", "等待回调", "持有", "持有 / 等待", "减仓", "保护利润"]:
                classification = "观察"
            elif permission == "GREEN" and "Stage 1" in stage:
                classification = "候选"
            else:
                classification = "清除"

            result["分类"] = classification
            results.append(result)

            print(f"{classification} | {stype} | StockQ={stock_q:.0f} EntryQ={entry_q:.0f} | "
                  f"TopWarn={top_score:.0f} | "
                  f"{stage} | 信号={signal} | Entry={entry} | AI={timing}")

        except Exception as e:
            errors.append(code)
            print(f"错误：{e}")

    if not results:
        print("\n⚠️ 没有符合条件的股票")
        return None

    df_results = pd.DataFrame(results)

    df_results["_timing_order"] = df_results["AI_时机"].map(TIMING_PRIORITY_ORDER).fillna(99)
    df_results["_stock_q"] = pd.to_numeric(df_results.get("股票质量分", 0), errors="coerce").fillna(0)
    df_results["_entry_q"] = pd.to_numeric(df_results.get("买点质量分", 0), errors="coerce").fillna(0)

    class_order = {"重点关注": 1, "观察": 2, "候选": 3, "清除": 4}
    df_results["_class"] = df_results["分类"].map(class_order).fillna(9)

    df_results = df_results.sort_values(
        ["_timing_order", "_class", "_entry_q", "_stock_q"],
        ascending=[True, True, False, False]
    ).drop(columns=["_timing_order", "_class", "_stock_q", "_entry_q"]).reset_index(drop=True)

    print("\n" + "=" * 120)
    print("📋 V4.2 批量扫描结果")
    print("=" * 120)

    display_cols = [
        "代码", "名称", "股票类型", "收盘价",
        "综合建议", "分类",
        "股票质量分", "买点质量分", "买点质量评级",
        "Top_Warning_Score",
        "阶段", "Trading Permission",
        "最终信号",
        "Entry Engine", "Exit Engine",
        "Setup_Type", "Setup_Age",
        "AI_时机", "AI_动作", "AI_减仓比例",
        "EMA5金叉年龄(天)", "突破年龄(天)",
        "ADX", "Volume Ratio",
        "ATR Distance", "ATR Health Score",
        "交易优先级",
    ]
    print(df_results[[c for c in display_cols if c in df_results.columns]].to_string(index=False))

    if output_excel:
        filename = (
            f"EMA_V4.2_批量扫描_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        dashboard_df = build_dashboard_df(df_results)
        overview_df = build_trading_overview_df(df_results)
        watchlist_df = build_watchlist_df(df_results)
        position_df = build_position_management_df(df_results)
        decision_df = build_decision_summary_df(df_results)
        ai_df = build_ai_analysis_df(df_results)

        all_indicators_df = df_results.copy()
        if "分类" in all_indicators_df.columns:
            all_indicators_df = all_indicators_df.drop(columns=["分类"])

        focus = df_results[df_results["分类"] == "重点关注"]
        observe = df_results[df_results["分类"].isin(["观察", "候选"])]
        clear = df_results[df_results["分类"] == "清除"]

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            dashboard_df.to_excel(writer, sheet_name=SHEET_DASHBOARD, index=False)
            overview_df.to_excel(writer, sheet_name=SHEET_OVERVIEW, index=False)
            watchlist_df.to_excel(writer, sheet_name=SHEET_WATCHLIST, index=False)
            position_df.to_excel(writer, sheet_name=SHEET_POSITION, index=False)

            decision_df.to_excel(writer, sheet_name=SHEET_DECISION, index=False)
            ai_df.to_excel(writer, sheet_name=SHEET_AI, index=False)

            all_indicators_df.to_excel(writer, sheet_name=SHEET_ALL, index=False)

            df_results.to_excel(writer, sheet_name=SHEET_FULL_SCAN, index=False)
            if not focus.empty:
                focus.to_excel(writer, sheet_name=SHEET_FOCUS, index=False)
            if not observe.empty:
                observe.to_excel(writer, sheet_name=SHEET_OBSERVE, index=False)
            if not clear.empty:
                clear.to_excel(writer, sheet_name=SHEET_CLEAR, index=False)
            if errors:
                pd.DataFrame({"代码": errors}).to_excel(
                    writer, sheet_name=SHEET_ERRORS, index=False
                )

        format_excel_file(filename)
        print(f"\n✅ 批量报告：{filename}")

    return df_results


# ================================================================
# 14. BACKTRACK
# ================================================================

def simulate_position_tracking(df_back):
    position = 0.0
    entry_price = np.nan
    highest_price = np.nan
    trailing_stop = np.nan

    positions = []
    actions = []
    trailing_stops = []
    stop_triggered = []
    position_values = []

    for i, row in df_back.iterrows():
        close = row.get("收盘价", np.nan)
        atr = row.get("ATR14", np.nan)
        ema5 = row.get("EMA5", np.nan)
        ema20 = row.get("EMA20", np.nan)
        signal = row.get("最终信号", "观望")
        structure_stop = row.get("结构防守价", np.nan)
        stage = row.get("阶段", "")
        vr = row.get("Volume Ratio", np.nan)
        top_score = float(row.get("Top_Warning_Score", 0) or 0)

        action = "无"
        stop_hit = False

        if position > 0 and pd.notna(close):
            if pd.isna(highest_price):
                highest_price = close
            else:
                highest_price = max(highest_price, close)

            if pd.notna(atr) and CONFIG["TRAIL_STOP_ENABLED"]:
                atr_trail = highest_price - _trail_mult(stage) * atr
                candidates = [atr_trail]

                if CONFIG["TRAIL_STOP_USE_EMA5"] and pd.notna(ema5):
                    candidates.append(ema5)

                if CONFIG["TRAIL_STOP_USE_STRUCTURE"] and pd.notna(structure_stop):
                    candidates.append(structure_stop)

                new_trail = max(candidates)

                if pd.isna(trailing_stop):
                    trailing_stop = new_trail
                else:
                    trailing_stop = max(trailing_stop, new_trail)

        if position > 0 and pd.notna(trailing_stop) and pd.notna(close):
            if close < trailing_stop:
                stop_hit = True
                action = f"移动止损清仓（止损价 {trailing_stop:.4f}）"
                position = 0.0
                entry_price = np.nan
                highest_price = np.nan
                trailing_stop = np.nan

        if not stop_hit and position > 0 and CONFIG["TOP_WARNING_ENABLED"]:
            if top_score >= CONFIG["TOP_WARNING_EXIT"]:
                position = 0.0
                entry_price = np.nan
                highest_price = np.nan
                trailing_stop = np.nan
                action = f"顶部预警清仓（分数{top_score:.0f}）"
            elif top_score >= CONFIG["TOP_WARNING_REDUCE_2"]:
                reduce = min(CONFIG["PROTECT_PROFIT_REDUCE_PCT"], position)
                position -= reduce
                action = f"顶部预警减仓 {reduce:.0%}，剩余 {position:.0%}"
            elif top_score >= CONFIG["TOP_WARNING_REDUCE_1"]:
                reduce = min(CONFIG["PROTECT_PROFIT_REDUCE_PCT"], position)
                position -= reduce
                action = f"顶部预警减仓 {reduce:.0%}，剩余 {position:.0%}"

        if not stop_hit and action == "无":
            if signal == "买入" and position == 0:
                position = CONFIG["INITIAL_POSITION"]
                entry_price = close
                highest_price = close
                if pd.notna(atr):
                    trailing_stop = close - _trail_mult(stage) * atr
                action = f"建仓 {position:.0%}"

            elif signal == "加码" and 0 < position < 1.0:
                add = min(CONFIG["ADD_POSITION"], 1.0 - position)
                position += add
                if pd.isna(entry_price):
                    entry_price = close
                action = f"加仓 {add:.0%}，总仓位 {position:.0%}"

            elif signal == "减仓" and position > 0:
                reduce = min(CONFIG["PROTECT_PROFIT_REDUCE_PCT"], position)
                position -= reduce
                action = f"减仓 {reduce:.0%}，剩余 {position:.0%}"
                if position <= 0.001:
                    position = 0.0
                    entry_price = np.nan
                    highest_price = np.nan
                    trailing_stop = np.nan

            elif signal == "卖出" and position > 0:
                position = 0.0
                entry_price = np.nan
                highest_price = np.nan
                trailing_stop = np.nan
                action = "清仓"

            elif signal == "买入观察" and position == 0:
                re_ok = True
                if CONFIG["RE_ENTRY_REQUIRE_STAGE"]:
                    if not ("Stage 1" in stage or "Stage 2" in stage):
                        re_ok = False
                if CONFIG["RE_ENTRY_REQUIRE_VOL"]:
                    if pd.isna(vr) or vr < CONFIG["RE_ENTRY_VOL_MIN"]:
                        re_ok = False
                if re_ok and (pd.notna(ema5) and pd.notna(ema20) and pd.notna(close)
                              and close > ema20 and ema5 > ema20):
                    position = CONFIG["RE_ENTRY_POSITION"]
                    entry_price = close
                    highest_price = close
                    if pd.notna(atr):
                        trailing_stop = close - _trail_mult(stage) * atr
                    action = f"再入场 {position:.0%}"

        positions.append(round(position, 4))
        actions.append(action)
        trailing_stops.append(round(trailing_stop, 4) if pd.notna(trailing_stop) else np.nan)
        stop_triggered.append(stop_hit)
        position_values.append(round(position * close, 4) if pd.notna(close) else np.nan)

    df_back["持仓比例"] = positions
    df_back["本次操作"] = actions
    df_back["移动止损价"] = trailing_stops
    df_back["移动止损触发"] = stop_triggered
    df_back["持仓市值"] = position_values
    return df_back


def single_stock_backtrack(code, days=60, output_excel=True):
    print(f"\n⏪ 回溯 {code} 最近 {days} 个交易日")

    calendar_days = max(CONFIG["HISTORY_CALENDAR_DAYS"], days * 3 + 400)
    df_raw, name = safe_fetch_stock(code, calendar_days)
    if df_raw is None:
        return None

    data = calculate_engine(df_raw)
    if data is None:
        return None

    back = data.tail(days).copy()
    records = []

    wait_map = {
        "BUY_ON_VOLUME": "等量放大",
        "WAIT_VOL": "等量放大",
        "BUY_ON_EMA5_RECLAIM": "EMA5 重新站上",
        "WAIT_EMA5": "EMA5 重新站上",
        "BUY_ON_PULLBACK_EMA20": "回 EMA20",
        "WAIT_PULLBACK": "回 EMA20",
        "BUY_ON_BREAKOUT": "突破买",
    }

    for idx, row in back.iterrows():
        timing_code, priority, action_level, reason = ai_timing_engine(row)
        entry_engine = row.get("Entry_Engine", "NONE")
        exit_engine = row.get("Exit_Engine", "N/A")

        wait_for = wait_map.get(timing_code, wait_map.get(entry_engine, ""))
        buy_zone = ai_buy_zone(row, timing_code)
        trading_priority = calculate_trading_priority(row, timing_code)
        stock_label = f"{name} ({clean_stock_code(code)})"
        stock_type = row.get("Stock_Type", "mixed")
        type_label = {"pullback": "回调型", "breakout": "突破型", "mixed": "混合型"}.get(stock_type, stock_type)
        reduce_ratio = ai_reduce_ratio(exit_engine, timing_code)

        records.append({
            "股票": stock_label,
            "日期": idx.strftime("%Y-%m-%d"),
            "股票类型": type_label,
            "阶段": row["阶段"],
            "最终信号": row["最终信号"],
            "股票质量分": row.get("Stock_Quality_Score", np.nan),
            "买点质量分": row.get("Entry_Quality_Score", np.nan),
            "买点质量评级": row.get("Entry_Quality_Rating", "N/A"),
            "AI时机": timing_code,
            "Entry Engine": entry_engine,
            "Setup_Type": row.get("Setup_Type", "NONE"),
            "Setup_Age": row.get("Setup_Age", 999),
            "Setup_Active": bool(row.get("Setup_Active", False)),
            "Breakout_Confirmed": bool(row.get("Breakout_Confirmed", False)),
            "等什么": wait_for,
            "交易优先级": trading_priority,
            "AI_参考买入区": buy_zone,
            "AI_减仓比例": reduce_ratio,

            "Top_Warning_Score": row.get("Top_Warning_Score", 0),
            "Top_Warning_Reasons": row.get("Top_Warning_Reasons", ""),

            "名称": name,
            "代码": clean_stock_code(code),
            "收盘价": round(row["Close"], 4),
            "EMA5": round(row["EMA5"], 4) if pd.notna(row.get("EMA5", np.nan)) else np.nan,
            "EMA20": round(row["EMA20"], 4) if pd.notna(row["EMA20"]) else np.nan,
            "EMA60": round(row["EMA60"], 4) if pd.notna(row["EMA60"]) else np.nan,
            "Deduct5": round(row["Deduct5"], 4) if pd.notna(row.get("Deduct5", np.nan)) else np.nan,
            "Deduct20": round(row["Deduct20"], 4) if pd.notna(row.get("Deduct20", np.nan)) else np.nan,
            "Deduct60": round(row["Deduct60"], 4) if pd.notna(row.get("Deduct60", np.nan)) else np.nan,
            "EMA5扣底状态": row.get("EMA5_扣底状态", "N/A"),
            "EMA20扣底状态": row.get("EMA20_扣底状态", "N/A"),
            "EMA60扣底状态": row.get("EMA60_扣底状态", "N/A"),
            "EMA5_Cross_Age": row.get("EMA5_Cross_Age", np.nan),
            "Breakout_Age": row.get("Breakout_Age", np.nan),
            "Trading Permission": row["Trading_Permission"],
            "Exit Engine": exit_engine,
            "AI_动作": action_level,
            "AI_优先级": priority,
            "AI_原因": reason,
            "ADX": round(row["ADX"], 2) if pd.notna(row["ADX"]) else np.nan,
            "+DI": round(row["+DI"], 2) if pd.notna(row["+DI"]) else np.nan,
            "-DI": round(row["-DI"], 2) if pd.notna(row["-DI"]) else np.nan,
            "Volume Ratio": round(row["Volume_Ratio"], 2) if pd.notna(row["Volume_Ratio"]) else np.nan,
            "Close_Position_Range": round(row["Close_Position_Range"], 4) if pd.notna(row.get("Close_Position_Range", np.nan)) else np.nan,
            "ATR Distance": round(row.get("ATR_Distance", np.nan), 2) if pd.notna(row.get("ATR_Distance", np.nan)) else np.nan,
            "ATR14": round(row.get("ATR14", np.nan), 4) if pd.notna(row.get("ATR14", np.nan)) else np.nan,
            "ATR Health Score": row.get("ATR_Health_Score", np.nan),
            "ATR Risk Score": row.get("ATR_Risk_Score", np.nan),
            "ATR Risk": row.get("ATR_Risk", "数据不足"),
            "结构防守价": round(row.get("结构防守价", np.nan), 4) if pd.notna(row.get("结构防守价", np.nan)) else np.nan,
            "ATR_Stop_2x": round(row.get("ATR_Stop_2x", np.nan), 4) if pd.notna(row.get("ATR_Stop_2x", np.nan)) else np.nan,
            "Raw Score": row.get("Raw_Score", np.nan),
            "Stage Adjustment": row.get("Stage_Adjustment", np.nan),
            "Final Score": row.get("Final_Score", np.nan),
        })

    df_back = pd.DataFrame(records)

    close_series = df_back["收盘价"].astype(float).reset_index(drop=True)

    def fwd_return(n):
        return (close_series.shift(-n) / close_series - 1.0) * 100.0

    df_back["信号后5日收益%"] = fwd_return(5).round(2)
    df_back["信号后10日收益%"] = fwd_return(10).round(2)
    df_back["信号后20日收益%"] = fwd_return(20).round(2)

    def fwd_max_dd(n=10):
        vals = []
        arr = close_series.to_numpy()
        for i in range(len(arr)):
            end = min(i + n, len(arr) - 1)
            if end <= i:
                vals.append(np.nan)
                continue
            future = arr[i:end + 1]
            peak = future[0]
            max_dd = 0.0
            for x in future:
                if x > peak:
                    peak = x
                if peak > 0:
                    dd = (x - peak) / peak * 100.0
                    if dd < max_dd:
                        max_dd = dd
            vals.append(round(float(max_dd), 2))
        return vals

    df_back["信号后10日最大回撤%"] = fwd_max_dd(10)

    df_back = simulate_position_tracking(df_back)

    df_back["上一日信号"] = df_back["最终信号"].shift(1)
    df_turns = df_back[df_back["最终信号"] != df_back["上一日信号"]].copy()

    print("\n最近交易状态：")
    show_cols = [
        "日期", "阶段", "最终信号", "AI时机", "AI_动作",
        "Setup_Type", "Setup_Age",
        "Top_Warning_Score",
        "持仓比例", "本次操作", "移动止损价", "移动止损触发",
        "信号后5日收益%", "信号后10日收益%",
    ]
    show_cols = [c for c in show_cols if c in df_back.columns]
    print(df_back[show_cols].tail(15).to_string(index=False))

    if output_excel:
        filename = (
            f"{safe_filename(name)}_EMA_V4.2_回溯{days}天_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df_back.to_excel(writer, sheet_name="逐日信号日志", index=False)
            df_turns.to_excel(writer, sheet_name="转折点", index=False)

        format_excel_file(filename)
        print(f"\n✅ 回溯报告：{filename}")

    return df_back


# ================================================================
# 15. ADX 对照
# ================================================================

def compare_adx_recent(code, rows=30, save_excel=True):
    df_raw, name = safe_fetch_stock(code)
    if df_raw is None:
        return None

    data = calculate_engine(df_raw)
    cols = ["Close", "High", "Low", "+DI", "-DI", "DI_Spread", "DX", "ADX", "ADX_Change"]
    result = data[cols].tail(rows).copy()

    print("\n" + "=" * 90)
    print(f"ADX 对照表（Wilder ADX 14,14，已过滤假期假数据）：{name} ({code})")
    print("=" * 90)
    print(result.round(4).to_string())

    if save_excel:
        filename = (
            f"{safe_filename(name)}_ADX_V4.2_对照_"
            f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        result.reset_index().to_excel(filename, index=False)
        format_excel_file(filename)
        print(f"\n✅ ADX对照文件：{filename}")
    return result


# ================================================================
# 16. MAIN ENGINE
# ================================================================

def apply_stage_confirmation(data, required_days=2):
    if "阶段_Raw" not in data.columns:
        return data

    raw = data["阶段_Raw"].tolist()
    if not raw:
        return data

    confirmed = []
    current = raw[0]
    candidate = None
    candidate_count = 0

    for s in raw:
        if s == current:
            confirmed.append(current)
            candidate = None
            candidate_count = 0
        else:
            # V4.2：同一 Stage 内的子标签切换（如 3A<->3B、2 的长期标注）立即生效，不需等确认
            if (get_stage_key(s) == get_stage_key(current)
                    or (get_stage_key(s) in ("Stage 3A", "Stage 3B")
                        and get_stage_key(current) in ("Stage 3A", "Stage 3B"))):
                current = s
                confirmed.append(current)
                candidate = None
                candidate_count = 0
                continue
            if s == candidate:
                candidate_count += 1
            else:
                candidate = s
                candidate_count = 1

            if candidate_count >= required_days:
                current = candidate
                candidate = None
                candidate_count = 0

            confirmed.append(current)

    data["阶段"] = confirmed
    return data


def calculate_engine(df):
    if df is None or df.empty:
        return None

    data = df.copy()

    # ---------- EMA ----------
    data["EMA5"] = data["Close"].ewm(span=CONFIG["EMA_FAST"], adjust=False, min_periods=CONFIG["EMA_FAST"]).mean()
    data["EMA20"] = data["Close"].ewm(span=CONFIG["EMA_SHORT"], adjust=False, min_periods=CONFIG["EMA_SHORT"]).mean()
    data["EMA60"] = data["Close"].ewm(span=CONFIG["EMA_LONG"], adjust=False, min_periods=CONFIG["EMA_LONG"]).mean()
    data["EMA200"] = data["Close"].ewm(span=CONFIG["EMA_BACKGROUND"], adjust=False, min_periods=CONFIG["EMA_BACKGROUND"]).mean()

    # ---------- 价格距离 ----------
    data["Gap5"] = data["Close"] - data["EMA5"]
    data["Gap20"] = data["Close"] - data["EMA20"]
    data["Gap60"] = data["Close"] - data["EMA60"]

    data["Gap5_Change"] = data["Gap5"].diff()
    data["Gap20_Change"] = data["Gap20"].diff()
    data["Gap60_Change"] = data["Gap60"].diff()

    data["Gap5_Change_5"] = data["Gap5"].diff(5)
    data["Gap20_Change_5"] = data["Gap20"].diff(5)
    data["Gap60_Change_5"] = data["Gap60"].diff(5)

    data["Gap5_Pct"] = data["Gap5"] / data["Close"] * 100
    data["Gap20_Pct"] = data["Gap20"] / data["Close"] * 100
    data["Gap60_Pct"] = data["Gap60"] / data["Close"] * 100

    data["EMA5_距离状态"] = [
        classify_distance(d, dc, p)
        for d, dc, p in zip(data["Gap5"], data["Gap5_Change"], data["Close"])
    ]
    data["EMA20_距离状态"] = [
        classify_distance(d, dc, p)
        for d, dc, p in zip(data["Gap20"], data["Gap20_Change"], data["Close"])
    ]
    data["EMA60_距离状态"] = [
        classify_distance(d, dc, p)
        for d, dc, p in zip(data["Gap60"], data["Gap60_Change"], data["Close"])
    ]

    # ---------- 真正的 EMA 扣底 ----------
    data["Deduct5"] = data["EMA5"] - data["Close"].shift(CONFIG["EMA_FAST"])
    data["Deduct20"] = data["EMA20"] - data["Close"].shift(CONFIG["EMA_SHORT"])
    data["Deduct60"] = data["EMA60"] - data["Close"].shift(CONFIG["EMA_LONG"])

    data["Deduct5_Change"] = data["Deduct5"].diff()
    data["Deduct20_Change"] = data["Deduct20"].diff()
    data["Deduct60_Change"] = data["Deduct60"].diff()

    data["EMA5_扣底状态"] = [
        classify_deduct(d, dc, p)
        for d, dc, p in zip(data["Deduct5"], data["Deduct5_Change"], data["Close"])
    ]
    data["EMA20_扣底状态"] = [
        classify_deduct(d, dc, p)
        for d, dc, p in zip(data["Deduct20"], data["Deduct20_Change"], data["Close"])
    ]
    data["EMA60_扣底状态"] = [
        classify_deduct(d, dc, p)
        for d, dc, p in zip(data["Deduct60"], data["Deduct60_Change"], data["Close"])
    ]

    # ---------- EMA 斜率 ----------
    data["EMA5_Actual_Slope_5"] = data["EMA5"].diff(5)
    data["EMA20_Actual_Slope_5"] = data["EMA20"].diff(5)
    data["EMA60_Actual_Slope_5"] = data["EMA60"].diff(5)
    data["EMA200_Actual_Slope_20"] = data["EMA200"].diff(20)

    data["Extension_EMA5_Pct"] = (data["Close"] - data["EMA5"]) / data["EMA5"] * 100
    data["Extension_EMA20_Pct"] = (data["Close"] - data["EMA20"]) / data["EMA20"] * 100
    data["Extension_EMA60_Pct"] = (data["Close"] - data["EMA60"]) / data["EMA60"] * 100

    # ---------- Volume ----------
    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=20).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)
    data["Breakout_Volume"] = data["Volume_Ratio"] >= CONFIG["BREAKOUT_VOLUME_RATIO"]

    # ---------- ATR ----------
    tr = pd.concat(
        [
            data["High"] - data["Low"],
            (data["High"] - data["Close"].shift(1)).abs(),
            (data["Low"] - data["Close"].shift(1)).abs(),
        ],
        axis=1,
    ).max(axis=1)
    data["TR"] = tr
    data["ATR14"] = wilder_rma(tr, 14)
    data["ATR_Pct"] = data["ATR14"] / data["Close"] * 100

    data["EMA20_ATR_Distance"] = (data["Close"] - data["EMA20"]) / data["ATR14"].replace(0, np.nan)
    data["EMA20_ATR_Distance_Abs"] = data["EMA20_ATR_Distance"].abs()

    def atr_risk_label(row):
        d = row.get("ATR_Distance", np.nan)
        ap = row.get("ATR_Pct", np.nan)
        if pd.isna(d) or pd.isna(ap):
            return "数据不足"
        ad = abs(float(d))
        if ad <= CONFIG["ATR_DISTANCE_HEALTHY"] and ap <= 4:
            return "低风险：接近EMA20"
        if ad <= CONFIG["ATR_DISTANCE_CAUTION"] and ap <= 6:
            return "中风险：正常波动"
        if ad <= CONFIG["ATR_DISTANCE_HIGH"]:
            return "偏高：不宜追高"
        return "高风险：过度延伸"

    data["ATR_Distance"] = data["EMA20_ATR_Distance"]
    data["ATR_Distance_Abs"] = data["EMA20_ATR_Distance_Abs"]
    data["ATR_Risk"] = data.apply(atr_risk_label, axis=1)

    # ---------- ADX ----------
    plus_di, minus_di, dx, adx = compute_adx_wilder(data, 14)
    data["+DI"] = plus_di
    data["-DI"] = minus_di
    data["DX"] = dx
    data["ADX"] = adx
    data["ADX_Change"] = data["ADX"].diff()
    data["ADX_Change_5"] = data["ADX"].diff(5)
    data["DI_Spread"] = data["+DI"] - data["-DI"]

    def adx_strength(x):
        if pd.isna(x):
            return "数据不足"
        if x < 20:
            return "弱趋势/震荡"
        if x < 25:
            return "趋势发展中"
        if x < 30:
            return "强趋势"
        return "强趋势 (>30)"
    data["ADX_强度"] = data["ADX"].apply(adx_strength)
    data["DI_方向"] = np.where(data["+DI"] > data["-DI"], "多方", "空方")

    # ---------- OBV ----------
    direction = np.sign(data["Close"].diff()).fillna(0)
    data["OBV"] = (direction * data["Volume"]).cumsum()
    data["OBV_MA5"] = data["OBV"].rolling(5, min_periods=5).mean()
    data["OBV_Change_5"] = data["OBV"].diff(5)
    data["OBV_MA_Distance_Pct"] = (
        (data["OBV"] - data["OBV_MA5"]) / data["OBV_MA5"].abs().replace(0, np.nan) * 100
    )
    data["OBV_趋势"] = np.where(
        data["OBV"].diff(5) > 0, "上升",
        np.where(data["OBV"].diff(5) < 0, "下降", "走平"),
    )

    # ---------- Bollinger ----------
    data["BB_mid"] = data["Close"].rolling(CONFIG["BB_PERIOD"], min_periods=CONFIG["BB_PERIOD"]).mean()
    data["BB_std"] = data["Close"].rolling(CONFIG["BB_PERIOD"], min_periods=CONFIG["BB_PERIOD"]).std()
    data["BB_high"] = data["BB_mid"] + CONFIG["BB_STD"] * data["BB_std"]
    data["BB_low"] = data["BB_mid"] - CONFIG["BB_STD"] * data["BB_std"]
    data["BB_Position_Value"] = (
        (data["Close"] - data["BB_low"]) / (data["BB_high"] - data["BB_low"]).replace(0, np.nan)
    )

    def bb_position(x):
        if pd.isna(x):
            return "数据不足"
        if x >= 1:
            return "突破上轨"
        if x <= 0:
            return "跌破下轨"
        if x >= 0.5:
            return "中轨上方"
        return "中轨下方"
    data["BB_位置"] = data["BB_Position_Value"].apply(bb_position)

    # ---------- Candlestick ----------
    data["Candle_Range"] = (data["High"] - data["Low"]).replace(0, np.nan)
    data["Candle_Body"] = (data["Close"] - data["Open"]).abs()
    data["Candle_Body_Pct"] = data["Candle_Body"] / data["Candle_Range"] * 100
    data["Close_Position_Range"] = (data["Close"] - data["Low"]) / data["Candle_Range"]

    data["Upper_Shadow"] = data["High"] - data[["Open", "Close"]].max(axis=1)
    data["Lower_Shadow"] = data[["Open", "Close"]].min(axis=1) - data["Low"]
    data["Upper_Shadow_Pct"] = data["Upper_Shadow"] / data["Candle_Range"] * 100
    data["Lower_Shadow_Pct"] = data["Lower_Shadow"] / data["Candle_Range"] * 100

    data["Return_1"] = data["Close"].pct_change() * 100
    data["Return_5"] = data["Close"].pct_change(5) * 100
    data["Return_20"] = data["Close"].pct_change(20) * 100

    # ---------- V4.2 价量变化 ----------
    data["Volume_MA5"] = data["Volume"].rolling(5, min_periods=3).mean()
    data["Volume_Change_1"] = (data["Volume"].pct_change() * 100).replace([np.inf, -np.inf], np.nan)
    data["Volume_Change_5"] = (
        (data["Volume_MA5"] / data["Volume_MA5"].shift(5).replace(0, np.nan) - 1) * 100
    ).replace([np.inf, -np.inf], np.nan)
    data["Volume_MA5_vs_MA20_Pct"] = (
        (data["Volume_MA5"] / data["Volume_MA20"].replace(0, np.nan) - 1) * 100
    ).replace([np.inf, -np.inf], np.nan)
    data["PV_关系_1日"] = [classify_price_volume(r, v) for r, v in zip(data["Return_1"], data["Volume_Change_1"])]
    data["PV_关系_5日"] = [classify_price_volume(r, v) for r, v in zip(data["Return_5"], data["Volume_Change_5"])]

    # ---------- EMA200 context ----------
    data["Price_vs_EMA200_Pct"] = (data["Close"] - data["EMA200"]) / data["EMA200"] * 100
    data["EMA200_Slope_20"] = data["EMA200"].diff(20)

    # ---------- EMA5 金叉年龄 ----------
    data["EMA5_Cross_Age"] = _ema5_cross_age(data["EMA5"], data["EMA20"])

    # ---------- 突破年龄 ----------
    data["Breakout_Age"] = _breakout_age(data["Close"], data["High"], lookback=20)

    # 股票类型识别（滚动窗口）
    stock_types = []
    for i in range(len(data)):
        temp = data.iloc[:i + 1]
        stock_types.append(identify_stock_type(temp, CONFIG["STOCK_TYPE_LOOKBACK"]))
    data["Stock_Type"] = stock_types

    # ---------- Trading Permission V2 ----------
    def permission_v2(row):
        ema60 = row["EMA60"]
        ema200 = row["EMA200"]
        close = row["Close"]
        ema20 = row["EMA20"]
        slope60 = row["EMA60_Actual_Slope_5"]
        slope200 = row.get("EMA200_Actual_Slope_20", np.nan)

        if pd.isna(ema60) or pd.isna(ema200) or pd.isna(slope60):
            return "YELLOW"

        ema200_ok = pd.isna(slope200) or slope200 >= 0

        if slope60 > 0 and ema200_ok and pd.notna(close) and close > ema60:
            # V4.2：EMA200 长趋势过滤——站不上 EMA200 只给 YELLOW
            if CONFIG["LONG_TREND_FILTER_ENABLED"] and pd.notna(close) and close <= ema200:
                return "YELLOW"
            return "GREEN"
        if slope60 < 0:
            return "RED"
        return "YELLOW"

    data["Trading_Permission"] = data.apply(permission_v2, axis=1)
    data["Long_Trend_OK"] = (
        (data["Close"] > data["EMA200"])
        & (data["EMA200_Actual_Slope_20"].fillna(0) >= 0)
    ) | data["EMA200"].isna()

    # ---------- Stage（Raw） ----------
    def identify_stage(row):
        close = row["Close"]
        ema5 = row["EMA5"]
        ema20 = row["EMA20"]
        ema60 = row["EMA60"]
        g20 = row["Gap20"]
        g60 = row["Gap60"]
        ema20s = row["EMA20_Actual_Slope_5"]
        ema60s = row["EMA60_Actual_Slope_5"]
        ema200s20 = row["EMA200_Actual_Slope_20"]
        g20c = row["Gap20_Change_5"]
        g60c = row["Gap60_Change_5"]
        adx = row["ADX"]

        if any(pd.isna(x) for x in [ema20, ema60, g20, g60]):
            return "Stage 0 (数据不足)"

        if ema60s < 0 and g60 < 0 and g60c < 0:
            return "Stage 4 (趋势破坏)"

        ema200_ok = (pd.isna(ema200s20) or ema200s20 >= 0)

        late_condition = (
            ema20s > 0
            and ema60s > 0
            and (
                (g20c < 0 and g20 > 0)
                or (pd.notna(adx) and adx > 25 and row["ADX_Change_5"] < 0)
                or row["Extension_EMA20_Pct"] > CONFIG["EXTENSION_CAUTION_PCT"]
                or (pd.notna(ema5) and pd.notna(ema20) and ema5 < ema20)
            )
        )
        if late_condition:
            # V4.2：Stage 3 拆分 3A（健康调整）/ 3B（趋势恶化）
            di_sp = row.get("DI_Spread", np.nan)
            adx_chg5 = row.get("ADX_Change_5", np.nan)
            weak = False
            if pd.notna(ema60) and close < ema60:
                weak = True
            elif pd.notna(ema20) and close < ema20 and (
                (pd.notna(di_sp) and di_sp < 0)
                or (pd.notna(ema5) and ema5 < ema20 and pd.notna(adx_chg5) and adx_chg5 < 0)
            ):
                weak = True
            return "Stage 3B (末升-趋势恶化)" if weak else "Stage 3A (末升-健康调整)"

        if ema20s > 0 and ema60s > 0 and g20 > 0 and g60 > 0 and close > ema60 and ema200_ok:
            e200 = row["EMA200"]
            if CONFIG["LONG_TREND_FILTER_ENABLED"] and pd.notna(e200) and close <= e200:
                return "Stage 2 (主升段-未确认长期)"
            return "Stage 2 (主升段)"

        stage1 = (
            (ema20s > 0 and ema60s >= 0 and close >= ema20)
            and (g20 > 0 or (g20 <= 0 and pd.notna(g20c) and g20c > 0))
        )
        if stage1:
            return "Stage 1 (初升段)"

        if (ema60s >= 0 and pd.notna(g60c) and g60c > 0
                and pd.notna(g20c) and g20c > 0):
            return "Stage 1 (初升准备)"

        return "Stage 0 (观察)"
    data["阶段_Raw"] = data.apply(identify_stage, axis=1)

    # ---------- Stage 确认 ----------
    data = apply_stage_confirmation(data, required_days=CONFIG["STAGE_CONFIRMATION_DAYS"])

    # ---------- 前一日收盘 / EMA20 供 Stage 3 判定 ----------
    data["Prev_Close"] = data["Close"].shift(1)
    data["Prev_EMA20"] = data["EMA20"].shift(1)

    # ---------- 前一 20 日最高（供 Breakout Setup 使用）----------
    data["Prev_High_20"] = data["High"].rolling(20, min_periods=5).max().shift(1)

    # ---------- V4.2 Setup 记忆（必须在 Entry Engine 之前）----------
    data = compute_setup_memory_series(data)

    # ---------- Entry / Exit Engine V2 ----------
    entry_list = []
    entry_raw_list = []
    exit_list = []
    confirm_records = []
    for i in range(len(data)):
        r = data.iloc[i]
        prev = data.iloc[i - 1] if i > 0 else None
        raw_entry = entry_engine_v2(r, prev)
        confirmed_entry, confirm_dict = entry_confirmation_layer(r, raw_entry)   # V4.2 确认层
        entry_raw_list.append(raw_entry)
        entry_list.append(confirmed_entry)
        confirm_records.append(confirm_dict)
        exit_list.append(exit_engine_v2(r, prev))
    data["Entry_Engine_Raw"] = entry_raw_list
    data["Entry_Engine"] = entry_list
    data["Exit_Engine"] = exit_list
    confirm_df = pd.DataFrame(confirm_records, index=data.index)
    for c in confirm_df.columns:
        data[c] = confirm_df[c]

    # ---------- 基础信号 ----------
    def generate_signal(row):
        stage = row["阶段"]
        permission = row["Trading_Permission"]
        close = row["Close"]
        ema5 = row["EMA5"]
        ema20 = row["EMA20"]
        ema60 = row["EMA60"]
        g20 = row["Gap20"]
        g60 = row["Gap60"]
        adx = row["ADX"]
        adx_change = row["ADX_Change_5"]
        di_spread = row["DI_Spread"]
        volume_ratio = row["Volume_Ratio"]
        extension = row["Extension_EMA20_Pct"]

        if "Stage 4" in stage:
            return "卖出"

        if permission == "RED":
            if close < ema60 and g60 < 0:
                return "卖出"
            return "减码"

        if "Stage 1" in stage:
            early_confirm = (pd.notna(g20) and g20 >= 0 and pd.notna(ema20) and close >= ema20)
            adx_confirm = (
                pd.notna(adx) and pd.notna(adx_change) and adx_change > 0
                and pd.notna(di_spread) and di_spread > 0
            )
            volume_confirm = (pd.notna(volume_ratio) and volume_ratio >= CONFIG["BREAKOUT_VOLUME_RATIO"])
            if early_confirm and adx_confirm:
                if pd.notna(extension) and extension > CONFIG["EXTENSION_MAX_PCT"]:
                    return "等待回调"
                if volume_confirm:
                    return "买入"
                return "买入观察"
            return "观望"

        if "Stage 2" in stage:
            if pd.notna(extension) and extension > CONFIG["EXTENSION_MAX_PCT"]:
                return "等待回调"
            if (pd.notna(adx) and adx > CONFIG["STAGE2_ADX_ADD_MIN"]
                    and pd.notna(di_spread) and di_spread > 0):
                if close >= ema20:
                    return "加码"
                return "持有"
            return "持有"

        if "Stage 3" in stage:
            if permission == "GREEN":
                if (pd.notna(g20) and g20 > 0 and pd.notna(adx_change) and adx_change < 0):
                    return "保护利润"
                if pd.notna(ema5) and pd.notna(ema20) and ema5 < ema20:
                    return "保护利润"
                return "持有"
            return "减码"

        return "观望"
    data["最终信号"] = data.apply(generate_signal, axis=1)

    # ---------- 综合信号（Stage 优先） ----------
    def combined_signal(row):
        entry = row.get("Entry_Engine", "NONE")
        exit_e = row.get("Exit_Engine", "持有")
        base = row.get("最终信号", "观望")
        permission = row.get("Trading_Permission", "YELLOW")
        stage = row.get("阶段", "Stage 0 (观察)")

        if exit_e == "EXIT":
            return "卖出"
        if exit_e in ("保护利润", "减仓"):
            return "减仓"
        if "Stage 4" in stage:
            return "观望"
        if "Stage 3" in stage:
            return base
        if "Stage 2" in stage:
            if base == "加码":
                return "加码"
            if entry in ["WAIT_VOL", "WAIT_EMA5", "WAIT_PULLBACK"]:
                return "持有 / 等待"
            return base
        if "Stage 1" in stage:
            if entry in ("BUY", "BUY_STRONG_BREAKOUT") and permission != "RED":
                return "买入"
            if entry in ["WAIT_VOL", "WAIT_EMA5", "WAIT_PULLBACK"]:
                return "买入观察"
            return base

        return base
    data["最终信号"] = data.apply(combined_signal, axis=1)

    # ---------- Risk / Defense ----------
    data["EMA20_推动力"] = data["Gap20"]
    data["EMA60_推动力"] = data["Gap60"]
    data["Recent_Low_20"] = data["Low"].rolling(20, min_periods=5).min()
    data["ATR_Stop_2x"] = data["Close"] - 2 * data["ATR14"]
    # V4.2：按阶段动态 ATR 倍数
    data["ATR_Multiplier"] = data["阶段"].apply(calculate_dynamic_atr_multiplier)
    data["ATR_Stop_Dynamic"] = data["Close"] - data["ATR_Multiplier"] * data["ATR14"]
    data["结构防守价"] = data[["Recent_Low_20", "ATR_Stop_Dynamic"]].max(axis=1)
    data["持有原则"] = np.select(
        [
            (data["Trading_Permission"] == "GREEN") & (data["Gap60"] > 0),
            (data["Trading_Permission"] == "YELLOW"),
            (data["Trading_Permission"] == "RED"),
        ],
        [
            "三层趋势绿色：允许Hold/交易",
            "三层趋势黄色：Hold但减少新增仓位",
            "三层趋势红色：停止新增多仓并考虑退出",
        ],
        default="观察",
    )

    # ---------- 双评分 ----------
    stock_q_rows = []
    entry_q_rows = []
    for i in range(len(data)):
        temp = data.iloc[:i + 1]
        sq = calculate_stock_quality_score(temp)
        stock_q_rows.append(sq)
        eq = calculate_entry_quality_score(temp.iloc[-1])
        entry_q_rows.append(eq)

    stock_q_df = pd.DataFrame(stock_q_rows, index=data.index)
    entry_q_df = pd.DataFrame(entry_q_rows, index=data.index)

    for c in stock_q_df.columns:
        data[c] = stock_q_df[c]
    for c in entry_q_df.columns:
        data[c] = entry_q_df[c]

    # ---------- 顶部预警分 ----------
    top_scores = []
    top_reasons_list = []
    for i in range(len(data)):
        row = data.iloc[i]
        prev_rows = data.iloc[max(0, i - 5):i]
        s, r = calculate_top_warning_score(row, prev_rows)
        top_scores.append(s)
        top_reasons_list.append("、".join(r))
    data["Top_Warning_Score"] = top_scores
    data["Top_Warning_Reasons"] = top_reasons_list

    # ---------- V4.2 AI 信心分 ----------
    data["AI_Confidence_Score"] = [calculate_ai_confidence_score(data.iloc[i]) for i in range(len(data))]

    # ---------- Human-readable trends ----------
    data["ADX_5日趋势"] = [arrow_trend(data["ADX"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Gap5_5日趋势"] = [arrow_trend(data["Gap5"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Gap20_5日趋势"] = [arrow_trend(data["Gap20"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Gap60_5日趋势"] = [arrow_trend(data["Gap60"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Deduct5_5日趋势"] = [arrow_trend(data["Deduct5"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Deduct20_5日趋势"] = [arrow_trend(data["Deduct20"].iloc[:i + 1], 5) for i in range(len(data))]
    data["Deduct60_5日趋势"] = [arrow_trend(data["Deduct60"].iloc[:i + 1], 5) for i in range(len(data))]
    data["EMA5_5日趋势"] = [arrow_trend(data["EMA5"].iloc[:i + 1], 5) for i in range(len(data))]
    data["OBV_5日趋势"] = [arrow_trend(data["OBV"].iloc[:i + 1], 5) for i in range(len(data))]

    # ================================================================
    # 最终信号统一到 AI 时机
    # ================================================================
    final_signals_unified = []
    for i in range(len(data)):
        r = data.iloc[i]
        timing_code, _, _, _ = ai_timing_engine(r)
        exit_e = r.get("Exit_Engine", "持有")
        final_signals_unified.append(map_ai_to_final_signal(timing_code, exit_e))
    data["最终信号"] = final_signals_unified

    return data


# ================================================================
# 17. MENU
# ================================================================

def print_menu():
    print("\n" + "=" * 70)
    print("   🇲🇾 马股 EMA Trading System V4.2")
    print("=" * 70)
    print("1. 单股分析（输出 9 个 Sheet 双层报告 + Master Prompt）")
    print("2. 批量扫描（默认价格 RM0.05–RM0.50）")
    print("3. Excel 股票池批量扫描")
    print("4. 单股历史回溯（含顶部预警 + 持仓跟踪）")
    print("5. ADX 对照测试（与 TradingView 对齐）")
    print("6. 退出")
    print("-" * 70)
    print("   （网页界面：streamlit run 本文件名.py）")


def main():
    print("=" * 70)
    print("🇲🇾 马股 EMA5 + EMA20 + EMA60 + EMA200 + 扣底 + ADX + ATR + OBV")
    print("      V4.2（V4.0.2 主干 + V4.1 移植）")
    print("=" * 70)
    print("核心逻辑：")
    print("  ★ 股票质量分 + 买点质量分（双评分）+ AI 信心分")
    print("  ★ Entry Engine：Setup + 1~3 天记忆 + Trigger，再经 5 因子确认层")
    print("  ★ EMA200 长趋势过滤：Close<=EMA200 → Permission 降为 YELLOW")
    print("  ★ Stage 3A（健康调整，2 个转弱条件减仓）/ 3B（趋势恶化，1 个条件减仓）")
    print("  ★ 动态 ATR 倍数：Stage1=2.5 / Stage2=3.0 / 3A=2.0 / 3B=1.5 / Stage4=1.0")
    print("  ★ 顶部预警分：≥45 减 1/3，≥70 再减 1/3，≥90 清仓")
    print("  ★ Master Prompt：含 +DI/-DI、ADX、OBV、价量变化、OBV 背离、防守/压力位")
    print()

    while True:
        print_menu()
        choice = input("请输入选项 (1-6): ").strip()

        if choice == "1":
            code = input("输入股票代码，例如 10 / 0010 / 10.KL：").strip()
            if code:
                single_stock_scan(code)

        elif choice == "2":
            print("\n请输入扫描范围，直接 Enter 使用默认值。")
            min_p = input("最低价格 RM [0.05]：").strip()
            max_p = input("最高价格 RM [0.50]：").strip()
            min_v = input("最低成交量 [100000]：").strip()

            min_p = float(min_p) if min_p else 0.05
            max_p = float(max_p) if max_p else 0.50
            min_v = int(min_v) if min_v else 100000

            codes_text = input("输入股票代码，以逗号分隔（例如 10,38,85）：").strip()
            if not codes_text:
                print("⚠️ 没有输入股票代码。")
                continue

            codes = [
                {"code": clean_stock_code(x), "name": ""}
                for x in codes_text.split(",") if x.strip()
            ]
            batch_scan(codes, min_price=min_p, max_price=max_p, min_volume=min_v)

        elif choice == "3":
            filepath = input("输入 Excel 文件名，例如 Final list.xlsx：").strip()
            if not filepath:
                filepath = "Final list.xlsx"
            if not os.path.exists(filepath):
                print(f"❌ 文件不存在：{filepath}")
                continue

            pool = load_stock_pool_from_excel(filepath)
            if pool:
                min_p = input("最低价格 RM [0.05]：").strip()
                max_p = input("最高价格 RM [0.50]：").strip()
                min_v = input("最低成交量 [100000]：").strip()
                min_p = float(min_p) if min_p else 0.05
                max_p = float(max_p) if max_p else 0.50
                min_v = int(min_v) if min_v else 100000
                batch_scan(pool, min_price=min_p, max_price=max_p, min_volume=min_v)

        elif choice == "4":
            code = input("输入股票代码：").strip()
            days = input("回溯多少交易日 [60]：").strip()
            days = int(days) if days.isdigit() else 60
            if code:
                single_stock_backtrack(code, days)

        elif choice == "5":
            code = input("输入股票代码：").strip()
            rows = input("显示最近多少天 [30]：").strip()
            rows = int(rows) if rows.isdigit() else 30
            if code:
                compare_adx_recent(code, rows)

        elif choice == "6":
            print("再见！")
            break

        else:
            print("⚠️ 无效选项，请重新输入。")


# ================================================================
# 18. STREAMLIT UI（V4.2 新增）
#     运行：streamlit run EMA_trading_system_V4_2.py
# ================================================================

def run_streamlit_app():
    import streamlit as st

    st.set_page_config(page_title="EMA Trading System V4.2", page_icon="📈", layout="wide")
    st.title("🇲🇾 马股 EMA Trading System V4.2")
    st.caption("Stage 3A/3B + EMA200 长趋势过滤 + 动态 ATR + 确认层 + AI 信心 + 完整 Master Prompt")

    tab1, tab2, tab3 = st.tabs(["📊 单股分析", "📈 批量扫描", "🔬 ADX 对照"])

    with tab1:
        colA, colB = st.columns([2, 1])
        code = colA.text_input("股票代码（如 10 / 0010 / 10.KL）", value="10")
        if colB.button("🚀 开始分析", use_container_width=True):
            with st.spinner("正在获取数据并计算..."):
                result = single_stock_scan(code, output_excel=False)
            if result:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("股票质量分", f"{result['股票质量分']:.1f}", result["股票质量评级"])
                c2.metric("买点质量分", f"{result['买点质量分']:.1f}", result["买点质量评级"])
                c3.metric("AI 信心", f"{result['AI_Confidence']:.0f}")
                c4.metric("阶段", str(result["阶段"]).split(" (")[0])

                st.markdown(f"### 综合建议：**{result['综合建议']}**")
                st.markdown(
                    f"**最终信号**：{result['最终信号']} ｜ **Entry**：{result['Entry Engine']} ｜ "
                    f"**Exit**：{result['Exit Engine']} ｜ **Permission**：{result['Trading Permission']}"
                )
                st.info(result["AI 自然语言摘要"])

                d1, d2, d3 = st.columns(3)
                d1.metric("ADX", result.get("ADX"), f"变化5日 {result.get('ADX变化5日')}")
                d2.metric("+DI / -DI", f"{result.get('+DI')} / {result.get('-DI')}", f"Spread {result.get('DI Spread')}")
                d3.metric("量比", result.get("Volume Ratio"), result.get("价量关系_1日"))

                st.text_area("机构级 Master Prompt（复制到 Claude / ChatGPT 使用）",
                             result["机构级MasterPrompt"], height=520)
            else:
                st.error("无法获取该股票数据，请检查代码。")

    with tab2:
        codes_text = st.text_input("股票代码（逗号分隔）", value="10,38,85")
        colA, colB, colC = st.columns(3)
        min_p = colA.number_input("最低价格 RM", value=0.05, step=0.01, format="%.2f")
        max_p = colB.number_input("最高价格 RM", value=0.50, step=0.01, format="%.2f")
        min_v = colC.number_input("最低成交量", value=100000, step=10000)
        if st.button("🔍 开始批量扫描", use_container_width=True):
            codes = [{"code": clean_stock_code(x), "name": ""} for x in codes_text.split(",") if x.strip()]
            with st.spinner(f"正在扫描 {len(codes)} 只股票..."):
                df = batch_scan(codes, output_excel=False,
                                min_price=min_p, max_price=max_p, min_volume=int(min_v))
            if df is not None and not df.empty:
                show_cols = [c for c in ["代码", "名称", "阶段", "综合建议", "AI_Confidence",
                                          "股票质量分", "买点质量分", "Top_Warning_Score",
                                          "Trading Permission", "最终信号"] if c in df.columns]
                st.dataframe(df[show_cols], use_container_width=True, hide_index=True)
                if "机构级MasterPrompt" in df.columns:
                    pick = st.selectbox("查看某只股票的 Master Prompt", df["代码"].tolist())
                    st.text_area("Master Prompt", df.loc[df["代码"] == pick, "机构级MasterPrompt"].iloc[0], height=480)
            else:
                st.warning("无符合条件的结果。")

    with tab3:
        code = st.text_input("ADX 测试代码", value="10", key="adx_code")
        rows = st.number_input("显示最近多少天", value=30, step=5, key="adx_rows")
        if st.button("测试", use_container_width=True, key="adx_btn"):
            with st.spinner("正在计算..."):
                res = compare_adx_recent(code, int(rows), save_excel=False)
            if res is not None:
                st.dataframe(res.round(4), use_container_width=True)


def _is_streamlit_runtime():
    try:
        import streamlit as st
        return bool(st.runtime.exists())
    except Exception:
        return False


if __name__ == "__main__":
    if _is_streamlit_runtime():
        run_streamlit_app()
    else:
        main()
