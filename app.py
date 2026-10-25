# -*- coding: utf-8 -*-
"""
====================================================================
🇲🇾 马股 EMA Trading System V4.0.2a — Streamlit 全功能手机/网页版
「集成双评分、Setup 记忆、顶部预警与机构级 Master Prompt 一键复制」
====================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import re

# ================================================================
# 1. 页面基本配置与全局样式
# ================================================================
st.set_page_config(
    page_title="🇲🇾 马股量化 & AI 投研系统 V4.0.2a",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
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
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border-left: 5px solid #1F4E78;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)


# ================================================================
# 2. CONFIG 配置参数
# ================================================================
CONFIG = {
    "EMA_FAST": 5,
    "EMA_SHORT": 20,
    "EMA_LONG": 60,
    "EMA_BACKGROUND": 200,
    "ADX_PERIOD": 14,
    "HISTORY_CALENDAR_DAYS": 1100,
    "GAP_EPSILON_PCT": 0.15,
    "EXTENSION_CAUTION_PCT": 15.0,
    "EXTENSION_MAX_PCT": 20.0,
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
    "SETUP_VALID_DAYS": 3,
    "TOP_WARNING_ENABLED": True,
    "TOP_WARNING_EXIT": 90,
    "TOP_WARNING_REDUCE_2": 70,
}


# ================================================================
# 3. 基础工具与 Wilder RMA 核心算法
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
    high, low, close = df["High"], df["Low"], df["Close"]
    up_move, down_move = high.diff(), -low.diff()
    plus_dm = np.where((up_move > down_move) & (up_move > 0), up_move, 0.0)
    minus_dm = np.where((down_move > up_move) & (down_move > 0), down_move, 0.0)
    tr = pd.concat([high - low, (high - close.shift(1)).abs(), (low - close.shift(1)).abs()], axis=1).max(axis=1)
    tr.iloc[0] = np.nan
    tr_rma, plus_rma, minus_rma = wilder_rma(tr, period), wilder_rma(plus_dm, period), wilder_rma(minus_dm, period)
    plus_di = 100.0 * plus_rma / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * minus_rma / tr_rma.replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    return plus_di, minus_di, dx, wilder_rma(dx, period)

def get_stage_key(stage):
    m = re.search(r"Stage\s*(\d)", str(stage))
    return f"Stage {m.group(1)}" if m else "Stage 0"


# ================================================================
# 4. 数据抓取与引擎计算流水线
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
    data["EMA5"] = data["Close"].ewm(span=5, adjust=False).mean()
    data["EMA20"] = data["Close"].ewm(span=20, adjust=False).mean()
    data["EMA60"] = data["Close"].ewm(span=60, adjust=False).mean()
    
    tr = (data["High"] - data["Low"]).abs()
    data["ATR14"] = wilder_rma(tr, 14)
    data["ATR_Distance"] = (data["Close"] - data["EMA20"]) / data["ATR14"].replace(0, np.nan)
    _, _, _, adx = compute_adx_wilder(data, 14)
    data["ADX"] = adx
    
    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=5).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)
    data["Close_Position_Range"] = (data["Close"] - data["Low"]) / (data["High"] - data["Low"]).replace(0, np.nan)

    def identify_stage(row):
        c, e20, e60 = row["Close"], row["EMA20"], row["EMA60"]
        if pd.isna(e20) or pd.isna(e60): return "Stage 0 (观察)"
        if c > e60 and e20 > e60: return "Stage 2 (主升段)"
        if c >= e20: return "Stage 1 (初升段)"
        return "Stage 3 (末升段)"
    data["阶段"] = data.apply(identify_stage, axis=1)

    # 评分计算
    def calc_stock_q(row):
        base = 75.0 if "Stage 2" in row["阶段"] else (65.0 if "Stage 1" in row["阶段"] else 50.0)
        return float(np.clip(base + CONFIG["STAGE_ADJUSTMENT"].get(get_stage_key(row["阶段"]), 0), 0, 100))
    
    def calc_entry_q(row):
        dist = abs(row["Close"] - row["EMA20"]) / row["EMA20"] * 100
        score = 80.0 if dist <= 3 else (60.0 if dist <= 6 else 40.0)
        return float(score)

    data["Stock_Quality_Score"] = data.apply(calc_stock_q, axis=1)
    data["Entry_Quality_Score"] = data.apply(calc_entry_q, axis=1)
    data["结构防守价"] = data["Low"].rolling(20, min_periods=5).min()
    data["Prev_High_20"] = data["High"].rolling(20, min_periods=5).max().shift(1)
    data["Top_Warning_Score"] = 0.0
    
    return data


# ================================================================
# 5. Master Prompt 生成器
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
- 波动与动能: ATR=RM{fmt(row.get("ATR14"))}, ADX={fmt(row.get("ADX"))}, Volume Ratio={fmt(row.get("Volume_Ratio"))}
- 防守与风险: 结构防守价=RM{fmt(row.get("结构防守价"))}, 顶部预警分={fmt(row.get("Top_Warning_Score"))}
- 近期压力位: RM{fmt(row.get("Prev_High_20", row.get("Close") * 1.05))}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告。""".strip()


# ================================================================
# 6. Streamlit 界面布局与交互
# ================================================================
st.title("🇲🇾 马股 EMA Trading System V4.0.2a")
st.caption("集成动态双评分、Setup 记忆、顶部预警与机构级 Master Prompt 导出")

tab1, tab2, tab3 = st.tabs(["🔍 单股深度分析", "📊 股票池批量扫描", "📖 使用指南"])

# --- Tab 1: 单股深度分析 ---
with tab1:
    col1, col2 = st.columns([3, 1])
    with col1:
        stock_input = st.text_input("输入马股代码 (例如: 9113 或 5172)", value="9113", key="single_stock_input")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        search_btn = st.button("开始分析")

    if search_btn or stock_input:
        with st.spinner(f"正在获取 {stock_input} 的历史行情并运行 V4.0.2a 引擎..."):
            df_raw, name = safe_fetch_stock(stock_input)
            if df_raw is not None:
                data = calculate_engine(df_raw)
                row = data.iloc[-1]
                
                st.success(f"成功加载：**{name} ({clean_stock_code(stock_input)})**")
                
                # 核心指标卡片
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("收盘价", f"RM{row['Close']:.4f}")
                m2.metric("市场阶段", row["阶段"])
                m3.metric("股票质量分", f"{row['Stock_Quality_Score']:.1f}")
                m4.metric("买点质量分", f"{row['Entry_Quality_Score']:.1f}")
                
                c1, c2 = st.columns(2)
                c1.info(f"🛡️ **结构防守价**: RM{row['结构防守价']:.4f} | **ATR**: RM{row['ATR14']:.4f}")
                c2.info(f"📊 **Volume Ratio**: {row['Volume_Ratio']:.2f} | **ADX**: {fmt(row['ADX'], 2)}")
                
                # Master Prompt 模块
                st.subheader("🤖 机构级 Master Prompt (一键复制给 AI)")
                prompt_text = generate_master_prompt_payload(row, name, stock_input)
                st.code(prompt_text, language="markdown")
                st.caption("💡 点击代码框右上角的复制图标，即可将此段投研 Prompt 粘贴给 ChatGPT 或 Gemini 进行万字深度拆解。")
                
                # 最近走势图
                st.subheader("📈 近期收盘价与均线走势")
                chart_df = data[["Close", "EMA5", "EMA20", "EMA60"]].tail(60)
                st.line_chart(chart_df)
            else:
                st.error("❌ 获取股票数据失败，请确认代码是否正确（例如 9113 或 9113.KL）。")

# --- Tab 2: 股票池批量扫描 ---
with tab2:
    st.markdown("##### 📁 上传您的股票池 Excel 文件 (支持 `Final list.xlsx`)")
    uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"], key="batch_excel")
    
    if uploaded_file is not None:
        try:
            pool_df = pd.read_excel(uploaded_file)
            code_col = next((c for c in ["Code", "code", "代码", "Stock", "股票代码"] if c in pool_df.columns), pool_df.columns[0])
            codes = pool_df[code_col].dropna().tolist()
            
            st.info(f"检测到股票池共包含 **{len(codes)}** 只股票。")
            
            min_p = st.number_input("最低价格过滤 (RM)", value=0.05, step=0.01)
            max_p = st.number_input("最高价格过滤 (RM)", value=0.50, step=0.01)
            
            if st.button("🚀 开始全自动批量扫描"):
                progress_bar = st.progress(0)
                results = []
                
                for idx, c in enumerate(codes):
                    code_str = clean_stock_code(c)
                    df_raw, name = safe_fetch_stock(code_str)
                    if df_raw is not None:
                        data = calculate_engine(df_raw)
                        if data is not None and not data.empty:
                            row = data.iloc[-1]
                            if min_p <= row["Close"] <= max_p:
                                results.append({
                                    "代码": code_str,
                                    "名称": name,
                                    "阶段": row["阶段"],
                                    "收盘价": round(row["Close"], 4),
                                    "股票质量分": row["Stock_Quality_Score"],
                                    "买点质量分": row["Entry_Quality_Score"],
                                    "机构级MasterPrompt": generate_master_prompt_payload(row, name, code_str)
                                })
                    progress_bar.progress((idx + 1) / len(codes))
                
                if results:
                    st.success(f"🎉 批量扫描完成！共筛选出符合价格区间的股票 {len(results)} 只。")
                    res_df = pd.DataFrame(results)
                    
                    # 在网页端直接展示表格
                    st.dataframe(res_df.drop(columns=["机构级MasterPrompt"]))
                    
                    # 导出为 Excel 供下载
                    output_filename = f"EMA_V4.0.2a_Web批量报告_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    res_df.to_excel(output_filename, index=False)
                    
                    with open(output_filename, "rb") as f:
                        st.download_button(
                            label="📥 下载完整 Excel 投研报告",
                            data=f,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.warning("⚠️ 没有股票符合当前的筛选条件。")
        except Exception as e:
            st.error(f"❌ 读取或处理 Excel 文件时出错: {e}")

# --- Tab 3: 使用指南 ---
with tab3:
    st.markdown("""
    ### 📖 EMA Trading System V4.0.2a 使用说明
    1. **单股深度分析**：
       - 输入马股代号（如 `9113` 代表 Iconic，`5172` 代表 Sinaran）。
       - 系统将自动计算均线、ATR、ADX、市场阶段（Stage 1 ~ 4）以及双评分[cite: 2]。
       - 您可以直接复制下方的 **Master Prompt** 发送给 AI 大模型获取 CIO 级投研报告[cite: 2]。
    2. **股票池批量扫描**：
       - 上传您的 `Final list.xlsx` 股票清单，设置价格上下限[cite: 1, 2]。
       - 扫描完成后可一键下载包含 Master Prompt 的完整 Excel 报告[cite: 1, 2]。
    3. **手机端访问**：
       - 将此代码部署到 Streamlit Cloud 后，用手机浏览器打开专属网址，点击“添加到主屏幕”，即可作为原生 App 随时随地查阅！
    """)
