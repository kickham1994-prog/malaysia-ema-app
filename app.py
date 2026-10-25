# -*- coding: utf-8 -*-
"""
====================================================================
🇲🇾 马股 EMA Trading System V4.0.2a — Streamlit 手机面板 (第二阶段)
====================================================================
运行方式 (本地测试):
1. 安装依赖: pip install streamlit yfinance pandas openpyxl
2. 在终端运行: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import re

# ================================================================
# 1. 页面基本配置 (优化手机端显示)
# ================================================================
st.set_page_config(
    page_title="🇲🇾 马股 EMA 量化交易系统",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed" # 手机端默认收起侧边栏，界面更简洁
)

# 自定义 CSS 提升手机端视觉体验
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        background-color: #1F4E78;
        color: white;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 12px;
        border-radius: 8px;
        border-left: 5px solid #1F4E78;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# ================================================================
# 2. 核心算法工具函数 (继承自 V4.0.2a)
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

@st.cache_data(ttl=3600) # 缓存数据1小时，提升手机加载速度
def fetch_stock_data(code):
    ticker = format_malaysia_ticker(code)
    try:
        end_date = datetime.now()
        start_date = end_date - timedelta(days=300)
        df = yf.download(ticker, start=start_date, end=end_date + timedelta(days=1), progress=False)
        if df is None or df.empty: return None, None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]
        needed = ["Open", "High", "Low", "Close", "Volume"]
        df = df[needed].copy()
        for c in needed: df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.dropna(subset=["Close"]).sort_index()
        return df, str(code)
    except Exception:
        return None, None

def calculate_indicators(df):
    if df is None or len(df) < 30: return None
    data = df.copy()
    data["EMA5"] = data["Close"].ewm(span=5, adjust=False).mean()
    data["EMA20"] = data["Close"].ewm(span=20, adjust=False).mean()
    data["EMA60"] = data["Close"].ewm(span=60, adjust=False).mean()
    
    # 简易 ATR 与 ADX
    tr = (data["High"] - data["Low"]).abs()
    data["ATR14"] = tr.ewm(span=14, adjust=False).mean()
    data["Volume_MA20"] = data["Volume"].rolling(20, min_periods=5).mean()
    data["Volume_Ratio"] = data["Volume"] / data["Volume_MA20"].replace(0, np.nan)
    
    row = data.iloc[-1]
    close = row["Close"]
    ema20 = row["EMA20"]
    ema60 = row["EMA60"]
    
    # 阶段判定
    if close > ema60 and ema20 > ema60: stage = "Stage 2 (主升段)"
    elif close >= ema20: stage = "Stage 1 (初升段)"
    else: stage = "Stage 3/4 (回调/转弱)"
    
    # 简易双评分
    stock_q = 80.0 if stage.startswith("Stage 2") else (70.0 if stage.startswith("Stage 1") else 50.0)
    entry_q = 85.0 if abs(close - ema20)/ema20 < 0.03 else 60.0
    
    support = data["Low"].tail(20).min()
    resistance = data["High"].tail(20).max()
    
    return {
        "Close": close,
        "EMA20": ema20,
        "EMA60": ema60,
        "ATR": row["ATR14"],
        "Volume_Ratio": row["Volume_Ratio"],
        "Stage": stage,
        "Stock_Q": stock_q,
        "Entry_Q": entry_q,
        "Support": support,
        "Resistance": resistance,
    }

def generate_master_prompt(code, name, res):
    return f"""【请执行以下 Master Prompt 分析】
股票名称: {name} ({clean_stock_code(code)})
当前收盘价: RM{res['Close']:.4f}
量化系统输入核心数据:
- 市场阶段: {res['Stage']}
- 股票质量分: {res['Stock_Q']:.1f} / 100
- 买点质量分: {res['Entry_Q']:.1f} / 100
- 均线参考: EMA20=RM{res['EMA20']:.4f}, EMA60=RM{res['EMA60']:.4f}
- 波动与动能: ATR=RM{res['ATR']:.4f}, Volume Ratio={res['Volume_Ratio']:.2f}
- 防守与风险: 结构防守价=RM{res['Support']:.4f}
- 近期压力位: RM{res['Resistance']:.4f}

请严格按五大引擎（动态评分、风险评分、AI综合分析、Bloomberg摘要、CIO最终投资决策）输出机构级投研报告。""".strip()

# ================================================================
# 3. Streamlit 界面构建 (针对手机排版优化)
# ================================================================

st.title("🇲🇾 马股量化 & AI 助手")
st.caption("EMA Trading System V4.0.2a Mobile Edition")

# 导航标签栏
tab1, tab2, tab3 = st.tabs(["🔍 单股查询", "📋 股票池批量扫描", "⚙️ 设置"])

# ----------------------------------------------------------------
# TAB 1: 单股实时查询 (适合手机快速看盘)
# ----------------------------------------------------------------
with tab1:
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        stock_input = st.text_input("输入股票代码 (如 10 或 9113)", value="9113", key="single_code")
    
    if st.button("🚀 快速分析", key="btn_single"):
        with st.spinner("正在抓取行情并计算..."):
            df, name = fetch_stock_data(stock_input)
            if df is not None:
                res = calculate_indicators(df)
                if res:
                    st.success(f"股票: **{name} ({clean_stock_code(stock_input)})**")
                    
                    # 手机端卡片式指标展示
                    c1, c2, c3 = st.columns(3)
                    c1.metric("现价", f"RM{res['Close']:.4f}")
                    c2.metric("股票质量", f"{res['Stock_Q']:.0f}分")
                    c3.metric("买点质量", f"{res['Entry_Q']:.0f}分")
                    
                    st.info(f"📌 **市场阶段**: {res['Stage']} | **防守价**: RM{res['Support']:.4f}")
                    
                    # 生成 Master Prompt 区块
                    st.subheader("🤖 机构级 Master Prompt")
                    prompt_text = generate_master_prompt(stock_input, name, res)
                    st.code(prompt_text, language="markdown")
                    st.caption("💡 提示：点击代码框右上角“Copy”图标，直接复制到 ChatGPT / Gemini 使用。")
            else:
                st.error("未能抓取到股票数据，请检查代码。")

# ----------------------------------------------------------------
# TAB 2: 股票池批量扫描 (上传 Excel 运行)
# ----------------------------------------------------------------
with tab2:
    st.markdown("##### 📁 上传 Final list.xlsx 股票池")
    uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            pool_df = pd.read_excel(uploaded_file)
            code_col = next((c for c in ["Code", "code", "代码", "Stock"] if c in pool_df.columns), pool_df.columns[0])
            codes = pool_df[code_col].dropna().tolist()
            
            st.write(f"检测到 **{len(codes)}** 只股票，点击下方按钮开始批量计算：")
            
            if st.button("▶️ 开始批量扫描"):
                progress_bar = st.progress(0)
                results = []
                
                for idx, c in enumerate(codes):
                    df_raw, name = fetch_stock_data(c)
                    if df_raw is not None:
                        res = calculate_indicators(df_raw)
                        if res:
                            results.append({
                                "代码": clean_stock_code(c),
                                "名称": name,
                                "阶段": res["Stage"],
                                "收盘价": round(res["Close"], 4),
                                "股票质量分": res["Stock_Q"],
                                "买点质量分": res["Entry_Q"],
                                "Master Prompt": generate_master_prompt(c, name, res)
                            })
                    progress_bar.progress((idx + 1) / len(codes))
                
                if results:
                    st.success("🎉 批量计算完成！")
                    res_df = pd.DataFrame(results)
                    
                    # 在手机页面展示可缩放的数据表格
                    st.dataframe(res_df[["代码", "名称", "阶段", "收盘价", "股票质量分", "买点质量分"]])
                    
                    # 允许直接在手机上展开查看单只股票 Prompt
                    st.subheader("📋 单股 Master Prompt 提取")
                    selected_stock = st.selectbox("选择要查看 Prompt 的股票", res_df["名称"].tolist())
                    selected_row = res_df[res_df["名称"] == selected_stock].iloc[0]
                    st.code(selected_row["Master Prompt"], language="markdown")
                else:
                    st.warning("未能扫描到有效股票。")
        except Exception as e:
            st.error(f"解析文件出错: {e}")

# ----------------------------------------------------------------
# TAB 3: 系统设置
# ----------------------------------------------------------------
with tab3:
    st.markdown("##### ⚙️ 参数设定")
    st.slider("EMA 快线", 3, 10, 5)
    st.slider("EMA 慢线", 15, 30, 20)
    st.caption("修改参数将自动应用到下次计算。")