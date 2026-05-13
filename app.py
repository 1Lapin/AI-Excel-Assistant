import streamlit as st
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
import os

load_dotenv()

# 本地 Ollama 客户端
client = OpenAI(
    api_key="ollama",
    base_url="http://localhost:11434/v1"
)

st.set_page_config(page_title="AI Excel 分析助手", layout="wide")
st.title("AI Excel 数据分析助手 (精简版)")
st.write("上传 Excel，自动清洗数据并用本地 AI 生成分析总结。")

# ---------- 数据清洗函数 ----------
def clean_dataframe(df):
    report = []
    original_shape = df.shape

    df = df.dropna(how='all')
    if original_shape[0] - df.shape[0] > 0:
        report.append(f"删除全空行: {original_shape[0] - df.shape[0]} 行")

    before = df.shape[0]
    df = df.drop_duplicates()
    if before - df.shape[0] > 0:
        report.append(f"删除重复行: {before - df.shape[0]} 行")

    for col in df.columns:
        if df[col].isna().any():
            if pd.api.types.is_numeric_dtype(df[col]):
                fill_val = df[col].median()
                df[col].fillna(fill_val, inplace=True)
                report.append(f"列 '{col}' 缺失值已用中位数 {fill_val:.2f} 填充")
            else:
                mode_vals = df[col].mode()
                if not mode_vals.empty:
                    fill_val = mode_vals[0]
                    df[col].fillna(fill_val, inplace=True)
                    report.append(f"列 '{col}' 缺失值已用众数 '{fill_val}' 填充")

    for col in df.select_dtypes(include='number').columns:
        Q1, Q3 = df[col].quantile(0.25), df[col].quantile(0.75)
        IQR = Q3 - Q1
        lower, upper = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
        before_out = df.shape[0]
        df = df[(df[col] >= lower) & (df[col] <= upper)]
        if before_out - df.shape[0] > 0:
            report.append(f"列 '{col}' 剔除异常值: {before_out - df.shape[0]} 行")

    return df, report

# ---------- AI 总结函数 ----------
def generate_summary(dataframe):
    desc = dataframe.describe(include='all').to_string()
    cols = list(dataframe.columns)
    shape = dataframe.shape
    na = dataframe.isna().sum().to_string()

    prompt = f"""
你是一位擅长向非技术用户解释数据的分析师。请根据以下数据信息，用**通俗易懂的中文**写一份数据分析总结，让一个完全不了解这个数据的人也能明白。

数据基本信息:
- 表格大小: {shape[0]} 行 × {shape[1]} 列
- 列名: {cols}
- 缺失值情况: {na}
- 统计摘要: {desc}

请按以下结构输出总结（用小标题分开）:

1. 这个数据是关于什么的？
   用一两句话说清楚这个表格记录了什么主题。

2. 里面主要有哪些信息？
   列出主要的指标类型和数据维度，不要罗列所有列名，而是归类说明。

3.这个数据能用来干什么？
   结合实际场景，说明这份数据可能被用来做什么分析或决策。

4. 从这个数据里能看到什么？
   基于统计摘要，指出 1-2 个值得注意的模式或现象。

要求:
- 语言通俗，避免太多专业术语。
- 每条结论尽量说明"所以呢？对普通人意味着什么"。
- 总字数控制在 400 字以内。
"""
    response = client.chat.completions.create(
        model="qwen2.5:7b",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.5,
        max_tokens=800
    )
    return response.choices[0].message.content

# ---------- 文件上传 ----------
uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"])
if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet = st.selectbox("选择工作表", excel_file.sheet_names)

    # --- 智能表头检测 ---
    preview = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, nrows=5)
    header_row = 0
    for i, row in preview.iterrows():
        row_text = [str(v).strip() for v in row if pd.notna(v)]
        if any("时间" in t or "季度" in t or "指标" in t for t in row_text) or \
           (len(row_text) == len(row) and all(not t.isdigit() for t in row_text)):
            header_row = i
            break

    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row)
    df_raw = df_raw.dropna(how='all').reset_index(drop=True)

    # --- 侧边栏清洗开关 ---
    st.sidebar.header("🔧 数据预处理")
    do_clean = st.sidebar.checkbox("启用自动数据清洗", value=False)
    if do_clean:
        df, report = clean_dataframe(df_raw.copy())
        st.sidebar.success("清洗完成")
        with st.sidebar.expander("查看清洗报告"):
            for line in report:
                st.write(f"- {line}")
    else:
        df = df_raw

    # --- 数据概览 ---
    st.subheader("数据预览")
    st.dataframe(df.head())
    c1, c2, c3 = st.columns(3)
    c1.metric("行数", df.shape[0])
    c2.metric("列数", df.shape[1])
    c3.metric("缺失值总数", df.isna().sum().sum())

    # --- AI 总结按钮 ---
    if st.button("生成 AI 分析总结"):
        with st.spinner("本地 AI 正在分析数据，请稍候..."):
            summary = generate_summary(df)
            st.subheader("📝 AI 分析总结")
            st.write(summary)