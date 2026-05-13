import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
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
st.title("AI Excel 数据分析助手")
st.write("上传 Excel，自动清洗数据，并用本地 AI 生成分析总结。支持多轮追问。")

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
    st.sidebar.header("数据预处理")
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

    # ---------- 初始化对话历史 ----------
    if "messages" not in st.session_state:
        st.session_state.messages = []

    if "last_df_shape" not in st.session_state:
        st.session_state.last_df_shape = None
    if st.session_state.last_df_shape != df.shape:
        st.session_state.messages = []
        st.session_state.last_df_shape = df.shape

    # ---------- 初次分析按钮 ----------
    if len(st.session_state.messages) == 0:
        if st.button("生成 AI 分析总结", type="primary"):
            with st.spinner("本地 AI 正在分析数据，请稍候..."):
                data_context = f"""你是一位擅长向非技术用户解释数据的分析师。
当前分析的数据表格包含 {df.shape[0]} 行 × {df.shape[1]} 列。
列名: {list(df.columns)}
前几行数据预览:
{df.head(3).to_string()}

请用通俗中文总结：1.这数据是什么 2.主要信息 3.用途 4.一个值得注意的现象。
总字数控制在 400 字以内。"""

                st.session_state.messages = [
                    {"role": "system", "content": data_context},
                    {"role": "user", "content": "请帮我分析这份数据。"}
                ]

                response = client.chat.completions.create(
                    model="qwen2.5:7b",
                    messages=st.session_state.messages,
                    temperature=0.5,
                    max_tokens=600
                )
                reply = response.choices[0].message.content
                st.session_state.messages.append({"role": "assistant", "content": reply})
                st.rerun()

    # ---------- 显示对话历史 ----------
    for msg in st.session_state.messages:
        if msg["role"] in ["user", "assistant"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    # ---------- 多轮对话输入框 ----------
    if len(st.session_state.messages) > 0:
        if prompt := st.chat_input("继续追问，例如：哪个季度数值最高？、这数据能用来做什么决策？"):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("思考中..."):
                    api_messages = [{"role": m["role"], "content": m["content"]}
                                    for m in st.session_state.messages
                                    if m["role"] in ["system", "user", "assistant"]]

                    response = client.chat.completions.create(
                        model="qwen2.5:7b",
                        messages=api_messages,
                        temperature=0.7
                    )
                    reply = response.choices[0].message.content
                    st.markdown(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})