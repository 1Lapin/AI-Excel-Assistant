# AI Excel 数据分析助手

一个基于大语言模型的智能 Excel 分析工具，上传表格即可自动清洗数据并生成通俗易懂的自然语言总结。

##  核心功能
-  **文件上传**：支持 `.xlsx` / `.xls`，自动识别工作表
-  **智能数据清洗**：处理缺失值、重复行、异常值（IQR 法），生成清洗报告
- **AI 解读**：调用本地/云端大模型，用“人话”解释数据含义和业务价值
-  **可视化预览**：数据概览、统计指标一目了然

##  技术栈
- **前端交互**：Streamlit
- **数据处理**：Pandas
- **AI 模型**：Ollama (本地 Qwen2.5-7B) / DeepSeek / Moonshot API
- **可视化**：Matplotlib

##  快速开始
1. 克隆仓库：`git clone https://github.com/你的用户名/ai-excel-analyst.git`
2. 安装依赖：`pip install -r requirements.txt`
3. 启动 Ollama 并下载模型：`ollama pull qwen2.5:7b`
4. 运行：`streamlit run app.py`
5. 上传 Excel，点击“生成 AI 分析总结”

##  项目结构
├── app.py            # Streamlit 主程序

├── requirements.txt  # 依赖列表

├── .env.example      # 环境变量模板

└── README.md

## 效果展示
<img width="2403" height="864" alt="image" src="https://github.com/user-attachments/assets/0f07b463-1cb7-443c-8fcf-923609156ec3" />
<img width="2451" height="1334" alt="image" src="https://github.com/user-attachments/assets/64365f95-93e0-4cb6-9757-0f6286f7c54f" />
