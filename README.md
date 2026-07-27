# 智能简历匹配与改写工具（网页版 v3）

上传岗位描述 → 自动分析核心要求 → 上传简历并匹配 Top 3 → 智能改写 → 导出 PDF。
核心匹配引擎（jieba 分词）跑在服务器端，无需任何 API Key 即可使用；LLM 智能改写为可选项，
由访问者自行在页面里输入自己的 Key，只保存在当次浏览器会话中，刷新页面即清除。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`。

可选增强（本地）：
- Word→PDF 效果更好：安装 [LibreOffice](https://zh-cn.libreoffice.org/download/)
- 图片提取文字（OCR）：macOS 执行 `brew install tesseract`

## 免费部署到 Streamlit Community Cloud

1. 把这个文件夹推送到你自己的 GitHub 仓库（public）。
2. 打开 https://share.streamlit.io ，用 GitHub 账号登录。
3. 点击 "New app" → 选择刚才的仓库 → Main file path 填 `app.py` → 点击 Deploy。
4. 等待几分钟首次构建完成（会根据 `requirements.txt` 装 Python 依赖、根据 `packages.txt` 装
   LibreOffice / Tesseract 等系统依赖），之后你会拿到一个形如
   `https://your-app-name.streamlit.app` 的公开链接，在任何电脑的浏览器打开即可使用。

无需在 Streamlit Cloud 后台配置任何 Secrets —— LLM 改写的 API Key 是每次访问者自己在
侧边栏输入的，不需要也不会预置任何人的密钥。

## 关于隐私和费用

- 这是一个公开链接，任何拿到链接的人都可以打开使用，也能在 GitHub 上看到源代码。
- 仓库里**不包含**任何真实简历文件或 API Key —— 简历通过网页上传，处理完即用即焚
  （不同访问者之间的上传文件互不可见，也不会持久化保存）。
- 如果不想让链接被搜到 / 想限制访问，可以后续加一层密码保护（未包含在当前版本，
  想要的话可以再加）。

## 目录说明

- `app.py` — 应用主程序（Streamlit）
- `requirements.txt` — Python 依赖
- `packages.txt` — Streamlit Cloud 构建时通过 apt 安装的系统依赖（LibreOffice、Tesseract、中文字体）
