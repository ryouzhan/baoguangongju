# SmartCustoms Pro | 跨境出口报关协同中台 (Streamlit Cloud 部署版)

专为直接部署在 **[streamlit.io (Streamlit Community Cloud)](https://streamlit.io/)** 打造的现代企业级出口报关协同平台。

---

## 🌟 核心业务架构（3 合 1 高效精简版）

### 1. 📑 报关单据生成 (采购单云端联动)
- **AirScript Webhook 云端直连**：自动读取 WPS 金山文档云端最新采购单数据（单价、品名、单位、HS编码、中英文品名、申报要素、境内货源地）。
- **零尾差数学精度保障**：严格遵循 `数量(整数) × 单价USD(2位) = 金额USD(2位)`，彻底消除由于汇率折算产生的尾差问题。
- **双表联动生成**：一键生成包含「明细」与「汇总」工作表的专业报关 Excel，汇总表具备 FBA 斑马纹交替着色。

### 2. 🔗 FBA 货件智能合并 (按地址归集)
- **智能去括号过滤**：自动清除物流中心地址中的备注字符（如 `ACY2(美东)` -> `ACY2`）。
- **多维度筛选归集**：支持下拉多选和直接粘贴多行 FBA 编号批量加入分组。
- **品名+单位自动聚合**：自动累加箱数、发货量、毛重、净重、体积、货值、金额USD，合并行无缝写入原汇总表尾部并一键导出。

### 3. 📋 报关套打直接生成 (集成智能清理与合同号规范重命名)
- **免二次手动清理**：将原有的第四个手动清理功能**直接深度融合**至第三个套版流程中。
- **就地清除空白模板行**：在填充数据后，系统自动检测该批单据实际品项条数（1~5 项），直接对【报关单】、【装箱单】、【发票】、【合同】多余的未填充行执行清空。
- **直接规范重命名**：自动从模板与数据源中提取「合同号」，直接将成品单据命名为 `{合同号}_报关单_{账号}_{FBA编号}.xlsx`，**不产生或保留任何未清理的过渡文件**。
- **一键 ZIP 打包**：所有生成的独立报关单一次性打包为 ZIP 供一键下载。

---

## 🚀 如何在 streamlit.io (Streamlit Community Cloud) 上部署？

部署到 `streamlit.io` 只需简单的 3 步：

### 第 1 步：将本文件夹代码推送到您的 GitHub 仓库
1. 在 GitHub 上新建一个仓库（例如命名为 `customs-streamlit-app`，设为 Public 或 Private 均可）。
2. 将本压缩包内的所有文件（包含 `streamlit_app.py`、`app.py`、`requirements.txt`、`.streamlit/config.toml` 等）上传或推送到该仓库的根目录。

### 第 2 步：登录 streamlit.io 并创建 App
1. 打开并登录 [share.streamlit.io](https://share.streamlit.io/)（使用 GitHub 账号授权登录）。
2. 点击右上角的 **「New app」**。
3. 选择刚才创建的 GitHub 仓库与分支（如 `main`）。
4. **Main file path** 填写：`streamlit_app.py`（或 `app.py`）。
5. 点击 **「Deploy!」**。

### 第 3 步：部署完成与在线使用
- Streamlit Cloud 将自动读取 `requirements.txt` 安装相关依赖，并在约 1~2 分钟内启动完毕。
- 启动后您将获得一个专属的公共访问网址（如 `https://your-customs-app.streamlit.app`），可在任何电脑、手机浏览器上随时随地协同使用！

---

## 💻 本地离线运行调试（可选）

如需在本地电脑运行测试：
```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动服务
streamlit run streamlit_app.py
```
浏览器将自动弹出并访问 `http://localhost:8501`。
