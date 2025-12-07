# 🎨 AI PPT 自動改版工具

使用 **Streamlit + Ollama 本地模型**，可以將你的 PPT 依照指定風格自動重新設計。
完全免費，不需要 OpenAI API Key。

**▶️ 線上 Demo：[點我打開](https://ppt-format-conversion-ovlfr8pbryqdary55uofrp.streamlit.app/)**

---

## 🛠 功能

* 上傳 `.pptx` 檔案
* 選擇預設風格或自訂風格
* 自動生成新 PPT
* 下載改版後的 PPT
* 使用本地 Ollama 模型（LLaMA3 / Qwen）

---

## 📂 專案結構

```
ai-ppt-redesign/
│
├─ app.py                 # 主程式
├─ requirements.txt       # Python 套件
└─ README.md              # 使用說明
```

---

## 💻 系統需求

* Python 3.10+
* Streamlit
* python-pptx
* **Ollama 已安裝並可在命令列使用**

> ❗❗ **Windows 使用者務必注意**
> 請確保 **`ollama.exe` 已加入系統 PATH**，或在程式裡指定完整路徑
> 否則程式無法順利呼叫本地模型。

---

## ⚡ 安裝與執行

1. Clone 專案

```bash
git clone https://github.com/你的帳號/ai-ppt-redesign.git
cd ai-ppt-redesign
```

2. 安裝 Python 套件

```bash
pip install -r requirements.txt
```

3. 安裝 Ollama

前往官方網站下載並安裝：

[Ollama 官方下載](https://ollama.com/download)

確認可以執行：

```bash
ollama --version
```

4. 執行專案

```bash
streamlit run app.py
```

---

## 🖌 使用說明

1. 選擇風格（或在自訂風格區輸入）
2. 上傳 PPT 檔案
3. 按下「開始轉換」
4. 等待 AI 重新生成 PPT
5. 下載全新 PPT

---

## ⚠️ 注意事項

* Ollama **必須安裝**，程式才會順利呼叫模型
* Windows 使用者如果遇到 `FileNotFoundError`，請確認 **`ollama.exe` 在系統 PATH**，或在程式裡指定完整路徑
* Streamlit 每次操作會 rerun app，所以風格選擇已改用 `st.session_state` 記錄
* 模型生成速度取決於本地電腦性能

---

## 📌 建議

* 建議使用 LLaMA3.1 或 Qwen 2.5 模型
* 可自行修改 `PRESET_STYLES`，新增更多自訂風格
* 適合用於快速 Demo 或簡報設計初稿

---

## ❤️ 作者

林義庭
