# Excel AI Assistant

A comprehensive Excel add-in combining local AI chat (LM Studio / Ollama), advanced data analysis, smart chart generation, and automatic anomaly & trend detection — all running 100 % on your machine for complete data privacy.

---

## Features

| Feature | Description |
|---|---|
| 🤖 **Local AI Chat** | Chat with your spreadsheet data via LM Studio or Ollama. No cloud. No data leaves your machine. |
| 📊 **Data Analysis** | Descriptive statistics (mean, median, std, IQR, skewness, kurtosis) per column, with type inference. |
| 📈 **Smart Charting** | Auto-selects the best chart type based on column types (Pie, Line, Scatter, Histogram, Bar, Column…). |
| 🔍 **Anomaly Detection** | IQR + Z-score outlier detection with severity ranking. One-click cell highlighting in Excel. |
| 📉 **Trend Detection** | Linear regression trend detection with R² strength rating per numeric column. |
| 🎤 **Voice Input** | Record voice messages (Web Speech API) transcribed directly into the chat input. |
| 📷 **Video Capture** | Capture a webcam frame and describe it to the AI as visual context. |
| 📎 **File Attach** | Attach CSV / TXT / JSON files to chat; previewed and sent as context. |
| 🌑 **Dark Theme** | Full dark-themed task-pane UI designed for extended spreadsheet work. |
| 🔒 **Complete Privacy** | All inference is local (LM Studio / Ollama). Settings are persisted in `localStorage`. |

---

## Architecture

```
excel-ai-assistant/
├── manifest.xml                  Excel add-in manifest
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html         Task-pane UI (dark theme, tabs, chat, media)
│   │   ├── taskpane.css          Dark-theme stylesheet
│   │   └── taskpane.js           Main controller (Office.js + AI + analysis)
│   ├── commands/
│   │   └── commands.html         Add-in ribbon command host page
│   ├── utils/
│   │   ├── ai-client.js          LM Studio / Ollama streaming client
│   │   ├── excel-api.js          Excel API read/write/chart helpers
│   │   ├── data-analyzer.js      Pure-JS stats, anomaly & trend detection
│   │   └── chart-manager.js      Smart chart recommendation + creation
│   ├── __tests__/                Jest unit tests
│   └── __mocks__/                CSS mock for Jest
├── backend/
│   ├── app.py                    FastAPI server (AI proxy + advanced analysis)
│   ├── analyzer.py               NumPy / pandas / SciPy analysis engine
│   ├── requirements.txt          Python dependencies
│   └── test_analyzer.py          pytest test suite
├── package.json
├── webpack.config.js
└── .gitignore
```

---

## Quick Start

### Prerequisites

- **Node.js** ≥ 18
- **Python** ≥ 3.10 (optional – only needed for advanced server-side analysis)
- **LM Studio** or **Ollama** running locally with at least one model loaded

### 1 – Install the add-in (JavaScript)

```bash
npm install
npm run build:dev          # development build
# or
npm start                  # start dev server at https://localhost:3000
```

#### Sideload into Excel

1. In Excel → **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**
2. Browse to `manifest.xml`
3. The **AI Assistant** button appears in the **Home** ribbon tab

### 2 – Start the Python backend (optional)

```bash
cd backend
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

The backend provides more powerful server-side analysis (Isolation Forest, full scipy stats).  
If it is not running, all analysis falls back to the built-in JavaScript engine.

### 3 – Configure LM Studio / Ollama

Open the **⚙️ Settings** panel inside the add-in and set:

| Setting | LM Studio default | Ollama default |
|---|---|---|
| AI Backend URL | `http://localhost:1234/v1` | `http://localhost:11434/v1` |
| Model name | *(auto-detected)* | e.g. `llama3` |

---

## Usage Guide

### Chat Tab 💬
- Type a question and press **Ctrl+Enter** or click **➤**
- Toggle **📋 Include data** to automatically inject the selected Excel range as context
- Use **🎤** for voice input, **📷** for a webcam snapshot, **📎** to attach a file

### Analyze Tab 📊
- **Analyse Selection** – statistics for the currently selected range
- **Analyse Entire Sheet** – statistics for the used range of the active sheet

### Charts Tab 📈
- **✨ Auto Chart** – the AI picks the best chart type for your selection
- Manual buttons for Column, Line, Pie, Bar, Area, Scatter, Histogram, Box Plot

### Detect Tab 🔍
- **Detect Anomalies** – finds outliers using IQR + Z-score
- **Detect Trends** – linear regression direction per numeric column
- **Highlight Anomalies** – colours anomalous cells 🔴 High / 🟠 Medium / 🟡 Low

---

## Running Tests

### JavaScript (Jest)
```bash
npm test
```
54 unit tests covering `data-analyzer.js`, `chart-manager.js`, and `ai-client.js`.

### Python (pytest)
```bash
cd backend
pytest test_analyzer.py -v
```
30 unit tests covering `_infer_type`, `_safe_float`, `analyze_dataframe`, `detect_anomalies`, and `detect_trends`.

---

## Privacy & Security

- All LLM inference runs locally via LM Studio or Ollama — no API keys, no cloud.
- No data is ever sent outside `localhost`.
- Settings (backend URL, model name, temperature) are stored in browser `localStorage` only.
- The Python backend, if used, also runs entirely on-premise.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Add-in UI | HTML5, CSS3 (dark theme), vanilla ES2022 modules |
| Office integration | Office.js Excel API |
| AI connectivity | OpenAI-compatible REST API (LM Studio / Ollama) |
| Client-side analysis | Pure JS (IQR, Z-score, linear regression, descriptive stats) |
| Server-side analysis | Python, FastAPI, NumPy, pandas, SciPy, scikit-learn |
| Build | webpack 5, Babel |
| Tests | Jest (JS), pytest (Python) |

