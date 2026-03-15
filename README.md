# Excel Expert · Financial Intel Suite 📊

**Excel Expert** is a high-performance AI toolset specifically engineered for finance professionals in Asset Management, Investment Banking, FP&A, and Equity Research. It integrates Large Language Model (LLM) reasoning directly into financial workflows to automate complex data reconciliation, logical audits, and advanced formula engineering.

The project provides two flexible deployment options:
1.  **Excel Add-in**: A deep integration that lives in the Office task pane, ideal for processing massive workbooks.
2.  **Web Agent**: A standalone, privacy-focused HTML tool (`excel-expert-agent-final.html`) for instant, local data processing.

---

## ✨ Core Features

* **Financial-Grade Rigor**: Pre-configured logic optimized for financial datasets like valuation models, general ledgers, and bank statement reconciliations.
* **Dual-Phase Execution Architecture**:
    * **Phase 1: Logic Verification**: Test the AI's algorithm on a small sample (approx. 20 rows) to ensure accuracy before full execution.
    * **Phase 2: Full Processing**: Once verified, apply the logic to the entire dataset with a single click.
* **Intelligence Knowledge Base (KB)**: Allows users to inject firm-specific terminology, internal accounting policies, or proprietary formula libraries.
* **Privacy-First Design**: Raw data remains in the local browser environment. Only anonymized structures are sent to the AI provider.
* **Multi-Engine Support**: Seamlessly switch between **Anthropic (Claude 3.5 Sonnet)**, **DeepSeek**, or **Tencent Yuanbao**.

---

## 📁 Project Structure

Based on the latest repository build:

```text
/excel-expert-suite
├── /excel-expert-addin      # Excel Add-in Source
│   ├── /src                 # Frontend UI and logic
│   ├── manifest.xml         # Office Add-in configuration
│   ├── server.js            # Node.js proxy (handles CORS for DeepSeek/Yuanbao)
│   ├── SETUP.bat            # Dependency installation script
│   └── START.bat            # Local server launch script
└── excel-expert-agent-final.html  # Standalone Web Agent
