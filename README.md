# Excel-Helper
A high-performance AI assistant for financial data processing. Features a deep-integrated Excel Add-in and a privacy-first Web agent for complex modeling, reconciliation, and automated reporting.

# Excel Expert · Financial Intel Suite 📊

Excel Expert 是一款面向泛金融领域（资产管理、投行、财务分析及投研）的 AI 增强工具集。它旨在将大模型的推理能力深度整合进金融办公流，解决高压力场景下的数据对齐、逻辑核查与复杂公式编写任务。

本项目包含两个版本：
1.  **Excel Add-in (插件版)**: 直接嵌入 Office 侧边栏，适合处理庞大的工作簿。
2.  **Web Agent (网页版)**: 独立单文件 HTML，支持离线运行，满足极高的数据保密需求。

## ✨ 核心特性

- **金融级严谨逻辑**: 内置针对金融数据（如估值表、科目明细、流水对账）优化的推理逻辑。
- **双阶段执行架构**: 
    - **逻辑验证 (Verification)**：先用 20 行样本数据测试 AI 的算法方案，确保逻辑无偏。
    - **全量处理 (Full Processing)**：验证通过后，一键自动化处理完整数据集。
- **动态知识库 (Intel KB)**: 允许注入特定行业术语、公司会计政策或复杂的金融公式库。
- **隐私与合规**: 原始数据保留在本地浏览器环境，仅传输脱敏后的上下文及表结构。
- **多引擎支持**: 灵活切换 Anthropic (Claude 3.5 Sonnet)、DeepSeek 或腾讯元宝等主流金融适配模型。

## 📁 项目结构

```text
/excel-expert-suite
├── /excel-expert-addin      # 插件源代码
│   ├── /src                 # 前端界面逻辑
│   ├── manifest.xml         # Office 加载项配置文件
│   ├── server.js            # 本地 Node.js 代理服务器
│   ├── SETUP.bat            # 快速安装依赖
│   └── START.bat            # 启动本地服务
└── excel-expert-agent-final.html  # 独立网页版（即开即用）
