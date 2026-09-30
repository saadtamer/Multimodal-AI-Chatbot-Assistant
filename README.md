<div align="center">

# 🤖 Multimodal AI Chatbot Assistant
### **Conversational & Document Intelligence Platform powered by Streamlit & OpenRouter**

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B.svg?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![OpenRouter](https://img.shields.io/badge/OpenRouter-API-7C3AED.svg?style=for-the-badge)](https://openrouter.ai)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg?style=for-the-badge)](LICENSE)

<br/>

<img src="https://capsule-render.vercel.app/api?type=waving&color=gradient&customColorList=6,14,26,45&height=180&section=header&text=Multimodal%20AI%20Chatbot%20Assistant&fontSize=32&fontColor=ffffff&animation=fadeIn" width="100%"/>

</div>

---

## 🌟 Overview
**Multimodal AI Chatbot Assistant** is a responsive conversational AI application engineered by **Eng. Saad Tamer Abo-Elazm**. Built using **Streamlit** and integrated with **OpenRouter API**, it provides an intuitive web interface for multimodal interaction across text, images, speech/audio, and complex document formats.

---

## ✨ Key Features

* 👁️ **Multimodal Vision Intelligence:**
  * Analyzes user-uploaded images and documents in real-time.
  * Native integration with vision models (`google/gemma-4-31b-it`, `google/gemma-4-26b-a4b-it`, `nvidia/nemotron-nano-12b-v2-vl`).
* 🎙️ **Voice & Audio Reasoning:**
  * Audio input support powered by `nvidia/nemotron-3-nano-omni-30b-a3b-reasoning`.
* 📄 **Intelligent Document Processing:**
  * **PDFs (`PyPDF2`):** Extracts text with configurable chunking up to 18,000 characters.
  * **Word Documents (`python-docx`):** Parses nested paragraphs, bullet points, and clinical tables.
  * **Tabular Data (`Pandas`):** Ingests CSV files, previews up to 25 sample rows, and extracts column distributions for precise statistical Q&A.
* 🔄 **Dynamic Fallback & Failover Routing:**
  * Automatic model failover sequence if primary models experience rate limits or network latency.
* 🧠 **Session Memory & Context Window:**
  * Sliding chat history window keeping the last 12 conversational turns for coherent long-term discussions without context overflow.

---

## 🛠️ Tech Stack

* **Frontend / UI:** Streamlit
* **AI Orchestration & LLMs:** OpenRouter API (Gemma, Nemotron Omni, Llama)
* **Document Parsing:** PyPDF2, python-docx, Pandas
* **Audio & Media:** Base64 binary serialization & Tempfile handlers

---

## 🚀 Getting Started

### 1. Clone & Setup
```bash
git clone https://github.com/saadtamer/Multimodal-AI-Chatbot-Assistant.git
cd Multimodal-AI-Chatbot-Assistant
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure Environment
Create a `.env` file in the root directory:
```env
OPENROUTER_API_KEY=your_openrouter_api_key_here
```

### 3. Launch the Application
```bash
streamlit run main1.py
```

---

<div align="center">
Developed with ❤️ by <b>Eng. Saad Tamer</b> | AI & Data Engineer
</div>
