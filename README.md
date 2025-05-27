# 📘 AI-Powered Learning Companion

An advanced desktop application that transforms your PDFs into an interactive learning experience using local AI models from Ollama. With PDF parsing, voice support, chat-based tutoring, vision analysis, and study material generation, this is your personal AI tutor—right on your desktop.

---

## ✨ Features

- 🔍 **PDF Analysis**: Load PDFs and extract text + images per page.
- 💬 **AI Chat**: Ask questions, explain concepts, or analyze text with your selected Ollama model.
- 🧠 **Personalities**: Choose from a wide range of AI tutor personalities (Socratic, Comedian, Motivator, etc.).
- 🎨 **Vision Support**: Use multimodal models to analyze diagrams and figures.
- 🧑‍🏫 **Explain Concepts**: Select text and get AI-powered explanations, summaries, and analogies.
- 🧪 **Study Material Generator**: Auto-generate summaries, quizzes, and key points.
- 🔊 **Text-to-Speech (TTS)**: Let AI responses or PDF pages be read aloud using edge-tts.
- 🎙️ **Voice Query**: Ask questions using your voice (requires `SpeechRecognition`).
- 📦 **Runs Locally**: No cloud dependencies – fully local with Ollama backend.

---

## 🛠 Installation

1. **Install Python 3.10+**
2. **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
3. **Install and start Ollama**:  
   [https://ollama.com](https://ollama.com)  
   Example:
   ```bash
   ollama serve
   ollama pull llama3
