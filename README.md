# Meeting Summarizer

AI-powered Streamlit web app: Upload a meeting video → auto-transcribe → extract key points → download PowerPoint.

## Features
- Upload any meeting video (.mp4, .mov, .mkv, .avi, .webm)
- Transcription using OpenAI Whisper (runs locally, 100% free)
- Summarization using Google Gemini Flash (free tier — 1,500 req/day)
- Auto-generated 6-slide PowerPoint: Title, Summary, Key Points, Decisions, Action Items, Next Steps
- Runs in any browser — works on VDI, no special tools required

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Install FFmpeg
```bash
# Windows
winget install ffmpeg

# Mac
brew install ffmpeg

# Linux
sudo apt install ffmpeg
```

### 3. Add your free Gemini API key
Create `.streamlit/secrets.toml`:
```toml
GEMINI_API_KEY = "your-key-here"
```
Get a free key (no credit card): https://aistudio.google.com/apikey

### 4. Run the app
```bash
streamlit run app.py
```
Open http://localhost:8501 in your browser.

## Architecture
```
Video Upload (browser)
    ↓
FFmpeg  →  extract WAV audio
    ↓
Whisper  →  speech to text (local, free)
    ↓
Gemini Flash  →  key points, decisions, action items (JSON)
    ↓
python-pptx  →  build 6-slide .pptx
    ↓
Download button in browser
```

## Built for IBM WatsonX Challenge
