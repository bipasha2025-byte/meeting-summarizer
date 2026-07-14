import streamlit as st
import tempfile
import json
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import google.generativeai as genai
import time

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Meeting Summarizer",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .main-header {
    background: #1f2328;
    padding: 2rem; border-radius: 12px; margin-bottom: 2rem; color: white;
  }
  .main-header h1 { margin: 0; font-size: 2rem; }
  .main-header p  { margin: 0.5rem 0 0; color: #aaa; font-size: 0.95rem; }
  .slide-preview {
    background: white; border: 1px solid #e5e7eb;
    border-radius: 8px; padding: 1rem; margin-bottom: 0.5rem;
    min-height: 120px;
  }
  .slide-title { font-weight: 700; color: #3b82d4; margin-bottom: 0.5rem; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>Meeting Summarizer</h1>
  <p>Upload a meeting video &rarr; Gemini transcribes &amp; analyses &rarr; Download PowerPoint</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    default_key = ""
    try:
        default_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass

    api_key = st.text_input(
        "Gemini API Key",
        value=default_key,
        type="password",
        help="Free key from aistudio.google.com/apikey",
        placeholder="Your Gemini API key",
    )

    st.subheader("Meeting Info (optional)")
    meeting_title = st.text_input("Meeting Title", placeholder="e.g. Q4 Planning Session")
    meeting_date  = st.text_input("Meeting Date",  placeholder="e.g. 15 Jan 2025")

    st.divider()
    st.caption("Built for IBM WatsonX Challenge")
    st.caption("Powered by Google Gemini Flash + python-pptx")
    st.caption("No FFmpeg or Whisper needed — runs entirely in the cloud")


# ── Helpers ───────────────────────────────────────────────────────────────────
def transcribe_and_summarize(video_path, video_mime, title, date, key, progress):
    """Upload video to Gemini Files API, then transcribe + summarise in one call."""

    genai.configure(api_key=key)

    # Step 1: Upload the video file to Gemini Files API
    progress.progress(15, "Uploading video to Gemini (this may take a moment for large files)...")
    video_file = genai.upload_file(path=video_path, mime_type=video_mime)

    # Step 2: Wait for Gemini to finish processing the file
    progress.progress(30, "Gemini is processing the video...")
    max_wait = 300  # 5 minutes max
    waited   = 0
    while video_file.state.name == "PROCESSING":
        if waited >= max_wait:
            raise TimeoutError("Gemini took too long to process the video. Try a shorter clip.")
        time.sleep(5)
        waited += 5
        video_file = genai.get_file(video_file.name)

    if video_file.state.name == "FAILED":
        raise ValueError("Gemini failed to process the video file. Check the file is a valid video.")

    # Step 3: Send the processed video to Gemini with a combined transcribe + summarise prompt
    progress.progress(55, "Gemini is transcribing and extracting key points...")

    model = genai.GenerativeModel(
        model_name="gemini-flash-latest",
        generation_config={"response_mime_type": "application/json", "temperature": 0.2},
    )

    prompt = f"""You are an expert meeting analyst.
Watch and listen to this meeting video carefully.
First transcribe what is said, then extract the most important information.
Return ONLY valid JSON — no markdown fences, no extra text:
{{
  "title": "meeting title",
  "date": "meeting date",
  "transcript": "full transcript of the meeting speech",
  "executive_summary": "2-3 sentence high-level summary",
  "key_points": ["key discussion point 1", "key discussion point 2"],
  "decisions": ["decision agreed upon 1"],
  "action_items": [
    {{"task": "what needs to be done", "owner": "person or TBD", "due": "date or TBD"}}
  ],
  "next_steps": ["follow-up item 1"],
  "attendees": ["Name 1", "Name 2"]
}}

Rules:
- key_points must cover ALL major topics discussed
- decisions = only things explicitly agreed upon
- Extract action items even if implied; use TBD for unknown owners/dates

Meeting Title: {title or "Untitled Meeting"}
Meeting Date:  {date  or "Unknown"}"""

    result  = model.generate_content([video_file, prompt])
    summary = json.loads(result.text)

    # Clean up uploaded file from Gemini storage
    try:
        genai.delete_file(video_file.name)
    except Exception:
        pass

    progress.progress(80, "Analysis complete.")
    return summary


def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_title_slide(prs, title, date):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = hex_rgb("1f2328")

    txb = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(12.3), Inches(1.5))
    tf  = txb.text_frame
    tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title
    run.font.size  = Pt(36)
    run.font.bold  = True
    run.font.color.rgb = hex_rgb("ffffff")

    if date:
        txb2 = slide.shapes.add_textbox(Inches(0.5), Inches(3.7), Inches(12.3), Inches(0.5))
        p2   = txb2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = date
        run2.font.size  = Pt(16)
        run2.font.color.rgb = hex_rgb("aaaaaa")


def add_content_slide(prs, title, items, accent):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = hex_rgb("ffffff")

    # Heading
    txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.65))
    p   = txb.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size  = Pt(26)
    run.font.bold  = True
    run.font.color.rgb = hex_rgb(accent)

    # Accent line
    ln = slide.shapes.add_shape(1, Inches(0.5), Inches(1.1), Inches(12.3), Emu(12700))
    ln.fill.background()
    ln.line.color.rgb = hex_rgb(accent)
    ln.line.width = Emu(25400)

    # Body bullets
    txb2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(12.3), Inches(5.5))
    tf2  = txb2.text_frame
    tf2.word_wrap = True
    for i, item in enumerate(items):
        para = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
        para.space_after = Pt(6)
        run2 = para.add_run()
        if isinstance(item, dict):
            run2.text = (
                f"  {item.get('task', '')}   |   "
                f"Owner: {item.get('owner', 'TBD')}   |   "
                f"Due: {item.get('due', 'TBD')}"
            )
        else:
            run2.text = f"  {item}"
        run2.font.size  = Pt(14)
        run2.font.color.rgb = hex_rgb("1f2328")


def generate_ppt(summary, progress):
    progress.progress(85, "Building PowerPoint slides...")

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs, summary.get("title", "Meeting Summary"), summary.get("date", ""))

    if summary.get("executive_summary"):
        add_content_slide(prs, "Executive Summary", [summary["executive_summary"]], "3b82d4")

    if summary.get("key_points"):
        add_content_slide(prs, "Key Points", summary["key_points"], "3b82d4")

    if summary.get("decisions"):
        add_content_slide(prs, "Decisions Made", summary["decisions"], "7c5cd8")

    if summary.get("action_items"):
        add_content_slide(prs, "Action Items", summary["action_items"], "15803d")

    if summary.get("next_steps"):
        add_content_slide(prs, "Next Steps", summary["next_steps"], "c2410c")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    tmp.close()
    data = Path(tmp.name).read_bytes()
    os.unlink(tmp.name)
    progress.progress(100, "Done!")
    return data


# ── Main UI ───────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("1  Upload Meeting Video")
    uploaded = st.file_uploader(
        "Drop your video here",
        type=["mp4", "mov", "mkv", "avi", "webm"],
        help="Zoom, Teams, Meet, or any screen recording",
    )
    if uploaded:
        st.video(uploaded)
        st.caption(f"{uploaded.name}  ({uploaded.size / 1024 / 1024:.1f} MB)")

with col2:
    st.subheader("2  Generate Summary + PPT")

    if not api_key:
        st.info("Enter your Gemini API key in the sidebar to get started.")
    elif not uploaded:
        st.info("Upload a meeting video on the left to begin.")
    else:
        if st.button("Generate Summary + PPT", type="primary", use_container_width=True):
            progress = st.progress(0, "Starting...")
            try:
                # Save uploaded file to a temp path on disk
                suffix   = Path(uploaded.name).suffix.lower()
                mime_map = {
                    ".mp4": "video/mp4", ".mov": "video/quicktime",
                    ".mkv": "video/x-matroska", ".avi": "video/x-msvideo",
                    ".webm": "video/webm",
                }
                video_mime = mime_map.get(suffix, "video/mp4")

                progress.progress(5, "Saving uploaded file...")
                tmp_video = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
                tmp_video.write(uploaded.getbuffer())
                tmp_video.close()

                # Transcribe + summarise in one Gemini call
                summary = transcribe_and_summarize(
                    tmp_video.name,
                    video_mime,
                    meeting_title or uploaded.name,
                    meeting_date or "",
                    api_key,
                    progress,
                )
                st.session_state["transcript"] = summary.get("transcript", "")
                st.session_state["summary"]    = summary

                # Build PPT
                ppt_bytes = generate_ppt(summary, progress)
                st.session_state["ppt_bytes"] = ppt_bytes
                os.unlink(tmp_video.name)

            except Exception as e:
                progress.empty()
                st.error(f"Error: {e}")
                st.stop()

        if "summary" in st.session_state:
            summary   = st.session_state["summary"]
            ppt_bytes = st.session_state["ppt_bytes"]
            fname = (summary.get("title") or "meeting_summary").replace(" ", "_") + ".pptx"

            st.success("Processing complete!")
            st.download_button(
                label="Download PowerPoint (.pptx)",
                data=ppt_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )

            if st.session_state.get("transcript"):
                with st.expander("View Transcript"):
                    st.text_area(
                        "Transcript",
                        st.session_state["transcript"],
                        height=200,
                        label_visibility="collapsed",
                    )

# ── Slide Preview ─────────────────────────────────────────────────────────────
if "summary" in st.session_state:
    s = st.session_state["summary"]
    st.divider()
    st.subheader("Slide Preview")

    slides = [
        ("Slide 1 — Title",            [s.get("title", ""), s.get("date", "")]),
        ("Slide 2 — Executive Summary", [s.get("executive_summary", "")]),
        ("Slide 3 — Key Points",        s.get("key_points", [])),
        ("Slide 4 — Decisions Made",    s.get("decisions", [])),
        ("Slide 5 — Action Items",
         [f"{a.get('task','')} | {a.get('owner','TBD')} | {a.get('due','TBD')}"
          for a in s.get("action_items", [])]),
        ("Slide 6 — Next Steps",        s.get("next_steps", [])),
    ]

    cols = st.columns(3)
    for i, (slide_title, items) in enumerate(slides):
        with cols[i % 3]:
            content = "".join(
                f"<div style='font-size:12px;margin:2px 0'>• {it}</div>"
                for it in items if it
            )
            st.markdown(
                f'<div class="slide-preview">'
                f'<div class="slide-title">{slide_title}</div>{content}</div>',
                unsafe_allow_html=True,
            )
