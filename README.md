---
title: SchoolSheet Studio
emoji: 📘
colorFrom: blue
colorTo: indigo
sdk: gradio
sdk_version: 4.44.0
app_file: app.py
pinned: false
---

# SchoolSheet Studio — Professional Worksheet Generator

A worksheet generator for Grades 1–10: curriculum-based or free-content AI
question generation (Anthropic Claude or Google Gemini), plus an offline
Question Bank mode. Exports to PDF or Word, with correct Urdu rendering.

## Deploying this Space

1. Create a new Space on Hugging Face → SDK: **Gradio**.
2. Upload `app.py`, `requirements.txt`, and this `README.md` to the Space
   (or push them to the Space's git repo).
3. Go to the Space's **Settings → Variables and secrets** and add (as
   **Secrets**, not public variables):
   - `ANTHROPIC_API_KEY` — if you want to use Claude
   - `GEMINI_API_KEY` — if you want to use Gemini instead
   - `AI_PROVIDER` — optional; set to `anthropic` or `gemini` to force a
     provider. If left unset, the app auto-picks Anthropic if that key is
     present, otherwise Gemini, otherwise runs Question-Bank-only.
   - `PIXABAY_API_KEY` — optional, for real photo illustrations on young
     grades' worksheets.
4. The Space will build and give you a permanent public URL like
   `https://huggingface.co/spaces/<your-username>/schoolsheet-studio` —
   that's the link to share with organizations.

No key is required to run the app — without one, users can still generate
worksheets fully offline via the built-in Question Bank (CSV/Excel upload).
