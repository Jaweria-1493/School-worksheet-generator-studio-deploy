# SchoolSheet Studio — Professional Worksheet Generator

A worksheet generator for Grades 1–10: curriculum-based or free-content AI
question generation (Anthropic Claude or Google Gemini), plus an offline
Question Bank mode. Exports to PDF or Word, with correct Urdu rendering.

## Deploying this app on Streamlit Community Cloud (free, no credit card)

1. Push `app.py`, `requirements.txt`, and this `README.md` to a GitHub
   repository (they can be the only files in the repo).
2. Go to **share.streamlit.io**, sign in with GitHub, and click
   **"New app"**.
3. Select your repository, branch (usually `main`), and set
   **Main file path** to `app.py`.
4. Before/after deploying, open **Settings → Secrets** on the app and add
   (in TOML format):

   ```
   ANTHROPIC_API_KEY = "your-key-here"
   # or:
   GEMINI_API_KEY = "your-key-here"
   AI_PROVIDER = "anthropic"   # or "gemini" — optional, auto-detected if omitted
   PIXABAY_API_KEY = "your-key-here"   # optional, for real photo illustrations
   ```

5. Click **Deploy**. In a couple of minutes you'll get a permanent URL like
   `https://your-app-name.streamlit.app` — that's the link to share with
   organizations.

No AI key is required to run the app — without one, users can still
generate worksheets fully offline via the built-in Question Bank
(CSV/Excel upload, Method 3).
