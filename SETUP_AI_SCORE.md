# AI Success Score — Setup

## Files
- `CloseDocGenV6.py` — your app, now with a **🎯 Success Score** tab.
- `ai_success.py` — the scoring engine (import is already wired in).
- `requirements.txt` — adds a pinned `openai>=1.0.0`.

## Provide the OpenAI key
The key is read from `st.secrets` first, then the environment — so the same
code runs locally and on Streamlit Cloud.

**Local:** create a `.env` file next to the app:
```
OPENAI_API_KEY=sk-...
```
(`python-dotenv` is in requirements; or just `export OPENAI_API_KEY=...`.)

Or use Streamlit's secrets locally — create `.streamlit/secrets.toml`:
```toml
OPENAI_API_KEY = "sk-..."
```

**Streamlit Cloud:** App → Settings → Secrets, paste:
```toml
OPENAI_API_KEY = "sk-..."
```

No key? The app still works — it falls back to an offline heuristic and
labels the result accordingly.

## How it scores
Six weighted criteria, each 0–100, combined into a Final Success Score:

| Criterion | Default weight | Scored by |
|---|---|---|
| Objectives Achievement | 30% | AI |
| Deliverables Completion | 25% | AI |
| Budget Adherence | 15% | Python (deterministic) |
| Schedule Adherence | 10% | Python (deterministic) |
| Risk Management | 10% | AI |
| Stakeholder & Quality | 10% | AI |

Budget and schedule are computed in code (no AI math, fully reproducible).
The AI only judges the qualitative sections and returns strict JSON.

**Bands:** 85+ Highly Successful · 70–84 Successful · 50–69 Partially
Successful · 30–49 Marginal/At Risk · <30 Unsuccessful.

Weights are adjustable via sliders in the tab and are auto-normalised, so they
need not sum to 100%. The computed score is saved into (and restored from) your
draft JSON, and—if enabled—added to the Word and PDF exports as a numbered
"Project Success Assessment" section.

## Run
```
pip install -r requirements.txt
streamlit run CloseDocGenV6.py
```
