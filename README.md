# Gen AI Champions Dashboard

A Streamlit dashboard that reads the `AI_Champions_Initiative_Report.xlsx`
tracking sheet and displays it as a live web dashboard, with a **Sync Now**
button to re-read the file on demand, an AI chatbot grounded in the sheet
data, and an AI-drafted monthly status summary — both powered by your org's
**LLM Farm**.

## 🎉 New Features (Enhanced Version)

### Additional Activity Information
- **Start Date, Completion Date, Duration**: Track timing for each activity
- **Docupedia Links**: Clickable links to documentation pages
- Automatic display of timeline information

### Activity Sessions (One-to-Many Relationship)
- Support for activities with multiple sessions/events
- Example: AI Tech Talk with multiple completed and upcoming sessions
- Expandable rows to view session details
- Status-based color coding (Completed, Ongoing, Planned)

### Generic & Reusable Implementation
- Works with any activity without hardcoding
- Uses separate "Activity Sessions" Excel sheet
- Backward compatible with existing Excel files

📚 **See [ENHANCEMENT_GUIDE.md](ENHANCEMENT_GUIDE.md) for detailed documentation on new features and Excel structure.**

## 1. Setup

```bash
cd gen-ai-dashboard
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configure

Copy the env template and fill in your values:

```bash
cp .env.example .env
```

Edit `.env`:

| Variable | Description |
|---|---|
| `EXCEL_SOURCE_PATH` | Path to the tracking Excel file. Point this at your SharePoint-synced local folder once that's set up. Defaults to `./data/AI_Champions_Initiative_Report.xlsx`. |
| `LLM_FARM_API_KEY` | Your LLM Farm API key. **Never commit this.** |
| `LLM_FARM_BASE_URL` | Your LLM Farm's OpenAI-compatible base URL. |
| `LLM_FARM_MODEL` | The model name available on your LLM Farm (e.g. whatever your gateway exposes). |

> **You'll need to get the exact `LLM_FARM_BASE_URL` and model name from your
> org's LLM Farm documentation/portal** — these can't be guessed, and the app
> ships with only a placeholder (`default-model-name`) in `llm_client.py`.

If you deploy somewhere that reads `st.secrets` instead of `.env` (e.g.
Streamlit Community Cloud), copy `.streamlit/secrets.toml.example` to
`.streamlit/secrets.toml` and fill it in there instead — `app.py` checks
both.

The app works fine with **no LLM Farm config at all** — the dashboard,
sync, and file upload all work normally; only the chatbot and monthly-report
tabs will show a friendly "not configured" warning until you add the keys.

## 3. Run

```bash
streamlit run app.py
```

Open the URL Streamlit prints (usually `http://localhost:8502`).

## 4. Run tests

```bash
pytest tests/ -v
```

## Project layout

```
gen-ai-dashboard/
  app.py                    # Streamlit UI: Overview / Team / Topics / Notes / Chatbot / Report
  theme.py                  # CSS theme (navy sidebar, Space Grotesk headers, badges, cards)
  components.py             # HTML builder helpers (stat cards, badges, progress bars, member cards)
  excel_parser.py           # parse_excel(filepath) -> DashboardData
  models.py                 # Pydantic models: ChampionMember, Subtopic, TopicGroup, DashboardData
  llm_client.py             # LLM Farm wrapper — chat + summary functions, single place to configure the model
  tests/
    test_excel_parser.py    # unit tests against a generated fixture matching the real file's layout
  data/
    AI_Champions_Initiative_Report.xlsx   # your sample/default source file
  .env.example               # copy to .env and fill in
  .streamlit/secrets.toml.example
  requirements.txt
  README.md
```

## Design

The look is a navy sidebar + card-based "executive dashboard" style (Space
Grotesk headings, Inter body text, colored status badges: green=Completed,
blue=In Progress, amber=Started, gray=Not Started). All of it is CSS
injected via `theme.py`, so you can retheme colors/fonts in one place
without touching `app.py`'s logic. `components.py` holds the HTML builders
(stat cards, progress bars, badges, member cards) that `app.py` calls with
your real parsed data — nothing is hardcoded/sample data.

## How the parser handles the sheet

The source file is one worksheet with three regions stacked vertically —
see the top of `excel_parser.py` for the full rules, but in short:

- **Region 1** (rows 1–7ish): initiative name / owner / supported-by / two
  champion lists. Names like `Jane Doe (BD/SWD-FSB1)` are split into
  `name` + `department`.
- **Region 2**: the `Topic | Sub topics | Progress | Lead | Key Points`
  table. A row with a Topic value and blank Sub-topics is a group header
  whose Progress is the already-computed rollup (not recalculated). Rows
  below with a Sub-topics value belong to that group — including a second,
  ungrouped batch of subtopic rows with no parent Topic row, which attaches
  to the most recent group. A stray value in a 6th column (e.g. `"July"`)
  is captured as a `note` on that row instead of causing an error.
- **Region 3**: plain sentences with no Progress/Lead/Key Points, captured
  as the Notes / Open Items list.

The parser always reads the **first sheet** in the workbook (name isn't
hardcoded), and is defensive about blank rows, missing Lead/Key Points
(rendered as "Unassigned" / "—"), and extra columns.

## Adding new AI features

Route any new AI feature through `llm_client.py` so there's one place that
knows how to talk to LLM Farm (`load_config()` + a thin wrapper function
per feature, following the pattern of `ask_chatbot` / `generate_monthly_summary`).
That keeps API-key handling, error handling, and the base URL/model config
in one spot as you build more on top of this.

## Non-goals for v1

(carried over from the build spec)

- No authentication/login
- No database — session-state cache is enough for now
- No multi-user concurrent editing
- No automated file-watching — manual "Sync Now" is enough for v1
- LLM Farm calls are for chat Q&A and summary drafting only — not for
  modifying the Excel file itself
