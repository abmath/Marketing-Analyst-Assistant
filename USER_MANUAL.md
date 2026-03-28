# User Manual — Marketing Analytics Assistant

**Audience:** Marketing strategists, campaign managers, and digital analysts
**Skill required:** No coding knowledge needed

---

## Table of Contents

1. [Getting started](#1-getting-started)
2. [The interface at a glance](#2-the-interface-at-a-glance)
3. [Step 1 — Loading your Excel file](#3-step-1--loading-your-excel-file)
4. [Step 2 — Asking questions](#4-step-2--asking-questions)
5. [Step 3 — Generating PowerPoint insights](#5-step-3--generating-powerpoint-insights)
6. [Writing better questions](#6-writing-better-questions)
7. [Question reference by use case](#7-question-reference-by-use-case)
8. [Understanding the results](#8-understanding-the-results)
9. [Common errors and fixes](#9-common-errors-and-fixes)

---

## 1. Getting started

Before opening the app, make sure two things are running in the background:

**Step A — Start Ollama**

Open a Terminal and run:
```
ollama serve
```
Leave this terminal open. The AI brain of the app runs here.

**Step B — Pull the Phi-3 model** *(first time only, ~2.3 GB download)*
```
ollama pull phi3
```

**Step C — Launch the app**
```
streamlit run app.py
```

Your browser will open automatically at `http://localhost:8501`.

> If the browser does not open, navigate there manually.

---

## 2. The interface at a glance

```
┌─────────────────┬──────────────────────────────────────────────────┐
│   SIDEBAR       │   MAIN AREA                                      │
│                 │                                                   │
│  ⚙️ Settings    │  📊 Marketing Analytics Assistant                │
│  • Model picker │                                                   │
│  • Status light │  ── Step 1: Select a Table ──────────────────    │
│                 │  Dropdown + Data Preview                          │
│  📂 Data        │                                                   │
│  • File uploader│  ── Step 2: Ask a Question ──────────────────    │
│                 │  Example buttons | Question box + Run button      │
│                 │  Result table + Quick chart                       │
│                 │                                                   │
│                 │  ── Step 3: PowerPoint Insights ─────────────    │
│                 │  Generate button + Slide cards + Download         │
└─────────────────┴──────────────────────────────────────────────────┘
```

**Sidebar status indicators**

| Indicator | Meaning |
|-----------|---------|
| Green "Ollama connected" | AI is ready |
| Red "Ollama not reachable" | Run `ollama serve` in Terminal |
| Yellow "No models found" | Run `ollama pull phi3` |

---

## 3. Step 1 — Loading your Excel file

### Uploading the file

1. In the **sidebar**, click **Browse files** under the Data section.
2. Select your `.xlsx` or `.xls` file and click Open.
3. The app parses the file automatically and shows a dropdown.

### Selecting a table

The dropdown lists every table the app found. Table names follow this pattern:

- **Single table per sheet:** `Sheet Name` (e.g., `Campaign Performance`)
- **Multiple tables on one sheet:** `Sheet Name › Table Label` (e.g., `Q3 Results › Top Campaigns`)
- **Named Excel Tables:** `Sheet Name › TableName` (e.g., `Traffic › tbl_weekly`)

**Tip:** Adobe Analytics exports often pack multiple report tables on a single sheet separated by blank rows. The app splits these automatically — you will see them as separate entries in the dropdown.

### Data Preview

Once a table is selected, a **Data Preview** card expands showing:
- Row count, column count, numeric column count, text column count
- First 100 rows of the table

Scan the column names here — you will need them when phrasing questions.

---

## 4. Step 2 — Asking questions

### Using the quick example buttons

Five example questions appear on the left. Clicking one fills the question box — you can then edit it or run it as-is.

### Writing your own question

Type your question in the text box on the right. Be as natural as you like — see [Section 6](#6-writing-better-questions) for tips.

### Running a query

Click **▶ Run Query**. The app will:

1. Send your question + the table's schema to Phi-3.
2. Phi-3 writes pandas code to answer it.
3. The code runs on your data.
4. If the code fails, the error is automatically sent back to Phi-3 to fix (up to 3 attempts).

**What you see after running:**

| Element | What it is |
|---------|-----------|
| **Generated code** expander | The Python/pandas code Phi-3 wrote. Click to inspect. Shows each attempt if retries were needed. |
| **Result table** | The answer rendered as a data table. |
| **Quick chart** expander | Auto-generated bar chart (appears when the result has ≤ 60 rows and at least one numeric column). |

### Clearing a question

Click **Clear** to reset the question box and start fresh.

---

## 5. Step 3 — Generating PowerPoint insights

Click **✨ Generate Slide Suggestions**.

Phi-3 reads the selected table's statistics and drafts **5 slide suggestions** in the format:

```
## Slide 1: [Title]
Key message: one-sentence summary
Bullets:
  - Insight 1
  - Insight 2
  - Insight 3
Recommended chart: bar / line / pie / scatter / heatmap
```

Each slide appears as a coloured card on screen.

### Downloading the insights

Click **⬇️ Download insights as .txt** to save the suggestions as a plain-text file. Paste the content directly into your PowerPoint speaker notes or share with your team.

**Tips for best results:**
- Run this on an aggregated or summary table rather than a raw daily log for higher-quality strategic recommendations.
- The Model Comparison or CRM & Lead Quality sheets tend to produce the most actionable slide ideas.

---

## 6. Writing better questions

### Do's

| Do | Example |
|----|---------|
| Use column names from the preview | "Show top 10 rows by `Booking Completed`" |
| Be specific about the metric | "What is the average `Bounce Rate (%)` by `Channel`?" |
| Name the aggregation | "Sum of `Revenue ($)` grouped by `State`" |
| Use comparison words | "Which `Model` has the highest `TD → Booking Rate (%)`?" |
| Ask for outliers explicitly | "Show outliers in `Cost per Booking (₹)` using IQR" |
| Ask for trends | "Show weekly trend of `Online Bookings`" |

### Don'ts

| Don't | Better alternative |
|-------|-------------------|
| "Analyse everything" | "Show me the top 5 campaigns by bookings" |
| Use vague terms | "bounce rate outliers" → "Show rows where `Bounce Rate (%)` is above 2 standard deviations" |
| Ask multi-part questions | Split into two separate questions |
| Reference columns not in the table | Check the Data Preview first |

### Phrasing for specific analysis types

**Ranking / Top N**
> "What are the top 10 `State`s by `Online Bookings`?"
> "Show the bottom 5 `Channel`s by `Engagement Rate (%)`"

**Filtering**
> "Show all rows where `Bounce Rate (%)` is greater than 60"
> "Filter to rows where `Model` is 'Varuna Bolt' and `Channel` is 'YouTube'"

**Aggregation**
> "Total `Ad Spend (₹)` and `Booking Completed` grouped by `Campaign Type`"
> "Average `Avg Session Duration (s)` by `Device` and `Channel`"

**Outlier detection**
> "Find outliers in `Cost per Booking (₹)` using IQR"
> "Which rows have unusually high `Bounce Rate (%)` compared to the rest?"
> "Show campaigns where `ROAS` is more than 2 standard deviations from the mean"

**Trends over time**
> "Show `Online Bookings` by `Week Start`"
> "Monthly trend of `Test Drive Requests`"

**Ratios and calculated fields**
> "Show `Booking Completed` divided by `Sessions` for each `Model`"
> "Which `State` has the best ratio of `Bookings` to `Test Drive Requests`?"

**Correlation / comparison**
> "Compare `Bounce Rate (%)` across all values of `State Tier`"
> "Show average `Conversion Rate (%)` side by side for each `Vehicle Type`"

---

## 7. Question reference by use case

### Campaign performance

| Goal | Question to ask |
|------|----------------|
| Best-performing campaigns | "Top 10 rows by `Booking Completed`" |
| Most efficient spend | "Sort by `ROAS` descending, show top 15" |
| Worst cost-per-booking | "Bottom 10 rows by `Cost per Booking (₹)` where `Booking Completed` > 0" |
| Channel efficiency | "Average `CTR (%)` and `Booking Completed` grouped by `Channel`" |
| Festive vs non-festive | "Compare average `Bookings / 1000 Sessions` where `Festival Week` is '—' vs others" |
| Outliers in spend | "Find outliers in `Ad Spend (₹)` using IQR" |

### Booking funnel

| Goal | Question to ask |
|------|----------------|
| Biggest drop-off step | "Show average of all funnel columns (`F1` through `F12`)" |
| Funnel by model | "Group by `Model`, average `Visit → Booking (%)`" |
| Best converting channel | "Top `Channel` by `Booking Start → Done (%)`" |
| Festive funnel lift | "Compare `Visit → Booking (%)` for festive vs non-festive weeks" |

### Geographic analysis

| Goal | Question to ask |
|------|----------------|
| Top states by bookings | "Sum `Online Bookings` by `State`, sort descending" |
| Tier comparison | "Average `Booking Rate (%)` grouped by `State Tier`" |
| City-level deep dive | "Filter to `State` = 'Maharashtra', group by `City`, sum `Online Bookings`" |
| Digital walkin attribution | "Top 10 states by `Digital-Attributed Walkins`" |

### Lead & CRM

| Goal | Question to ask |
|------|----------------|
| Best segment to target | "Average `TD → Booking Rate (%)` by `Buyer Segment`" |
| WhatsApp effectiveness | "Correlation of `WhatsApp Reply Rate (%)` and `Online Bookings`" |
| Lead quality by channel | "Average `Lead Quality Rate (%)` grouped by `Channel`" |
| Churn risk | "Show rows where `Churn Rate (%)` > 15, sorted by `Total Customers` descending" |

### Model comparison

| Goal | Question to ask |
|------|----------------|
| EV vs ICE performance | "Filter to `Vehicle Type` in ['Electric SUV', 'Compact SUV'], compare `Booking Rate (%)`" |
| Model share trend | "Show `Model` and `Share of Model PV (%)` by `Month`" |
| Best test-drive conversion | "Average `TD Request Rate (%)` by `Model`" |

---

## 8. Understanding the results

### Result table

Results are displayed as interactive tables. You can:
- **Sort** any column by clicking its header
- **Resize** columns by dragging their edges
- **Scroll** horizontally for wide tables

### Quick chart

Appears automatically when the result has 60 or fewer rows. Use the **Metric** dropdown to switch which numeric column is plotted on the Y-axis. The X-axis uses the first (leftmost) column of your result.

### Generated code expander

Click **🛠 Generated code** to see exactly what pandas code Phi-3 wrote. This is useful for:
- Understanding how the question was interpreted
- Copying the code into a notebook for further customisation
- Debugging if the result looks unexpected

If there were multiple attempts, each attempt's code and its error are shown in order.

---

## 9. Common errors and fixes

### "Ollama not reachable"
Ollama is not running. Open a new terminal and run `ollama serve`.

### "No models found"
Ollama is running but Phi-3 hasn't been downloaded yet. Run `ollama pull phi3`.

### "Could not execute after 3 attempts"
Phi-3 was unable to write working code after 3 tries. Try one of these:

1. **Rephrase using exact column names** — copy a column name from the Data Preview and paste it into your question.
2. **Be more specific** — instead of "show outliers", try "show rows where `Bounce Rate (%)` > 65".
3. **Break it down** — split a complex question into two simpler ones.
4. **Switch the model** — `phi3:medium` in the sidebar dropdown is more capable than `phi3:mini`.

### Query ran but result looks wrong
- Check the Generated Code expander — the code may have interpreted your question differently than intended.
- Try rephrasing. For example, if you asked "top campaigns" and got only 1 row, add "show top 10".

### Excel file shows "No tables found"
The file likely has only metadata rows without a proper header + data structure. Make sure:
- Row 1 (or the first non-blank row) contains column headers.
- At least one data row follows the header.
- The file is saved as `.xlsx` or `.xls`.

### Slide suggestions are too generic
Run the insight generator on a more specific table (e.g., Model Comparison or CRM Segments) rather than a raw daily log table. Aggregated data produces more strategic recommendations.

### App is slow
- Response time depends on your machine's RAM and CPU.
- Use `phi3:mini` (sidebar model picker) for faster but slightly less accurate responses.
- Phi-3 Medium takes ~30–60 seconds per query on a standard laptop.
