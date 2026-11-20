
# NovaRetail Executive Deck Automation

This repository generates the NovaRetail Q2 2026 executive deck from the included workbook and PowerPoint template. It now includes a Google Agent Development Kit (ADK) agent that reads the workbook, generates executive insights as structured JSON, and passes those insights to the PowerPoint generator.

## Repository contents

```text
novaretail_agent/
  __init__.py
  agent.py
run_pipeline.py
generate_novaretail_ppt.py
generated_insights.json
novaretail_dataset.xlsx
novaretail_template.pptx
NovaRetail_Q2_2026_Executive_Deck.pptx
growth_chart.png
customer_star.png
risk_alert.png
lightbulb_action.png
requirements.txt
.env.example
README.md
```

Every file above is used. The input workbook, template, required icons, generated insight JSON, and final PowerPoint are included.

## How the automation works

1. `novaretail_agent/agent.py` defines the ADK `root_agent` and the `load_novaretail_data` tool.
2. The agent reads all three workbook sheets and generates headlines, bullets, recommendations, and an executive summary in strict JSON.
3. `run_pipeline.py` saves the agent output as `generated_insights.json`.
4. `generate_novaretail_ppt.py` calculates the numeric KPIs, creates native PowerPoint charts, inserts icons, and populates the template with the ADK insights.
5. The pipeline writes `NovaRetail_Q2_2026_Executive_Deck.pptx`.

## Requirements

- Python 3.11 or later
- A Google AI Studio API key

Install dependencies:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

macOS or Linux:

```bash
source .venv/bin/activate
```

```bash
pip install -r requirements.txt
```

## Configure the API key

Copy `.env.example` to `.env` and replace the placeholder value:

```text
GOOGLE_API_KEY=your_real_api_key
GOOGLE_GENAI_USE_VERTEXAI=FALSE
```

Do not commit `.env`.

## Run the complete ADK pipeline

```bash
python run_pipeline.py
```

This command generates fresh AI insights and rebuilds the PowerPoint.

## Regenerate only the PowerPoint

The repository includes `generated_insights.json`, so the deck can be rebuilt without another model call:

```bash
python generate_novaretail_ppt.py
```

## Output

```text
NovaRetail_Q2_2026_Executive_Deck.pptx
```

## KPI calculations

- Q2 revenue: sum of `Revenue_USD_M` for the latest quarter.
- QoQ growth: `(latest quarter revenue / previous quarter revenue) - 1`.
- Revenue-weighted NPS: `sum(region revenue × region NPS) / total revenue`.

