# NovaRetail Q2 2026 Executive Deck Generator

This project generates the NovaRetail Q2 2026 Executive Deck from the supplied Excel dataset and PowerPoint template. The script calculates the executive KPIs, creates native PowerPoint charts, populates the template placeholders, inserts the supplied icons, and saves the completed presentation as a new `.pptx` file.

## Project structure

Keep the following files in the same folder:

```text
NovaRetail_Capstone/
├── generate_novaretail_ppt.py
├── novaretail_template.pptx
├── novaretail_dataset.xlsx
├── growth_chart.png
├── customer_star.png
├── risk_alert.png
├── lightbulb_action.png
└── README.md
```

## Prerequisites

- Python 3.9 or later
- `pip`
- The supplied NovaRetail PowerPoint template
- The supplied NovaRetail Excel dataset
- The four required PNG icon files

## Setup

### 1. Open a terminal in the project folder

Windows PowerShell example:

```powershell
cd "C:\path\to\NovaRetail_Capstone"
```

macOS or Linux example:

```bash
cd /path/to/NovaRetail_Capstone
```

### 2. Create a virtual environment

Windows:

```powershell
python -m venv .venv
.venv\Scripts\activate
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install the dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Run the script

```bash
python generate_novaretail_ppt.py
```

On systems where Python is invoked as `python3`, run:

```bash
python3 generate_novaretail_ppt.py
```

## Output

A successful run creates the following file in the project folder:

```text
NovaRetail_Q2_2026_Executive_Deck.pptx
```

The terminal also displays:

```text
Created: <project-folder>/NovaRetail_Q2_2026_Executive_Deck.pptx
```

## Input workbook requirements

The script expects these worksheets:

```text
Regional_Performance
Customer_Experience
Operational_Risk
```

The KPI and chart logic also expects the field names used by the supplied dataset, including:

```text
Quarter
Region
Revenue_USD_M
NPS
Stockout_Rate_pct
```

Do not rename the worksheets, template shapes, or required columns unless the corresponding references in the Python script are also updated.

## KPI calculations

### Q2 2026 revenue

```text
Sum of Revenue_USD_M for all regions in the latest quarter
```

For the supplied dataset:

```text
465 + 345 + 330 + 140 = 1,280 million = $1.28B
```

### Quarter-over-quarter revenue growth

```text
(Current-quarter revenue / Previous-quarter revenue) - 1
```

For the supplied dataset:

```text
(1,280 / 1,225) - 1 = 4.49%, displayed as +4.5%
```

### Revenue-weighted NPS

```text
Sum(Regional revenue × Regional NPS) / Total revenue
```

For the supplied dataset, the result is displayed as:

```text
53.4
```

## Presentation contents

The script generates five slides:

1. Executive overview and KPI cards
2. Revenue and commercial performance
3. Customer experience trends
4. Operational risk
5. Recommended actions and next steps

## Troubleshooting

### `FileNotFoundError`

Confirm that the template, dataset, script, and required icons are in the same folder and use these exact filenames:

```text
novaretail_template.pptx
novaretail_dataset.xlsx
growth_chart.png
customer_star.png
risk_alert.png
lightbulb_action.png
```

### `KeyError: Template shape not found`

The PowerPoint template uses named shapes such as `kpi1_value`, `chart_performance`, and `icon_1`. Use the supplied template without renaming or deleting its placeholders.

### `KeyError` for a worksheet or column

Confirm that the workbook contains the expected worksheets and column headers. Column and worksheet names are case-sensitive in the script.

### Dependency errors

Reinstall the project dependencies:

```bash
pip install --upgrade --force-reinstall -r requirements.txt
```

## Recreating the output

To regenerate the presentation after changing the dataset, replace `novaretail_dataset.xlsx` with the updated workbook while preserving the expected worksheet and column names, then rerun:

```bash
python generate_novaretail_ppt.py
```

The existing output file with the same name will be overwritten.
