# Nova-Retail-Capstone-Project

This project generates a five-slide Quarterly Business Review deck from the provided PowerPoint template and Excel dataset.

## Files

- `generate_novaretail_deck.py` - reads the workbook, calculates KPIs, populates named PowerPoint shapes, and creates native PowerPoint charts.
- `novaretail_template.pptx` - fixed five-slide template.
- `novaretail_dataset.xlsx` - source data with regional performance, customer experience, and operational risk sheets.
- `NovaRetail_Q2_2026_Executive_Deck.pptx` - generated deliverable.

## Requirements

```bash
pip install python-pptx openpyxl
```
# Required files in the same directory:
- novaretail_template.pptx
- novaretail_dataset.xlsx
- growth_chart.png
- customer_star.png
- risk_alert.png
- lightbulb_action.png
  
## Run

Place all files in the same folder and run:

```bash
python generate_novaretail_deck.py
```

The script writes `NovaRetail_Q2_2026_Executive_Deck.pptx` to the same folder.

## Design Overview

- Slide 1 uses total Q2 revenue, quarter-over-quarter revenue growth, and revenue-weighted NPS.
- Slide 2 uses a clustered column chart for revenue by region and quarter.
- Slide 3 uses a multi-series line chart for regional NPS trends.
- Slide 4 uses a horizontal bar chart for Q2 stockout rates.
- Slide 5 links recommendations directly to the issues identified earlier.

The template layout and slide count remain unchanged. All fillable regions are located by their intentional shape names rather than by guessed coordinates.
