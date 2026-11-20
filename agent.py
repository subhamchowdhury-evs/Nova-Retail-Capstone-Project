"""Google ADK agent for automated NovaRetail executive insight generation."""
from pathlib import Path
import json
import openpyxl
from google.adk.agents import Agent

BASE = Path(__file__).resolve().parent.parent
DATASET = BASE / "novaretail_dataset.xlsx"


def load_novaretail_data() -> dict:
    """Loads the NovaRetail workbook and returns all rows grouped by worksheet."""
    workbook = openpyxl.load_workbook(DATASET, data_only=True)
    result = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        rows = list(sheet.iter_rows(values_only=True))
        headers = rows[0]
        result[sheet_name] = [
            dict(zip(headers, row)) for row in rows[1:] if any(value is not None for value in row)
        ]
    return result


root_agent = Agent(
    model="gemini-2.5-flash",
    name="novaretail_insight_agent",
    description="Analyzes NovaRetail quarterly data and writes concise executive-deck insights.",
    instruction="""
Use load_novaretail_data before answering. Analyze the latest quarter against prior quarters.
Return ONLY valid JSON using this exact schema:
{
  "executive_summary": "3 concise sentences",
  "performance_headline": "insight-led headline",
  "performance_insights": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
  "customer_headline": "insight-led headline",
  "customer_insights": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
  "risk_headline": "insight-led headline",
  "risk_insights": ["bullet 1", "bullet 2", "bullet 3", "bullet 4"],
  "recommendations": ["action 1", "action 2", "action 3", "action 4"],
  "closing_statement": "one sentence"
}
Every claim must be traceable to the workbook. Keep each bullet under 24 words. Do not use Markdown.
""",
    tools=[load_novaretail_data],
)
