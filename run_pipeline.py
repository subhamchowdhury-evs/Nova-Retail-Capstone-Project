"""Run the ADK insight agent, save JSON, then generate the PowerPoint."""
import asyncio
import json
from google.adk.runners import InMemoryRunner
from google.genai import types
from novaretail_agent.agent import root_agent
from generate_novaretail_ppt import build_deck

OUTPUT_JSON = "generated_insights.json"


async def generate_insights() -> dict:
    runner = InMemoryRunner(agent=root_agent, app_name="novaretail_capstone")
    session = await runner.session_service.create_session(
        app_name="novaretail_capstone", user_id="capstone_user"
    )
    message = types.Content(
        role="user",
        parts=[types.Part.from_text(text="Generate the executive insights for the NovaRetail Q2 2026 deck.")],
    )
    final_text = None
    async for event in runner.run_async(
        user_id="capstone_user", session_id=session.id, new_message=message
    ):
        if event.content and event.content.parts:
            for part in event.content.parts:
                if part.text:
                    final_text = part.text
    if not final_text:
        raise RuntimeError("ADK agent returned no final text.")
    text = final_text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    result = json.loads(text)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, ensure_ascii=False)
    return result


async def main():
    insights = await generate_insights()
    build_deck(insights)
    print("Created generated_insights.json and NovaRetail_Q2_2026_Executive_Deck.pptx")


if __name__ == "__main__":
    asyncio.run(main())
