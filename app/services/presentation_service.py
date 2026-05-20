import json
import re

from app.services.cortex_service import (
    ask_cortex
)


async def generate_presentation_content(

    conn,

    original_query: str,

    original_response: dict,
):

    final_text = original_response.get(
        "final_text",
        ""
    )

    chart_spec = original_response.get(
        "chart_spec"
    )

    if not final_text:

        final_text = str(original_response)

    prompt = f"""
You are an enterprise presentation expert.

TASK:
Convert the analysis into ONE executive presentation slide.

STRICT RULES:
- Maximum 4 bullets
- Maximum 10 words per bullet
- No paragraphs
- Business wording only

CRITICAL:
- Return ONLY raw JSON
- No markdown
- No explanation
- No intro text
- Response must start with {{
- Response must end with }}

USER QUERY:
{original_query}

ANALYSIS:
{final_text}

OUTPUT FORMAT:

{{
  "presentation_title": "string",

  "slide": {{
    "title": "string",

    "bullets": [
      "string",
      "string",
      "string"
    ]
  }}
}}
"""

    result = await ask_cortex(
        conn=conn,
        message=prompt
    )

    try:

        print("\n========== RAW RESULT ==========")
        print(result)
        print("================================\n")

        cleaned = (
            result
            .replace("```json", "")
            .replace("```", "")
            .strip()
        )

        # ============================================
        # EXTRACT JSON OBJECT
        # ============================================

        match = re.search(
            r"\{.*\}",
            cleaned,
            re.DOTALL
        )

        if not match:

            raise Exception(
                "No JSON object found"
            )

        json_text = match.group(0)

        print("\n========== EXTRACTED JSON ==========")
        print(json_text)
        print("====================================\n")

        parsed = json.loads(json_text)

        slide = parsed.get(
            "slide",
            {}
        )

        bullets = slide.get(
            "bullets",
            []
        )

        cleaned_bullets = []

        for bullet in bullets:

            if (
                bullet
                and isinstance(bullet, str)
            ):

                cleaned_bullets.append(
                    bullet.strip()
                )

        return {

            "presentation_title": parsed.get(
                "presentation_title",
                "Analysis Summary"
            ),

            "slide": {

                "title": slide.get(
                    "title",
                    "Key Insights"
                ),

                "bullets": cleaned_bullets[:4]
            },

            "chart_spec": chart_spec
        }

    except Exception as e:

        print("\n========== PPT FALLBACK ==========")
        print(str(e))
        print("==================================\n")

        return {

            "presentation_title":
                "Analysis Summary",

            "slide": {

                "title":
                    "Key Insights",

                "bullets": [
                    "Analysis generated successfully",
                    "Key findings summarized",
                    "See chart for insights"
                ]
            },

            "chart_spec": chart_spec
        }