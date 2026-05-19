# app/services/presentation_service.py

import json

from app.services.cortex_service import ask_cortex


async def generate_presentation_content(
    agent_id: str,
    thread_id: str,
    original_query: str,
    original_response: dict,
):
    """
    Generates executive-style PPT content.

    OUTPUT:
    - Single slide
    - Concise bullets
    - Reuses original chart_spec
    """

    # ============================================
    # EXTRACT RESPONSE
    # ============================================

    final_text = original_response.get(
        "final_text",
        ""
    )

    chart_spec = original_response.get(
        "chart_spec"
    )

    if not final_text:
        final_text = str(original_response)

    # ============================================
    # STRICT PROMPT
    # ============================================

    prompt = f"""
You are an enterprise presentation expert.

TASK:
Convert the analysis into ONE executive presentation slide.

STRICT RULES:
- Maximum 4 bullets
- Maximum 10-12 words per bullet
- No long sentences
- No paragraphs
- Concise executive wording only
- Return VALID JSON ONLY
- No markdown
- No explanations
- Create EXACTLY ONE slide
- Maximum 4 bullets
- Each bullet MAX 12 words
- Bullets MUST be concise
- No paragraphs
- No long sentences
- Business-friendly wording only
- Rewrite professionally
- Do NOT copy raw analysis directly

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

    # ============================================
    # CALL CORTEX
    # ============================================

    result = await ask_cortex(
        message=prompt
    )

    # ============================================
    # PARSE
    # ============================================

    try:

        cleaned = (
            result
            .replace("```json", "")
            .replace("```", "")
            .strip()
        )

        parsed = json.loads(cleaned)

        slide = parsed.get("slide", {})

        bullets = slide.get("bullets", [])

        # safety cleanup
        cleaned_bullets = []

        for bullet in bullets:

            if bullet and isinstance(bullet, str):

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

    except Exception:

        # ============================================
        # FALLBACK
        # ============================================

        return {
            "presentation_title": "Analysis Summary",

            "slide": {
                "title": "Key Insights",

                "bullets": [
                    "Analysis generated successfully",
                    "Key findings summarized",
                    "See visualization for trends"
                ]
            },

            "chart_spec": chart_spec
        }