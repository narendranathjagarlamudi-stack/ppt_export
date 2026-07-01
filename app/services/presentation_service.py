import json


from app.services.Heuristic_extractor import (
    extract_bullets,
    extract_bullets_from_response
)


def _parse_response(original_response):

    if isinstance(original_response, str):

        try:
            return json.loads(original_response)

        except Exception:
            return {
                "final_text": original_response
            }

    if isinstance(original_response, dict):
        return original_response

    return {
        "final_text": str(original_response)
    }


def _extract_first_chart_spec(response_data: dict):

    chart_spec = response_data.get(
        "chart_spec"
    )

    if chart_spec:
        return chart_spec

    chart_specs = response_data.get(
        "chart_specs",
        []
    )

    if chart_specs:
        return chart_specs[0]

    for item in response_data.get(
        "content",
        []
    ):

        event_type = item.get(
            "event_type",
            ""
        )

        if event_type not in [
            "chart",
            "response.chart"
        ]:
            continue

        data = item.get(
            "data",
            {}
        )

        chart_spec = data.get(
            "chart_spec"
        )

        if chart_spec:
            return chart_spec

    return None


def _get_chart_title(chart_spec):

    if not chart_spec:
        return None

    if isinstance(chart_spec, str):

        try:
            chart_spec = json.loads(chart_spec)

        except Exception:
            return None

    if not isinstance(chart_spec, dict):
        return None

    title = chart_spec.get(
        "title"
    )

    if isinstance(title, str) and title.strip():
        return title.strip()

    return None


def _parse_chart_spec(chart_spec):

    if isinstance(chart_spec, str):

        try:
            return json.loads(chart_spec)

        except Exception:
            return None

    if isinstance(chart_spec, dict):
        return chart_spec

    return None


def _is_informative_chart_spec(chart_spec):

    parsed = _parse_chart_spec(chart_spec)

    if not parsed:
        return False

    values = (
        parsed
        .get("data", {})
        .get("values", [])
    )

    if not values:
        return False

    labels = set()

    for row in values:

        label = (
            row.get("ENTITY")
            or row.get("US_REGION")
            or row.get("QID")
            or row.get("QID_LABEL")
            or row.get("BRAND")
            or row.get("ATTRIBUTE")
            or ""
        )

        label = str(label).strip()

        if not label:
            continue

        if label.lower() in [
            "sigma",
            "total",
            "all",
        ]:
            continue

        labels.add(label)

    return len(labels) > 1


def _select_chart_spec(chart_specs):

    for spec in chart_specs:

        if _is_informative_chart_spec(spec):
            return spec

    if chart_specs:
        return chart_specs[0]

    return None


def _fallback_bullets():

    return [
        "Analysis generated successfully",
        "Key findings summarized",
        "See chart for insights"
    ]


def _build_slide(
    title: str,
    text: str,
    chart_spec=None,
    event_metadata=None
):

    bullets = extract_bullets(
        final_text=text,
        chart_spec=chart_spec
    )

    if not bullets:
        bullets = _fallback_bullets()

    return {

        "presentation_title":
            "Analysis Summary",

        "slide": {

            "title":
                _get_chart_title(chart_spec) or title or "Key Insights",

            "bullets":
                bullets
        },

        "chart_spec":
            chart_spec,

        "event_metadata":
            event_metadata or {}
    }


def _build_slides_from_stream(
    response_data: dict,
    original_query: str
):

    slides = []
    pending_text_parts = []
    last_text_parts = []

    for item in response_data.get(
        "content",
        []
    ):

        event_type = item.get(
            "event_type",
            ""
        )

        data = item.get(
            "data",
            {}
        )

        if not isinstance(data, dict):
            continue

        if (
            "text" in event_type.lower()
            or event_type == ""
        ):

            text = data.get(
                "text",
                ""
            )

            if text:
                pending_text_parts.append(text)
                last_text_parts = pending_text_parts[:]

            continue

        if event_type not in [
            "chart",
            "response.chart"
        ]:
            continue

        chart_spec = data.get(
            "chart_spec"
        )

        if not chart_spec:
            continue

        text_parts = pending_text_parts or last_text_parts
        text = "\n".join(text_parts)

        slides.append(
            _build_slide(
                title=original_query,
                text=text,
                chart_spec=chart_spec,
                event_metadata={
                    "content_index": data.get("content_index"),
                    "sequence_number": data.get("sequence_number"),
                    "tool_use_id": data.get("tool_use_id"),
                    "event_type": event_type,
                }
            )
        )

        pending_text_parts = []

    return slides


def _build_slides_from_final_text(
    response_data: dict,
    original_query: str
):

    final_text = response_data.get(
        "final_text",
        ""
    )

    chart_specs = response_data.get(
        "chart_specs",
        []
    )

    chart_spec = response_data.get(
        "chart_spec"
    )

    if not chart_specs and chart_spec:
        chart_specs = [chart_spec]

    if chart_specs:
        chart_spec = _select_chart_spec(
            chart_specs
        )

    return [
        _build_slide(
            title=original_query,
            text=final_text,
            chart_spec=chart_spec
        )
    ]


def _build_single_slide_from_response(
    response_data: dict,
    original_query: str
):

    chart_spec = _extract_first_chart_spec(
        response_data
    )

    bullets = extract_bullets_from_response(
        response_data
    )

    if not bullets:
        bullets = _fallback_bullets()

    return [{

        "presentation_title":
            "Analysis Summary",

        "slide": {

            "title":
                _get_chart_title(chart_spec) or original_query or "Key Insights",

            "bullets":
                bullets
        },

        "chart_spec":
            chart_spec
    }]


async def generate_presentation_content(

    conn,

    original_query: str,

    original_response: dict,
):

    response_data = _parse_response(
        original_response
    )

    if isinstance(
        response_data.get("content"),
        list
    ):

        slides = _build_slides_from_stream(
            response_data=response_data,
            original_query=original_query
        )

        if slides:
            return slides

    if "final_text" in response_data:
        return _build_slides_from_final_text(
            response_data=response_data,
            original_query=original_query
        )

    return _build_single_slide_from_response(
        response_data=response_data,
        original_query=original_query
    )
