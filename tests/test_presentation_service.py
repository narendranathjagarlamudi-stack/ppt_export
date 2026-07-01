import json

from app.services.presentation_service import (
    _build_slides_from_final_text,
    _build_slides_from_stream,
)


def _chart(title):
    return json.dumps({
        "title": title,
        "data": {
            "values": [
                {
                    "BRAND": "A",
                    "VALUE": 1,
                },
                {
                    "BRAND": "B",
                    "VALUE": 2,
                },
            ]
        },
    })


def test_stream_response_builds_one_slide_per_chart():
    response_data = {
        "content": [
            {
                "event_type": "response.text",
                "data": {
                    "text": "Chart one says A is lower than B.",
                    "content_index": 1,
                },
            },
            {
                "event_type": "response.chart",
                "data": {
                    "chart_spec": _chart("Chart One"),
                    "content_index": 2,
                    "sequence_number": 10,
                },
            },
            {
                "event_type": "response.text",
                "data": {
                    "text": "Chart two adds another comparison.",
                    "content_index": 3,
                },
            },
            {
                "event_type": "response.chart",
                "data": {
                    "chart_spec": _chart("Chart Two"),
                    "content_index": 4,
                    "sequence_number": 20,
                },
            },
        ]
    }

    slides = _build_slides_from_stream(
        response_data=response_data,
        original_query="compare brands"
    )

    assert len(slides) == 2
    assert slides[0]["slide"]["title"] == "Chart One"
    assert slides[1]["slide"]["title"] == "Chart Two"
    assert slides[0]["event_metadata"]["content_index"] == 2
    assert slides[1]["event_metadata"]["sequence_number"] == 20


def test_stream_response_with_no_chart_returns_no_slides():
    response_data = {
        "content": [
            {
                "event_type": "response.text",
                "data": {
                    "text": "Text only response.",
                },
            },
        ]
    }

    slides = _build_slides_from_stream(
        response_data=response_data,
        original_query="explain data"
    )

    assert slides == []


def test_final_text_response_keeps_single_chart_slide():
    response_data = {
        "final_text": "A is lower than B.",
        "chart_spec": _chart("Final Text Chart"),
    }

    slides = _build_slides_from_final_text(
        response_data=response_data,
        original_query="compare brands"
    )

    assert len(slides) == 1
    assert slides[0]["slide"]["title"] == "Final Text Chart"
    assert slides[0]["chart_spec"]
