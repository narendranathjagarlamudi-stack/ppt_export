# services/ppt_service.py

import json
import tempfile

from pptx import Presentation
from pptx.util import Inches, Pt

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


# ============================================
# CHART PARSER
# ============================================

def extract_chart_data(chart_spec):

    if not chart_spec:
        return None

    # ============================================
    # STRING -> JSON
    # ============================================

    if isinstance(chart_spec, str):

        try:
            chart_spec = json.loads(chart_spec)

        except Exception:
            return None

    # ============================================
    # SIMPLE FORMAT
    # ============================================

    if (
        "labels" in chart_spec
        and "values" in chart_spec
    ):

        return {

            "chart_type": chart_spec.get(
                "type",
                "bar"
            ),

            "orientation": "vertical",

            "title": chart_spec.get(
                "title",
                "Chart"
            ),

            "categories": chart_spec["labels"],

            "values": chart_spec["values"],
        }

    # ============================================
    # VEGA-LITE FORMAT
    # ============================================

    try:

        mark = chart_spec.get(
            "mark",
            "bar"
        )

        if isinstance(mark, dict):

            mark = mark.get(
                "type",
                "bar"
            )

        data_values = (
            chart_spec
            .get("data", {})
            .get("values", [])
        )

        encoding = chart_spec.get(
            "encoding",
            {}
        )

        x_encoding = encoding.get(
            "x",
            {}
        )

        y_encoding = encoding.get(
            "y",
            {}
        )

        x_field = x_encoding.get(
            "field"
        )

        y_field = y_encoding.get(
            "field"
        )

        x_type = x_encoding.get(
            "type"
        )

        y_type = y_encoding.get(
            "type"
        )

        categories = []
        values = []

        # ============================================
        # DETERMINE ORIENTATION
        # ============================================

        if (
            x_type == "quantitative"
            and y_type == "nominal"
        ):

            orientation = "horizontal"

            category_field = y_field
            value_field = x_field

        else:

            orientation = "vertical"

            category_field = x_field
            value_field = y_field

        # ============================================
        # EXTRACT VALUES
        # ============================================

        for row in data_values:

            categories.append(
                str(
                    row.get(
                        category_field
                    )
                )
            )

            try:

                values.append(
                    float(
                        row.get(
                            value_field,
                            0
                        )
                    )
                )

            except Exception:

                values.append(0)

        return {

            "chart_type": mark,

            "orientation": orientation,

            "title": chart_spec.get(
                "title",
                "Chart"
            ),

            "categories": categories,

            "values": values,
        }

    except Exception as e:

        print(
            "Chart parsing error:",
            str(e)
        )

        return None


# ============================================
# CHART TYPE
# ============================================

def get_chart_type(parsed_chart):

    chart_type = parsed_chart.get(
        "chart_type",
        "bar"
    )

    orientation = parsed_chart.get(
        "orientation",
        "vertical"
    )

    # ============================================
    # LINE
    # ============================================

    if chart_type == "line":

        return XL_CHART_TYPE.LINE

    # ============================================
    # PIE / DONUT
    # ============================================

    if chart_type in ["pie", "arc"]:

        return XL_CHART_TYPE.PIE

    # ============================================
    # HORIZONTAL BAR
    # ============================================

    if orientation == "horizontal":

        return XL_CHART_TYPE.BAR_CLUSTERED

    # ============================================
    # VERTICAL BAR
    # ============================================

    return XL_CHART_TYPE.COLUMN_CLUSTERED


# ============================================
# PPT GENERATOR
# ============================================

def generate_pptx(
    presentation_json: dict,
    include_charts: bool = True
):

    prs = Presentation()

    # ============================================
    # TITLE SLIDE
    # ============================================

    title_slide_layout = prs.slide_layouts[0]

    title_slide = prs.slides.add_slide(
        title_slide_layout
    )

    title_slide.shapes.title.text = (
        presentation_json.get(
            "presentation_title",
            "Analysis Summary"
        )
    )

    title_slide.placeholders[1].text = (
        "Generated by Cortex PPT Export"
    )

    # ============================================
    # LOOP THROUGH SLIDES
    # ============================================

    for slide_data in presentation_json.get(
        "slides",
        []
    ):

        slide_layout = prs.slide_layouts[6]

        slide = prs.slides.add_slide(
            slide_layout
        )

        # ============================================
        # TITLE
        # ============================================

        title_box = slide.shapes.add_textbox(
            Inches(0.5),
            Inches(0.3),
            Inches(9),
            Inches(0.6)
        )

        title_tf = title_box.text_frame

        title_tf.text = slide_data["slide"].get(
            "title",
            "Key Insights"
        )

        title_tf.paragraphs[0].font.size = Pt(24)

        # ============================================
        # BULLETS
        # ============================================

        body_box = slide.shapes.add_textbox(
            Inches(0.5),
            Inches(1.2),
            Inches(3.6),
            Inches(4.8)
        )

        tf = body_box.text_frame

        tf.clear()

        tf.word_wrap = True

        bullets = slide_data["slide"].get(
            "bullets",
            []
        )

        for idx, bullet in enumerate(bullets):

            if idx == 0:

                p = tf.paragraphs[0]

            else:

                p = tf.add_paragraph()

            p.text = f"• {bullet}"

            p.font.size = Pt(16)

            p.space_after = Pt(10)

        # ============================================
        # CHART
        # ============================================

        if include_charts:

            parsed_chart = extract_chart_data(
                slide_data.get(
                    "chart_spec"
                )
            )

            if parsed_chart:

                chart_data = CategoryChartData()

                chart_data.categories = (
                    parsed_chart["categories"]
                )

                chart_data.add_series(
                    "Series 1",
                    parsed_chart["values"]
                )

                chart = slide.shapes.add_chart(

                    get_chart_type(
                        parsed_chart
                    ),

                    Inches(4.4),

                    Inches(1.3),

                    Inches(4.8),

                    Inches(3.8),

                    chart_data

                ).chart

                chart.has_title = True

                chart.chart_title.text_frame.text = (
                    parsed_chart.get(
                        "title",
                        "Chart"
                    )
                )

                chart.has_legend = False

    # ============================================
    # SAVE
    # ============================================

    temp_file = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pptx"
    )

    prs.save(temp_file.name)

    return temp_file.name