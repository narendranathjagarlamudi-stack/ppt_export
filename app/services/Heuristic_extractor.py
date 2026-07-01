

import re
import json


# =====================================================
# CONFIG
# =====================================================

NOISE_LINES = [
    "tooltip",
    "$schema",
    "vega-lite",
    "encoding",
    "transform",
    "config",
    "chart shown below",
    "please find",
    "visualization below",
    "donut chart",
    "bar chart",
    "heatmap",
    "grouped bar chart",
    "analysis details",
    "crosstab type",
    "row type",
    "fresh comparative charts",
    "key cross-chart takeaways",
    "about this dataset",
    "dataset at a glance",
    "what the data contains",
    "country-wise",
    "closest geographic comparison",
    "shown in the charts below",
]

FILLER_PHRASES = [
    "it is",
    "making it",
    "therefore",
    "significantly",
    "clearly",
    "overall",
    "in the market",
    "holding",
]


# =====================================================
# CLEAN MARKDOWN
# =====================================================

def clean_markdown(text: str):

    replacements = [
        ("**", ""),
        ("`", ""),
        ("#", ""),
        (">", ""),
        ("📌", ""),
        ("🔑", ""),
        ("🏆", ""),
        ("🗺️", ""),
        ("📊", ""),
        ("📌", ""),
        ("🔑", ""),
        ("🏆", ""),
        ("🗺️", ""),
        ("📊", ""),
        ("---", ""),
    ]

    for old, new in replacements:
        text = text.replace(old, new)

    text = re.sub(r'\s+', ' ', text)

    return text.strip()


# =====================================================
# REMOVE TABLES
# =====================================================

def remove_markdown_tables(text: str):

    cleaned = []

    for line in text.splitlines():

        if "|" in line:
            continue

        cleaned.append(line)

    return "\n".join(cleaned)


# =====================================================
# SPLIT SENTENCES
# =====================================================

def split_sentences(text: str):

    text = re.sub(
        r'\bvs\.\s+',
        'vs ',
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r'#{1,6}\s+',
        '. ',
        text
    )

    text = re.sub(
        r'\n{2,}',
        '. ',
        text
    )

    text = text.replace("\n", ". ")

    text = re.sub(
        r'(?<=\))\s+(?=[A-Z])',
        '. ',
        text
    )

    parts = re.split(
        r'(?<=[.!?])\s+',
        text
    )

    return [
        p.strip()
        for p in parts
        if len(p.strip()) > 20
    ]


# =====================================================
# NOISE FILTER
# =====================================================

def is_noise(line: str):

    lower = line.lower()

    if any(n in lower for n in NOISE_LINES):
        return True

    if re.match(r'^\s*chart\s+\d+\b', lower):
        return True

    if lower.startswith("here are "):
        return True

    if "{" in line or "}" in line:
        return True

    if len(re.findall(r'\d+', line)) > 20:
        return True

    return False


# =====================================================
# NORMALIZE BULLETS
# =====================================================

def normalize_bullet(line: str):

    line = line.replace("—", "-")

    line = re.sub(r'\*\*', '', line)

    line = re.sub(r'\s+', ' ', line)

    for filler in FILLER_PHRASES:

        line = re.sub(
            rf'\b{re.escape(filler)}\b',
            '',
            line,
            flags=re.IGNORECASE
        )

    line = re.sub(
        r'(\d+\.\d{1,2})\d*%',
        r'\1%',
        line
    )

    line = re.sub(r'\s+', ' ', line)

    line = line.strip(" ,.-")

    if len(line) > 150:

        cutoff = max(
            line[:150].rfind("."),
            line[:150].rfind(";"),
            line[:150].rfind(",")
        )

        if cutoff < 80:
            cutoff = line[:150].rfind(" ")

        line = line[:cutoff].strip() + "..."

    return line


# =====================================================
# DUPLICATE CHECK
# =====================================================

def is_duplicate(candidate, existing):

    c_words = set(candidate.lower().split())

    for ex in existing:

        e_words = set(ex.lower().split())

        overlap = len(c_words & e_words)

        similarity = overlap / max(len(c_words), 1)

        if similarity > 0.70:
            return True

    return False


def is_worthy_text_insight(sentence: str):

    lower = sentence.lower()

    if re.search(r'\d+(?:\.\d+)?%', sentence):
        return True

    if re.search(r'\b\d+(?:,\d{3})*(?:\.\d+)?\b', sentence):

        if any(
            word in lower
            for word in [
                "record",
                "records",
                "respondent",
                "respondents",
                "count",
                "base",
                "total",
                "share",
                "volume",
                "rank",
                "ranks",
                "region",
                "country",
                "qid",
                "qids",
                "pp",
                "point",
                "points",
            ]
        ):
            return True

    if any(
        phrase in lower
        for phrase in [
            "highest",
            "lowest",
            "largest",
            "smallest",
            "strongest",
            "weakest",
            "leads",
            "lead",
            "ahead",
            "behind",
            "followed by",
            "ranks",
            "ranked",
            "more than",
            "less than",
            "compared",
            "versus",
            "vs",
        ]
    ):
        return True

    return False


# =====================================================
# TEXT INSIGHTS
# =====================================================

def extract_text_insights(text: str):

    sentences = split_sentences(text)

    scored = []

    KEYWORDS = [
        "highest",
        "lowest",
        "largest",
        "smallest",
        "strongest",
        "weakest",
        "best",
        "worst",
        "lead",
        "leads",
        "leader",
        "dominates",
        "dominant",
        "ahead",
        "behind",
        "outperforms",
        "underperforms",
        "gains",
        "drops",
        "declines",
        "increases",
        "decreases",
        "peaks",
        "falls",
        "cliff",
        "blind spot",
        "advantage",
        "disadvantage",
        "opportunity",
        "risk",
        "loyalty",
        "usage",
        "consideration",
        "awareness",
        "familiarity",
        "nps",
        "share",
        "market",
        "regional",
        "region",
        "customer",
        "brand"
    ]

    COMPARISON_WORDS = [
        "vs",
        "versus",
        "compared",
        "compare",
        "relative",
        "than",
        "while",
        "whereas",
        "against"
    ]

    for sentence in sentences:

        if is_noise(sentence):
            continue

        score = 0

        lower = sentence.lower()

        for kw in KEYWORDS:
            if kw in lower:
                score += 3

        for kw in COMPARISON_WORDS:
            if kw in lower:
                score += 4

        score += len(
            re.findall(
                r'\d+(?:\.\d+)?%',
                sentence
            )
        ) * 3

        score += len(
            re.findall(
                r'\d+(?:\.\d+)?',
                sentence
            )
        )

        capitals = re.findall(
            r'\b[A-Z][a-zA-Z]+\b',
            sentence
        )

        if len(capitals) >= 2:
            score += 4

        if any(
            w in lower
            for w in [
                "only",
                "strongest",
                "weakest",
                "highest",
                "lowest",
                "best",
                "worst",
                "catastrophic",
                "sharp",
                "significant",
                "key finding",
                "important",
            ]
        ):
            score += 5

        scored.append(
            (score, sentence)
        )

    scored.sort(
        key=lambda x: x[0],
        reverse=True
    )

    if not scored:
        return []

    scored.sort(
        key=lambda x: x[0],
        reverse=True
    )

    top = [
        sentence
        for _, sentence in scored[:20]
    ]

    if not top:
        top = sentences[:10]

    return top


# =====================================================
# MARKDOWN TABLE INSIGHTS
# =====================================================

def extract_markdown_table_insights(text: str):

    bullets = []

    lines = text.splitlines()

    table_rows = []

    for line in lines:

        if "|" not in line:
            continue

        if "---" in line:
            continue

        cells = [
            c.strip()
            for c in line.split("|")
            if c.strip()
        ]

        if len(cells) < 3:
            continue

        table_rows.append(cells)

    if not table_rows:
        return bullets

    headers = [
        h.lower()
        for h in table_rows[0]
    ]

    is_org_profile_table = (
        any("org size" in h for h in headers)
        and any("revenue" in h for h in headers)
        and any("enterprise" in h for h in headers)
    )

    if not is_org_profile_table:
        return bullets

    # Skip header
    data_rows = table_rows[1:]

    for row in data_rows:

        try:

            org = row[0]
            revenue = row[1]
            enterprise = row[2]

            if "Org Size" in org:
                continue

            bullet = (
                f"{org} organizations are primarily in "
                f"{revenue}; Enterprise share is {enterprise}."
            )

            bullets.append(bullet)

        except Exception:
            continue

    return bullets


# =====================================================
# CHART INSIGHTS
# =====================================================

def apply_fold_transforms(chart_spec, values):

    rows = list(values)

    for transform in chart_spec.get(
        "transform",
        []
    ):

        fold_fields = transform.get(
            "fold"
        )

        as_fields = transform.get(
            "as",
            []
        )

        if (
            not fold_fields
            or len(as_fields) < 2
        ):
            continue

        key_field = as_fields[0]
        value_field = as_fields[1]
        folded_rows = []

        for row in rows:

            for field in fold_fields:

                folded = dict(row)
                folded[key_field] = field
                folded[value_field] = row.get(
                    field,
                    0
                )
                folded_rows.append(folded)

        rows = folded_rows

    return rows


def clean_label(value):

    label = str(value).strip()

    replacements = {
        "_PCT": "",
        "_NPS": "",
        "PCT": "%",
        "_": " ",
    }

    for old, new in replacements.items():
        label = label.replace(old, new)

    return label.title()


def format_chart_value(value, title):

    try:
        number = float(value)

    except Exception:
        return str(value)

    lower_title = title.lower()

    if (
        "%" in title
        or "pct" in lower_title
        or "top 2 box" in lower_title
    ):
        return f"{number:.1f}%"

    if number >= 1000:
        return f"{number:,.0f}"

    if number == int(number):
        return f"{number:.0f}"

    return f"{number:.1f}"


def generic_chart_insights(chart_spec, values):

    bullets = []

    title = chart_spec.get(
        "title",
        ""
    )

    rows = apply_fold_transforms(
        chart_spec,
        values
    )

    if not rows:
        return bullets

    first = rows[0]

    label_field = None

    for field in [
        "BRAND",
        "ATTRIBUTE",
        "US_REGION",
        "ENTITY",
        "QID_LABEL",
        "QID",
    ]:

        if field in first:
            label_field = field
            break

    if not label_field:
        return bullets

    value_field = None

    for field in [
        "VALUE",
        "PCT",
        "RECORD_COUNT",
        "RESPONDENT_COUNT",
    ]:

        if field in first:
            value_field = field
            break

    series_field = None

    for field in [
        "METRIC",
        "METRIC_KEY",
        "ROLE",
        "ROLE_KEY",
        "REGION",
        "REG_KEY",
        "INDUSTRY",
        "IND_KEY",
    ]:

        if field in first:
            series_field = field
            break

    if not value_field:

        numeric_fields = []

        for field, value in first.items():

            if field == label_field:
                continue

            if isinstance(
                value,
                (int, float)
            ):
                numeric_fields.append(field)

        if not numeric_fields:
            return bullets

        rows = [
            {
                label_field: row.get(label_field),
                "_SERIES": field,
                "_VALUE": row.get(field)
            }
            for row in values
            for field in numeric_fields
        ]

        value_field = "_VALUE"
        series_field = "_SERIES"

    points = []

    for row in rows:

        label = row.get(
            label_field
        )

        if not label:
            continue

        if str(label).strip().lower() in [
            "sigma",
            "total",
            "all",
        ]:
            continue

        try:
            value = float(
                row.get(
                    value_field,
                    0
                )
            )

        except Exception:
            continue

        series = (
            row.get(series_field)
            if series_field
            else None
        )

        points.append({
            "label": str(label),
            "series": clean_label(series) if series else None,
            "value": value,
        })

    if not points:
        return bullets

    points.sort(
        key=lambda item: item["value"],
        reverse=True
    )

    for rank, point in enumerate(points[:4], start=1):

        context = (
            f' on {point["series"]}'
            if point["series"]
            else ""
        )

        if rank == 1:

            bullets.append(
                f'{point["label"]} leads{context} at '
                f'{format_chart_value(point["value"], title)}.'
            )

        else:

            bullets.append(
                f'{point["label"]} ranks #{rank}{context} at '
                f'{format_chart_value(point["value"], title)}.'
            )

    if len(points) > 4:

        lowest = points[-1]

        low_context = (
            f' on {lowest["series"]}'
            if lowest["series"]
            else ""
        )

        bullets.append(
            f'{lowest["label"]} is lowest{low_context} at '
            f'{format_chart_value(lowest["value"], title)}.'
        )

    by_label = {}

    for point in points:

        if not point["series"]:
            continue

        by_label.setdefault(
            point["label"],
            []
        ).append(point)

    spread_candidates = []

    for label, label_points in by_label.items():

        if len(label_points) < 2:
            continue

        high = max(
            label_points,
            key=lambda item: item["value"]
        )

        low = min(
            label_points,
            key=lambda item: item["value"]
        )

        spread_candidates.append(
            (
                high["value"] - low["value"],
                label,
                low,
                high,
            )
        )

    if spread_candidates:

        spread, label, low, high = max(
            spread_candidates,
            key=lambda item: item[0]
        )

        if spread > 0:

            bullets.append(
                f'{label} varies most: '
                f'{clean_label(low["series"])} '
                f'{format_chart_value(low["value"], title)} to '
                f'{clean_label(high["series"])} '
                f'{format_chart_value(high["value"], title)}.'
            )

    return bullets


def extract_chart_insights(chart_spec):

    bullets = []

    try:

        if isinstance(chart_spec, str):
            chart_spec = json.loads(chart_spec)

        values = (
            chart_spec
            .get("data", {})
            .get("values", [])
        )

        if not values:
            return bullets

        # =================================================
        # SIMPLE PERCENTAGE CHARTS
        # =================================================

        if "PCT" in values[0]:

            cleaned_values = []

            for row in values:

                entity = (
                    row.get("ENTITY")
                    or row.get("US_REGION")
                    or row.get("QID")
                    or row.get("QID_LABEL")
                )

                if entity is None:
                    continue

                entity = str(entity).strip()

                if entity == "":
                    continue

                if entity.lower() in [
                    "sigma",
                    "total",
                    "all",
                ]:
                    continue

                row["_label"] = entity

                cleaned_values.append(row)

            values = cleaned_values

            if values:

                sorted_vals = sorted(
                    values,
                    key=lambda x: x.get("PCT", 0),
                    reverse=True
                )

                for rank, row in enumerate(sorted_vals[:5], start=1):

                    if rank == 1 and len(sorted_vals) > 1:

                        second = sorted_vals[1]

                        gap = round(
                            row["PCT"] - second["PCT"],
                            1
                        )

                        bullets.append(
                            f'{row["_label"]} leads at '
                            f'{row["PCT"]:.1f}%, ahead of '
                            f'{second["_label"]} by {gap} pp.'
                        )

                    else:

                        bullets.append(
                            f'{row["_label"]} ranks #{rank} at '
                            f'{row["PCT"]:.1f}%.'
                        )

                if len(sorted_vals) > 5:

                    lowest = sorted_vals[-1]

                    bullets.append(
                        f'{lowest["_label"]} has the lowest share at '
                        f'{lowest["PCT"]:.1f}%.'
                    )

        # =================================================
        # RECORD COUNTS
        # =================================================

        if "RECORD_COUNT" in values[0]:

            cleaned_values = []

            for row in values:

                label = (
                    row.get("QID")
                    or row.get("QID_LABEL")
                    or row.get("US_REGION")
                    or row.get("ENTITY")
                    or ""
                )

                if str(label).strip().lower() in [
                    "sigma",
                    "total",
                    "all",
                ]:
                    continue

                cleaned_values.append(row)

            if cleaned_values:
                values = cleaned_values

            sorted_vals = sorted(
                values,
                key=lambda x: x.get("RECORD_COUNT", 0),
                reverse=True
            )

            total = sum(
                v.get("RECORD_COUNT", 0)
                for v in sorted_vals
            )

            for rank, row in enumerate(sorted_vals[:5], start=1):

                pct = round(
                    (row["RECORD_COUNT"] / total) * 100,
                    1
                ) if total else 0

                label = (
                    row.get("QID")
                    or row.get("QID_LABEL")
                    or row.get("US_REGION")
                    or row.get("ENTITY")
                    or f"Segment {rank}"
                )

                if rank == 1:

                    bullets.append(
                        f'{label} contributes the highest volume with '
                        f'{row["RECORD_COUNT"]:,} records '
                        f'(~{pct}% of total).'
                    )

                else:

                    bullets.append(
                        f'{label} ranks #{rank} with '
                        f'{row["RECORD_COUNT"]:,} records '
                        f'(~{pct}% of total).'
                    )

            if len(sorted_vals) > 5:

                lowest = sorted_vals[-1]

                low_label = (
                    lowest.get("QID")
                    or lowest.get("QID_LABEL")
                    or lowest.get("US_REGION")
                    or "Lowest segment"
                )

                bullets.append(
                    f'{low_label} has the smallest volume with '
                    f'{lowest["RECORD_COUNT"]:,} records.'
                )

        if not bullets:

            bullets.extend(
                generic_chart_insights(
                    chart_spec,
                    values
                )
            )

    except Exception as e:

        print("Chart insight extraction error:", e)

    return bullets


# =====================================================
# MAIN ENGINE
# =====================================================

def extract_bullets(
    final_text: str,
    chart_spec=None,
    top_n=None
):

    text_without_tables = remove_markdown_tables(
        final_text
    )

    cleaned = clean_markdown(
        text_without_tables
    )

    bullets = []

    # =================================================
    # CHART INSIGHTS
    # =================================================

    if chart_spec:

        bullets.extend(
            normalize_bullet(bullet)
            for bullet in extract_chart_insights(chart_spec)
        )

    # =================================================
    # TEXT INSIGHTS
    # =================================================

    text_insights = extract_text_insights(
        cleaned
    )

    # =================================================
    # TABLE INSIGHTS
    # =================================================

    table_insights = extract_markdown_table_insights(
        final_text
    )

    text_insights.extend(table_insights)

    # =================================================
    # ADD TEXT BULLETS
    # =================================================

    for insight in text_insights:

        insight = normalize_bullet(insight)

        if len(insight) < 25:
            continue

        if not is_worthy_text_insight(insight):
            continue

        if is_duplicate(
            insight,
            bullets
        ):
            continue

        bullets.append(insight)

    # =================================================
    # FINAL DEDUPE
    # =================================================

    final_bullets = []

    for bullet in bullets:

        bullet = normalize_bullet(bullet)

        if len(bullet) < 20:
            continue

        if is_duplicate(
            bullet,
            final_bullets
        ):
            continue

        final_bullets.append(bullet)

        if top_n and len(final_bullets) >= top_n:
            break

    return final_bullets


# =====================================================
# RESPONSE WRAPPER
# =====================================================

def extract_bullets_from_response(
    response_data,
    top_n=None
):

    if isinstance(response_data, str):

        response_data = json.loads(
            response_data
        )

    # ============================================
    # OLD FORMAT SUPPORT
    # ============================================

    if "final_text" in response_data:

        final_text = response_data.get(
            "final_text",
            ""
        )

        chart_specs = response_data.get(
            "chart_specs",
            []
        )

    # ============================================
    # NEW STREAM FORMAT SUPPORT
    # ============================================

    else:

        text_parts = []

        chart_specs = []

        for item in response_data.get("content", []):

            event_type = item.get(
                "event_type",
                ""
            )

            data = item.get(
                "data",
                {}
            )

            # text events
            if (
                "text" in event_type.lower()
                or event_type == ""
            ):

                text = data.get(
                    "text",
                    ""
                )

                if text:
                    text_parts.append(text)

            # chart events
            elif event_type in [
                "chart",
                "response.chart"
            ]:

                chart_spec = data.get(
                    "chart_spec"
                )

                if chart_spec:
                    chart_specs.append(chart_spec)

        final_text = "\n".join(text_parts)

    if not final_text and isinstance(response_data.get("content"), list):

        fallback = []

        for item in response_data["content"]:

            data = item.get("data", {})

            if isinstance(data, dict):

                text = data.get("text")

                if text:
                    fallback.append(text)

        final_text = "\n".join(fallback)

    all_bullets = []

    # chart bullets
    for spec in chart_specs:

        all_bullets.extend(
            extract_chart_insights(spec)
        )

    # text bullets
    all_bullets.extend(
        extract_bullets(
            final_text=final_text,
            chart_spec=None,
            top_n=None
        )
    )

    # dedupe
    final = []
    if not all_bullets:

        sentences = split_sentences(final_text)

        for s in sentences:
            s = normalize_bullet(s)

            if (
                len(s) > 20
                and is_worthy_text_insight(s)
            ):
                final.append(s)

            if top_n and len(final) >= top_n:
                break

        return final

    for bullet in all_bullets:

        bullet = normalize_bullet(bullet)

        if is_duplicate(bullet, final):
            continue

        final.append(bullet)

        if top_n and len(final) >= top_n:
            break

    return final
