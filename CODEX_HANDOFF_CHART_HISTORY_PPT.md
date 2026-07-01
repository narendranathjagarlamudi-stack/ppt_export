# Chart Editing + PPT Export Handoff

## Context

The project has chat/history responses stored in `di_chat_history.response`. Newer responses are stored as ordered events, for example:

- `response.text`
- `response.chart`

Each chart event contains a Vega-Lite `chart_spec` and metadata like `content_index`.

Collections/reports are built from selected charts and later exported to PPT. A single `dih_id` / message can contain multiple charts, and users may choose only some charts for a collection/export.

## Main Decision

Keep `di_chat_history.response` immutable as the original raw response.

Store chart-level records separately in `di_chart`, including the original Vega-Lite spec and the latest edited Vega-Lite spec.

History API and PPT export should both use the same resolved chart logic:

```text
original history response
+ edited chart specs from di_chart
= resolved response returned to frontend / used for PPT
```

This keeps frontend history, collection reports, and PPT exports consistent.

## Chart Identity

For now, support the new response format only.

Use this chart-level key:

```text
dih_id + content_index
```

Reason:

- New `response.chart` events contain `data.content_index`.
- One `dih_id` can contain multiple charts.
- `dih_id` alone is not enough.

Old response format does not have a reliable identifier, so old-format chart edit persistence can be handled later.

## PostgreSQL Schema Changes

Current DB is PostgreSQL, not Snowflake.

Use `INTEGER` instead of `NUMBER`, and `JSONB` instead of `VARIANT`.

```sql
ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS message_id VARCHAR;
ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS content_index INTEGER;

ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS original_chart_spec JSONB;
ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS edited_chart_spec JSONB;
ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS is_edited BOOLEAN DEFAULT FALSE;

ALTER TABLE di_chart ADD COLUMN IF NOT EXISTS modified_by VARCHAR;
```

Add uniqueness so the same chart under one `dih_id` is not duplicated:

```sql
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conname = 'uq_di_chart_dih_content_index'
    ) THEN
        ALTER TABLE di_chart
        ADD CONSTRAINT uq_di_chart_dih_content_index
        UNIQUE (dih_id, content_index);
    END IF;
END $$;
```

## Why `is_edited`

If `di_chart` stores one row per discovered chart, `is_edited` is useful.

Behavior:

```text
is_edited = false
edited_chart_spec = null
=> use original_chart_spec

is_edited = true
edited_chart_spec = {...}
=> use edited_chart_spec
```

If the user resets/reverts to original:

```sql
UPDATE di_chart
SET edited_chart_spec = NULL,
    is_edited = FALSE,
    modified_by = :user_id,
    modified_at = CURRENT_TIMESTAMP
WHERE dih_id = :dih_id
  AND content_index = :content_index;
```

## History API Resolution Flow

When history API is called:

```text
1. Fetch original row from di_chat_history by dih_id/message.
2. Parse response JSON.
3. Find all response.chart events.
4. For each chart, read data.content_index.
5. Look up di_chart using dih_id + content_index.
6. If is_edited = true and edited_chart_spec exists, replace event.data.chart_spec with edited spec.
7. Otherwise keep original chart_spec from history/original_chart_spec.
8. Return resolved response to frontend.
```

## PPT Export Flow

PPT export should not parse only raw `di_chat_history.response`.

It should use the same resolved chart service as history API.

That means:

- Edited chart type is reflected in PPT.
- Edited colors/palette are reflected in PPT.
- Edited legends/labels/title/axis settings are reflected where the PPT renderer supports them.
- Unedited charts use the original chart spec.

Recommended service shape:

```text
get_resolved_history(dih_id)
get_resolved_charts(dih_ids or collection/report_id)
```

Both frontend history and PPT export should call this shared service.

## Collection Report Flow

`di_chart_report_mapping` should map reports/collections to `di_chart.chart_id`.

When user adds a chart to a collection:

```text
1. Ensure chart exists in di_chart for dih_id + content_index.
2. Insert mapping into di_chart_report_mapping.
3. Use display_order for PPT slide order.
```

When exporting a collection:

```text
1. Fetch mapped chart rows in display_order.
2. For each chart, use edited_chart_spec if is_edited = true.
3. Otherwise use original_chart_spec.
4. Generate one PPT slide per selected chart.
```

## Vega-Lite Spec Editing

The frontend can edit the Vega-Lite spec directly.

`edited_chart_spec` can include:

- chart type: `mark`
- colors/palette: `encoding.color`, `scale`, `config`, `transform`
- legends: `encoding.color.legend`
- labels: `mark`, `encoding`, `config`
- title: `title`
- axes: `encoding.x.axis`, `encoding.y.axis`
- tooltip: `encoding.tooltip`

No separate `chart_config` column is needed for now. `edited_chart_spec` is the source of truth.

## Undo / Redo

For frontend-only undo/redo during a configure session:

```text
Keep undo/redo stack in frontend state.
Persist only the final edited_chart_spec when user applies/saves.
```

If server-side audit/version history is needed later, add a separate version table instead of overloading `di_chart`.

Possible future table:

```text
di_chart_spec_version
- version_id
- chart_id
- chart_spec JSONB
- created_by
- created_at
- change_type
```

## Important Notes

- Do not overwrite `di_chat_history.response`.
- Do not identify charts only by `dih_id`.
- For current implementation, only new response format supports persisted chart edits.
- Old response format can still be displayed/exported as-is, but edit persistence can be added later with generated chart indexes.
- PPT export should use resolved charts, not raw history, to stay consistent with frontend.

