# Workbook Schema

The cost proposal workbook contains three sheets. This document defines the exact structure so agents can produce compatible markdown input.

## Sheet 1: Summary

Totals by phase/category. All cells are formulas referencing the Detail sheet.

| Column | Content | Formula |
|---|---|---|
| A | Phase / Category | Text from Detail!A |
| B | Total Hours | `=SUMIFS(Detail!D:D, Detail!A:A, A{row})` |
| C | Customer Total | `=SUMIFS(Detail!F:F, Detail!A:A, A{row})` |
| D | Kutlu Total | `=SUMIFS(Detail!H:H, Detail!A:A, A{row})` |

Last row: **GRAND TOTAL** with `=SUM()` of each column.

## Sheet 2: Detail

Full line-item breakdown. This is the data entry sheet.

| Column | Header | Content | Notes |
|---|---|---|---|
| A | Phase | Phase / category name | Groups into Summary rows |
| B | Task/Deliverable | Description of work item | Primary description |
| C | Role | Role or resource type | Must match a Rate Card role |
| D | Hours | Estimated hours | **Editable input** (blue, yellow bg) |
| E | Customer Rate ($/hr) | `=VLOOKUP(C{row}, 'Rate Card'!A:B, 2, FALSE)` | Formula; falls back to $225 |
| F | Customer Total | `=D{row}*E{row}` | Formula |
| G | Kutlu Rate ($/hr) | `=VLOOKUP(C{row}, 'Rate Card'!A:C, 3, FALSE)` | Formula; falls back to $150 |
| H | Kutlu Total | `=D{row}*G{row}` | Formula |

Last row: **TOTAL** with `=SUM()` of columns D, F, H.

## Sheet 3: Rate Card

Role/rate assumptions. All rates are editable inputs (blue text, yellow background).

| Column | Header | Content |
|---|---|---|
| A | Role | Role name (must match Detail!C values) |
| B | Customer Rate ($/hr) | Default $225.00, editable |
| C | Kutlu Cost ($/hr) | Default $150.00, editable |
| D | Notes | "Default" or "Custom" |

## Formatting conventions

- **Segoe UI** throughout (matches Shandiin/Microsoft brand templates)
- **Blue text** on yellow background = editable input cells (hours, rates on Rate Card)
- **Dark blue header fill** (#2B579A) with white text
- **Currency**: `$#,##0.00`
- **Hours**: `#,##0.0`
- **Totals row**: double-top border
- `fullCalcOnLoad=True` set so Excel recalculates all formulas on open

## Default rates

| Role | Customer Rate | Kutlu Cost |
|---|---|---|
| General | $225.00 | $150.00 |

The agent should add role-specific rates when the proposal specifies multiple roles (e.g., Senior Consultant $250/$175, Analyst $175/$100).

## Markdown input format

The build script expects markdown tables. At minimum, tables should include:

### Minimal table (auto-fills rates from defaults)

```markdown
| Task | Hours |
|---|---|
| Phase 1: Assessment | 40 |
| Phase 2: Implementation | 80 |
```

### Full table (explicit rates and roles)

```markdown
| Phase | Task | Role | Hours | Customer Rate | Kutlu Rate |
|---|---|---|---|---|---|
| Discovery | Requirements gathering | Senior Consultant | 20 | $250 | $175 |
| Discovery | Stakeholder interviews | Analyst | 16 | $175 | $100 |
| Implementation | System configuration | Senior Consultant | 40 | $250 | $175 |
```

### With cost markers (for proposal integration)

Wrap tables in markers so the proposal agent can separate cost data from the technical narrative:

```markdown
<!-- cost-start -->
## Hours and cost summary

| Phase | Task | Role | Hours |
|---|---|---|---|
| Phase 1 | Assessment | Senior Consultant | 40 |
| Phase 2 | Implementation | Consultant | 80 |
<!-- cost-end -->
```

The script recognizes these column header keywords:

| Semantic role | Matched headers (case-insensitive) |
|---|---|
| Phase | phase, category |
| Task | task, description, service, item, deliverable |
| Role | role, resource, position |
| Hours | hour, hours |
| Customer rate | customer rate, customer price, customer /hr |
| Kutlu rate | kutlu rate, our rate, internal rate, cost rate |
