---
name: cost-proposal-kutlu
description: "Generate a formatted Excel cost proposal workbook (.xlsx) for Kutlu with three sheets: Summary (by phase), Detail (line items with Phase/Task/Role/Hours/Customer Rate/Kutlu Cost), and Rate Card (editable rate assumptions). Default rates are $225/hr customer, $150/hr Kutlu. Use whenever a proposal, quote, bid, or cost estimate needs an accompanying .xlsx cost workbook for Kutlu. Triggers on: 'cost proposal', 'cost workbook', 'pricing spreadsheet for Kutlu', 'Kutlu cost proposal', 'generate xlsx cost', 'build cost workbook', 'bid pricing', 'quote spreadsheet', or any proposal workflow that needs a separate cost/pricing Excel alongside the docx. Do NOT use for NHA-specific RFP responses (use shandiin-nha-rfp-response instead) or for general spreadsheet tasks (use xlsx skill instead)."
---

# Cost Proposal Workbook Builder

Generate a 3-sheet Excel workbook that presents proposal pricing in Kutlu's preferred format: customer-facing rates alongside internal cost, with formulas throughout so the workbook stays dynamic when rates or hours change.

## When to use

This skill produces the **cost/pricing xlsx** that accompanies a proposal docx. Every quote to Kutlu should pair a docx technical proposal with an xlsx cost workbook built by this skill.

For NHA-specific RFPs, the `shandiin-nha-rfp-response` skill handles both the docx and xlsx together. This skill is for **all other proposals**.

## Workbook structure

| Sheet | Purpose |
|---|---|
| **Summary** | Totals by phase/category — all formulas, no hardcodes |
| **Detail** | Line items: Phase, Task, Role, Hours, Customer Rate/Total, Kutlu Rate/Total |
| **Rate Card** | Role-based rate assumptions — editable inputs that flow to Detail via VLOOKUP |

Default rates: **$225/hr** customer, **$150/hr** Kutlu. The Rate Card sheet uses blue text on yellow background for editable cells.

Load `references/workbook-schema.md` for the complete column-by-column schema and formula definitions.

## Workflow

### 1. Define line items

Cost data can live in **either** of two places in the proposal markdown file. The build script tries YAML first, then falls back to markdown tables.

#### Option A: YAML pricing block (preferred for structured proposals)

Include a `pricing` block in the YAML frontmatter:

```yaml
---
title: "Project Title"
pricing:
  rates:
    customer_facing_rate: 225
    subcontractor_bill_rate: 150
  phases:
    - name: "Phase 1 — Assessment and design"
      hours: 60
    - name: "Phase 3 — Implementation"
      hours: 40
    - name: "Hyper-V — Assessment"
      hours: 50
      role: "Senior Consultant"    # optional — defaults to "General"
      customer_rate: 250           # optional — overrides default
      kutlu_rate: 175              # optional — overrides default
  travel:                          # optional
    description: "Airfare, rental car, hotel, and meals"
    amount: 5000
---
```

Phase names are auto-split into Phase/Task on ` — ` (em-dash with spaces). "Hyper-V — Assessment" becomes Phase="Hyper-V", Task="Assessment".

#### Option B: Markdown tables with cost markers

Include cost data in the proposal body using `<!-- cost-start -->` / `<!-- cost-end -->` markers:

Tables inside these markers become the workbook's Detail rows. At minimum include **Task** and **Hours** columns. The build script auto-detects columns by header keywords. A full table looks like:

```markdown
<!-- cost-start -->
## Staffing and Cost

| Phase | Task | Role | Hours |
|---|---|---|---|
| Discovery | Requirements analysis | Senior Consultant | 24 |
| Discovery | Stakeholder interviews | Analyst | 16 |
| Build | System implementation | Senior Consultant | 60 |
| Build | Testing & QA | Analyst | 32 |
<!-- cost-end -->
```

If rates differ from defaults ($225/$150), add `Customer Rate` and `Kutlu Rate` columns. Otherwise the Rate Card defaults apply via VLOOKUP formulas.

Phase/category groupings in the table drive the Summary sheet rows. Use consistent phase names across line items.

### 2. Build the workbook

**Always run the bundled script — do NOT rewrite the logic as inline Python.** The script handles YAML extraction, column classification, VLOOKUP formulas, rate card generation, and edge cases. Writing a one-off script will miss these.

```bash
python {skill_path}/scripts/build_cost_workbook.py <markdown_file> [output.xlsx]
```

On Windows, if WSL2 is available, prefer running through WSL for better LibreOffice compatibility during recalculation:

```bash
wsl python3 {skill_path_wsl}/scripts/build_cost_workbook.py <markdown_file> [output.xlsx]
```

Convert Windows paths to WSL paths with `wslpath -u` or use `/mnt/c/...` format. The script itself is cross-platform — openpyxl works identically on both.

The script:
- Parses YAML frontmatter for metadata (RFP number, title, vendor name)
- Extracts tables from `<!-- cost-start/end -->` markers (or treats entire body as cost data)
- Generates a 3-sheet workbook with formulas (`SUMIFS`, `VLOOKUP`, multiplication)
- Sets `fullCalcOnLoad=True` so Excel recalculates on open
- Auto-derives the output filename from frontmatter if no output path given

If the xlsx skill's `scripts/recalc.py` is available, run it afterward to pre-populate formula values:

```bash
python {xlsx_skill_path}/scripts/recalc.py output.xlsx
```

### 3. Verify output

Confirm:
- [ ] Three sheets exist: Summary, Detail, Rate Card
- [ ] Detail has correct number of line items (no totals rows counted as items)
- [ ] Summary phases match the phases in Detail
- [ ] Rate Card has entries for every role mentioned in Detail
- [ ] No formula errors (#REF!, #VALUE!, etc.)
- [ ] Currency formatting applied to dollar columns, hours formatting on hours column
- [ ] Grand total formulas reference all data rows

### 4. Deliver alongside the proposal docx

The xlsx is a companion to the docx proposal. When reporting to the user, list both files:

```
Deliverables:
  Technical proposal: [proposal name].docx
  Cost workbook:      Cost Proposal - [title].xlsx
```

## Frontmatter fields

The build script reads YAML frontmatter for both pricing data and metadata. It supports two patterns:

**Proposal-style** (with pricing block — used by proposal-writing agents):
```yaml
---
title: "NTUA SCADA Infrastructure Modernization"
opportunity_id: OPP-2026-03-021
pricing:
  rates:
    customer_facing_rate: 225
    subcontractor_bill_rate: 150
  phases:
    - name: "Phase 1 — Assessment"
      hours: 60
  travel:
    description: "Airfare, rental car, hotel"
    amount: 5000
---
```

**RFP-style** (metadata only — cost data in markdown tables):
```yaml
---
RFP:
  Number: "#123"
  Title: "IT Services Engagement"
Proposal:
  Title: "IT Modernization Project"
Vendor:
  Name: "Shandiin Solutions, LLC"
---
```

If no frontmatter, the workbook still builds — it just skips the title/header metadata.

## Handling multiple roles

When the proposal involves different billing rates by role:

1. Define each role with its rates in the markdown table (add Customer Rate / Kutlu Rate columns), OR
2. Let the build script use defaults and manually adjust the Rate Card sheet after generation

The Rate Card sheet acts as the single source of truth — changing a rate there updates all Detail rows for that role via VLOOKUP.

## Edge cases

- **No role column**: All items map to "General" role on the Rate Card
- **No phase column**: Section headings (`## Phase Name`) above tables are used as the phase
- **Mixed tables**: Multiple tables in the cost block are merged into one Detail sheet
- **Bold total rows**: Rows with `**bold**` cells are detected as totals and excluded from Detail (they get recalculated by formulas)

## Common mistakes to avoid

- **Do NOT create a new Python script** to generate the workbook. The bundled `scripts/build_cost_workbook.py` already handles all input formats, formulas, formatting, and edge cases. Call it directly.
- **Do NOT hardcode calculated values** in cells. The workbook uses Excel formulas (`VLOOKUP`, `SUMIFS`, multiplication) so it stays dynamic when hours or rates change.
- **Do NOT write the xlsx to the skill directory.** Write it to the proposal/output directory alongside the docx.

## References

| File | Load when |
|---|---|
| `references/workbook-schema.md` | Need exact column definitions, formula patterns, or markdown input format |
