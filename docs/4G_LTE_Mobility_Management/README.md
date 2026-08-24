# Mobility Management — eRAN21.1 PDF summary

Exclusive Excel summary of the three Huawei feature-parameter documents. No load-balance agent content.

## File

[4G_LTE_Mobility_Management_eRAN21.1_v2.0.xlsx](./4G_LTE_Mobility_Management_eRAN21.1_v2.0.xlsx)

```bash
python3 tools/build_mobility_workbook.py
```

## How to read it

1. Open **Mobility Management** (cover). Click a feature name to jump to that sheet.
2. Each feature sheet is **SN-1 to SN-11** in connected order, matching the original template.
3. Click an SN number in the SN list to jump to that section.
4. Column **H** is Chart / Doc Ref on the right of every SN.
5. Use the Excel outline (+/−) at the left of each SN header to collapse a section.

| Sheet | Source PDF |
|---|---|
| Mobility Management | Cover / folder / how the three books fit |
| Idle Mode Management | Idle Mode Management, eRAN21.1, Issue 04 |
| Connected Mode | Mobility Management in Connected Mode, eRAN21.1, Issue 08 |
| Intra-RAT MLB | Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10 |

## SN list (every feature)

| SN | Item |
|---|---|
| 1 | Working Principal (activities in connected sequence + process chart) |
| 2 | Major highlighted Point |
| 3 | Benefit and Limitations |
| 4 | Selection criteria / Trigger |
| 5 | Activation parameter / Switch (table: value + Conditional Parameter) |
| 6 | Prerequisite (table) |
| 7 | Mutually impacted (table) |
| 8 | Relation with Other Feature |
| 9 | License |
| 10 | All Parameter List (sequence; Conditional Parameter = relation with other feature) |
| 11 | Final MML (sequence) |

SN-11 columns A–G are exactly:

`Parameter Sequence | MO | Activation Value | Conditional Parameter | Remarks | Parameter Description | More Notes`

Column H is Chart / Doc Ref.

## Format (operator Excel sample)

| Element | Background | Text |
|---|---|---|
| Title bar (row 1, A–H) | Dark blue `#005596` | White, bold, centered |
| Colour key (row 2) | Mixed | Label of each colour |
| SN number (column A of SN header) | Dark blue `#005596` | White, bold |
| SN title | Yellow `#FFFF00` | Bold green `#008000` |
| Table header | Light blue `#DDEBF7` | Black, bold |
| Data rows | Light grey `#F2F2F2` | Black |
| MAJOR badge | Green / light yellow `#FFF2CC` | Bold green |
| Process chart boxes | Dark blue | White |

Eight columns A–H. White spacer rows. Calibri. Gridlines on. Landscape A3.

The original CSV sample MML (`SymbolShutdownSwitch`) is Symbol Power Saving and is not used.

MML uses placeholder `LocalCellId` / `DlEarfcn`. Confirm enums, defaults and syntax in MAE-Access / the version-matched parameter reference.
