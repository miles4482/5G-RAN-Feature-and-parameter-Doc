# Mobility Management — eRAN21.1 PDF summary

Exclusive Excel summary of the three Huawei feature-parameter documents. No load-balance agent content.

## File

[Mobility_Management_eRAN21.1_Workbook.xlsx](./Mobility_Management_eRAN21.1_Workbook.xlsx)

```bash
python3 tools/build_mobility_workbook.py
```

## Format (attached Excel sample)

| Element | Background | Text |
|---|---|---|
| Title bar (row 1, A–F) | Dark blue `#005596` | White, bold, centered |
| Section header | Yellow `#FFFF00` | Bold green `#008000` |
| Table header | Light blue `#DDEBF7` | Black, bold |
| Data rows | Light grey `#F2F2F2` | Black |

Six columns A–F. White spacer rows. Calibri. Gridlines on.

Each feature follows SN-1 to SN-11 from the original template, grouped as:

1. Feature Introduction  
2. Triggering Conditions  
3. eNodeB Actions  
4. Prerequisites  
5. Mutual impact and related features  
6. License  
7. All parameter list  
8. Final MML (sequence)

## Sheets

| Sheet | Source PDF |
|---|---|
| Mobility Management | Cover / how the three books fit |
| Idle Mode Management | Idle Mode Management, eRAN21.1, Issue 04 |
| Connected Mode | Mobility Management in Connected Mode, eRAN21.1, Issue 08 |
| Intra-RAT MLB | Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10 |

The original CSV sample MML (`SymbolShutdownSwitch`) is Symbol Power Saving and is not used.

MML uses placeholder `LocalCellId` / `DlEarfcn`. Confirm enums, defaults and syntax in MAE-Access / the version-matched parameter reference.
