# 4G LTE Mobility Management Workbook (Huawei eRAN21.1)

Excel file uses the **attached operator format** (same colors as the Symbol Power Saving sample):

| Element | Background | Text |
|---|---|---|
| Title bar (row 1, A–F) | Dark blue `#005596` | White, bold, centered |
| Section header | Yellow `#FFFF00` | Bold green `#008000` |
| Table header | Light blue `#DDEBF7` | Black, bold |
| Data rows | Light grey `#F2F2F2` | Black |
| Spacing | White empty rows | — |

Six columns A–F. Calibri. Gridlines on. No dark background.

## File

[Mobility_Management_eRAN21.1_Workbook.xlsx](./Mobility_Management_eRAN21.1_Workbook.xlsx)

```bash
python3 tools/build_mobility_workbook.py
```

## Sheets

| Sheet | Content |
|---|---|
| Mobility Management | Introduction + Robi layer map |
| Idle Mode Management | Feature 1 — SN-1 to SN-11 + MML |
| Connected Mode | Feature 2 — A1–A5 / FreqPri + MML |
| Intra-RAT MLB | Feature 3 — MLB + MML |
| Daily KPI | Date / Open / High / Low / Close tracker |

Each feature sheet is split like the sample: Section 1 Introduction, Section 2 Triggering Conditions, Section 3 eNodeB Actions, then prerequisites, impacts, license, parameter list, MML.

## Daily KPI

- **Date** = busy-hour date  
- **Open** = L1800 DL user throughput (Mbps)  
- **High** = max capacity-layer TP  
- **Low** = min capacity-layer TP  
- **Close** = High − Low (investigate if > 2 Mbps)  

L900 is not included in High / Low.

## Notes

- Proposed values are Robi engineering recommendations, not Huawei defaults.
- Original CSV `SymbolShutdownSwitch` row is power saving and is not used.
- MML uses placeholder `LocalCellId` / `DlEarfcn`. Validate in MAE-Access.
