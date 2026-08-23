# 4G LTE Mobility Management — document summary

Excel summary of the three Huawei eRAN21.1 Mobility Management feature books. Operator Excel format (dark-blue title, yellow-green sections, light-blue headers, grey rows).

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
| Intra-RAT MLB | Feature 3 — Huawei MLB + MML |

Each feature sheet follows the original template: Introduction, Triggering Conditions, eNodeB Actions, prerequisites, mutual impact, license, full parameter list, sequential MML.

## Sources

1. Idle Mode Management, eRAN21.1, Issue 04  
2. Mobility Management in Connected Mode, eRAN21.1, Issue 08  
3. Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10  

## Notes

- Proposed values are Robi engineering recommendations, not Huawei defaults.
- Original CSV `SymbolShutdownSwitch` row is power saving and is not used.
- MML uses placeholder `LocalCellId` / `DlEarfcn`. Validate in MAE-Access.
