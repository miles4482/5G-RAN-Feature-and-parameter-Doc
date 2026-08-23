# 4G LTE Mobility Management Workbook (Huawei eRAN21.1)

Excel deliverable for Robi Axiata PLC 4G RNO daily work: idle mobility, connected-mode mobility (A1–A5 / frequency priority), intra-RAT MLB, and a guarded AI change-request logic for capacity-layer DL user-throughput gaps **> 2 Mbps**.

## File

- [Mobility_Management_eRAN21.1_Workbook.xlsx](./Mobility_Management_eRAN21.1_Workbook.xlsx)

Regenerate after editing the builder:

```bash
pip3 install -r tools/requirements-mobility-workbook.txt
python3 tools/build_mobility_workbook.py
```

`Sample_file1_template.csv` is the original 11-item template that was filled. The power-saving MML example in that template (`SymbolShutdownSwitch`) is intentionally not used.

## Sheets (read in this order)

| Sheet | Content |
|---|---|
| `00_Cover` | Documents, Robi layer architecture, objective, safety notes |
| `01_Idle_Mode` | Feature 1 — all 11 template items + idle MML |
| `02_Connected_Mode` | Feature 2 — A1–A5, FreqPri, MML |
| `03_IntraRAT_MLB` | Feature 3 — MLB algorithms, L900 protection, MML |
| `04_AI_CR_Logic` | RCA classes, CR approval rule, rollback |
| `05_Parameter_Index` | Filterable cross-feature parameter list |

## Source documents

1. Idle Mode Management, eRAN21.1, Issue 04  
2. Mobility Management in Connected Mode, eRAN21.1, Issue 08  
3. Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10  

## Important

- Green cells are **Robi engineering proposals**, not official Huawei defaults. Calibrate on live KPI before CR.
- The sample MML row in the original template (`SymbolShutdownSwitch`) is a power-saving example and is **not** used here.
- L900 is a coverage layer: coverage HO in is allowed; routine MLB targeting is not.
- MML uses placeholder `LocalCellId` / `DlEarfcn`. Validate syntax on the running eRAN version in MAE-Access.
