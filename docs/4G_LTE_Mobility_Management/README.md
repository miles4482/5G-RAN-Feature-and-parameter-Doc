# 4G LTE Mobility Management Workbook (Huawei eRAN21.1)

Excel file prepared in the **required picture format**:

- Black background
- Header order and colors (mandatory):
  - **Date** — white
  - **Open** — yellow
  - **High** — yellow
  - **Low** — yellow
  - **Close** — light blue

Plus the original 11-item template (SN-1 to SN-11) for each Mobility Management feature.

## File

[Mobility_Management_eRAN21.1_Workbook.xlsx](./Mobility_Management_eRAN21.1_Workbook.xlsx)

Regenerate:

```bash
pip3 install -r tools/requirements-mobility-workbook.txt
python3 tools/build_mobility_workbook.py
```

## Sheets (same tree as Mobility Management folder)

| Sheet | Content |
|---|---|
| `Mobility Management` | Format legend + introduction + layer map |
| `Idle Mode Management` | Feature 1 — SN-1 to SN-11 + MML |
| `Connected Mode` | Feature 2 — A1–A5 / FreqPri + MML |
| `Intra-RAT MLB` | Feature 3 — MLB + MML |
| `Daily KPI` | Date / Open / High / Low / Close tracker + RCA |

## Daily KPI column meaning

| Picture field | Color | Meaning in this network |
|---|---|---|
| Date | white | Busy-hour date |
| Open | yellow | L1800 (or anchor) DL user throughput Mbps |
| High | yellow | Maximum capacity-layer throughput |
| Low | yellow | Minimum capacity-layer throughput |
| Close | light blue | Gap = High − Low. Investigate if Close > 2 Mbps |

L900 is not included in High / Low.

On feature sheets the same five fields are the first five table columns:

| Date | Open | High | Low | Close |
|---|---|---|---|---|
| SN / Sequence | MO | Activation / proposed | Conditional parameter | Remarks / chart / result |

MML columns after that stay: Parameter Description, More Notes.

## Important

- Yellow / light-blue proposed values are Robi engineering recommendations, not Huawei defaults.
- Template sample MML `SymbolShutdownSwitch` is power saving and is not used.
- MML uses placeholder `LocalCellId` / `DlEarfcn`. Validate in MAE-Access.
