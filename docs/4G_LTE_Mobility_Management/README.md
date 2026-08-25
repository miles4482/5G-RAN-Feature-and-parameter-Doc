# 4G LTE Mobility Management — eRAN21.1 v3.0

Step-by-step summary of the three Huawei feature-parameter books. Not the SN sample template.

## File

[4G_LTE_Mobility_Management_eRAN21.1_v3.0.xlsx](./4G_LTE_Mobility_Management_eRAN21.1_v3.0.xlsx)

## How to read it

Open the sheets in order:

| Sheet | What it is |
|---|---|
| Read me | Three books, one chain, how to use the file |
| 1. End-to-end chain | Power-on → idle → RRC → coverage → FreqPri → MLB → next idle |
| 2. Idle Mode | Idle Mode Management, Issue 04, procedure order |
| 3. Connected Mode | Connected Mode, Issue 08, procedure order |
| 4. Intra-RAT MLB | Intra-RAT MLB, Issue 10, procedure order |
| 5. Activation order | One MML sequence across all three books |

Idle decides the next access cell. Connected is the HO engine. MLB decides who/when to move for load. Frequency-priority HO is not MLB.

MML uses placeholder `LocalCellId` / `DlEarfcn`. Confirm enums and defaults in MAE. Example dBm in the Connected book are not design values.
