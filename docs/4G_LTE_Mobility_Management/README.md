# 4G LTE Mobility Management — eRAN21.1 v3.3

Step-by-step summary of the three Huawei feature-parameter books, with the document charts on the sheets.

## File

[4G_LTE_Mobility_Management_eRAN21.1_v3.3.xlsx](./4G_LTE_Mobility_Management_eRAN21.1_v3.3.xlsx)

## How to read it

Open the sheets in order: Read me → End-to-end chain → Idle → Connected → Intra-RAT MLB → Activation order.

Each feature sheet starts with the **document chart** (Idle Fig 4-1 / 5-1, Connected Fig 4-1 and Fig 11-1 / 11-2, MLB Fig 3-1 / 4-1 / 4-4 / 4-5). The original PDF page bitmaps are not in this workspace, so the drawings are rebuilt from the documented procedures and labelled with the same figure numbers.

Parameter tables: one parameter per row. **Core parameter** = required to activate. **Basic and Optimized** = tune after activation. Parameter Value and Source use the Huawei document name.

The last section of every sheet is **Combined MML Command (all Together)** (SN / MML / Purpose / Note).

MML uses documented values where the book gives them (CFG, NORMAL, ALLOWED, A4, UE_NUMBER_ONLY, ONLY_STRONGEST_CELL, SNonIntraSearch=10). Remaining thresholds: calibrate from MR. Example A1/A2/A4 dBm in the Connected book are not design values.
