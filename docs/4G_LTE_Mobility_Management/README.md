# 4G LTE Mobility Management — eRAN21.1 v3.2

Step-by-step summary of the three Huawei feature-parameter books.

## File

[4G_LTE_Mobility_Management_eRAN21.1_v3.2.xlsx](./4G_LTE_Mobility_Management_eRAN21.1_v3.2.xlsx)

## How to read it

Open the sheets in order: Read me → End-to-end chain → Idle → Connected → Intra-RAT MLB → Activation order.

Parameter tables: one parameter per row. **Core parameter** = required to activate. **Basic and Optimized** = tune after activation. Parameter Value and Source use the Huawei document name.

The last section of every sheet is **Combined MML Command (all Together)** (SN / MML / Purpose / Note).

MML uses documented values where the book gives them (CFG, NORMAL, ALLOWED, A4, UE_NUMBER_ONLY, ONLY_STRONGEST_CELL, SNonIntraSearch=10). Remaining thresholds: calibrate from MR. Example A1/A2/A4 dBm in the Connected book are not design values.
