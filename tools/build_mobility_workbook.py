#!/usr/bin/env python3
"""Build Mobility Management Excel using the attached operator template colors.

Format from the user's Excel screenshot (must follow):
  Title bar:     dark blue #005596, white bold, centered A-F
  Section header: bright yellow #FFFF00, bold green #008000
  Table header:  light sky blue #DDEBF7, black bold
  Data rows:     light grey #F2F2F2, black
  Spacing rows between sections
  Six columns A-F
  Calibri, visible gridlines

Content: document summary of Mobility Management —
Idle Mode, Connected Mode, Intra-RAT MLB (Huawei eRAN21.1).
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/4G_LTE_Mobility_Management/Mobility_Management_eRAN21.1_Workbook.xlsx"

# Attached-picture palette
TITLE_BG = "005596"
TITLE_FG = "FFFFFF"
SECTION_BG = "FFFF00"
SECTION_FG = "008000"
THEAD_BG = "DDEBF7"
THEAD_FG = "000000"
DATA_BG = "F2F2F2"
DATA_FG = "000000"
WHITE = "FFFFFF"
GRID = "B4B4B4"

thin = Border(
    left=Side(style="thin", color=GRID),
    right=Side(style="thin", color=GRID),
    top=Side(style="thin", color=GRID),
    bottom=Side(style="thin", color=GRID),
)
wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_c = Alignment(wrap_text=True, vertical="center", horizontal="center")
top = Alignment(wrap_text=True, vertical="top", horizontal="left")


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(size=10, bold=False, color=DATA_FG, italic=False):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)


def put(ws, r, c, val, size=10, bold=False, color=DATA_FG, bg=DATA_BG, align=None, h=None):
    cell = ws.cell(r, c, val)
    cell.font = font(size, bold, color)
    cell.fill = fill(bg)
    cell.alignment = align or wrap
    cell.border = thin
    if h:
        ws.row_dimensions[r].height = h
    return cell


def merge_put(ws, r, c1, c2, val, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    put(ws, r, c1, val, **kw)
    bg = kw.get("bg", DATA_BG)
    color = kw.get("color", DATA_FG)
    size = kw.get("size", 10)
    bold = kw.get("bold", False)
    for c in range(c1 + 1, c2 + 1):
        x = ws.cell(r, c)
        x.fill = fill(bg)
        x.border = thin
        x.font = font(size, bold, color)


def widths(ws, sizes):
    for i, w in enumerate(sizes, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def setup(ws, footer):
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5)
    ws.oddHeader.left.text = "Mobility Management"
    ws.oddFooter.left.text = footer
    ws.oddFooter.right.text = "Huawei eRAN21.1  |  Page &P of &N"
    ws.sheet_properties.tabColor = TITLE_BG
    ws.freeze_panes = "A3"
    ws.sheet_format.defaultRowHeight = 18


def title(ws, r, cols, text):
    merge_put(ws, r, 1, cols, text, size=16, bold=True, color=TITLE_FG, bg=TITLE_BG,
              align=Alignment(wrap_text=True, vertical="center", horizontal="center"), h=28)
    return r + 1


def spacer(ws, r, cols=6, h=12):
    for c in range(1, cols + 1):
        put(ws, r, c, "", bg=WHITE, h=h)
    return r + 1


def section(ws, r, cols, text):
    merge_put(ws, r, 1, cols, text, size=12, bold=True, color=SECTION_FG, bg=SECTION_BG,
              align=Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1), h=22)
    return r + 1


def headers(ws, r, names):
    for c, name in enumerate(names, 1):
        put(ws, r, c, name, size=10, bold=True, color=THEAD_FG, bg=THEAD_BG, align=wrap_c, h=22)
    return r + 1


def data_row(ws, r, values, h=28):
    for c, v in enumerate(values, 1):
        put(ws, r, c, v, size=10, color=DATA_FG, bg=DATA_BG, align=top, h=h)
    return r + 1


# ---------------------------------------------------------------------------
# Cover
# ---------------------------------------------------------------------------
def build_cover(wb):
    ws = wb.active
    ws.title = "Mobility Management"
    cols = 6
    widths(ws, [22, 14, 28, 32, 32, 52])
    setup(ws, "Mobility Management")

    r = 1
    r = title(ws, r, cols, "Mobility Management - Detailed Notes")
    r = spacer(ws, r, cols)

    r = section(ws, r, cols, "Section 1: Feature Introduction")
    r = headers(ws, r, ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"])
    intro_rows = [
        ["Purpose", "LTE FDD", "Document summary of Mobility Management", "Idle Mode + Connected Mode + Intra-RAT MLB", "Huawei eRAN21.1 feature books",
         "Parameter and feature summary from the three eRAN21.1 Mobility Management documents. SN-1 to SN-11 on each feature sheet."],
        ["Vendor / release", "LTE FDD", "Huawei eRAN21.1", "Idle Issue 04 / Connected Issue 08 / MLB Issue 10", "Documents already reviewed",
         "All parameter names are from those feature books. Confirm ranges in the matching parameter reference / MAE."],
        ["Capacity layers", "LTE FDD", "L1800 / L2100 / L2600 C1-C4", "20 / 15 / 4x20 MHz", "Capacity-balancing pool",
         "Nominal PRB share 100/75/100. Use Load=N/C with ActiveUe + SpectralEff. Do not equalize raw UE count."],
        ["Coverage layer", "LTE FDD", "L900 5 MHz indoor", "Coverage HO in = YES", "Routine MLB target = NO",
         "A5/coverage fallback to L900. A1/FreqPri escape from L900 when high-band is strong. Do not fill 5 MHz for TP equality."],
        ["Workbook tree", "LTE FDD", "Mobility Management", "Idle Mode Management / Connected Mode / Intra-RAT MLB", "Same folder order as requested",
         "Each feature sheet uses this same color format: dark-blue title, yellow-green section, light-blue header, grey data."],
        ["Template leftover", "LTE FDD", "Original CSV sample MML", "ENodeBAlgoSwitch / SymbolShutdownSwitch", "NOT USED",
         "That row is Symbol Power Saving, not mobility. It is removed from all MML tables in this book."],
        ["Safety", "LTE FDD", "Change control", "One parameter family per cluster", "Calibrate before live change",
         "Rollback if drop, VoLTE, re-est, HO success or L900 indoor KPI degrades. Huawei learned MLB thresholds collect 7 days and refresh every 7 days — do not overwrite them every day."],
    ]
    for row in intro_rows:
        r = data_row(ws, r, row, h=32)
    r = spacer(ws, r, cols)
    r = spacer(ws, r, cols)

    r = section(ws, r, cols, "Section 2: Robi Layer Map")
    r = headers(ws, r, ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"])
    layers = [
        ["L2600 C1", "LTE FDD", "Primary capacity", "20 MHz / 100 PRB / idle prio 7", "MLB source+target, A4", "Equal common priority with C2-C4. Use LOADPRIORITY among the four carriers."],
        ["L2600 C2", "LTE FDD", "Primary capacity", "20 MHz / 100 PRB / idle prio 7", "MLB source+target, A4", "Do not stack C1>C2>C3>C4 as static common priority."],
        ["L2600 C3", "LTE FDD", "Primary capacity", "20 MHz / 100 PRB / idle prio 7", "MLB source+target, A4", "CA SCell role must be checked before calling a carrier unloaded."],
        ["L2600 C4", "LTE FDD", "Primary capacity", "20 MHz / 100 PRB / idle prio 7", "MLB source+target, A4", "Same as other L2600 carriers."],
        ["L1800", "LTE FDD", "Broad capacity + mobility anchor", "20 MHz / 100 PRB / idle prio 6", "MLB source+target, A4 + coverage", "Typical fallback below L2600 when high-band load is high."],
        ["L2100", "LTE FDD", "Capacity (smaller BW)", "15 MHz / 75 PRB / idle prio 5 or 6", "MLB source+target, SE-normalized", "Do not copy 20 MHz UE-number threshold onto L2100."],
        ["L900", "LTE FDD", "Indoor / deep coverage only", "5 MHz / 25 PRB / idle prio 2", "NOT in capacity pool", "Coverage HO permitted. Connected/idle MLB targeting prohibited. CIO toward L900 = 0."],
    ]
    for row in layers:
        r = data_row(ws, r, row, h=28)
    r = spacer(ws, r, cols)
    r = spacer(ws, r, cols)

    r = section(ws, r, cols, "Section 3: How to read every sheet")
    r = headers(ws, r, ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"])
    how = [
        ["Title bar", "LTE FDD", "Row 1", "Dark blue #005596 / white bold", "Centered A-F", "Copied from the attached Symbol Power Saving Excel format."],
        ["Section header", "LTE FDD", "Section N: ...", "Yellow #FFFF00 / bold green #008000", "Full width", "One empty row before/after each section."],
        ["Table header", "LTE FDD", "Column titles", "Light blue #DDEBF7 / black bold", "Six columns A-F", "Headers change per section, same as the attached file."],
        ["Data row", "LTE FDD", "Content", "Grey #F2F2F2 / black", "Wrapped text", "Comfortable contrast. No dark background."],
        ["SN-1 to SN-11", "LTE FDD", "Original template sequence", "Working / Highlight / Benefit / Selection / Activation / Prerequisite / Mutual / Relation / License / Parameter list / MML",
         "Filled from eRAN21.1 books", "Chart / doc ref is in the last column (Notes)."],
    ]
    for row in how:
        r = data_row(ws, r, row, h=28)
    return ws


# ---------------------------------------------------------------------------
# Shared feature-sheet builder
# ---------------------------------------------------------------------------
def build_feature_sheet(wb, name, title_text, sections):
    ws = wb.create_sheet(name)
    cols = 6
    widths(ws, [22, 12, 30, 36, 36, 54])
    setup(ws, title_text)

    r = 1
    r = title(ws, r, cols, title_text)
    for sec in sections:
        r = spacer(ws, r, cols)
        r = section(ws, r, cols, sec["name"])
        r = headers(ws, r, sec["headers"])
        for row in sec["rows"]:
            r = data_row(ws, r, row, h=sec.get("h", 30))
        r = spacer(ws, r, cols)
    return ws


def idle_sections():
    return [
        {
            "name": "Section 1: Feature Introduction",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 36,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD", "Idle camping and reselection",
                 "Criterion S; SIB3/SIB5 priority; ThreshXhigh / ThrshServLow / ThreshXlow; equal-prio rank; T320 dedicated prio",
                 "Next RRC starts on the camped cell",
                 "CHART: PLMN → Criterion S → Camp + SIB3/5 → Measure/Reselect → RRC on camped cell. Ref [IM] Fig 4-1 p.6; Fig 5-1 pp.13-14."],
                ["SN-2 Major highlighted Point", "LTE FDD", "Priority is frequency-level",
                 "CellReselPriority / CellReselPriorityCfgInd / MeasPerformanceDemand / SNonIntraSearch",
                 "L2600=7 equal; L1800=6; L2100=5 or 6; L900=2",
                 "No SIB5 priority = no reselection to that freq (max 16). UNDELIVER blocks idle MLB. SI change is not instant ([IM] §7.1.3). QRxLevMin is a floor, not an MLB knob. Huawei example SNonIntraSearch=10."],
                ["Document", "LTE FDD", "Idle Mode Management", "eRAN21.1 Issue 04 (2026-05-30)", "Related: MLB idle transfer, Connected FreqPri",
                 "Initial cell selection is not priority balancing. Evaluate camping after reselection has run."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 Benefit / SN-4 Selection)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts / applies", "Parameter Detail", "User Experience Consideration"],
            "h": 34,
            "rows": [
                ["Benefit", "LTE FDD", "Idle transfer / dedicated priority", "At RRC release (idle MLB / SPID / PCC)",
                 "T320-bounded; low signalling vs connected HO", "Sets next access/PCC layer. Does not rebalance a long connected session."],
                ["Limitation", "LTE FDD", "Static common priority", "Always",
                 "Highest usable tier collects most idle UEs", "Four L2600 carriers need idle MLB + connected MLB, not C1>C2>C3>C4 priority."],
                ["Limitation", "LTE FDD", "Adaptive-proportion idle MLB", "Do not enable",
                 "Huawei not recommended without Huawei support [MLB] §5.5.2.1 p.80", "Fixed-proportion idle + user-number connected MLB = ping-pong."],
                ["Criterion S", "LTE FDD", "CELLSEL / CELLRESEL QRxLevMin, QQualMin, PMax", "Cell selection / reselection suitability",
                 "Srxlev>0 and Squal>0 if QQualMin configured", "Do not tighten L900 floors to push indoor users off coverage."],
                ["Higher-priority reselection", "LTE FDD", "EUTRANINTERNFREQ ThreshXhigh / ThreshXhighQ", "L900 → L2600/L1800/L2100",
                 "Target S > ThreshXhigh for EutranReselTime; camped >1 s", "Protects indoor users from premature weak high-band camping."],
                ["Lower-priority reselection", "LTE FDD", "CELLRESEL ThrshServLow + target ThreshXlow", "Capacity → L900 indoor fallback",
                 "Serving poor AND L900 usable; no higher-prio target", "Too-low ThrshServLow = sticky dying L2600. Align with connected A2/A5."],
                ["Equal-priority rank", "LTE FDD", "Qhyst / QoffsetFreq / CellQoffset", "Among L2600 C1-C4",
                 "Rn = Qmeas,n − Qoffset ; Rs = Qmeas,s + Qhyst", "Start QoffsetFreq=0. Do not chase hourly load with static offset."],
                ["Idle MLB trigger", "LTE FDD", "InterFreqIdleMlbUeNumThd + IdleUE", "Source idle load ≥ thd (+offset)",
                 "RRC release with IdleModeMobilityControlInfo", "Steers NEXT session only. See Intra-RAT MLB sheet."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5 Activation parameter / Switch)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 32,
            "rows": [
                ["Broadcast priority", "LTE FDD", "CELLRESEL / EUTRANINTERNFREQ",
                 "Set CellReselPriority and CfgInd=CFG", "L2600=7, L1800=6, L2100=5/6, L900=2; MeasPerformanceDemand=NORMAL",
                 "Common priority = coverage intent, not proportional load."],
                ["Search start", "LTE FDD", "CELLRESEL",
                 "SNonIntraSearchCfgInd=CFG, SNonIntraSearch=10 (Huawei example)", "SIntraSearch > SNonIntraSearch",
                 "Does not stop L900 UEs searching higher-priority L2600."],
                ["Idle MLB", "LTE FDD", "CELLALGOSWITCH",
                 "InterFreqIdleMlbSwitch-1 after target policy", "InterFreqMlbSwitch; capacity MlbTargetInd=ALLOWED; L900 without idle+connect MLB",
                 "Not with fixed-proportion idle balancing."],
                ["Hold camping", "LTE FDD", "EnhancedMlb / RRCCONNSTATETIMER",
                 "DediPrioManageOnLowLoadSw ON; T320ForLoadBalance=MIN30 or MIN60", "Idle MLB ON",
                 "SPID/PCC anchoring uses T320=180 min always. Do not start at 180 min for MLB."],
                ["L900 protect", "LTE FDD", "EUTRANINTERNFREQ MlbTargetInd",
                 "L900 = ALLOWED_WITHOUT_CONNECT_MLB (and WITHOUT_IDLE_MLB if enum exists)", "NoHoFlag remains PERMIT for coverage HO",
                 "Verify exact enum in eRAN21.1 parameter reference before MML."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["SIB5 complete", "LTE FDD", "EUTRANINTERNFREQ", "All 7 layers present, NORMAL", "Required", "Missing freq cannot be reselected or idle-MLB targeted."],
                ["Neighbor list", "LTE FDD", "EUTRANINTERFREQNCELL", "Cosited neighbors listed; ≤16 / freq in SIB", "Audit", "Truncation looks like a threshold problem."],
                ["Blacklist", "LTE FDD", "InterFreqBlkCell", "Do not load-blacklist a good layer", "65535 = idle+connected", "Blacklist is for invalid cells, not MLB."],
                ["SI quality", "LTE FDD", "SIB broadcast", "BER ≤ 1% [IM] §5.3.4", "Required", "Failed SIB5 decode looks like failed layer balance."],
                ["UE band mix", "LTE FDD", "Capability / SPID", "Know % without L2600 or L2100", "Segment KPI", "Idle priority cannot move incapable UEs."],
                ["Forbidden combo", "LTE FDD", "Adaptive / fixed-proportion idle", "OFF", "Huawei warning", "Ping-pong with user-number connected MLB."],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 28,
            "rows": [
                ["Cyclic steering", "LTE FDD", "Idle prio + connected MLB + release prio", "Do not change all same day", "Conflict", "Idle sets next access; connected owns the session."],
                ["ThreshXhigh too low", "LTE FDD", "L900 indoor UL / VoLTE", "Require robust high-band", "Risk", "Premature L2600 camp and setup fail."],
                ["ThrshServLow too low", "LTE FDD", "Connected A2/A5", "Align with Feature 2", "Risk", "Sticky dying L2600."],
                ["RSRQ reselection", "LTE FDD", "Load-varying RSRQ", "RSRP primary", "Oscillation", "Busy → RSRQ bad → leave → return."],
                ["Related: Intra-RAT MLB", "LTE FDD", "Idle transfer method", "Dedicated prio + T320", "Feature 3", "L900 not an idle MLB target."],
                ["Related: Connected Mode", "LTE FDD", "Dedicated idle prio discarded at RRC connect", "A1↔ThreshXhigh; A5↔ThrshServLow", "Feature 2", "Session safety after access."],
                ["Related: CA / ES", "LTE FDD", "PCC anchoring; sleeping carrier", "Do not idle-steer onto ES-off L2600", "Check PCell/SCell", "A busy SCell is not an unloaded layer."],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["Basic idle / SIB reselection", "LTE FDD", "CELLRESEL / SIB3/SIB5", "Basic LTE eNodeB", "Usually on", "No extra license."],
                ["Idle MLB", "LTE FDD", "InterFreqIdleMlbSwitch", "Intra-RAT Mobility Load Balancing (confirm name)", "Check MAE", "If missing: switch ON but dedicated-pri counters stay 0."],
                ["Blind idle MLB", "LTE FDD", "InterFreqBlindMlbSwitch", "Blind option if sold", "Keep OFF", "Only with verified full containment."],
                ["16 dedicated frequencies", "LTE FDD", "UE incMonEUTRA", "UE capability, not only eNB license", "Legacy ≤8 NORMAL", "Keep all primary layers NORMAL."],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence / connected order)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["1", "LTE FDD", "CELLSEL", "QRxLevMin / QQualMin", "Keep coverage-safe; QQualMin 0/absent until planned", "Not a daily MLB knob. [IM] §5.1.2"],
                ["2", "LTE FDD", "CELLRESEL", "CellReselPriority", "L2600=7; L1800=6; L2100=5/6; L900=2", "Equal L2600 C1-C4. [IM] §5.1.3.1"],
                ["3", "LTE FDD", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd + CellReselPriority", "CFG + same hierarchy", "No CFG = no reselection to that freq."],
                ["4", "LTE FDD", "EUTRANINTERNFREQ", "MeasPerformanceDemand", "NORMAL", "UNDELIVER forbidden on capacity. [IM] §5.3.2.3"],
                ["5", "LTE FDD", "CELLRESEL", "SIntraSearch / SNonIntraSearch", "CFG; SNonIntraSearch example 10; SIntra > SNonIntra", "[IM] §5.4.1.1; [MLB] §5.1.2.1"],
                ["6", "LTE FDD", "EUTRANINTERNFREQ", "ThreshXhigh / ThreshXhighQ", "Calibrate L900→high-band; ±1-2 dB per trial", "Indoor protect. [IM] Tables 5-1/5-2"],
                ["7", "LTE FDD", "CELLRESEL", "ThrshServLow / ThrshServLowQ", "Allow fallback before access collapse", "Align A2/A5. [IM] Tables 5-3/5-4"],
                ["8", "LTE FDD", "EUTRANINTERNFREQ", "ThreshXlow / ThreshXlowQ", "L900 must be actually usable", "Do not open weak L900 fallback."],
                ["9", "LTE FDD", "CELLRESEL / EUTRANINTERNFREQ / NCELL", "Qhyst / QoffsetFreq=0 / CellQoffset=0", "Offset only for stable RF asymmetry", "Not hourly load. [IM] §5.1.3.4"],
                ["10", "LTE FDD", "EUTRANINTERNFREQ", "EutranReselTime", "Keep / slightly longer if ping-pong", "Must persist for the timer."],
                ["11", "LTE FDD", "CELLRESEL", "SpeedDepReselCfgInd", "OFF for static layer imbalance", "Highway only. [IM] §5.1.3.6"],
                ["12", "LTE FDD", "RRCCONNSTATETIMER", "T320ForLoadBalance", "MIN30 or MIN60 to start", "[IM] §5.1.3.1"],
                ["13", "LTE FDD", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch", "ON capacity after target policy", "License required."],
                ["14", "LTE FDD", "EnhancedMlb", "DediPrioManageOnLowLoadSw", "ON with idle MLB", "[MLB] Table 5-10"],
                ["15", "LTE FDD", "EUTRANINTERNFREQ", "MlbTargetInd", "Capacity ALLOWED; L900 without connect+idle MLB", "[MLB] pp.28, 129. Coverage HO stays."],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 36,
            "rows": [
                ["0", "LTE FDD", "—",
                 "LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLALGOSWITCH: LocalCellId=<x>; LST RRCCONNSTATETIMER:;",
                 "Read-only", "Dump current before change. Keep LST output with the CR."],
                ["1", "LTE FDD", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<7|6|5|2>, SNonIntraSearchCfgInd=CFG, SNonIntraSearch=10, SIntraSearchCfgInd=CFG;",
                 "SIntraSearch > SNonIntraSearch", "L2600=7, L1800=6, L2100=5/6, L900=2. One layer per command."],
                ["2", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;",
                 "EARFCN exists; neighbors listed", "Repeat for every paired capacity + L900 frequency. UNDELIVER forbidden on capacity."],
                ["3", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<high-band>, ThreshXhigh=<calib>, ThreshXlow=<calib>, QoffsetFreq=0;",
                 "SN-4 rules", "L900 source: ThreshXhigh protects indoor. L2600 mutual: QoffsetFreq=0. ±1-2 dB per trial."],
                ["4", "LTE FDD", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<calib>;",
                 "Align with connected A2/A5", "Capacity cells: allow L900 fallback before access collapse."],
                ["5", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;",
                 "Add WITHOUT_IDLE_MLB if enum exists — verify on NE; coverage HO stays PERMIT",
                 "L900 is not a capacity MLB target. Verify exact enum in eRAN21.1 parameter reference."],
                ["6", "LTE FDD", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "License present; L900 not idle target; Feature 3 connected settings designed",
                 "Blind bit stays 0. Turn idle MLB on only after target indications are correct."],
                ["7", "LTE FDD", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;",
                 "Idle MLB ON", "Start conservative. Increase only if bounce-back is proven."],
                ["8", "LTE FDD", "Verify",
                 "LST all above; wait SI modification period; monitor L.RRCRel.load.DedicatedPri.LTE.High and RRC-setup by layer",
                 "Do not judge in the same 15-min", "SI delay [IM] §7.1.3. Placeholder LocalCellId/DlEarfcn must be replaced."],
            ],
        },
    ]


def connected_sections():
    return [
        {
            "name": "Section 1: Feature Introduction",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 36,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD", "Connected-mode HO engine",
                 "A1/A2/A3/A4/A5; coverage vs FreqPri vs MLB execution", "Coverage HO preempts load/FreqPri",
                 "CHART: RRC connected → A2/MLB/FreqPri meas → A3/A4/A5 report → Admit+HO → A1 stop. Ref [CM] Fig 4-1 pp.29-65. Full MLB algorithm is on the Intra-RAT MLB sheet."],
                ["Events", "LTE FDD", "Measurement events",
                 "A1 serving good; A2 serving poor; A3 relative better; A4 absolute target good; A5 serving poor AND target good",
                 "A4 = main capacity MLB event; A5 = L900 fallback; A1 = L900 escape gate",
                 "[CM] Table 4-8 pp.41-46; Tables 5-18/5-19; [MLB] pp.137-139."],
                ["SN-2 Major highlighted Point", "LTE FDD", "Safeguards",
                 "RSRP trigger; A4 better than coverage A2; A4 TTT≠5120 ms; object cap ≥6; no reverse MLB on FreqPri pair",
                 "Example −85/−87/−103 dBm are NOT design values",
                 "Equal-priority freqs may be randomly measured — 4×L2600 same static prio is not load balance. [CM] p.125, Table 4-9 p.48, pp.317-318."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 / SN-4)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts / applies", "Parameter Detail", "User Experience Consideration"],
            "h": 32,
            "rows": [
                ["Coverage path", "LTE FDD", "INTERFREQHOGROUP A2 then A3/A4/A5", "Serving crosses A2",
                 "CovBasedInterFreqHoMode = IMMEDIATE / SIGNAL / FREQPRIORITY", "Must preempt MLB. L2600 A2 early enough for indoor rescue."],
                ["Frequency priority", "LTE FDD", "A1 + A4 target; FreqPriIFHoWaitingTimer", "Serving good enough to leave L900 / steer high-band",
                 "LoadTriggerFreqPriHoSwitch: overlap, load info, neighbor not UE-number MLB triggered, no PCI conflict",
                 "Do not FreqPri toward L900 as capacity. [CM] Fig 11-1/11-2 pp.299-300."],
                ["MLB execution", "LTE FDD", "MlbInterFreqHoEventType A4 (or A5)", "Feature 3 source trigger already true",
                 "ONLY_STRONGEST_CELL recommended [MLB] Table 6-5 p.159", "A5 MLB needs Intra-LTE Load Balancing for Non-cosited Cells license."],
                ["Blind", "LTE FDD", "BlindHoPriority / InterFreqMlbBlindHo", "Immediate mobility + neighbor contains serving",
                 "Higher access-failure risk [CM] Fig 4-2", "OFF for daily throughput work."],
                ["Benefit / limit", "LTE FDD", "Connected mobility", "Always",
                 "A4 moves to 'good enough' target; unnecessary HO must admit ALL QCIs", "Not a load algorithm. Meas gaps steal TTI. PCC move ≠ SCC traffic."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5 Activation)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 32,
            "rows": [
                ["Coverage safety", "LTE FDD", "INTERFREQHOGROUP",
                 "Calibrate A1/A2/A5 for L2600→L900 rescue — do not paste example dBm", "Correct A2 family for A3 vs A4/A5 vs blind",
                 "Indoor rescue first. L900 A2 not hyper-aggressive."],
                ["Capacity MLB gate", "LTE FDD", "INTERFREQHOGROUP",
                 "InterFreqLoadBasedHoA4ThdRsrp calibrate; A4 Hyst 2-4; A4 TTT e.g. MS320", "A4 better than coverage A2; TTT ≠ 5120 ms",
                 "5120 ms disables FreqPri/CQI/service IFHO. [CM] Table 4-9 p.48"],
                ["Event type", "LTE FDD", "EUTRANINTERNFREQ",
                 "MlbInterFreqHoEventType=A4; IfMlbThdRsrpOffset=0 start", "Co-sited overlap", "Use A5 only if designed and licensed."],
                ["L900 CIO", "LTE FDD", "EUTRANINTERFREQNCELL",
                 "CellIndividualOffset=0 toward L900", "Do not use +CIO to help HO success into L900", "Positive CIO fills the 5 MHz layer."],
                ["FreqPri vs MLB", "LTE FDD", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch",
                 "ON wherever connected MLB is ON; FreqPriInHoProtectionTimer non-zero", "No reverse MLB target on a FreqPri pair [CM] p.303",
                 "Stops FreqPri fighting MLB and immediate bounce."],
                ["Meas eligibility", "LTE FDD", "EUTRANINTERNFREQ / CELLUEMEASCONTROLCFG",
                 "FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected; MaxNonIntraMeasObjNum ≥6", "7-layer site needs 6 inter-freq objects",
                 "Most 'A4 not working' cases are flags/NRT, not dBm."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["NRT", "LTE FDD", "EUTRANINTERFREQNCELL", "Both directions; no PCI conflict", "Required", "Looks like bad A4 if missing."],
                ["Object capacity", "LTE FDD", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "≥6 on 7-layer site", "Required", "[CM] Tables 4-3/4-4"],
                ["SMeasure", "LTE FDD", "HOMEASCOMM", "Must not silently suppress A4/FreqPri", "Verify RRC", "[CM] §4.1.5 p.55"],
                ["Gap", "LTE FDD", "AutoGapSwitch / GapPatternType", "Prefer A1/A2 gated meas", "VoLTE / old UE", "Gap steals TTI."],
                ["Admission / X2", "LTE FDD", "HoAdmitSwitch / X2RoHoAdmitSwitch", "Healthy", "Unnecessary HO = all QCI", "Prep fail is often admission, not RF."],
                ["CA PCC", "LTE FDD", "PCC anchoring", "Do not FreqPri against PCC policy", "Check CA book", "[CM] §11.3.2.3 ping-pong warning."],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["FreqPri A→B", "LTE FDD", "MLB B→A", "Remove reverse target", "Ping-pong [CM] p.303", "Documented Huawei warning."],
                ["A4 near A2", "LTE FDD", "Coverage meas", "Keep A4 better than A2", "Ping-pong", "[CM] Table 5-22 pp.147-148"],
                ["Large CIO", "LTE FDD", "RF problem", "Fix RF first", "Masks overshoot", "Do not tune CIO to hide RF."],
                ["Coverage meas already on", "LTE FDD", "MLB A4", "LOAD_COVERAGE_MEAS_DECOUPLE_SW if needed", "MLB meas blocked", "[MLB] p.139"],
                ["Related: Idle (F1)", "LTE FDD", "Access layer", "Align ThreshXhigh↔A1, ThrshServLow↔A5", "Session vs next access", ""],
                ["Related: MLB (F3)", "LTE FDD", "Who/when vs how", "ONLY_STRONGEST_CELL + A4", "This sheet executes HO", ""],
                ["Related: VoLTE", "LTE FDD", "QCI admit", "Never offload QCI1 if target cannot admit all QCIs", "Unnecessary HO rule", "[CM] Tables 4-16/4-17"],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["Coverage A1-A5", "LTE FDD", "Basic mobility", "Basic LTE mobility", "Usually on", "Confirm lean license package."],
                ["Frequency-priority HO", "LTE FDD", "A1-escape / high-band steer", "FreqPri / service-based package — verify", "Check MAE", "If missing, L900 A1-escape will not run."],
                ["MLB A4", "LTE FDD", "Co-sited load HO", "Intra-RAT Mobility Load Balancing", "See MLB sheet", ""],
                ["MLB A5", "LTE FDD", "Non-cosited load HO", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay on A4 if missing", "[MLB] pp.142-143"],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["1", "LTE FDD", "EUTRANINTERNFREQ", "DlEarfcn / MeasBandWidth / FREQ_MEAS_FLAG", "All 6 non-serving layers; flag selected", "[CM] Table 4-2"],
                ["2", "LTE FDD", "EUTRANINTERNFREQ", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Deselected for L900 + capacity", "Silent no-HO if set."],
                ["3", "LTE FDD", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "≥6", "[CM] Tables 4-3/4-4"],
                ["4", "LTE FDD", "HOMEASCOMM", "SMeasure", "Verify vs intended A4/FreqPri", "[CM] §4.1.5"],
                ["5", "LTE FDD", "CELLHOPARACFG", "EutranFilterCoeffRsrp / Rsrq", "Do not over-smooth L900 rescue", "[CM] Table 4-14"],
                ["6", "LTE FDD", "INTERFREQHOGROUP", "A1/A2 Hyst + TTT + A2 families", "Calibrate; not example dBm", "[CM] Tables 5-3, 5-10, 4-9"],
                ["7", "LTE FDD", "INTERFREQHOGROUP", "A3 offset / hyst / TTT", "Small; similar-coverage only", "[CM] Table 5-16"],
                ["8", "LTE FDD", "EUTRANINTERNFREQ / NCELL", "QoffsetFreqConn=0; CIO=0 toward L900", "No global +CIO to L900", "[CM] pp.46-48"],
                ["9", "LTE FDD", "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp + A4 Hyst/TTT", "A4 better than A2; TTT≠5120 ms", "[CM] Table 11-5; [MLB] p.137"],
                ["10", "LTE FDD", "INTERFREQHOGROUP", "Coverage A5 Thd1/Thd2", "Capacity→L900 fallback", "[CM] Tables 5-18/5-19"],
                ["11", "LTE FDD", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType", "A4", "[MLB] Table 6-3"],
                ["12", "LTE FDD", "INTRARATHOCOMM", "FreqPriInHoProtectionTimer / FreqPriIFHoWaitingTimer", "Protect non-zero", "[CM] pp.300, 308-309"],
                ["13", "LTE FDD", "FreqPri switches", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch", "ON if MLB ON", "[CM] Table 11-7"],
                ["14", "LTE FDD", "CELLOPHOCFG", "MLB_HO_FORBID_SW / FREQ_PRI_HO_FORBID_SW", "Optional high-speed cells only", "NLOS can misclassify. [MLB] pp.136-141"],
                ["15", "LTE FDD", "HOMEASCOMM", "HO fail punish timers", "Do not lower A4 after admission fail", "[CM] Tables 4-16/4-17"],
                ["16", "LTE FDD", "EUTRANINTERFREQNCELL", "BlindHoPriority", "OFF", "[CM] Table 5-22 p.149"],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 34,
            "rows": [
                ["0", "LTE FDD", "—",
                 "LST INTERFREQHOGROUP / EUTRANINTERNFREQ / EUTRANINTERFREQNCELL / CELLHOPARACFG / HOMEASCOMM;",
                 "Read-only", "Dump first. Attach to CR."],
                ["1", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>; ensure FREQ_MEAS_FLAG selected and HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for L900 and capacity;",
                 "NRT exists; object cap sufficient", "Audit before any threshold CR."],
                ["2", "LTE FDD", "CELLUEMEASCONTROLCFG",
                 "MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<≥6>, MaxEutranFddMeasFreqNum=<≥6>;",
                 "7-layer site", "Otherwise some L2600 never measured."],
                ["3", "LTE FDD", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; calibrate coverage A1/A2/A5 for L2600→L900 — do not paste example dBm;",
                 "Correct A2 family", "Protect indoor first."],
                ["4", "LTE FDD", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<calib>, InterFreqHoA4Hyst=2, InterFreqHoA4TimeToTrig=MS320;",
                 "A4 better than coverage A2; TTT ≠ 5120 ms", "MS320 is a starting TTT, not a mandate."],
                ["5", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<capa>, MlbInterFreqHoEventType=A4, IfMlbThdRsrpOffset=0;",
                 "Co-sited; A5 license not required", "A4 for capacity MLB."],
                ["6", "LTE FDD", "EUTRANINTERFREQNCELL",
                 "MOD EUTRANINTERFREQNCELL: LocalCellId=<x>, DlEarfcn=<L900>, CellId=<id>, CellIndividualOffset=0;",
                 "No +CIO into L900", "Avoid filling 5 MHz."],
                ["7", "LTE FDD", "FreqPri MO",
                 "Enable MlbBasedFreqPriHoSwitch-1 and LoadTriggerFreqPriHoSwitch-1; keep FreqPriInHoProtectionTimer non-zero;",
                 "Connected MLB will be ON", "Confirm exact MO in MAE."],
                ["8", "LTE FDD", "Verify",
                 "L.HHO.InterFreq.Coverage.*; L.HHO.InterFreq.FreqPri.*; L.RRC.ReEst.ReconfFail.Att; drop; VoLTE. PrepSucc=ExecAtt/PrepAtt; ExecSucc=ExecSucc/ExecAtt.",
                 "Busy hour; pair-level", "Counters [CM] Table 5-24, Tables 11-9/11-10."],
            ],
        },
    ]


def mlb_sections():
    return [
        {
            "name": "Section 1: Feature Introduction",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 36,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD", "Intra-RAT MLB",
                 "Equalisation vs offload; connected HO + idle dedicated prio; Load=N/C",
                 "Primary: UE_NUMBER_ONLY + SynchronizedUE + ActiveUe + SpectralEff",
                 "CHART: Eval N/C → Trigger → Admit target → Select UEs → A4 HO or idle release. Ref [MLB] Fig 3-1 p.15; §5.1.1 pp.24-27; Table 6-2 p.127."],
                ["SN-2 Major highlighted Point", "LTE FDD", "Rules that decide success",
                 "ONLY_STRONGEST_CELL; L900 not routine target; 5s eval + MlbMaxUeNum≥40 over-transfers; PRB MLB skips CA UEs",
                 "Smart thd learns 7 days / refreshes 7 days",
                 "Raw UE MLB on 15 vs 20 MHz can reduce TP [MLB] p.141. Do not daily overwrite learned neighbor thresholds."],
                ["Modes", "LTE FDD", "Algorithm choice",
                 "UE-number connected = primary; idle = secondary; PRB = supplement; blind = OFF",
                 "PRB gain falls if CA ≳ 60%",
                 "[MLB] Fig 6-2, §6.5 pp.207-215."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 / SN-4)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When Power Saving Starts", "Parameter Detail", "User Experience Consideration"],
            "h": 32,
            "rows": [
                ["Source trigger (UE-number)", "LTE FDD", "CELLMLB InterFreqMlbUeNumThd + MlbUeNumOffset",
                 "N ≥ Thd+Offset for whole MlbTrigJudgePeriod", "Stop when N < Thd",
                 "Use ACTIVE UEs, not raw UL-sync, on mixed BW. Header text 'When Power Saving Starts' is the attached template column name — here it means When MLB starts."],
                ["PRB trigger (supplement)", "LTE FDD", "InterFreqMlbThd + LoadOffset + min UE",
                 "PRB ≥ thd + offset", "Does not transfer CA UEs", "Do not use as the only algorithm on high-CA L2600."],
                ["Target admit", "LTE FDD", "MlbTargetInd / OverlapInd / NoHoFlag / HO SR / HW+transport",
                 "Low/Med HW+transport; not punished; pair HO SR ≥ NCellHoSuccRateThld", "L900 not admitted as routine target",
                 "Rejected target punished for CellPunishPrdNum × InterFreqLoadEvalPrd."],
                ["UE selection", "LTE FDD", "CELLMLBUESEL ARP/PRB/MCS/QCI/SNR + protect timers",
                 "UL-sync, not emergency, pass SPID/QCI", "ONLY_STRONGEST_CELL",
                 "Do not pick indoor-edge UEs only because they eat PRB. Else coverage HO bounce [MLB] Table 6-5 p.159."],
                ["Volume / freq pick", "LTE FDD", "MlbMaxUeNum / FreqSelectStrategy",
                 "min(needed delta, hyst, MlbMaxUeNum)", "LOADPRIORITY among 4×L2600",
                 "Never MlbMaxUeNum≥40 with 5 s eval [MLB] p.136."],
                ["HO event", "LTE FDD", "MlbInterFreqHoEventType",
                 "After UE selected and target admitted", "A4 on co-sited capacity", "A5 needs extra license."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5 Activation)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 30,
            "rows": [
                ["Master bits", "LTE FDD", "CELLALGOSWITCH",
                 "InterFreqMlbSwitch-1 and InterFreqIdleMlbSwitch-1; Blind-0", "License + NRT + target policy first",
                 "Table 6-2 p.127."],
                ["Mode", "LTE FDD", "CELLMLB",
                 "MlbTriggerMode=UE_NUMBER_ONLY; InterFreqUeTrsfType=SynchronizedUE", "PRB_ONLY skips CA UEs",
                 "Primary algorithm for DL TP fairness."],
                ["Load model", "LTE FDD", "ActiveUeBasedLoadEvalSw / SpectralEffBasedLoadEvalSw / LoadTransferEnhSw",
                 "All ON before chasing UE-number thd", "Unequal BW (L2100 15 MHz); 4×L2600 targets",
                 "Huawei recommends ActiveUe when BW differs and SE when SE differs >~30%. [MLB] Table 5-5."],
                ["Strategy", "LTE FDD", "CELLMLB",
                 "FreqSelectStrategy=LOADPRIORITY; MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL", "Reliable load exchange",
                 "PRIORITYBASED only if PCC architecture requires a preferred carrier."],
                ["CA", "LTE FDD", "CaUserLoadTransferSw",
                 "ON only after PCell/SCell/active-CC baseline", "Target CA capability ≥ serving on one path",
                 "Else CA UEs filtered [MLB] p.157."],
                ["L900 / event", "LTE FDD", "EUTRANINTERNFREQ",
                 "Capacity MlbTargetInd=ALLOWED + A4; L900 WITHOUT connect+idle MLB", "Coverage HO to L900 stays PERMIT",
                 "Verify L900 enum on the NE."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["License", "LTE FDD", "Intra-RAT MLB", "Present on every eNB in the cluster", "Required", "Switch ON but no transfer if missing."],
                ["Load exchange", "LTE FDD", "X2 / intra-eNB", "Working on Huawei co-sited sectors", "Required for equalisation", "Else only unsafe offload/blind."],
                ["NRT / overlap", "LTE FDD", "OverlapInd + PERMIT_HO", "Capacity pairs", "Required", "Target never admitted."],
                ["Pair HO health", "LTE FDD", "NCellHoSuccRateThld", "Fix coverage HO first if pair already bad", "Do not lower thd to force MLB", ""],
                ["HW / transport", "LTE FDD", "LowLoad / MediumLoad", "Alarms clear", "High/OverLoad illegal target", "[MLB] Fig 4-3"],
                ["A4 actually delivered", "LTE FDD", "Feature 2 flags / object cap / SMeasure", "See Connected Mode Section 4", "Else trigger with 0 meas success", ""],
                ["Smart thd counters", "LTE FDD", "MAE 15-min subscription", "If NCellTrigThldSmartOptAlgoSw ON", "[MLB] §§5.1.3.4, 6.5.3.4", ""],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["UE-number connected", "LTE FDD", "Fixed-proportion idle", "Keep fixed-proportion OFF", "Ping-pong [MLB] §5.4.2.2", ""],
                ["PRB_ONLY sole mode", "LTE FDD", "High CA", "Keep UE_NUMBER_ONLY primary", "CA UEs stuck; no fairness", "[MLB] §6.5"],
                ["PRB_USAGE vs PRB_VALUATION", "LTE FDD", "Same mode", "Exclusive", "Cannot both [MLB] pp.216, 245", ""],
                ["MLB vs FreqPri reverse pair", "LTE FDD", "Feature 2", "No reverse target", "Ping-pong [CM] p.303", ""],
                ["MLB toward L900", "LTE FDD", "Indoor VoLTE / 5 MHz", "Forbidden", "Coverage layer becomes capacity", "Hard guardrail"],
                ["Huawei learned neighbor thd", "LTE FDD", "BW change / upgrade / eval SW change", "Allow 7-day relearn after those events", "Stale or reset [MLB] pp.29-31", "Huawei native NCellTrigThldSmartOptAlgoSw. Collects 7 days, recalculates every 7 days."],
                ["Related: Idle / Connected / CA / ES", "LTE FDD", "F1 idle method; F2 HO engine; CA PCC; ES target exclusion",
                 "A4 + ONLY_STRONGEST + L900 block", "Do not dump 4G overflow to 2G/3G for this TP KPI", ""],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["Connected/idle equalisation", "LTE FDD", "InterFreqMlbSwitch / IdleMlbSwitch", "Intra-RAT Mobility Load Balancing", "Check MAE", "No Load HO counters if missing."],
                ["MLB event A5", "LTE FDD", "MlbInterFreqHoEventType=A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay on A4 if missing", "[MLB] pp.142-143"],
                ["Blind MLB", "LTE FDD", "InterFreqBlindMlbSwitch", "Blind option if sold", "Keep OFF", ""],
                ["CA user transfer", "LTE FDD", "CaUserLoadTransferSw", "CA + MLB CA-transfer — verify", "Else CA UEs filtered", "[MLB] p.157"],
                ["Huawei learned n-cell thd", "LTE FDD", "NCellTrigThldSmartOptAlgoSw", "Huawei option + MAE 15-min counters", "Optional; 7-day learn / 7-day refresh", "Do not overwrite learned pair thresholds every day."],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["1", "LTE FDD", "CELLALGOSWITCH", "InterFreqMlbSwitch", "ON capacity", "Table 6-2"],
                ["2", "LTE FDD", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch", "ON capacity after targets", ""],
                ["3", "LTE FDD", "CELLALGOSWITCH", "InterFreqBlindMlbSwitch", "OFF", ""],
                ["4", "LTE FDD", "CELLMLB", "MlbTriggerMode", "UE_NUMBER_ONLY", "Primary"],
                ["5", "LTE FDD", "CELLMLB", "InterFreqUeTrsfType", "SynchronizedUE", "Add IdleUE if idle ON"],
                ["6", "LTE FDD", "eval SW", "ActiveUeBasedLoadEvalSw", "ON", "Table 5-5"],
                ["7", "LTE FDD", "eval SW", "SpectralEffBasedLoadEvalSw", "ON", "Table 5-5"],
                ["8", "LTE FDD", "eval SW", "LoadTransferEnhSw / CaUserLoadTransferSw", "Enh ON; CA after audit", "pp.129-136"],
        ["9", "LTE FDD", "CELLMLB", "InterFreqMlbUeNumThd + MlbUeNumOffset", "Calibrate per layer AFTER Active+SE ON", "Do not copy a 20 MHz threshold onto L2100."],
                ["10", "LTE FDD", "CELLMLB", "MlbMaxUeNum / MlbTrigJudgePeriod / InterFreqLoadEvalPrd", "Conservative; not ≥40 with 5 s", "p.136"],
                ["11", "LTE FDD", "CELLMLB", "FreqSelectStrategy / MlbHoCellSelectStrategy", "LOADPRIORITY / ONLY_STRONGEST_CELL", "p.137, Table 6-5"],
                ["12", "LTE FDD", "EUTRANINTERNFREQ", "MlbTargetInd / MlbInterFreqHoEventType / IfMlbThdRsrpOffset", "Capacity ALLOWED+A4; L900 blocked; offset 0", "pp.28, 129"],
                ["13", "LTE FDD", "EUTRANINTERFREQNCELL", "NoHoFlag / NCellHoSuccRateThld", "PERMIT coverage; do not lower SR thd to force MLB", ""],
                ["14", "LTE FDD", "CELLMLBUESEL", "ARP/PRB/MCS/QCI pick + protect timers", "Protect QCI1; no edge-PRB hunting", "pp.130-135"],
                ["15", "LTE FDD", "CELLMLB / related", "NCellTrigThldSmartOptAlgoSw / learned pair thds", "Optional Huawei function", "7-day data collection, then refresh every 7 days."],
                ["16", "LTE FDD", "RRCCONNSTATETIMER", "T320ForLoadBalance", "MIN30/60", "Idle path"],
                ["17", "LTE FDD", "CELLPRBVALMLB / PRB mode", "PRB_ONLY / valuation", "OFF unless GBR problem", "Not a TP-fairness tool"],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 36,
            "rows": [
                ["0", "LTE FDD", "—",
                 "LST CELLALGOSWITCH / CELLMLB / EUTRANINTERNFREQ / CELLMLBUESEL; check license Intra-RAT MLB;",
                 "MAE license", "Baseline dump. Attach to CR."],
                ["1", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L1800|L2100|L2600>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;",
                 "Overlap valid; NoHo=PERMIT; FREQ_MEAS_FLAG; not forbid-meas", "Capacity may be targets. From L900 toward capacity also ALLOWED."],
                ["2", "LTE FDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;",
                 "Add WITHOUT_IDLE_MLB if enum exists — verify; coverage HO stays PERMIT", "L900 is coverage, not a capacity target."],
                ["3", "LTE FDD", "eval MO",
                 "Turn ON ActiveUeBasedLoadEvalSw, SpectralEffBasedLoadEvalSw, LoadTransferEnhSw;",
                 "Unequal BW (L2100 15 MHz)", "Do this BEFORE chasing UE-number thd. Confirm bit names in MAE."],
                ["4", "LTE FDD", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, FreqSelectStrategy=LOADPRIORITY, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL;",
                 "InterFreqMlbSwitch will be ON", "Core connected equalisation. Huawei ONLY_STRONGEST_CELL."],
                ["5", "LTE FDD", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<calib>, MlbUeNumOffset=<hyst>, MlbMaxUeNum=<conservative>, MlbTrigJudgePeriod=<stable>, InterFreqLoadEvalPrd=<not 5s-if-maxUE-large>;",
                 "Active+SE already ON", "Do not copy a 20 MHz threshold onto L2100."],
                ["6", "LTE FDD", "CA / CELLALGOSWITCH",
                 "Enable CaUserLoadTransferSw only after PCell/SCell/active-CC baseline;",
                 "Target CA capability path", "PRB MLB still will not move CA UEs."],
                ["7", "LTE FDD", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "Steps 1-5 done; Blind bit stays 0", "Master ON last after targets."],
                ["8", "LTE FDD", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;",
                 "Idle MLB ON", "Start conservative."],
                ["9", "LTE FDD", "CELLMLBUESEL",
                 "MOD CELLMLBUESEL: LocalCellId=<x>; protect QCI1/emergency; avoid aggressive edge PRB pick;",
                 "VoLTE cells", "Do not move weak indoor users to L2600."],
                ["10", "LTE FDD", "Verify",
                 "L.HHO.InterFreq.Load.* and UeNumLoad.*; HighLoad.Dur; Load.Meas(Succ); ActiveUser.DL; PCell/SCell; PRB.DL; Thrp.bits.DL / Thrp.Time.DL; idle DedicatedPri. SON logs: Inter-Frequency Handover Statistics + Idle Mode Release Statistics.",
                 "15-min subscribed; judge sector-cluster", "Tables 6-6, 6-21, 6-28 pp.164, 227-228, 248-249."],
            ],
        },
    ]


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    build_cover(wb)
    build_feature_sheet(wb, "Idle Mode Management",
                        "Idle Mode Management - Detailed Notes", idle_sections())
    build_feature_sheet(wb, "Connected Mode",
                        "Mobility Management in Connected Mode - Detailed Notes", connected_sections())
    build_feature_sheet(wb, "Intra-RAT MLB",
                        "Intra-RAT Mobility Load Balancing - Detailed Notes", mlb_sections())
    wb.properties.title = "Mobility Management Detailed Notes"
    wb.properties.creator = "Robi Axiata PLC RNO"
    wb.properties.subject = "eRAN21.1 Idle / Connected / Intra-RAT MLB — operator Excel format"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
