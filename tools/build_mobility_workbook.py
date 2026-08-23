#!/usr/bin/env python3
"""Build Mobility Management Excel in the user-required format.

Visual format from the attached Excel picture (must follow):
  black background
  Date  = white
  Open  = yellow
  High  = yellow
  Low   = yellow
  Close = light blue

Template structure from Sample_file1:
  Mobility Management
    Feature 1 Idle Mode Management          SN-1 .. SN-11
    Feature 2 Mobility Management in Connected Mode
    Feature 3 Intra-RAT Mobility Load Balancing

MML columns (exact):
  Parameter Sequence | MO | Activation Value | Conditional Parameter |
  Remarks | Parameter Description | More Notes
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.page import PageMargins
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.layout import Layout, ManualLayout
from copy import copy
import os

OUT = "/workspace/docs/4G_LTE_Mobility_Management/Mobility_Management_eRAN21.1_Workbook.xlsx"

# Picture palette
BLACK = "000000"
WHITE = "FFFFFF"
YELLOW = "F7D046"      # Open / High / Low
CYAN = "7FDBFF"        # Close
DARK = "0A0A0A"
ROW_ALT = "141414"
GRID = "333333"

# Column meaning on every data table (must keep this order / colors)
FMT = ["Date", "Open", "High", "Low", "Close"]
FMT_COLOR = [WHITE, YELLOW, YELLOW, YELLOW, CYAN]

thin = Border(
    left=Side(style="thin", color=GRID),
    right=Side(style="thin", color=GRID),
    top=Side(style="thin", color=GRID),
    bottom=Side(style="thin", color=GRID),
)
wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_c = Alignment(wrap_text=True, vertical="center", horizontal="center")
top = Alignment(wrap_text=True, vertical="top", horizontal="left")


def F(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(size=10, bold=False, color=WHITE, italic=False):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)


def widths(ws, wsizes):
    for i, w in enumerate(wsizes, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def paint_row(ws, r, cols, color=BLACK):
    for c in range(1, cols + 1):
        cell = ws.cell(r, c)
        if cell.fill.fgColor is None or cell.fill.fgColor.rgb in (None, "00000000"):
            cell.fill = F(color)
        if not cell.border.left.style:
            cell.border = thin


def put(ws, r, c, val, size=10, bold=False, color=WHITE, fill=BLACK, align=None, h=None, italic=False):
    x = ws.cell(r, c, val)
    x.font = font(size, bold, color, italic)
    x.fill = F(fill)
    x.alignment = align or wrap
    x.border = thin
    if h:
        ws.row_dimensions[r].height = h
    return x


def merge_put(ws, r, c1, c2, val, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    put(ws, r, c1, val, **kw)
    fill = kw.get("fill", BLACK)
    for c in range(c1 + 1, c2 + 1):
        ws.cell(r, c).fill = F(fill)
        ws.cell(r, c).border = thin
        ws.cell(r, c).font = font(kw.get("size", 10), kw.get("bold", False), kw.get("color", WHITE))


def setup(ws, title, cols=10):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5)
    ws.oddHeader.left.text = "Mobility Management"
    ws.oddFooter.left.text = title
    ws.oddFooter.right.text = "Date | Open | High | Low | Close"
    ws.sheet_properties.tabColor = "F7D046"
    # black sheet
    ws.sheet_format.defaultRowHeight = 18
    ws.sheet_view.view = "normal"
    # fill a large used background later via tables
    ws.freeze_panes = "A7"


def format_block(ws, start_row, cols=10):
    """Exact picture: vertical Date / Open / High / Low / Close in A, black field."""
    labels = FMT
    colors = FMT_COLOR
    meanings = [
        "Index / SN / Parameter Sequence  (white)",
        "MO / Item Name / start value     (yellow)",
        "Activation / proposed / max      (yellow)",
        "Conditional / constraint / min   (yellow)",
        "Result / remarks / gap / MML     (light blue)",
    ]
    put(ws, start_row, 1, "EXCEL FORMAT (must follow)", size=11, bold=True, color=WHITE, fill=BLACK, h=20)
    merge_put(ws, start_row, 2, cols, "Picture order and colors are mandatory. All tables in this file use this header system.",
              size=9, color=CYAN, fill=BLACK, h=20)
    r = start_row + 1
    for lab, col, mean in zip(labels, colors, meanings):
        put(ws, r, 1, lab, size=14, bold=True, color=col, fill=BLACK, align=wrap_c, h=22)
        merge_put(ws, r, 2, cols, mean, size=10, color=col, fill=BLACK, h=22)
        r += 1
    return r


def title_bar(ws, r, cols, text):
    merge_put(ws, r, 1, cols, text, size=18, bold=True, color=WHITE, fill=BLACK, h=28)
    return r + 1


def feature_bar(ws, r, cols, text):
    merge_put(ws, r, 1, cols, text, size=14, bold=True, color=YELLOW, fill=BLACK, h=24)
    return r + 1


def sn_headers(ws, r):
    """SN table header using Date/Open/High/Low/Close colors in order."""
    headers = [
        ("Date", "SN", WHITE),
        ("Open", "Item Name", YELLOW),
        ("High", "Details / Working content", YELLOW),
        ("Low", "Conditional / constraint", YELLOW),
        ("Close", "Chart / Ref (right side of SN)", CYAN),
    ]
    # two-row header: format name then template name
    for c, (fmt, name, col) in enumerate(headers, 1):
        put(ws, r, c, fmt, size=9, bold=True, color=col, fill=BLACK, align=wrap_c, h=16)
        put(ws, r + 1, c, name, size=10, bold=True, color=col, fill=BLACK, align=wrap_c, h=20)
    return r + 2


def sn_row(ws, r, sn, item, detail, cond, chart, h=36):
    vals = [sn, item, detail, cond, chart]
    cols_c = [WHITE, YELLOW, YELLOW, YELLOW, CYAN]
    for c, (v, col) in enumerate(zip(vals, cols_c), 1):
        put(ws, r, c, v, size=9, color=col, fill=DARK if r % 2 else BLACK, align=top, h=h)
    return r + 1


def table_headers(ws, r, names):
    """First five headers always colored Date/Open/High/Low/Close; extras cyan/white."""
    for c, name in enumerate(names, 1):
        if c == 1:
            col = WHITE
        elif c in (2, 3, 4):
            col = YELLOW
        else:
            col = CYAN
        put(ws, r, c, name, size=9, bold=True, color=col, fill=BLACK, align=wrap_c, h=22)
    return r + 1


def table_row(ws, r, values, h=28):
    for c, v in enumerate(values, 1):
        if c == 1:
            col = WHITE
        elif c in (2, 3, 4):
            col = YELLOW
        else:
            col = CYAN
        put(ws, r, c, v, size=9, color=col, fill=DARK if r % 2 else BLACK, align=top, h=h)
    return r + 1


def blank(ws, r, cols=10, h=8):
    merge_put(ws, r, 1, cols, "", fill=BLACK, h=h)
    return r + 1


# ---------------------------------------------------------------------------
# COVER / Mobility Management
# ---------------------------------------------------------------------------
def build_cover(wb):
    ws = wb.active
    ws.title = "Mobility Management"
    cols = 10
    widths(ws, [16, 28, 36, 28, 36, 22, 22, 22, 22, 22])
    setup(ws, "Mobility Management", cols)

    r = 1
    r = title_bar(ws, r, cols, "Mobility Management")
    merge_put(ws, r, 1, cols, "Robi Axiata PLC  |  Senior RNO  |  Huawei eRAN21.1  |  4G L1800 / L2100 / L2600 C1-C4 / L900",
              size=10, color=CYAN, fill=BLACK, h=18)
    r += 1
    r = format_block(ws, r, cols)
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Introduction")
    merge_put(
        ws, r, 1, cols,
        "1. State details of each feature in this workbook along with MAJOR notes (highlighted in yellow / Close in light blue). "
        "A sequence is a set of things / events / parameters / switches that come one after another in a connected order. "
        "Objective: reduce capacity-layer DL user-throughput gap toward 1–2 Mbps. Gap > 2 Mbps = investigate, not auto-CR. "
        "L900 5 MHz is indoor coverage only — coverage HO allowed, routine MLB target forbidden. "
        "Sources: Idle Mode Management Issue 04; Connected Mode Issue 08; Intra-RAT MLB Issue 10.",
        size=10, color=WHITE, fill=BLACK, h=56, align=top,
    )
    r += 1
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Workbook tree  (same as Mobility Management folder)")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    r = table_row(ws, r, ["0", "Mobility Management", "This sheet — format + introduction", "Read first", "Master"])
    r = table_row(ws, r, ["1", "Idle Mode Management", "Feature 1 — camping / reselection / idle MLB", "SIB3/SIB5 / T320", "Sheet: Idle Mode Management"])
    r = table_row(ws, r, ["2", "Mobility Management in Connected Mode", "Feature 2 — A1 A2 A3 A4 A5 / FreqPri", "Coverage preempts load", "Sheet: Connected Mode"])
    r = table_row(ws, r, ["3", "Intra-RAT Mobility Load Balancing", "Feature 3 — connected + idle MLB", "Active UE + SE load model", "Sheet: Intra-RAT MLB"])
    r = table_row(ws, r, ["4", "Daily KPI", "Date / Open / High / Low / Close tracker", "Gap = High − Low", "Sheet: Daily KPI  |  CR if Close > 2 Mbps + RCA pass"])
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Robi layer map")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "PRB", "Idle priority", "Connected", "MLB role", "Note"])
    layers = [
        ["L2600C1", "20 MHz", "capacity", "equal C1-C4", "A4 MLB", "100", "7 highest equal", "A4 / A3", "source+target", "17.4% nominal"],
        ["L2600C2", "20 MHz", "capacity", "equal C1-C4", "A4 MLB", "100", "7", "A4 / A3", "source+target", "17.4%"],
        ["L2600C3", "20 MHz", "capacity", "equal C1-C4", "A4 MLB", "100", "7", "A4 / A3", "source+target", "17.4%"],
        ["L2600C4", "20 MHz", "capacity", "equal C1-C4", "A4 MLB", "100", "7", "A4 / A3", "source+target", "17.4%"],
        ["L1800", "20 MHz", "capacity+anchor", "below L2600", "A4 MLB", "100", "6", "A4 + coverage", "source+target", "17.4%"],
        ["L2100", "15 MHz", "capacity", "SE-normalized", "A4 MLB", "75", "5 or 6", "A4", "source+target", "13.0% — not 20 MHz"],
        ["L900", "5 MHz", "indoor coverage", "NOT capacity pool", "A5 in / A1 out", "25", "2 lowest", "coverage only", "NO routine MLB target", "Protect VoLTE / indoor"],
    ]
    for row in layers:
        r = table_row(ws, r, row, h=22)
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "MAJOR highlighted notes")
    notes = [
        ["N1", "Load = N/C", "Use ActiveUe + SpectralEff", "Do not equalize raw UE on 15 vs 20 MHz", "[MLB] Table 5-5"],
        ["N2", "L900 protect", "Coverage HO in = YES", "Idle/connected MLB target = NO", "Yellow layer"],
        ["N3", "Event map", "Capacity MLB = A4", "L900 fallback = A5; L900 escape = A1/FreqPri", "[CM] + [MLB]"],
        ["N4", "ONLY_STRONGEST_CELL", "Huawei recommended", "Else coverage HO bounce", "[MLB] Table 6-5 p.159"],
        ["N5", "PRB MLB", "Supplement only", "Does not move CA UEs; weak if CA>~60%", "[MLB] §6.5"],
        ["N6", "Smart MLB thd", "Learn 7 days / refresh 7 days", "AI must not overwrite daily", "[MLB] pp.29–31"],
        ["N7", "CSV sample MML", "SymbolShutdownSwitch", "POWER SAVING — deleted, not used", "Template leftover"],
        ["N8", "MML IDs", "LocalCellId / DlEarfcn = placeholders", "Validate in MAE before live", "Close = result after CR"],
    ]
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    for row in notes:
        r = table_row(ws, r, row, h=20)

    # black fill leftover
    for i in range(1, r + 2):
        for c in range(1, cols + 1):
            if ws.cell(i, c).value is None:
                ws.cell(i, c).fill = F(BLACK)
                ws.cell(i, c).border = thin
    ws.sheet_properties.tabColor = "FFFFFF"
    return ws


# ---------------------------------------------------------------------------
# FEATURE BUILDER
# ---------------------------------------------------------------------------
def build_feature(wb, sheet_name, feature_title, intro, items, core_table, pre_table, impact_table,
                  relation_rows, license_rows, param_rows, mml_rows, tab_color):
    ws = wb.create_sheet(sheet_name)
    cols = 7
    widths(ws, [22, 28, 42, 36, 42, 36, 36])
    setup(ws, feature_title, cols)
    ws.sheet_properties.tabColor = tab_color

    r = 1
    r = title_bar(ws, r, cols, "Mobility Management")
    r = feature_bar(ws, r, cols, feature_title)
    r = format_block(ws, r, cols)
    r = blank(ws, r, cols)

    merge_put(ws, r, 1, cols, intro, size=10, color=WHITE, fill=BLACK, h=52, align=top)
    r += 1
    merge_put(
        ws, r, 1, cols,
        "** Must list down all activities in sequence (a sequence is a set of things / events / Parameters / Switch that come one after another in a connected order)",
        size=9, italic=True, color=YELLOW, fill=BLACK, h=20,
    )
    r += 1
    r = blank(ws, r, cols)

    # SN 1-9 as the required item table
    r = feature_bar(ws, r, cols, "SN list  (Item Name)   +   Chart from doc on Close / right side")
    r = sn_headers(ws, r)
    for it in items:
        r = sn_row(ws, r, *it, h=it[-1] if False else 40)
        # items are 5-tuples
    r = blank(ws, r, cols)

    # SN5 core activation table
    r = feature_bar(ws, r, cols, "SN-5  Core Parameter Setting   |   Parameter description  |  Proposed Value  |  Conditional Parameter")
    merge_put(ws, r, 1, cols, "Must in table format. Conditional Parameter = disable or activate to enable main activation. Extra columns allowed.",
              size=9, color=CYAN, fill=BLACK, h=18)
    r += 1
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Parameter Description", "More Notes"])
    # explain mapping row
    r = table_row(ws, r, ["SN / Seq", "MO", "Activation / Proposed", "Conditional Parameter", "Remarks", "Description", "Notes"], h=18)
    for row in core_table:
        r = table_row(ws, r, row, h=30)
    r = blank(ws, r, cols)

    # SN6
    r = feature_bar(ws, r, cols, "SN-6  Prerequisite functions   |   Parameter / Switch  |  Description  |  Proposed Value  |  Notes")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Parameter Description", "More Notes"])
    r = table_row(ws, r, ["SN", "MO / check", "Proposed / required", "Conditional", "If missing", "Description", "Notes"], h=18)
    for row in pre_table:
        r = table_row(ws, r, row, h=26)
    r = blank(ws, r, cols)

    # SN7
    r = feature_bar(ws, r, cols, "SN-7  Mutually impacted   |   Must in table format")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Parameter Description", "More Notes"])
    r = table_row(ws, r, ["SN", "This control", "Couples with", "Conditional", "What goes wrong / rule", "Description", "Notes"], h=18)
    for row in impact_table:
        r = table_row(ws, r, row, h=26)
    r = blank(ws, r, cols)

    # SN8-9 short tables
    r = feature_bar(ws, r, cols, "SN-8  Relation with Other Feature")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    for row in relation_rows:
        r = table_row(ws, r, row, h=24)
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "SN-9  License Requirements")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    for row in license_rows:
        r = table_row(ws, r, row, h=24)
    r = blank(ws, r, cols)

    # SN10
    r = feature_bar(ws, r, cols, "SN-10  All Parameter List in table format   —   sequence and connected order")
    merge_put(ws, r, 1, cols, "Parameter Name | description | Proposed Value | Conditional Parameter (relation with other feature) | Notes | more columns as needed",
              size=9, color=CYAN, fill=BLACK, h=18)
    r += 1
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Parameter Description", "More Notes"])
    r = table_row(ws, r, ["Seq", "MO", "Parameter / Proposed", "Conditional Parameter", "Notes / other feature", "Description", "Doc ref"], h=18)
    for row in param_rows:
        r = table_row(ws, r, row, h=24)
    r = blank(ws, r, cols)

    # SN11 exact MML columns
    r = feature_bar(ws, r, cols, "SN-11  Final MML Command for activations   (maintain sequence)")
    merge_put(ws, r, 1, cols, "Template columns: Parameter Sequence | MO | Activation Value | Conditional Parameter | Remarks | Parameter Description | More Notes",
              size=9, color=CYAN, fill=BLACK, h=18)
    r += 1
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Parameter Description", "More Notes"])
    r = table_row(ws, r, ["Parameter Sequence", "MO", "Activation Value", "Conditional Parameter", "Remarks", "Parameter Description", "More Notes"], h=20)
    for row in mml_rows:
        r = table_row(ws, r, row, h=34)

    for i in range(1, r + 2):
        for c in range(1, cols + 1):
            if ws.cell(i, c).value is None:
                ws.cell(i, c).fill = F(BLACK)
                ws.cell(i, c).border = thin
    return ws


def idle_data():
    intro = (
        "Feature 1: Idle Mode Management. An RRC_IDLE UE does PLMN select → Criterion S cell select → camp → read SIB3/SIB5 → "
        "measure / reselect by priority + thresholds + timers → next RRC starts on the camped cell. "
        "Chart: PLMN → Select(S) → Camp+SIB → Measure/Reselect → RRC on camped cell. "
        "Ref: [IM] Fig 4-1 p.6; Fig 5-1 pp.13–14; §5.1.3."
    )
    items = [
        ["1", "Working Principal",
         "Suitability Srxlev>0 (and Squal>0 if QQualMin set). Higher-priority freq always measured. Equal/lower measured after SNonIntraSearch. Higher-prio reselect if target > ThreshXhigh. Lower-prio if serving < ThrshServLow AND target > ThreshXlow. Equal-prio rank Rn=Qmeas-n−Qoffset, Rs=Qmeas-s+Qhyst. Dedicated priority in RRC release replaces SIB until T320 / next RRC.",
         "QRxLevMin is a floor, not an MLB knob",
         "CHART: PLMN → S → Camp → SIB3/5 → Resel → RRC  |  [IM] §5.1"],
        ["2", "Major highlighted Point",
         "Priority is frequency-level. No SIB5 priority = no reselection to that freq (max 16). Initial select is NOT priority balancing. L900 UE always searches higher-priority L2600. SI change applies next modification period (or 3 h). UNDELIVER removes freq from SIB5 and idle MLB. Huawei example SNonIntraSearch=10. Offset table vs Criterion-S formula conflict — do not use that table sentence.",
         "Do not change QRxLevMin daily",
         "CHART: L2600=7 equal; L1800=6; L2100=5/6; L900=2  |  [IM] §5.1.3.1"],
        ["3", "Benefit and Limitations",
         "Benefit: low signalling; sets next PCC/access layer; T320-bounded dedicated prio. Limitation: not real-time; static priority packs highest usable tier; SI delay; legacy UE sees fewer freqs; load-based redirect is not RF-measured; adaptive-proportion idle MLB not recommended.",
         "Must pair with connected MLB",
         "[MLB] §4.3, §5.5.2.1 p.80"],
        ["4", "Selection criteria (UE selection / Trigger Condition / Etc.)",
         "Idle has no connected-style UE pick. UE applies Criterion S + reselection. eNB writes dedicated priorities only at RRC release (idle MLB / SPID / PCC). Idle MLB trigger: idle-user load ≥ InterFreqIdleMlbUeNumThd.",
         "L900→capacity only if ThreshXhigh pass",
         "CHART: S pass → priority test → timer persist  |  [IM] Tables 5-1..5-4"],
        ["5", "Activation parameter / Switch  (parameter) — Notes",
         "Core switches in table below. Main: CellReselPriority, CellReselPriorityCfgInd=CFG, MeasPerformanceDemand=NORMAL, SNonIntraSearch CFG, InterFreqIdleMlbSwitch, DediPrioManageOnLowLoadSw, T320ForLoadBalance.",
         "Idle MLB ON only after MlbTargetInd set",
         "See SN-5 table  |  [IM] §5.3.2.3"],
        ["6", "Pre-requiste functions",
         "Complete SIB5 for all 7 layers; neighbor list not truncated (≤16 listed/freq); no load-blacklist; SI BER≤1%; UE band inventory; Overlap/NoHo ready for MLB; adaptive-proportion OFF.",
         "Missing SIB5 looks like bad threshold",
         "[IM] §5.1.3.4, §5.3.4"],
        ["7", "Mutually impacted",
         "High L2600 prio + connected MLB + release back = cyclic steer. Fixed-proportion idle + user-number connected = ping-pong. Aggressive ThreshXhigh pulls indoor to weak L2600. RSRQ oscillates with load.",
         "One family per trial",
         "[MLB] §5.4.2.2"],
        ["8", "Relation with Other Feature",
         "Idle MLB is Feature 3 idle method. Dedicated prio discarded at RRC connect (Feature 2 takes over). CA PCC anchoring changes who can be idle-steered. ES sleeping carrier must not stay highest idle prio.",
         "Align ThreshXhigh with A1; ThrshServLow with A5",
         "F1 ↔ F2 ↔ F3 ↔ CA"],
        ["9", "License Requirements",
         "Basic reselection = basic LTE. Idle MLB = Intra-RAT Mobility Load Balancing (confirm on NE). Blind idle path extra. Enhanced 16-freq dedicated list is UE capability (incMonEUTRA), not only license.",
         "Check MAE license before InterFreqIdleMlbSwitch",
         "Verify exact license ID on site"],
    ]
    core = [
        ["5.1", "CELLRESEL", "CellReselPriority  L2600=7 L1800=6 L2100=5/6 L900=2", "SIB3 broadcast", "Equal L2600 C1–C4", "Serving common priority", "[IM] §5.1.3.1"],
        ["5.2", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd=CFG + matching priority", "Else no reselection to that freq", "All 6+1 layers", "SIB5 priority present", "[IM] §5.1.3.1"],
        ["5.3", "EUTRANINTERNFREQ", "MeasPerformanceDemand=NORMAL", "UNDELIVER blocks idle MLB target", "Never UNDELIVER capacity", "Visibility / meas performance", "[IM] §5.3.2.3"],
        ["5.4", "CELLRESEL", "SNonIntraSearchCfgInd=CFG, SNonIntraSearch=10 (Huawei example)", "SIntraSearch > SNonIntraSearch", "Does not stop higher-prio search from L900", "Equal/lower search start", "[IM] §5.4.1.1; [MLB] §5.1.2.1"],
        ["5.5", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch-1", "InterFreqMlbSwitch + MlbTargetInd allow idle on capacity only", "Not with fixed-proportion", "Idle MLB master bit", "[IM] §5.3.2.3"],
        ["5.6", "EnhancedMlb", "DediPrioManageOnLowLoadSw ON", "T320 set", "Holds steered camping", "Low-load dedicated prio hold", "[MLB] Table 5-10"],
        ["5.7", "RRCCONNSTATETIMER", "T320ForLoadBalance=MIN30 or MIN60", "Idle MLB ON", "SPID/PCC always 180 min", "Dedicated-priority lifetime", "[IM] §5.1.3.1"],
    ]
    pre = [
        ["6.1", "EUTRANINTERNFREQ set", "All principal freqs present NORMAL", "SIB5", "UE cannot reselect missing freq", "Complete inter-freq list", "Max 16 non-serving E-UTRAN"],
        ["6.2", "EUTRANINTERFREQNCELL", "Cosited neighbors listed", "≤16 listed / freq", "Looks like threshold problem", "NRT / SIB truncation", "[IM] §5.1.3.4"],
        ["6.3", "InterFreqBlkCell", "Do not load-blacklist a good layer", "ApplicationScope 65535 = idle+connected", "Hidden valid target", "Blacklist ≠ MLB", "[IM] §5.1.3.2"],
        ["6.4", "SI quality", "SIB BER ≤ 1%", "SIB1/3/5", "Failed SIB5 decode", "Broadcast health", "[IM] §5.3.4"],
        ["6.5", "UE band class", "Know % without L2600/L2100", "Capability / SPID", "Idle cannot move incapable UE", "Segment KPI", "Device mix"],
        ["6.6", "Adaptive-proportion idle", "OFF", "Huawei not recommended", "Ping-pong with user-number MLB", "Do not enable", "[MLB] §5.5.2.1"],
    ]
    impact = [
        ["7.1", "L2600 highest prio", "Connected MLB + release dedicated prio", "Do not change all same day", "Cyclic steering", "Idle vs connected", "[MLB] §5.4.2.2"],
        ["7.2", "Fixed-proportion idle", "UE-number connected MLB", "Keep fixed-proportion OFF", "Ping-pong", "Mutually avoid", "[MLB]"],
        ["7.3", "ThreshXhigh too low", "L900 indoor UL / VoLTE", "Require robust high-band", "Premature L2600 camp", "Coverage", "Yellow"],
        ["7.4", "ThrshServLow too low", "Connected A2/A5", "Align with Feature 2", "Sticky dying L2600", "Fallback timing", "[CM]"],
        ["7.5", "Negative QoffsetFreq", "CA PCC / FreqPri", "Offset only for stable RF", "One L2600 hoards idle UEs", "Prefer idle MLB", "[IM] §5.1.3.4"],
    ]
    relation = [
        ["8.1", "Intra-RAT MLB (F3)", "Idle transfer method = dedicated prio + T320", "L900 not idle MLB target", "See Feature 3"],
        ["8.2", "Connected Mode (F2)", "Dedicated idle prio discarded at RRC connect", "Align A1 with ThreshXhigh; A5 with ThrshServLow", "Session vs next access"],
        ["8.3", "Carrier Aggregation", "PCC anchoring + CaUserLoadTransferSw", "Do not call a busy SCell 'unloaded'", "CA book"],
        ["8.4", "Energy saving", "Sleeping L2600 must not stay highest idle target", "Coordinate ES", "SIB5 vs ES state"],
    ]
    lic = [
        ["9.1", "Basic idle / SIB reselection", "Basic LTE eNodeB", "Always on", "No extra license"],
        ["9.2", "Idle MLB / dedicated load prio", "Intra-RAT Mobility Load Balancing (confirm name)", "Switch ON but counters 0", "Check license"],
        ["9.3", "Blind idle MLB", "Blind MLB option if sold", "Keep OFF", "Full containment only"],
        ["9.4", "16 dedicated freqs", "UE incMonEUTRA", "Legacy ≤8 NORMAL", "Keep primary layers NORMAL"],
    ]
    params = [
        ["1", "CELLSEL", "QRxLevMin = keep coverage-safe", "UePowerMax / Pcompensation", "NOT a daily MLB knob", "Selection floor", "[IM] §5.1.2"],
        ["2", "CELLSEL", "QQualMin = keep 0/absent until planned", "RSRQ moves with load", "Optional", "RSRQ selection", "[IM] §5.1.2"],
        ["3", "CELLRESEL", "CellReselPriority per layer map", "SIB3", "Equal L2600", "Serving prio", "[IM] §5.1.3.1"],
        ["4", "EUTRANINTERNFREQ", "CellReselPriority + CfgInd=CFG", "SIB5", "Mirror hierarchy", "Target prio", "[IM]"],
        ["5", "EUTRANINTERNFREQ", "QRxLevMin/QqualMin/Pmax target", "UL/PRACH KPI", "Do not block L900", "Target suitability", "[IM] §5.1.3.4"],
        ["6", "CELLRESEL", "SIntraSearch CFG > SNonIntraSearch", "Battery vs meas", "Huawei CFG", "Intra search", "[IM] §5.4.1.1"],
        ["7", "CELLRESEL", "SNonIntraSearch=10 example", "Does not stop higher-prio meas", "Ping-pong note", "Inter search", "[MLB] §5.1.2.1"],
        ["8", "EUTRANINTERNFREQ", "ThreshXhigh calibrate L900→HB", "EutranReselTime", "±1–2 dB / trial", "Higher-prio qualify", "[IM] T5-1/5-2"],
        ["9", "CELLRESEL", "ThrshServLow calibrate HB→L900", "A2/A5 align", "Not too low", "Serving-low", "[IM] T5-3/5-4"],
        ["10", "EUTRANINTERNFREQ", "ThreshXlow L900 usable", "ThrshServLow", "L900 protect", "Lower-prio target", "[IM]"],
        ["11", "CELLRESEL", "Qhyst keep default unless ping-pong", "QoffsetFreq", "Stickiness", "Equal rank", "[IM] §5.1.3.4"],
        ["12", "EUTRANINTERNFREQ", "QoffsetFreq=0 among L2600", "CellQoffset", "No hourly chase", "Freq offset", "[IM]"],
        ["13", "EUTRANINTERFREQNCELL", "CellQoffset=0 unless overshoot", "SIB max 16", "Not daily load tool", "Cell offset", "[IM]"],
        ["14", "EUTRANINTERNFREQ", "EutranReselTime keep / slightly longer if ping-pong", "ThreshX*", "Persist", "Inter timer", "[IM]"],
        ["15", "EUTRANINTERNFREQ", "MeasPerformanceDemand=NORMAL", "Idle MLB eligibility", "No UNDELIVER capacity", "Meas demand", "[IM]"],
        ["16", "CELLRESEL", "SpeedDepResel OFF for static imbalance", "Highway only", "Amber", "Speed scale", "[IM] §5.1.3.6"],
        ["17", "RRCCONNSTATETIMER", "T320ForLoadBalance MIN30/60", "Idle MLB", "Not 180 on day-1", "T320", "[IM]"],
        ["18", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch-1", "License + target ind", "Capacity only", "Idle MLB", "[MLB]"],
        ["19", "EnhancedMlb", "DediPrioManageOnLowLoadSw ON", "T320", "Hold camping", "Dedi prio", "[MLB] T5-10"],
        ["20", "EUTRANINTERNFREQ", "MlbTargetInd capacity ALLOWED; L900 WITHOUT idle+connect MLB", "NoHoFlag still PERMIT coverage", "Verify enum", "Target ind", "[MLB] pp.28,129"],
    ]
    mml = [
        ["0", "—", "LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLALGOSWITCH: LocalCellId=<x>; LST RRCCONNSTATETIMER:;",
         "Read-only", "Dump current before change", "Baseline LST", "Keep output with CR"],
        ["1", "CELLRESEL", "MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<7|6|5|2>, SNonIntraSearchCfgInd=CFG, SNonIntraSearch=10, SIntraSearchCfgInd=CFG;",
         "SIntraSearch > SNonIntraSearch", "L2600=7 L1800=6 L2100=5/6 L900=2", "Serving priority + search", "One layer per command"],
        ["2", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;",
         "EARFCN exists; neighbors listed", "Repeat all paired freqs", "SIB5 prio + NORMAL", "UNDELIVER forbidden on capacity"],
        ["3", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<hb>, ThreshXhigh=<calib>, ThreshXlow=<calib>, QoffsetFreq=0;",
         "SN-4 rules", "L900 source ThreshXhigh protects indoor", "Qualify + equal offset", "±1–2 dB per trial"],
        ["4", "CELLRESEL", "MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<calib>;",
         "Align Feature 2 A2/A5", "Allow L900 fallback before access collapse", "Serving-low", "L900 protect"],
        ["5", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;",
         "Add WITHOUT_IDLE_MLB if enum exists — verify", "L900 not capacity target; coverage HO stays PERMIT", "L900 MlbTargetInd", "Verify eRAN21.1 enum"],
        ["6", "CELLALGOSWITCH", "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
         "License; Feature 3 designed; L900 not idle target", "Master idle+connected bits; Blind stays 0", "Idle MLB ON", "Last step after targets"],
        ["7", "RRCCONNSTATETIMER", "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;",
         "Idle MLB ON", "Start conservative", "T320", "Increase only if bounce-back proven"],
        ["8", "Verify", "LST all above; wait SI modification; monitor L.RRCRel.load.DedicatedPri.LTE.High and RRC-setup by layer",
         "Not same 15-min", "SI delay [IM] §7.1.3", "Close-loop", "Close = post-check result"],
    ]
    return intro, items, core, pre, impact, relation, lic, params, mml


def connected_data():
    intro = (
        "Feature 2: Mobility Management in Connected Mode. Flow: initiate → meas or blind → deliver meas → report → pick target → HO → punish/retry. "
        "Necessary coverage HO preempts unnecessary load/FreqPri HO. This book is the measurement/HO engine; full MLB algorithm is Feature 3. "
        "Chart: RRC connected → A2/MLB/FreqPri meas → A3/A4/A5 report → Admit+HO → A1 stop. "
        "Ref: [CM] Fig 4-1 pp.29–65; Table 3-1 pp.21–23."
    )
    items = [
        ["1", "Working Principal",
         "A1 serving good (stop coverage meas; L900→capacity gate). A2 serving poor (start inter-freq meas). A3 neighbor relatively better. A4 neighbor absolutely good enough (MLB/FreqPri). A5 serving poor AND neighbor good (L900 fallback). Blind only if neighbor contains serving.",
         "Coverage > load > static FreqPri",
         "CHART: A1/A2 gate ↔ A3/A4/A5 decide  |  [CM] Table 4-8"],
        ["2", "Major highlighted Point",
         "RSRP recommended (RSRQ follows load). Equal-priority freqs may be randomly measured — 4×L2600 same static prio ≠ load balance. A4 must be better than coverage A2 else ping-pong. A4 TTT=5120 ms disables FreqPri. Example −85/−87/−103 dBm are NOT design values. 7-layer site needs ≥6 meas objects. MlbBasedFreqPriHoSwitch lets MLB own heavy load. No reverse MLB target on a FreqPri pair.",
         "Audit flags/NRT before changing dBm",
         "[CM] pp.125, 147–148, 301–303, 317–318"],
        ["3", "Benefit and Limitations",
         "Benefit: indoor rescue; A4 offload without beating serving; A5 L900 semantics; A1 escape from L900. Limitation: not a load algorithm; meas gaps steal TTI; unnecessary HO must admit ALL QCIs; CA PCC ≠ SCC; high-speed forbid ~30 km/h can misclassify NLOS.",
         "Need Feature 3 for TP fairness",
         "[CM] Tables 4-16/4-17"],
        ["4", "Selection criteria (UE selection / Trigger Condition / Etc.)",
         "Coverage: A2 then A3/A4/A5. FreqPri: A1/good serving + A4 target; LoadTriggerFreqPri needs overlap, load info, neighbor not UE-number MLB triggered, no PCI conflict. MLB UE pick is Feature 3. Blind: BlindHoPriority + containment.",
         "Do not FreqPri toward L900 as capacity",
         "[CM] §5.3.1; §11"],
        ["5", "Activation parameter / Switch  (parameter) — Notes",
         "Core: A1/A2/A5 coverage set, A4 MLB thd, MlbInterFreqHoEventType=A4, MlbBasedFreqPriHoSwitch, meas flags, object cap, A4 TTT ≠ 5120 ms, CIO to L900 = 0.",
         "A5 MLB needs extra license",
         "See SN-5 table"],
        ["6", "Pre-requiste functions",
         "Symmetric NRT; no PCI conflict; MaxNonIntraMeasObjNum ≥ needed; SMeasure not suppressing; gap acceptable; X2/admit healthy; CA PCC policy known; punish timers understood.",
         "Prep fail often admission not RF",
         "[CM] Tables 4-3/4-4"],
        ["7", "Mutually impacted",
         "FreqPri A→B vs MLB B→A ping-pong. Incoming unnecessary HO vs immediate FreqPri (use protect timer). A4 too close to A2. Large CIO hides RF. Virtual-grid + FreqPri ping-pong. LOAD_COVERAGE_MEAS_DECOUPLE_SW if coverage meas blocks MLB.",
         "One family per trial",
         "[CM] p.303, p.312"],
        ["8", "Relation with Other Feature",
         "Idle F1 sets access layer. MLB F3 decides who/when; this feature executes HO. CA/NSA PCC can ping-pong with FreqPri. VoLTE QCI-specific hyst/TTT. ES must drop sleeping targets.",
         "Align with F1 thresholds and F3 A4",
         "F1/F3/CA/VoLTE/ES"],
        ["9", "License Requirements",
         "Basic A1–A5 coverage usually included. FreqPri package — verify. MLB A4 = Intra-RAT MLB. MLB A5 = Intra-LTE Load Balancing for Non-cosited Cells [MLB] pp.142–143. Blind option separate.",
         "A5 MLB only if event type A5",
         "Confirm MAE"],
    ]
    core = [
        ["5.1", "INTERFREQHOGROUP", "A1/A2 Hyst+TTT calibrate (NOT example −85/−87)", "Correct A2 family for A3 vs A4/A5 vs blind", "L2600 A2 early; L900 A2 not hyper-aggressive", "Coverage start/stop", "[CM] T5-3, T5-10"],
        ["5.2", "INTERFREQHOGROUP", "A3 offset/hyst/TTT small — similar coverage only", "QoffsetFreqConn + CIO", "Not for L900 capacity offload", "Relative HO", "[CM] T5-16"],
        ["5.3", "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp calibrate; A4 better than A2; TTT≠5120ms", "IfMlbThdRsrpOffset / FreqPri offset", "Main capacity MLB event", "A4 absolute", "[CM] T11-5; [MLB] p.137"],
        ["5.4", "INTERFREQHOGROUP", "A5 Thd1/Thd2 for capacity→L900", "L900 suitable", "Strongest L900 protect", "A5 dual", "[CM] T5-18/5-19"],
        ["5.5", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType=A4", "A5 needs non-cosited license", "Co-sited capacity", "MLB event", "[MLB] T6-3"],
        ["5.6", "FreqPri SW", "MlbBasedFreqPriHoSwitch-1 + LoadTriggerFreqPriHoSwitch-1", "MLB will be ON", "FreqPri yields in heavy load", "Coordination", "[CM] T11-7"],
        ["5.7", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for L900+capacity", "Object cap ≥6", "Silent no-HO if wrong", "Meas flags", "[CM] §4.1.4.1.2"],
        ["5.8", "EUTRANINTERFREQNCELL", "CellIndividualOffset=0 toward L900", "Do not +CIO to help HO success into L900", "Fills 5 MHz", "CIO", "[CM] pp.46–48"],
    ]
    pre = [
        ["6.1", "NRT both directions", "All capacity pairs + L900 coverage", "No PCI conflict", "Looks like bad A4", "Neighbor", "[CM]"],
        ["6.2", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum ≥6", "7-layer site", "Some L2600 never measured", "Object cap", "[CM] T4-3/4-4"],
        ["6.3", "HOMEASCOMM SMeasure", "Verify vs intended A4/FreqPri", "A1/A2", "A4 never delivered", "Silent suppress", "[CM] §4.1.5"],
        ["6.4", "Gap pattern", "Prefer A1/A2 gated meas", "VoLTE / old UE", "TP drop", "Gap", "[CM] Fig 4-10"],
        ["6.5", "Admit / X2", "Healthy", "All-QCI for unnecessary HO", "PrepAtt high ExecAtt low", "Admission", "[CM] T4-16"],
        ["6.6", "CA PCC policy", "Do not FreqPri against PCC anchor", "CA book", "Ping-pong", "CA", "[CM] §11.3.2.3"],
    ]
    impact = [
        ["7.1", "FreqPri target A→B", "MLB target B→A", "Remove reverse", "Documented ping-pong", "Pair design", "[CM] p.303"],
        ["7.2", "FreqPri", "Connected MLB", "MlbBasedFreqPriHoSwitch ON", "Both move same UE", "Coord", "[CM] T11-7"],
        ["7.3", "A4 near A2", "Coverage meas", "Keep A4 better", "Ping-pong", "Threshold gap", "[CM] T5-22"],
        ["7.4", "Large CIO", "RF problem", "Fix RF first", "Masks overshoot", "CIO", "[CM]"],
        ["7.5", "Coverage meas already on", "MLB A4", "LOAD_COVERAGE_MEAS_DECOUPLE_SW", "MLB meas blocked", "Decouple", "[MLB] p.139"],
    ]
    relation = [
        ["8.1", "Idle (F1)", "Idle discarded at connect", "Align ThreshXhigh↔A1, ThrshServLow↔A5", "Next access vs session"],
        ["8.2", "MLB (F3)", "This feature executes A4/A5 HO", "ONLY_STRONGEST_CELL + A4", "Who/when vs how"],
        ["8.3", "CA", "PCC move ≠ SCC traffic", "Judge aggregate TP", "CA book"],
        ["8.4", "VoLTE QCI", "Unnecessary HO needs all QCI admit", "Never offload QCI1 if target cannot admit", "[CM]"],
    ]
    lic = [
        ["9.1", "Coverage A1–A5", "Basic LTE mobility", "Usually on", "Confirm lean package"],
        ["9.2", "Frequency-priority HO", "FreqPri / service-based package — verify", "No L900 A1-escape", "Check license"],
        ["9.3", "MLB A4", "Intra-RAT MLB", "See F3", "Co-sited"],
        ["9.4", "MLB A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay on A4 if missing", "[MLB] pp.142–143"],
    ]
    params = [
        ["1", "EUTRANINTERNFREQ", "DlEarfcn/MeasBW all 6 non-serving", "NRT exists", "Object list", "Meas object", "[CM] T4-2"],
        ["2", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG selected", "Forbid-meas deselected", "Audit first", "Meas flag", "[CM]"],
        ["3", "CELLUEMEASCONTROLCFG", "Max objects ≥6", "4×L2600 equal prio random if over cap", "Missing carrier", "Object cap", "[CM] T4-3"],
        ["4", "HOMEASCOMM", "SMeasure verify", "Can suppress A4", "Read RRC", "SMeasure", "[CM] p.55"],
        ["5", "CELLHOPARACFG", "FilterCoeff RSRP keep; do not over-smooth L900 rescue", "TTT", "Indoor delay", "L3 filter", "[CM] T4-14"],
        ["6", "INTERFREQHOGROUP", "A1/A2 hyst/TTT calibrate", "A2 family", "Not example dBm", "A1A2", "[CM] T4-9"],
        ["7", "INTERFREQHOGROUP", "Coverage A2 families", "A3 vs A4/A5 vs blind", "Wrong family = wrong HO", "A2", "[CM] T5-3"],
        ["8", "INTERFREQHOGROUP", "A3 offset small", "CIO / QoffsetFreqConn", "Similar layers only", "A3", "[CM] T5-16"],
        ["9", "EUTRANINTERNFREQ", "QoffsetFreqConn=0 start", "A3/A4/A5", "Positive = harder HO", "Ofn", "[CM]"],
        ["10", "EUTRANINTERFREQNCELL", "CIO=0 toward L900", "NRT", "No global +CIO", "Ocn", "[CM]"],
        ["11", "INTERFREQHOGROUP", "LoadBased A4 thd calibrate", "A4 > A2", "Main MLB event", "A4", "[MLB] p.137"],
        ["12", "INTERFREQHOGROUP", "A4 Hyst/TTT; NEVER 5120 ms if FreqPri needed", "Disables FreqPri/CQI/service IFHO", "Silent killer", "A4 TTT", "[CM] T4-9 p.48"],
        ["13", "INTERFREQHOGROUP", "A5 coverage Thd1/Thd2", "L900 usable", "Capacity→L900", "A5", "[CM] T5-18"],
        ["14", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType=A4", "A5 license", "Co-sited", "MLB event", "[MLB] T6-3"],
        ["15", "CELLALGOSWITCH", "CovBasedInterFreqHoMode keep unless designed", "Waiting timer", "Exec timing", "Coverage mode", "[CM] §5.3.1"],
        ["16", "INTRARATHOCOMM", "FreqPriInHoProtectionTimer non-zero", "FreqPri ON", "Stop bounce", "Protect", "[CM] p.300"],
        ["17", "FreqPri", "MlbBasedFreqPriHoSwitch ON if MLB ON", "MLB license", "Coord", "FreqPri vs MLB", "[CM] T11-7"],
        ["18", "CELLOPHOCFG", "MLB_HO_FORBID_SW optional high-speed", "NLOS misclassify", "Not indoor default", "30 km/h", "[MLB] pp.136–141"],
        ["19", "HOMEASCOMM", "HO fail punish timers", "Do not lower A4 after admit fail", "RCA first", "Penalty", "[CM] T4-16"],
        ["20", "EUTRANINTERFREQNCELL", "BlindHoPriority OFF", "Containment required", "Higher access fail", "Blind", "[CM] T5-22 p.149"],
    ]
    mml = [
        ["0", "—", "LST INTERFREQHOGROUP / EUTRANINTERNFREQ / EUTRANINTERFREQNCELL / CELLHOPARACFG / HOMEASCOMM;",
         "Read-only", "Dump first", "Baseline", "Attach to CR"],
        ["1", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>;  (FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for L900 and capacity);",
         "NRT exists", "Audit before any dBm CR", "Meas eligibility", "Most 'A4 dead' = flags"],
        ["2", "CELLUEMEASCONTROLCFG", "MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<≥6>, MaxEutranFddMeasFreqNum=<≥6>;",
         "7-layer site", "Else some L2600 never measured", "Object cap", "Equal-prio random"],
        ["3", "INTERFREQHOGROUP", "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>;  (calibrate coverage A1/A2/A5 for L2600→L900 — do not paste example dBm);",
         "Correct A2 family", "Indoor rescue first", "Coverage safety", "L900 protect"],
        ["4", "INTERFREQHOGROUP", "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<calib>, InterFreqHoA4Hyst=2, InterFreqHoA4TimeToTrig=MS320;",
         "A4 better than coverage A2; TTT≠5120ms", "Capacity MLB/FreqPri gate", "A4", "MS320 is a start, not a mandate"],
        ["5", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<capa>, MlbInterFreqHoEventType=A4, IfMlbThdRsrpOffset=0;",
         "Co-sited; A5 license not required", "A4 for capacity MLB", "Event type", "A5 only if designed"],
        ["6", "EUTRANINTERFREQNCELL", "MOD EUTRANINTERFREQNCELL: LocalCellId=<x>, DlEarfcn=<L900>, CellId=<id>, CellIndividualOffset=0;",
         "No +CIO into L900", "Avoid filling 5 MHz", "CIO", "L900"],
        ["7", "FreqPri MO", "Enable MlbBasedFreqPriHoSwitch-1 and LoadTriggerFreqPriHoSwitch-1; keep FreqPriInHoProtectionTimer non-zero;",
         "Feature 3 MLB ON", "FreqPri yields in heavy load", "Coordination", "Confirm MO in MAE"],
        ["8", "Verify", "L.HHO.InterFreq.Coverage.* ; L.HHO.InterFreq.FreqPri.* ; L.RRC.ReEst.ReconfFail.Att ; drop ; VoLTE. PrepSucc=ExecAtt/PrepAtt; ExecSucc=ExecSucc/ExecAtt.",
         "Busy hour", "Pair-level not only source", "Counters [CM] T5-24, T11-9", "Close = result"],
    ]
    return intro, items, core, pre, impact, relation, lic, params, mml


def mlb_data():
    intro = (
        "Feature 3: Intra-RAT Mobility Load Balancing. Equalisation uses peer load and reduces source–target difference. Offload protects source even without full target load. "
        "User-number: Load=N/C ; trigger N ≥ InterFreqMlbUeNumThd + MlbUeNumOffset for MlbTrigJudgePeriod. "
        "Chart: Eval N/C → Trigger → Admit target → Select UEs → HO or idle release. "
        "Ref: [MLB] Fig 3-1 p.15; §5.1.1 pp.24–27; Table 6-2 p.127."
    )
    items = [
        ["1", "Working Principal",
         "Connected HO and idle dedicated-priority release. Primary for Robi: UE_NUMBER_ONLY + SynchronizedUE + ActiveUeBasedLoadEvalSw + SpectralEffBasedLoadEvalSw. Idle MLB secondary. PRB MLB supplement only (no CA UE, no fairness). Blind/offload not for co-sited Huawei pool.",
         "C must include BW + SE (L2100 15 MHz)",
         "CHART: N/C → Thd → Target → UE pick → A4 HO  |  [MLB] §4–§6.1"],
        ["2", "Major highlighted Point",
         "Raw UE MLB on unequal BW can reduce TP [p.141]. Target needs permit, overlap, HO SR, meas flags, Low/Med HW+transport, MlbTargetInd. ONLY_STRONGEST_CELL recommended. 5s eval + MlbMaxUeNum≥40 over-transfers. PRB MLB gain dies if CA≳60%. Smart thd 7-day learn. Fixed-proportion idle + user-number connected = ping-pong. L900 not routine target.",
         "Use MlbTargetInd to block L900",
         "[MLB] pp.28, 136, 159, 207–215"],
        ["3", "Benefit and Limitations",
         "Benefit: true load equalisation; BW/SE aware; idle+connected pair; optional CA PCC transfer. Limitation: cannot fix RF/CA-scheduler/transport; coverage mismatch kills gain; blind can overload; ES interaction; L900 erosion not quantified in Huawei book.",
         "RCA first if low TP + low PRB",
         "[MLB] function-impact tables"],
        ["4", "Selection criteria (UE selection / Trigger Condition / Etc.)",
         "Source: N≥Thd+Offset for judge period. Target: Low/Med HW+transport, load diff, not punished, HO SR, MlbTargetInd, overlap. UE: UL-sync, not emergency, QCI/SPID, protect timers, optional ARP/PRB/MCS/SNR — do not pick indoor-edge just for PRB. Volume min(delta, hyst, MlbMaxUeNum). Freq pick LOADPRIORITY. Event A4.",
         "Idle trigger uses InterFreqIdleMlbUeNumThd",
         "[MLB] §6.1.1 pp.128–140"],
        ["5", "Activation parameter / Switch  (parameter) — Notes",
         "InterFreqMlbSwitch, InterFreqIdleMlbSwitch, Blind OFF, UE_NUMBER_ONLY, SynchronizedUE, ActiveUe ON, SpectralEff ON, LoadTransferEnh ON, CaUserLoadTransfer after audit, ONLY_STRONGEST_CELL, LOADPRIORITY, A4, capacity ALLOWED, L900 without connect/idle MLB.",
         "License + NRT + A4 delivery (F2)",
         "See SN-5 table  |  Table 6-2"],
        ["6", "Pre-requiste functions",
         "MLB license; X2/intra-eNB load exchange; Overlap+PERMIT; pair HO SR; target not High/OverLoad; Feature 2 A4 actually delivered; SIB5 NORMAL if idle ON; MAE 15-min counters if smart ON; CA baseline before CA-transfer SW.",
         "No load exchange ⇒ do not run equalisation",
         "[MLB] pp.27–29, 46"],
        ["7", "Mutually impacted",
         "User-number vs fixed-proportion idle. PRB_ONLY vs CA. PRB_USAGE vs PRB_VALUATION exclusive. MLB vs FreqPri reverse pair. MLB vs carrier shutdown. Flexible CA avoids offload cell. Smart thd stale after BW/upgrade. Aggressive max UE + 5s = overshoot.",
         "Do not daily MOD learned thds",
         "[MLB] §5.4.2.2, pp.216, 245"],
        ["8", "Relation with Other Feature",
         "F1 = idle method. F2 = HO engine. CA = PCC vs SCell. ES = target exclusion. Admission = all-QCI for unnecessary HO. SON smart thd = monitor only. Inter-RAT MLB out of scope (do not dump to 2G/3G for this TP KPI).",
         "A4 + ONLY_STRONGEST + L900 block",
         "F1/F2/CA/ES"],
        ["9", "License Requirements",
         "Intra-RAT Mobility Load Balancing. Idle same family (confirm). A5 extra non-cosited license. Blind option. CA-transfer option. Smart/SON option + MAE counter subscription.",
         "A5 unused if event=A4",
         "[MLB] pp.142–143"],
    ]
    core = [
        ["5.1", "CELLALGOSWITCH", "InterFreqMlbSwitch-1", "License; NRT; MlbTargetInd", "Master bit Table 6-2", "Connected/idle family", "[MLB] p.127"],
        ["5.2", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch-1", "T320; SIB5 NORMAL; L900 not idle target", "After target policy", "Idle transfer", "[MLB]/[IM]"],
        ["5.3", "CELLALGOSWITCH", "InterFreqBlindMlbSwitch-0", "Containment if ever ON", "OFF for this campaign", "Blind", "[MLB]"],
        ["5.4", "CELLMLB", "MlbTriggerMode=UE_NUMBER_ONLY", "PRB_ONLY skips CA UEs", "Primary algorithm", "Trigger mode", "[MLB] T6-2"],
        ["5.5", "CELLMLB", "InterFreqUeTrsfType=SynchronizedUE", "Match trigger mode", "IdleUE additional if idle ON", "UE type", "[MLB] T6-2"],
        ["5.6", "eval SW", "ActiveUeBasedLoadEvalSw ON", "Unequal BW (L2100 15 MHz)", "Huawei recommended", "Load N", "[MLB] T5-5"],
        ["5.7", "eval SW", "SpectralEffBasedLoadEvalSw ON", "≥10 UL-sync for SE refresh", "If SE differs >~30%", "Load C", "[MLB] T5-5"],
        ["5.8", "eval SW", "LoadTransferEnhSw ON", "4×L2600 targets", "Recommended", "Multi-target", "[MLB] p.134"],
        ["5.9", "eval SW", "CaUserLoadTransferSw ON after PCell/SCell audit", "Target CA C ≥ serving (one path)", "Else CA filtered p.157", "CA transfer", "[MLB] pp.129–136"],
        ["5.10", "CELLMLB", "MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL", "A4 meas success", "Huawei recommended p.159", "Cell pick", "[MLB] T6-5"],
        ["5.11", "CELLMLB", "FreqSelectStrategy=LOADPRIORITY", "Reliable load exchange", "PRIORITYBASED only if PCC design requires", "Freq pick", "[MLB] p.137"],
        ["5.12", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType=A4 ; MlbTargetInd ALLOWED on capacity; L900 WITHOUT connect+idle MLB", "OverlapInd; NoHoFlag PERMIT coverage", "Verify L900 enum", "Target + event", "[MLB] pp.28,129"],
    ]
    pre = [
        ["6.1", "License Intra-RAT MLB", "Present every eNB", "MAE", "Switch ON no transfer", "License", ""],
        ["6.2", "X2 / intra-eNB load", "Working on Huawei co-sited", "Else offload only", "Unsafe equalisation", "Load exchange", ""],
        ["6.3", "Overlap + PERMIT", "Capacity pairs", "NRT", "Target never admitted", "NRT", ""],
        ["6.4", "NCellHoSuccRateThld", "Fix coverage HO first if pair already bad", "F2", "Target filtered", "HO SR", ""],
        ["6.5", "HW/transport Low/Med", "Alarms clear", "High/OverLoad illegal target", "Rejected / punished", "State Fig 4-3", ""],
        ["6.6", "Feature 2 A4 delivered", "Flags, object cap, SMeasure", "F2 SN-6", "Trigger with 0 meas succ", "Meas", ""],
        ["6.7", "MAE 15-min counters if smart ON", "≥1 period subscribed", "SON", "No learned thd", "[MLB] §5.1.3.4", ""],
    ]
    impact = [
        ["7.1", "UE-number connected", "Fixed-proportion idle", "Idle fixed OFF", "Ping-pong", "[MLB] §5.4.2.2", ""],
        ["7.2", "PRB_ONLY sole mode", "High CA", "Keep UE_NUMBER_ONLY primary", "CA UEs stuck", "[MLB] §6.5", ""],
        ["7.3", "PRB_USAGE", "PRB_VALUATION", "Exclusive same mode", "Cannot both", "pp.216,245", ""],
        ["7.4", "MLB", "FreqPri reverse pair", "No reverse target", "Ping-pong", "[CM] p.303", ""],
        ["7.5", "MLB toward L900", "Indoor VoLTE / 5 MHz", "Forbidden", "Coverage layer becomes capacity", "Hard guardrail", ""],
        ["7.6", "Smart learned thd", "BW change / upgrade / eval SW change", "Freeze AI CRs in relearn week", "Stale or reset", "pp.29–31", ""],
    ]
    relation = [
        ["8.1", "Idle (F1)", "Idle MLB = dedicated prio + T320", "L900 not idle target", "F1 SN-11"],
        ["8.2", "Connected (F2)", "A4/A5, meas, admit, punish", "ONLY_STRONGEST + A4", "F2 SN-11"],
        ["8.3", "CA", "PCC transfer vs SCell scheduling", "Read PCell+SCell before CR", "[MLB] p.136"],
        ["8.4", "Energy saving", "Mutual target exclusion", "No MLB onto sleeping carrier", "pp.153–156"],
        ["8.5", "SON smart thd", "7-day learn / 7-day refresh", "AI monitors, no daily MOD", "pp.29–31"],
    ]
    lic = [
        ["9.1", "Connected/idle equalisation", "Intra-RAT Mobility Load Balancing", "No Load HO counters", "Core"],
        ["9.2", "MLB event A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay A4", "pp.142–143"],
        ["9.3", "Blind MLB", "Blind option if sold", "Keep OFF", ""],
        ["9.4", "CA user transfer", "CA + MLB CA-transfer — verify", "CA UEs filtered", "p.157"],
        ["9.5", "Smart n-cell thd", "SON MLB option + MAE counters", "No learned pair thd", ""],
    ]
    params = [
        ["1", "CELLALGOSWITCH", "InterFreqMlbSwitch ON capacity", "License NRT", "Master", "Activate", "[MLB] T6-2"],
        ["2", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch ON capacity", "T320 SIB5", "After targets", "Idle", "[MLB]"],
        ["3", "CELLALGOSWITCH", "InterFreqBlindMlbSwitch OFF", "Containment", "Amber", "Blind", "[MLB]"],
        ["4", "CELLMLB", "MlbTriggerMode=UE_NUMBER_ONLY", "CA", "Primary", "Mode", "[MLB]"],
        ["5", "CELLMLB", "InterFreqUeTrsfType=SynchronizedUE", "Mode match", "Add IdleUE if idle ON", "Type", "[MLB]"],
        ["6", "eval", "ActiveUeBasedLoadEvalSw ON", "Unequal BW", "Once then leave", "N", "[MLB] T5-5"],
        ["7", "eval", "SpectralEffBasedLoadEvalSw ON", "SE refresh", "Once then leave", "C", "[MLB] T5-5"],
        ["8", "CELL / CELLMLB", "CellCapacityScaleFactor=1 unless known MIMO/HW delta", "SE SW", "Do not fake L2100=20 MHz", "Scale", "[MLB]"],
        ["9", "eval", "LoadTransferEnhSw ON", "4×L2600", "Multi-target", "Enh", "[MLB]"],
        ["10", "eval", "CaUserLoadTransferSw after audit", "CA license", "PCC path", "CA", "[MLB]"],
        ["11", "CELLMLB", "MlbTrigJudgePeriod conservative", "Eval prd", "Stability", "Timing", "[MLB] p.128"],
        ["12", "CELLMLB", "InterFreqLoadEvalPrd not 5s if maxUE large", "MlbMaxUeNum", "Over-transfer p.136", "Timing", "[MLB]"],
        ["13", "CELLMLB", "InterFreqMlbUeNumThd calibrate per layer AFTER Active+SE ON", "Offset hyst", "MAIN daily CR — not copied from 20 MHz onto L2100", "Trigger", "[MLB] p.128"],
        ["14", "CELLMLB", "MlbUeNumOffset >0", "Thd", "Chatter guard", "Hyst", "[MLB]"],
        ["15", "CELLMLB", "MlbMaxUeNum conservative; never ≥40 with 5s", "Eval prd", "Volume", "Max UE", "[MLB] p.136"],
        ["16", "CELLMLB", "FreqSelectStrategy=LOADPRIORITY", "Load exchange", "4×L2600", "Freq", "[MLB] p.137"],
        ["17", "CELLMLB", "MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL", "A4", "Set once", "Cell", "[MLB] T6-5"],
        ["18", "EUTRANINTERNFREQ", "MlbTargetInd ALLOWED capacity; L900 without connect/idle MLB", "Overlap NoHo", "Audit once", "Target", "[MLB] p.28"],
        ["19", "EUTRANINTERFREQNCELL", "NoHoFlag PERMIT capacity + L900 coverage", "Blacklist", "MLB exclude ≠ coverage block", "NRT", "[MLB]"],
        ["20", "CELLMLB", "NCellHoSuccRateThld keep; do not lower to force MLB", "Fix RF first", "Admit", "HO SR", "[MLB]"],
        ["21", "CELLMLBUESEL", "Protect QCI1/emergency; no edge-PRB hunting", "SPID eMBMS", "UE pick", "Sel", "[MLB] pp.130–135"],
        ["22", "CELLMLB", "MlbHoInProtectTimer / UeSelectPunishTimer non-zero", "FreqPri protect", "Ping-pong", "Protect", "[MLB]"],
        ["23", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType=A4; IfMlbThdRsrpOffset=0 start", "F2 A4 thd", "Co-sited", "Event", "[MLB]"],
        ["24", "smart", "NCellTrigThldSmartOptAlgoSw monitor only daily", "7-day learn", "NO daily overwrite", "SON", "[MLB] pp.29–31"],
        ["25", "RRCCONNSTATETIMER", "T320ForLoadBalance MIN30/60", "Idle MLB", "F1", "T320", "[MLB] §5.1.1.5"],
        ["26", "CELLPRBVALMLB / PRB mode", "OFF unless GBR problem", "Exclusive vs PRB_USAGE", "Not TP-fairness tool", "PRB", "[MLB] T6-18/T6-25"],
    ]
    mml = [
        ["0", "—", "LST CELLALGOSWITCH / CELLMLB / EUTRANINTERNFREQ / CELLMLBUESEL; check license Intra-RAT MLB;",
         "MAE license", "Baseline", "Dump", "Attach to CR"],
        ["1", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L1800|L2100|L2600>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;",
         "Overlap valid; NoHo PERMIT; FREQ_MEAS_FLAG; not forbid-meas", "Capacity may be targets", "Target+event", "L900→capacity also ALLOWED"],
        ["2", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;",
         "WITHOUT_IDLE_MLB if enum exists — verify; coverage HO stays PERMIT", "L900 is coverage not capacity target", "L900 protect", "Verify enum in parameter reference"],
        ["3", "eval MO", "Turn ON ActiveUeBasedLoadEvalSw, SpectralEffBasedLoadEvalSw, LoadTransferEnhSw;",
         "Unequal BW", "BEFORE chasing UE-number thd", "Load model", "Confirm bit names in MAE"],
        ["4", "CELLMLB", "MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, FreqSelectStrategy=LOADPRIORITY, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL;",
         "InterFreqMlbSwitch will be ON", "Core equalisation", "Mode+strategy", "Huawei ONLY_STRONGEST"],
        ["5", "CELLMLB", "MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<calib>, MlbUeNumOffset=<hyst>, MlbMaxUeNum=<conservative>, MlbTrigJudgePeriod=<stable>, InterFreqLoadEvalPrd=<not 5s-if-maxUE-large>;",
         "Active+SE already ON", "Do not copy 20 MHz thd to L2100", "Trigger+volume", "MAIN guarded CR"],
        ["6", "CA SW", "Enable CaUserLoadTransferSw only after PCell/SCell/active-CC baseline;",
         "Target CA capability path", "High CA on L2600", "CA transfer", "PRB MLB still will not move CA UEs"],
        ["7", "CELLALGOSWITCH", "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
         "Steps 1–5 done; Blind bit 0", "Master ON last", "Master bits", "Idle SIB5 already NORMAL"],
        ["8", "RRCCONNSTATETIMER", "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;",
         "Idle MLB ON", "Start conservative", "T320", "F1"],
        ["9", "CELLMLBUESEL", "MOD CELLMLBUESEL: LocalCellId=<x>; protect QCI1/emergency; avoid aggressive edge PRB pick;",
         "VoLTE cells", "Do not move weak indoor to L2600", "UE pick", "ONLY_STRONGEST already on CELLMLB"],
        ["10", "Verify", "L.HHO.InterFreq.Load.* and UeNumLoad.*; HighLoad.Dur; Load.Meas(Succ); ActiveUser.DL; PCell/SCell; PRB.DL; Thrp.bits.DL/Thrp.Time.DL; idle DedicatedPri. SON logs: Inter-Frequency Handover Statistics + Idle Mode Release Statistics.",
         "15-min subscribed", "Judge sector-cluster not only source", "Tables 6-6, 6-21, 6-28", "Close = cluster TP gap"],
    ]
    return intro, items, core, pre, impact, relation, lic, params, mml


# ---------------------------------------------------------------------------
# Daily KPI sheet — literal Date Open High Low Close
# ---------------------------------------------------------------------------
def build_daily(wb):
    ws = wb.create_sheet("Daily KPI")
    cols = 10
    widths(ws, [16, 16, 16, 16, 16, 22, 22, 28, 22, 28])
    setup(ws, "Daily KPI", cols)
    ws.sheet_properties.tabColor = "7FDBFF"

    r = 1
    r = title_bar(ws, r, cols, "Mobility Management")
    r = feature_bar(ws, r, cols, "Daily KPI   —   Date | Open | High | Low | Close   (picture format, mandatory order)")
    r = format_block(ws, r, cols)
    r = blank(ws, r, cols)

    merge_put(
        ws, r, 1, cols,
        "Use this sheet for daily capacity-layer DL user-throughput monitoring. "
        "Date = KPI date (busy hour). Open = L1800 (or first capacity layer) DL user TP Mbps. "
        "High = maximum capacity-layer TP. Low = minimum capacity-layer TP. "
        "Close = gap Mbps = High − Low. Investigate when Close > 2. Change-request only after RCA pass.",
        size=10, color=WHITE, fill=BLACK, h=40, align=top,
    )
    r += 1

    # Literal picture headers in A1-style vertical already done; now horizontal table
    r = feature_bar(ws, r, cols, "Capacity-layer throughput tracker")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close", "Site / Sector", "Worst layer", "RCA class 1-8", "CR? Y/N", "Remarks"])
    # mapping reminder
    r = table_row(ws, r, [
        "KPI date BH",
        "L1800 TP Mbps",
        "Max layer TP",
        "Min layer TP",
        "Gap = High−Low",
        "Site-sector",
        "Layer of Low",
        "See RCA table below",
        "Only if approval rule pass",
        "L900 is safeguard, not a peer",
    ], h=28)

    # empty input rows for the user
    for i in range(15):
        r = table_row(ws, r, ["", "", "", "", "", "", "", "", "", ""], h=20)

    # formulas hint on first data row after header mapping - row numbers will vary; add a note
    r = blank(ws, r, cols)
    merge_put(
        ws, r, 1, cols,
        "Excel formula for Close:  =High-Low   |   Conditional: if Close>2 highlight. "
        "Do not include L900 in High/Low. Open may be changed to L2100 or a chosen anchor if L1800 is missing on that sector.",
        size=9, color=CYAN, fill=BLACK, h=28,
    )
    r += 1
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "RCA class  (Close > 2 is not a CR by itself)")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    rcas = [
        ["1", "Load / PCell user imbalance", "Low TP + high PRB + high active UE; target spare + good overlap", "YES — Feature 3", "Active-UE MLB / LOADPRIORITY / modest thd"],
        ["2", "Heavy-user PRB imbalance", "Few UEs, very high PRB, non-CA", "YES — PRB supplement only", "Never sole algo if CA high"],
        ["3", "CA / SCell", "PCell bad but SCell healthy, or SCell never on", "NO first", "CA combo / activation / scheduler"],
        ["4", "RF / interference", "Low TP + low CQI, PRB not high", "NO", "RF / PCI / overshoot / external"],
        ["5", "UE capability", "L2600 empty — no band", "NO", "Device mix"],
        ["6", "Mobility / meas fail", "High Prep, low Exec, meas succ low", "NO until F2 flags/NRT/A4 fixed", "Feature 2 audit"],
        ["7", "HW / transport / alarm", "Low TP, resources free, alarm or HighLoad transport", "NO — illegal target", "Clear alarm"],
        ["8", "L900 indoor dependence", "Survive only on L900; high-band RSRP poor", "NO forced offload", "RF / indoor; A1-escape only if overlap strong"],
    ]
    for row in rcas:
        r = table_row(ws, r, row, h=22)
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Approval rule for CR   (all must be true)")
    merge_put(
        ws, r, 1, cols,
        "IF Close > 2 Mbps AND sustained (not one 15-min) AND source normalized load high AND target spare C "
        "AND target RF/overlap OK AND pair HO success healthy AND no HW/transport/RF alarm "
        "AND drop / VoLTE / re-est / L900 indoor guardrails pass AND sample enough "
        "THEN CR (one parameter family). ELSE RCA only.",
        size=10, color=CYAN, fill=BLACK, h=48, align=top,
    )
    r += 1
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Least-invasive CR family")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    crs = [
        ["1", "Missing NRT / meas flag / L900 wrongly MLB target", "Eligibility / MlbTargetInd", "F2+F3 SN-11", "Fix first"],
        ["2", "Raw UE MLB on 15 vs 20 MHz", "ActiveUe + SpectralEff ON", "F3 SN-5", "Once"],
        ["3", "Proven load delta, good overlap, healthy HO", "InterFreqMlbUeNumThd / offset / MlbMaxUeNum", "F3 SN-11 step 5", "MAIN daily"],
        ["4", "4×L2600 RF-similar, load exchange OK", "LOADPRIORITY; equal idle prio", "F3+F1", "Set once"],
        ["5", "CA UEs stuck on busy PCC", "CaUserLoadTransferSw after audit", "F3", "After baseline"],
        ["6", "Released UEs bounce to busy layer", "Idle MLB + T320 + DediPrio hold", "F1", "After targets"],
        ["7", "L900 high AND high-band MR good", "A1/FreqPri / ThreshXhigh — not weaker L900 floors", "F1+F2", "Escape only"],
        ["8", "Ping-pong", "A4 vs A2 gap, TTT, no reverse target, ONLY_STRONGEST", "F2+F3", "Stability"],
        ["9", "Drops / late indoor", "Earlier A2/A5 — NOT more offload", "F2", "Safety"],
    ]
    for row in crs:
        r = table_row(ws, r, row, h=20)
    r = blank(ws, r, cols)

    r = feature_bar(ws, r, cols, "Hard rollback / block")
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    for row in [
        ["B1", "Target alarmed or High/OverLoad", "Block CR", "Illegal target", "Feature 3 admit"],
        ["B2", "Pair HO success below NCellHoSuccRateThld", "Block CR", "Fix RF/HO first", "F2"],
        ["B3", "L900 PRB/users/VoLTE worsen", "Rollback", "Coverage protect", "Yellow"],
        ["B4", "MLB Prep/Exec fall, ping-pong, drop, re-est", "Rollback", "Mobility health", "Counters"],
        ["B5", "Learned SON pair thresholds", "Do not daily MOD", "7-day cycle", "[MLB] pp.29–31"],
        ["B6", "Trigger + event + target + volume same night", "Forbidden", "One family only", "Process"],
    ]:
        r = table_row(ws, r, row, h=20)

    # sample chart data (hidden mapping) for a small bar of High vs Low vs Close
    r = blank(ws, r, cols)
    r = feature_bar(ws, r, cols, "Example chart data  (replace with live BH)   —   High / Low / Close")
    chart_start = r
    r = table_headers(ws, r, ["Date", "Open", "High", "Low", "Close"])
    sample = [
        ["Day-1", 18.0, 22.0, 9.5, 12.5],
        ["Day-2", 17.5, 21.0, 10.0, 11.0],
        ["Day-3", 19.0, 20.5, 11.2, 9.3],
        ["Day-4", 18.2, 19.8, 16.0, 3.8],
        ["Day-5", 18.5, 19.4, 17.6, 1.8],
        ["Day-6", 18.8, 19.2, 17.9, 1.3],
        ["Day-7", 18.6, 19.0, 17.8, 1.2],
    ]
    for row in sample:
        r = table_row(ws, r, [row[0], row[1], row[2], row[3], row[4]], h=18)
    chart_end = r - 1

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "High / Low / Close  (example)"
    chart.y_axis.title = "Mbps"
    chart.x_axis.title = "Date"
    chart.style = 10
    chart.y_axis.scaling.min = 0
    data = Reference(ws, min_col=3, min_row=chart_start, max_col=5, max_row=chart_end)
    cats = Reference(ws, min_col=1, min_row=chart_start + 1, max_row=chart_end)
    chart.add_data(data, from_rows=False, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 18
    chart.height = 8
    # try to color series: High yellow, Low yellow, Close cyan — openpyxl series graphicals
    try:
        from openpyxl.chart.series import SeriesLabel
        from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
        colors = ["F7D046", "F7D046", "7FDBFF"]
        for i, series in enumerate(chart.series):
            series.graphicalProperties.solidFill = colors[i]
    except Exception:
        pass
    ws.add_chart(chart, f"A{r + 1}")

    for i in range(1, r + 22):
        for c in range(1, cols + 1):
            if ws.cell(i, c).value is None:
                ws.cell(i, c).fill = F(BLACK)
                ws.cell(i, c).border = thin
    return ws


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto")

    build_cover(wb)

    intro, items, core, pre, impact, relation, lic, params, mml = idle_data()
    build_feature(wb, "Idle Mode Management", "Feature 1: Idle Mode Management",
                  intro, items, core, pre, impact, relation, lic, params, mml, "F7D046")

    intro, items, core, pre, impact, relation, lic, params, mml = connected_data()
    build_feature(wb, "Connected Mode", "Feature 2: Mobility Management in Connected Mode",
                  intro, items, core, pre, impact, relation, lic, params, mml, "F7D046")

    intro, items, core, pre, impact, relation, lic, params, mml = mlb_data()
    build_feature(wb, "Intra-RAT MLB", "Feature 3: Intra-RAT Mobility Load Balancing",
                  intro, items, core, pre, impact, relation, lic, params, mml, "7FDBFF")

    build_daily(wb)

    wb.properties.title = "Mobility Management — Date Open High Low Close format"
    wb.properties.creator = "Robi Axiata PLC RNO workbook"
    wb.properties.subject = "eRAN21.1 Idle / Connected / Intra-RAT MLB"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()