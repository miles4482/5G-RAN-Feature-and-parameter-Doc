#!/usr/bin/env python3
"""Huawei eRAN21.1 Mobility Management — operator Excel summary.

Presentation (user sample + original SN template):
  Title:   #005596 white bold centered
  Section: #FFFF00 bold green #008000
  Header:  #DDEBF7 black bold
  Data:    #F2F2F2 black
  MAJOR:   yellow row, green bold
  Flow:    dark-blue boxes, white text
  Last column = Chart / Doc Ref (right side of every SN)

Sheets:
  Mobility Management | Idle Mode Management | Connected Mode | Intra-RAT MLB
"""

import os
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/4G_LTE_Mobility_Management/Mobility_Management_eRAN21.1_Workbook.xlsx"
COLS = 8

BLUE = "005596"
YELLOW = "FFFF00"
GREEN = "008000"
HDR = "DDEBF7"
GREY = "F2F2F2"
WHITE = "FFFFFF"
BLACK = "000000"
ARROW = "1F4E79"
GRID = "B4B4B4"
MAJOR_BG = "FFF2CC"

thin = Border(
    left=Side(style="thin", color=GRID),
    right=Side(style="thin", color=GRID),
    top=Side(style="thin", color=GRID),
    bottom=Side(style="thin", color=GRID),
)
L = Alignment(wrap_text=True, vertical="center", horizontal="left")
C = Alignment(wrap_text=True, vertical="center", horizontal="center")
T = Alignment(wrap_text=True, vertical="top", horizontal="left")
LI = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)


def fl(h):
    return PatternFill("solid", fgColor=h)


def ft(size=10, bold=False, color=BLACK, underline=None):
    return Font(name="Calibri", size=size, bold=bold, color=color, underline=underline)


def put(ws, r, c, v, size=10, bold=False, color=BLACK, bg=GREY, align=None, h=None):
    x = ws.cell(r, c, v)
    x.font = ft(size, bold, color)
    x.fill = fl(bg)
    x.alignment = align or T
    x.border = thin
    if h:
        ws.row_dimensions[r].height = h
    return x


def merge(ws, r, c1, c2, v, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    put(ws, r, c1, v, **kw)
    bg = kw.get("bg", GREY)
    color = kw.get("color", BLACK)
    size = kw.get("size", 10)
    bold = kw.get("bold", False)
    al = kw.get("align", T)
    for c in range(c1 + 1, c2 + 1):
        y = ws.cell(r, c)
        y.fill = fl(bg)
        y.border = thin
        y.font = ft(size, bold, color)
        y.alignment = al


def widths(ws, xs):
    for i, w in enumerate(xs, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def setup(ws, footer, cols=COLS, tab=BLUE):
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(0.35, 0.35, 0.5, 0.45)
    ws.oddHeader.left.text = "Mobility Management"
    ws.oddFooter.left.text = footer
    ws.oddFooter.right.text = "Huawei eRAN21.1  |  Page &P of &N"
    ws.sheet_properties.tabColor = tab
    ws.freeze_panes = "A3"
    ws.page_setup.horizontalCentered = False
    ws.print_title_rows = "1:2"
    ws.sheet_format.defaultRowHeight = 18
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_view.showOutlineSymbols = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A3


def title(ws, r, cols, text):
    merge(ws, r, 1, cols, text, size=16, bold=True, color=WHITE, bg=BLUE, align=C, h=32)
    return r + 1


def spacer(ws, r, cols=COLS, h=8):
    for c in range(1, cols + 1):
        put(ws, r, c, "", bg=WHITE, h=h)
    return r + 1


def legend_row(ws, r, cols=COLS):
    keys = [
        (BLUE, WHITE, "Title"),
        (YELLOW, GREEN, "SN-1 to SN-11"),
        (HDR, BLACK, "Table header"),
        (GREY, BLACK, "Data row"),
        (MAJOR_BG, GREEN, "MAJOR note"),
        (ARROW, WHITE, "Process chart"),
        (WHITE, BLUE, "Col H = Chart / Doc Ref"),
        (WHITE, BLACK, "Click SN index to jump"),
    ]
    for c, (bg, fg, lab) in enumerate(keys, 1):
        put(ws, r, c, lab, size=8, bold=True, color=fg, bg=bg, align=C, h=18)
    return r + 1


def section(ws, r, cols, text):
    m = re.match(r"(SN-\d+)\s+(.*)", text.strip())
    if m:
        put(ws, r, 1, m.group(1), size=12, bold=True, color=WHITE, bg=BLUE, align=C, h=24)
        merge(ws, r, 2, cols, "  " + m.group(2), size=12, bold=True, color=GREEN, bg=YELLOW, align=LI, h=24)
    else:
        merge(ws, r, 1, cols, "  " + text, size=12, bold=True, color=GREEN, bg=YELLOW, align=LI, h=22)
    return r + 1


def major(ws, r, cols, text):
    put(ws, r, 1, "MAJOR", size=9, bold=True, color=WHITE, bg=GREEN, align=C, h=22)
    merge(ws, r, 2, cols, "  " + text, size=10, bold=True, color=GREEN, bg=MAJOR_BG, align=LI, h=22)
    return r + 1


def pad8_heads(names):
    names = list(names)
    if len(names) >= COLS:
        return names[:COLS]
    if names and str(names[-1]).startswith("Chart"):
        names = names[:-1] + ["More Notes", names[-1]]
        return names[:COLS]
    while len(names) < COLS - 1:
        names.append("")
    names.append("Chart / Doc Ref")
    return names[:COLS]


def pad8(values, mml=False):
    v = list(values)
    if mml:
        while len(v) < 7:
            v.append("")
        if len(v) == 7:
            v.append("MML sequence")
        return v[:COLS]
    if len(v) == 7:
        return v[:6] + ["", v[6]]
    while len(v) < COLS:
        v.append("")
    return v[:COLS]


def auto_h(values, min_h=22, max_h=78):
    longest = 0
    for i, v in enumerate(values):
        s = "" if v is None else str(v)
        # wider MML / description columns wrap ~42 chars
        width = 18 if i <= 1 else 42
        lines = max(1, (len(s) + width - 1) // width)
        n = s.count("\n") + 1
        longest = max(longest, lines, n)
    return min(max_h, max(min_h, 16 + longest * 12))


def heads(ws, r, names, cols=COLS):
    names = pad8_heads(names)
    for c in range(1, cols + 1):
        n = names[c - 1] if c <= len(names) else ""
        put(ws, r, c, n, size=9, bold=True, color=BLACK, bg=HDR, align=C, h=22)
    return r + 1


def rec(ws, r, values, h=None, mml=False, cols=COLS):
    values = pad8(values, mml=mml)
    computed = auto_h(values, min_h=26 if mml else 22, max_h=96 if mml else 72)
    h = max(h or 0, computed)
    for c in range(1, cols + 1):
        v = values[c - 1] if c <= len(values) else ""
        put(ws, r, c, v, size=9 if mml and c == 3 else 10, bg=GREY, align=T, h=h)
    return r + 1


def _is_ref(s):
    s = str(s)
    return s.startswith(("Fig", "Doc", "IM ", "MLB", "CM ", "pp.", "§", "Table", "Ch.")) or s in ("MML", "Right column")


def flow(ws, r, cols, boxes):
    """Process chart: box → box → box, then ↓ to the next row. Last token can be a doc ref."""
    steps = list(boxes)
    ref = ""
    if steps and _is_ref(steps[-1]):
        ref = steps.pop()
    box_spans = [(1, 2), (4, 5), (7, 8)]
    arrow_cols = [3, 6]
    i = 0
    while i < len(steps):
        chunk = steps[i:i + 3]
        for c in range(1, cols + 1):
            put(ws, r, c, "", bg=WHITE, h=30)
        for j, step in enumerate(chunk):
            c1, c2 = box_spans[j]
            merge(ws, r, c1, c2, step, size=9, bold=True, color=WHITE, bg=BLUE, align=C, h=30)
            if j < len(chunk) - 1:
                put(ws, r, arrow_cols[j], "→", size=16, bold=True, color=GREEN, bg=WHITE, align=C, h=30)
        r += 1
        i += 3
        if i < len(steps):
            for c in range(1, cols + 1):
                put(ws, r, c, "", bg=WHITE, h=14)
            merge(ws, r, 1, cols, "↓", size=14, bold=True, color=GREEN, bg=WHITE, align=C, h=14)
            r += 1
    if ref:
        merge(ws, r, 1, cols, "  Chart / Doc Ref (right side of SN):  " + ref, size=9, bold=True, color=BLUE, bg=HDR, align=LI, h=18)
        r += 1
    return r


def arrow_row(ws, r, cols, n_boxes):
    for i in range(cols):
        if i < n_boxes - 1:
            put(ws, r, i + 1, "↓  next", size=8, bold=True, color=GREEN, bg=WHITE, align=C, h=14)
        else:
            put(ws, r, i + 1, "", bg=WHITE, h=14)
    return r + 1


def link_cell(cell, target, sheet):
    loc = f"'{sheet}'!{target}"
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc, display=str(cell.value) if cell.value else loc)
    cell.font = ft(10, True, "0563C1", underline="single")
    cell.alignment = T


# Shared 8-col header sets. Column H is always Chart / Doc Ref.
H_SN = ["SN", "RAT", "Item Name", "What the document says", "Conditional / switch to enable", "Notes / Rationale", "More Notes", "Chart / Doc Ref"]
H_STEP = ["Step", "RAT", "Activity (connected order)", "MO / Parameter", "Document value / rule", "Notes / Rationale", "More Notes", "Chart / Doc Ref"]
H_ACT = ["Seq", "RAT", "MO", "Parameter / Switch", "Document value", "Conditional Parameter (must enable first)", "Notes", "Chart / Doc Ref"]
H_PRE = ["Seq", "RAT", "MO / Check", "Parameter / Switch", "Document value", "If missing, what fails", "Notes", "Chart / Doc Ref"]
H_IMP = ["Seq", "RAT", "This function", "Impacts / couples with", "Document rule", "Conditional Parameter / related feature", "Notes", "Chart / Doc Ref"]
H_PAR = ["Seq", "RAT", "MO", "Parameter Name", "Parameter description", "Document value / Conditional Parameter", "Notes", "Chart / Doc Ref"]
H_MML = ["Parameter Sequence", "MO", "Activation Value", "Conditional Parameter", "Remarks", "Parameter Description", "More Notes", "Chart / Doc Ref"]

DEFAULT_WIDTHS = [12, 14, 34, 36, 28, 32, 26, 26]


def build_sheet(wb, name, title_text, blocks, w=None, tab=BLUE):
    ws = wb.create_sheet(name)
    cols = COLS
    widths(ws, w or DEFAULT_WIDTHS)
    setup(ws, title_text, cols, tab=tab)
    r = 1
    r = title(ws, r, cols, title_text)
    r = legend_row(ws, r, cols)
    r = spacer(ws, r, cols, 6)

    bookmarks = {}
    index_rows = []
    sn_starts = []
    in_index = False
    in_mml = False

    for b in blocks:
        kind = b[0]
        if kind == "section":
            in_index = "SN list" in b[1]
            in_mml = b[1].strip().startswith("SN-11") or "Final MML" in b[1]
            m = re.match(r"SN-(\d+)", b[1].strip())
            r = section(ws, r, cols, b[1])
            if m:
                bookmarks[m.group(1)] = r - 1
                sn_starts.append(r - 1)
        elif kind == "major":
            r = major(ws, r, cols, b[1])
        elif kind == "flow":
            r = flow(ws, r, cols, b[1])
        elif kind == "arrows":
            r = arrow_row(ws, r, cols, b[1])
        elif kind == "heads":
            r = heads(ws, r, b[1], cols)
        elif kind == "row":
            if in_index:
                index_rows.append((r, str(b[1][0])))
            r = rec(ws, r, b[1], h=b[2] if len(b) > 2 else None, mml=in_mml)
        elif kind == "space":
            r = spacer(ws, r, cols, b[1] if len(b) > 1 else 8)
        elif kind == "note":
            merge(ws, r, 1, cols, "  " + b[1], size=9, bold=False, color=BLUE, bg=WHITE, align=LI, h=20)
            r += 1
        elif kind == "links":
            r = heads(ws, r, ["Open sheet", "RAT", "Feature", "What you will find", "SN structure", "How to use", "More Notes", "Chart / Doc Ref"])
            for item in b[1]:
                rec(ws, r, item, h=24)
                link_cell(ws.cell(r, 1), "A1", item[0])
                r += 1

    for ir, sn in index_rows:
        if sn in bookmarks:
            target = f"A{bookmarks[sn]}"
            for col in (1, 3):
                link_cell(ws.cell(ir, col), target, name)

    if sn_starts:
        ends = sn_starts[1:] + [ws.max_row + 1]
        for start, end in zip(sn_starts, ends):
            if end - 1 > start:
                ws.row_dimensions.group(start + 1, end - 1, hidden=False, outline_level=1)
    return ws


# ===========================================================================
# COVER
# ===========================================================================
def cover_blocks():
    return [
        ("section", "Folder  —  Mobility Management"),
        ("note", "Open a feature sheet below. Each feature is SN-1 to SN-11 in connected order. Column H is Chart / Doc Ref on the right of every SN. Click an SN number on a feature sheet to jump to that section."),
        ("links", [
            ["Idle Mode Management", "LTE", "Feature 1", "Cell selection, reselection, SIB3/SIB5, dedicated priority, idle MLB interface", "SN-1 to SN-11", "Start here for camping / next RRC cell", "—", "Fig 4-1, Fig 5-1"],
            ["Connected Mode", "LTE", "Feature 2", "A1–A5, coverage HO, frequency-priority HO, measurement, admission", "SN-1 to SN-11", "HO engine. MLB algorithm is not in this book.", "—", "Fig 4-1, Fig 11-1/11-2"],
            ["Intra-RAT MLB", "LTE", "Feature 3", "Idle dedicated-priority transfer and connected load HO (equalisation / offload)", "SN-1 to SN-11", "Who/when to move for load. Uses Connected Mode A4/A5.", "—", "Fig 3-1, Fig 4-1 to 4-5"],
        ]),
        ("space", 8),
        ("section", "Introduction"),
        ("note", "Feature details and MAJOR notes are highlighted in yellow/green. Last column (H) is Chart / Doc Ref (right side of SN), as required in the original template. One idea per row."),
        ("space", 8),
        ("heads", ["Topic", "RAT", "Feature Part", "Document", "Issue / date", "What this book covers", "Chart / Doc Ref"]),
        ("row", ["Feature 1", "LTE", "Idle Mode Management", "Idle Mode Management Feature Parameter Description", "eRAN21.1 Issue 04, 2026-05-30",
                 "Cell selection, reselection, dedicated priority, SIB3/SIB5, idle MLB interface", "Fig 4-1, Fig 5-1"]),
        ("row", ["Feature 2", "LTE", "Mobility Management in Connected Mode", "Mobility Management in Connected Mode Feature Parameter Description", "eRAN21.1 Issue 08, 2026-06-30",
                 "A1/A2/A3/A4/A5, coverage HO, frequency-priority HO, measurement, admission", "Fig 4-1, Fig 11-1/11-2"]),
        ("row", ["Feature 3", "LTE", "Intra-RAT Mobility Load Balancing", "Intra-RAT Mobility Load Balancing Feature Parameter Description", "eRAN21.1 Issue 10, 2026-06-30",
                 "Idle dedicated-priority transfer and connected load HO (equalisation / offload)", "Fig 3-1, Fig 4-1 to 4-5"]),
        ("space", 10),
        ("major", "These three documents are one mobility chain. Idle decides the next access cell. Connected Mode executes measurement and HO. MLB decides which UEs move for load. Frequency-priority HO is not the MLB algorithm."),
        ("major", "The sample MML row ENodeBAlgoSwitch / SymbolShutdownSwitch in the original CSV is Symbol Power Saving. It is not a mobility parameter and is not used here."),
        ("space", 10),
        ("section", "How the three features connect (sequence)"),
        ("flow", ["1 Idle camp / reselect", "2 RRC connect", "3 Coverage A2/A5 protect", "4 FreqPri A1/A4 steer", "5 MLB load HO / idle release", "6 Next idle (T320)", "Doc chain"]),
        ("space", 8),
        ("heads", ["Step", "RAT", "Activity (connected order)", "Document", "Trigger", "Notes / Rationale", "Chart / Doc Ref"]),
        ("row", ["1", "LTE", "Idle selection and reselection", "Idle Mode Management", "Power-on, release, SI", "Sets the cell of the next RRC setup", "Fig 5-1"]),
        ("row", ["2", "LTE", "Enter RRC_CONNECTED", "Connected Mode", "Service request", "Dedicated idle priorities are discarded", "IM §5.1.3.1"]),
        ("row", ["3", "LTE", "Coverage mobility", "Connected Mode", "Event A2", "Necessary HO. Preempts load HO.", "Table 4-5, Fig 5-3"]),
        ("row", ["4", "LTE", "Frequency-priority HO", "Connected Mode Ch.11", "A1 / A4 when serving is good", "Puts service on high band; keeps low band for coverage", "Fig 11-1/11-2"]),
        ("row", ["5", "LTE", "Intra-RAT MLB", "MLB book", "UE-number or PRB threshold", "HO cause Reduce Load in Serving Cell, or idle dedicated priority", "Fig 3-1"]),
        ("row", ["6", "LTE", "Return to idle", "Idle + MLB", "RRC release + T320", "Idle MLB can steer the next camp", "MLB §5.1.1.5"]),
        ("space", 10),
        ("section", "Original template SN list  (every feature sheet follows this order)"),
        ("heads", H_SN),
        ("row", ["1", "LTE", "Working Principal", "Mechanism in connected order", "—", "Must include chart from doc on the right", "Right column"]),
        ("row", ["2", "LTE", "Major highlighted Point", "Huawei cautions and recommended settings", "—", "Highlighted MAJOR rows", "Right column"]),
        ("row", ["3", "LTE", "Benefit and Limitations", "What the feature gives and where it fails", "—", "Table", "Right column"]),
        ("row", ["4", "LTE", "Selection criteria / Trigger", "UE pick, event, threshold", "—", "Table", "Right column"]),
        ("row", ["5", "LTE", "Activation parameter / Switch", "Core switch + description + value + conditional parameter", "Conditional = what must be ON to enable the main switch", "Must be a table", "Right column"]),
        ("row", ["6", "LTE", "Prerequisite functions", "Parameter / Switch + description + value + notes", "Table", "Must be a table", "Right column"]),
        ("row", ["7", "LTE", "Mutually impacted", "Parameter setting + description + value", "Table", "Must be a table", "Right column"]),
        ("row", ["8", "LTE", "Relation with Other Feature", "Idle / Connected / MLB / CA / ES", "—", "—", "Right column"]),
        ("row", ["9", "LTE", "License Requirements", "Function to license mapping", "Verify in MAE", "Feature books point to license file", "Right column"]),
        ("row", ["10", "LTE", "All Parameter List", "Sequence / connected order", "Conditional Parameter = relation with other feature", "Must be a table", "Right column"]),
        ("row", ["11", "LTE", "Final MML Command", "Parameter Sequence, MO, Activation Value, Conditional Parameter, Remarks, Parameter Description, More Notes", "Maintain sequence", "Placeholders LocalCellId / DlEarfcn", "Right column"]),
    ]


# ===========================================================================
# IDLE
# ===========================================================================
def idle_blocks():
    return [
        ("section", "Process chart from document  (right-side chart for SN-1)"),
        ("flow", ["PLMN select", "Criterion S", "Camp + read SI", "Measure", "Reselect", "RRC on camped cell", "Fig 5-1"]),
        ("note", "Idle Mode Management Fig 4-1 (idle functions) and Fig 5-1 (selection/reselection sequence). A sequence is a set of things that come one after another in connected order."),
        ("space", 8),
        ("section", "SN list"),
        ("heads", H_SN),
        ("row", ["1", "LTE", "Working Principal", "Selection → SI → measurement → reselection → dedicated priority at release", "SIB3/SIB5 must be complete", "See SN-1 steps below", "Fig 4-1, Fig 5-1"]),
        ("row", ["2", "LTE", "Major highlighted Point", "Priority is per frequency; SI is delayed; UNDELIVER blocks idle MLB", "CfgInd=CFG required", "See MAJOR rows", "§5.1.3, §7.1.3"]),
        ("row", ["3", "LTE", "Benefit and Limitations", "Low signalling vs connected HO; not real-time", "T320 / next RRC ends dedicated prio", "See SN-3 table", "MLB §4.3"]),
        ("row", ["4", "LTE", "Selection criteria / Trigger", "Criterion S, ThreshXhigh/low, ranking, idle MLB thd", "Timers must persist", "See SN-4 table", "Tables 5-1 to 5-4"]),
        ("row", ["5", "LTE", "Activation parameter / Switch", "Priority, search CFG, idle MLB bit, T320, MlbTargetInd", "License + SIB5", "See SN-5 table", "§5.3.2.3"]),
        ("row", ["6", "LTE", "Prerequisite functions", "SIB5 list, neighbor 16-cap, SI BER, UE capability", "—", "See SN-6 table", "§5.1.3.4, §5.3.4"]),
        ("row", ["7", "LTE", "Mutually impacted", "Connected MLB, CA PCC, RSRQ, energy saving, fixed-proportion idle", "Do not combine forbidden pairs", "See SN-7 table", "MLB §5.4.2.2"]),
        ("row", ["8", "LTE", "Relation with Other Feature", "MLB idle method; Connected discards dedicated prio; CA; ES; GERAN/UTRAN class", "—", "See SN-8 table", "§5.1.3.1"]),
        ("row", ["9", "LTE", "License Requirements", "Basic idle vs Intra-RAT MLB vs blind vs enhanced dedi-prio", "Check MAE", "See SN-9 table", "License file"]),
        ("row", ["10", "LTE", "All Parameter List", "Suitability → priority → meas → reselection → idle MLB", "Conditional = other-feature relation", "See SN-10 table", "Connected order"]),
        ("row", ["11", "LTE", "Final MML Command", "LST → priority/search → SIB5 → thresholds → MlbTargetInd → idle switch → T320 → verify", "Maintain sequence", "See SN-11 table", "MML"]),
        ("space", 10),

        ("section", "SN-1  Working Principal  (connected order)"),
        ("heads", H_STEP),
        ("row", ["1", "LTE", "PLMN selection", "NAS / PLMN", "Before cell selection", "Not controlled by SIB reselection priority", "Fig 5-1"]),
        ("row", ["2", "LTE", "Initial / stored / release-directed cell selection", "CELLSEL ; RRCConnectionRelease redirected frequency", "Last camped cell, stored info, or strongest cell per band", "SIB priority does not fully control first camp after power-up", "§5.1.2 pp.15–16"]),
        ("row", ["3", "LTE", "Criterion S (suitable cell)", "QRxLevMin, QQualMin, UePowerMax / PMax",
                 "Srxlev = Qrxlevmeas−(Qrxlevmin+offset)−Pcompensation > 0; Squal>0 if QQualMin set",
                 "If QQualMin is 0/absent, RSRP-only. This is a floor, not a load knob.", "§5.1.2"]),
        ("row", ["4", "LTE", "Camp and read system information", "SIB1 / SIB3 / SIB4 / SIB5", "SIB3 intra + serving prio; SIB5 inter-freq prio, thd, offset, neighbor list", "SIB1 period 80 ms. Other SIB periods configurable.", "Table 7-1"]),
        ("row", ["5", "LTE", "Intra-frequency measurement", "SIntraSearch / SIntraSearchQ", "Measure unless Srxlev>SIntraSearchP and Squal>SIntraSearchQ", "Huawei: set CfgInd=CFG to avoid always-on intra meas.", "§5.1.3.3"]),
        ("row", ["6", "LTE", "Inter-frequency measurement", "SNonIntraSearch ; CellReselPriority", "Higher-priority freq: always measured. Equal/lower: only after search threshold", "Thresholds do not stop high-priority search.", "§5.1.3.3"]),
        ("row", ["7", "LTE", "Reselection to higher-priority frequency", "ThreshXhigh / ThreshXhighQ / EutranReselTime", "Camped >1 s AND target S > high thd for the timer", "Tables 5-1 / 5-2", "pp.29–32"]),
        ("row", ["8", "LTE", "Reselection to equal-priority frequency", "Qhyst, QoffsetFreq, CellQoffset", "Rn = Qmeas,n − Qoffset ; Rs = Qmeas,s + Qhyst ; Rn>Rs for the timer", "Positive target offset makes reselection harder.", "§5.1.3.4"]),
        ("row", ["9", "LTE", "Reselection to lower-priority frequency", "ThrshServLow / ThreshXlow (+ Q forms)", "No higher target AND serving S < serving-low AND target S > X-low", "Tables 5-3 / 5-4", "pp.33–35"]),
        ("row", ["10", "LTE", "Dedicated priority at RRC release (idle MLB / SPID / PCC)", "IdleModeMobilityControlInfo ; T320", "Replaces common SIB priorities until T320 / next RRC / PLMN select", "Missing serving freq in dedicated list ⇒ treated as lowest.", "§5.1.3.1"]),
        ("row", ["11", "LTE", "Next RRC setup", "Camped cell", "UE establishes RRC on the cell where it is then camped", "Idle distribution = next access layer", "Fig 5-1"]),
        ("space", 8),

        ("section", "SN-2  Major highlighted Point"),
        ("major", "Priority is frequency-level, not per-cell. No SIB5 priority ⇒ no reselection to that frequency. Maximum 16 non-serving E-UTRAN frequencies. [IM] §5.1.3.1"),
        ("major", "Huawei recommends SNonIntraSearchCfgInd=CFG, SIntraSearch > SNonIntraSearch, and example SNonIntraSearch=10. [IM] §5.4.1.1 ; [MLB] §5.1.2.1"),
        ("major", "Huawei recommends capacity/hotspot frequencies above the low-band coverage layer in broadcast reselection priority. [MLB] §5.1.2.1 pp.38–39"),
        ("major", "MeasPerformanceDemand=UNDELIVER removes the frequency from SIB5 and it cannot be an idle-MLB target. [IM] §5.3.2.3"),
        ("major", "SI change is not instant. Applied in the next SI modification period; otherwise after change-paging or 3 hours. SIB BER must be ≤1%. [IM] §7.1.3, §5.3.4"),
        ("major", "Do not use the optimization-table sentence on QRxLevMinOffset. It conflicts with the Criterion-S formula. [IM] §5.1.2 vs §5.4.1.1.3"),
        ("major", "Adaptive-proportion idle balancing is not recommended without Huawei engineering support. Fixed-proportion idle + user-number connected MLB causes ping-pong. [MLB] §5.5.2.1 p.80 ; §5.4.2.2"),
        ("space", 8),

        ("section", "SN-3  Benefit and Limitations"),
        ("heads", ["Type", "RAT", "Item", "Document statement", "Condition", "User impact", "Chart / Doc Ref"]),
        ("row", ["Benefit", "LTE", "Idle transfer", "Avoids gap-assisted inter-frequency measurement and connected HO", "Idle MLB / dedicated prio", "Lower immediate signalling and gap loss than connected HO", "MLB §4.3 Figs 4-4/4-5"]),
        ("row", ["Benefit", "LTE", "Sets next access cell", "Next RRC starts on the camped cell", "After reselection or dedicated prio", "Long-term layer distribution", "Fig 5-1"]),
        ("row", ["Benefit", "LTE", "Temporary dedicated priority", "T320-bounded; discarded at next RRC", "Load-balance release uses T320ForLoadBalance", "Does not permanently rewrite SIB", "§5.1.3.1"]),
        ("row", ["Limitation", "LTE", "Not real-time", "Dedicated priorities discarded when UE enters connected", "Long RRC session", "Cannot rebalance an already-connected UE", "§5.1.3.1"]),
        ("row", ["Limitation", "LTE", "Static common priority", "Highest usable priority tier collects most idle UEs", "Equal-priority still ranks by RF, not load", "Need idle MLB for dynamic load", "§5.1.3.5"]),
        ("row", ["Limitation", "LTE", "SI delay", "Next SI modification period / 3 hours", "Judge camping only after SI + turnover", "Do not evaluate in the same 15-min as the change", "§7.1.3"]),
        ("row", ["Limitation", "LTE", "UE capability", "Legacy UEs: only NORMAL freqs; ≤8 dedicated freqs without enhanced meas", "incMonEUTRA", "Enhanced UEs up to 16 dedicated freqs", "§5.1.3.1 / §5.1.3.3"]),
        ("row", ["Limitation", "LTE", "Load-based redirection", "Selects by QoS/priority, not measured RF", "No location quality check", "More disruptive than HO", "§6.1.1 pp.69–70"]),
        ("space", 8),

        ("section", "SN-4  Selection criteria / Trigger Condition"),
        ("heads", ["Rule", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration", "Chart / Doc Ref"]),
        ("row", ["Criterion S RSRP", "LTE", "QRxLevMin, QRxLevMinOffset, UePowerMax/PMax", "Cell selection / reselection suitability", "Srxlev>0", "Too-high floor causes access failure", "§5.1.2"]),
        ("row", ["Criterion S RSRQ", "LTE", "QQualMin, QQualMinOffset", "Only if QQualMin configured ≠ 0/absent", "Squal>0", "RSRQ moves with load — can oscillate", "§5.1.2"]),
        ("row", ["Higher-priority target", "LTE", "ThreshXhigh / ThreshXhighQ / EutranReselTime", "Always measuring higher-priority freqs", "Target S > high thd for timer; camped >1 s", "Protects against premature high-band camp", "Tables 5-1/5-2"]),
        ("row", ["Lower-priority target", "LTE", "ThrshServLow / ThreshXlow", "Serving becomes poor and no higher target", "Both serving-low and target-low must pass", "Too-low serving-low = sticky poor serving cell", "Tables 5-3/5-4"]),
        ("row", ["Equal-priority rank", "LTE", "Qhyst / QoffsetFreq / CellQoffset", "Same priority frequencies", "Rn>Rs for reselection time", "Do not use static offset as hourly load control", "§5.1.3.4"]),
        ("row", ["Equal/lower search start", "LTE", "SNonIntraSearch / SNonIntraSearchQ", "Serving S ≤ search threshold", "Does not apply to higher-priority freqs", "CFG recommended (battery)", "§5.1.3.3"]),
        ("row", ["Idle MLB trigger", "LTE", "InterFreqIdleMlbUeNumThd + IdleUE", "Idle-user load ≥ thd through judge period", "Dedicated prio to low-load E-UTRAN freqs", "Class order NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "MLB §5.1.1.5 p.33"]),
        ("space", 8),

        ("section", "SN-5  Activation parameter / Switch   (core table)"),
        ("note", "Conditional Parameter = what must be set/enabled before the main switch takes effect."),
        ("heads", H_ACT),
        ("row", ["5.1", "LTE", "CELLRESEL", "CellReselPriority", "SIB3 serving common priority. Larger = higher. Huawei: capacity/hotspot above coverage layer.", "SIB3 broadcast", "§5.1.3.1 ; MLB §5.1.2.1"]),
        ("row", ["5.2", "LTE", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd", "CFG — otherwise UE does not reselect to that frequency", "Frequency listed in SIB5 plan", "§5.1.3.1"]),
        ("row", ["5.3", "LTE", "EUTRANINTERNFREQ", "CellReselPriority", "SIB5 target priority (same hierarchy as serving design)", "CfgInd=CFG", "§5.1.3.1"]),
        ("row", ["5.4", "LTE", "EUTRANINTERNFREQ", "MeasPerformanceDemand", "NORMAL for main capacity and intensive-coverage frequencies", "UNDELIVER ⇒ not in SIB5 and not idle-MLB target", "§5.1.3.3 ; §5.3.2.3"]),
        ("row", ["5.5", "LTE", "CELLRESEL", "SIntraSearchCfgInd + SIntraSearch", "CFG. Huawei: SIntraSearch > SNonIntraSearch", "SIB3", "§5.4.1.1"]),
        ("row", ["5.6", "LTE", "CELLRESEL", "SNonIntraSearchCfgInd + SNonIntraSearch", "CFG. Huawei example value 10", "Does not stop higher-priority measurement", "§5.4.1.1 ; MLB §5.1.2.1"]),
        ("row", ["5.7", "LTE", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch", "ON to enable intra-LTE idle load equalisation", "InterFreqMlbSwitch + license + MlbTargetInd allows idle", "§5.3.2.3"]),
        ("row", ["5.8", "LTE", "EnhancedMlbAlgoSwitch", "DediPrioManageOnLowLoadSw (+ EnhSw)", "ON to keep released UEs off higher-load frequencies", "T320 set; idle MLB path", "MLB Table 5-10"]),
        ("row", ["5.9", "LTE", "RRCCONNSTATETIMER", "T320ForLoadBalance", "Lifetime of load-balance dedicated priorities", "Idle MLB ON. SPID/PCC T320 is always 180 min.", "§5.1.3.1"]),
        ("row", ["5.10", "LTE", "EUTRANINTERNFREQ", "MlbTargetInd", "ALLOWED, or ALLOWED_WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB", "OverlapInd / NoHoFlag still control coverage HO", "MLB pp.28, 129"]),
        ("row", ["5.11", "LTE", "GlobalProcSwitch", "CellReselectionOptSwitch", "ON recommended with idle equalisation (preferential LTE freq delivery)", "Dedicated-priority path", "§5.3.2.3"]),
        ("space", 8),

        ("section", "SN-6  Prerequisite functions"),
        ("heads", H_PRE),
        ("row", ["6.1", "LTE", "EUTRANINTERNFREQ", "Complete SIB5 inter-frequency set", "All required LTE freqs present, normally NORMAL", "UE cannot reselect a missing frequency", "§5.1.3.1"]),
        ("row", ["6.2", "LTE", "EUTRANINTERFREQNCELL", "Neighbor list in SIB4/SIB5", "Max 16 listed neighbors per frequency", "Truncation looks like a threshold problem", "§5.1.3.4"]),
        ("row", ["6.3", "LTE", "InterFreqBlkCell", "ApplicationScope", "Do not use blacklist as load control. 65535 = idle+connected. Only 16 idle blacklists delivered.", "Hidden valid target", "§5.1.3.2"]),
        ("row", ["6.4", "LTE", "SIB broadcast", "SI BER ≤ 1%", "Required for reselection", "Failed SIB5 decode looks like failed layer balance", "§5.3.4"]),
        ("row", ["6.5", "LTE", "UE capability", "Band support / incMonEUTRA", "Legacy may see only NORMAL and ≤8 dedicated freqs", "Idle cannot move incapable UEs", "§5.1.3.1"]),
        ("row", ["6.6", "LTE", "Idle MLB features", "Adaptive-proportion / fixed-proportion idle", "Adaptive: not recommended. Fixed-proportion: do not combine with user-number connected MLB.", "Ping-pong", "MLB §5.5.2.1 ; §5.4.2.2"]),
        ("space", 8),

        ("section", "SN-7  Mutually impacted"),
        ("heads", H_IMP),
        ("row", ["7.1", "LTE", "Idle dedicated priority", "Connected MLB / A1–A5", "Dedicated prio discarded at RRC connect", "Do not fight idle vs connected on the same day without design", "§5.1.3.1"]),
        ("row", ["7.2", "LTE", "Fixed-proportion idle MLB", "User-number connected MLB", "Huawei ping-pong warning — do not combine", "Keep fixed-proportion OFF if connected user-number MLB is ON", "MLB §5.4.2.2"]),
        ("row", ["7.3", "LTE", "ThreshXhigh too low", "UL / access on high band", "Premature high-band camping", "Coverage quality of target", "Tables 5-1/5-2"]),
        ("row", ["7.4", "LTE", "ThrshServLow too low", "Connected A2/A5", "UE stays on a dying serving cell in idle", "Align idle serving-low with coverage A2 philosophy", "Tables 5-3/5-4"]),
        ("row", ["7.5", "LTE", "RSRQ-based reselection", "Scheduler load", "RSRQ falls when cell is busy → oscillation", "RSRP is the more stable quantity (same caution as connected book)", "CM Table 4-15"]),
        ("row", ["7.6", "LTE", "Energy saving / carrier shutdown", "SIB5 highest-priority list", "Sleeping cell must not remain a high idle target", "ES feature interaction", "MLB pp.153–156"]),
        ("space", 8),

        ("section", "SN-8  Relation with Other Feature"),
        ("heads", H_IMP),
        ("row", ["8.1", "LTE", "Intra-RAT MLB", "Idle transfer method", "Dedicated priority + T320 + InterFreqIdleMlbSwitch", "MlbTargetInd can forbid idle targeting of a frequency", "MLB §5.1.1.5"]),
        ("row", ["8.2", "LTE", "Connected Mode", "A1/A2/A5 vs idle thds", "Dedicated idle prio discarded at connect", "Align ThreshXhigh with A1-escape; ThrshServLow with A5-return", "CM Ch.5 / Ch.11"]),
        ("row", ["8.3", "LTE", "Carrier Aggregation", "PCC anchoring + CaUserLoadTransferSw", "Changes whether CA-capable UEs are idle-steered", "PCC load ≠ SCC load", "MLB Fig 5-3"]),
        ("row", ["8.4", "LTE", "SPID / operator / RAN share", "Dedicated priorities from SPID or operator", "SPID/PCC T320 always 180 min", "Can override MLB dedicated prio design", "§5.1.3.1"]),
        ("row", ["8.5", "LTE", "GERAN / UTRAN", "Dedicated class order", "NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "Idle MLB class, not a 4G-to-2G dump tool", "MLB §5.1.1.5 p.33"]),
        ("space", 8),

        ("section", "SN-9  License Requirements"),
        ("heads", ["Function", "RAT", "Switch / function", "License (verify in MAE / license file)", "If missing", "Notes", "Chart / Doc Ref"]),
        ("row", ["Basic idle / SIB reselection", "LTE", "CELLSEL / CELLRESEL / SIB3/SIB5", "Basic LTE eNodeB", "N/A", "Normally included", "—"]),
        ("row", ["Idle MLB / dedicated load priority", "LTE", "InterFreqIdleMlbSwitch", "Intra-RAT Mobility Load Balancing family (confirm exact name)", "Switch may look ON; dedicated-pri counters stay 0", "Same family as connected MLB", "MLB book"]),
        ("row", ["Blind idle / blind MLB", "LTE", "InterFreqBlindMlbSwitch", "Blind MLB option if sold separately", "Blind path unavailable", "Higher access risk", "MLB"]),
        ("row", ["DediPrioManageOnLowLoad(Enh)", "LTE", "EnhancedMlbAlgoSwitch", "Enhanced MLB / idle-priority option — verify", "Released UEs bounce back to high-load freq", "Table 5-10", "MLB §5.3.1"]),
        ("space", 8),

        ("section", "SN-10  All Parameter List  (sequence / connected order)"),
        ("note", "Conditional Parameter column = relation with other feature, as requested in the original template."),
        ("heads", H_PAR),
        ("row", ["1", "LTE", "CELLSEL", "QRxLevMin / QQualMin / offsets / UePowerMax", "Cell-selection Criterion S floors", "Not a load knob. Pcompensation uses UePowerMax.", "§5.1.2"]),
        ("row", ["2", "LTE", "CELLRESEL", "QRxLevMin / QQualMin / PMax", "Serving/intra reselection suitability", "Same philosophy as selection", "§5.1.3.4"]),
        ("row", ["3", "LTE", "CELLRESEL", "CellReselPriority", "SIB3 serving common priority", "Capacity/hotspot above coverage layer (Huawei note)", "§5.1.3.1"]),
        ("row", ["4", "LTE", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd", "Whether SIB5 carries priority", "Without CFG: no reselection to that freq", "§5.1.3.1"]),
        ("row", ["5", "LTE", "EUTRANINTERNFREQ", "CellReselPriority", "SIB5 target priority", "CfgInd=CFG", "§5.1.3.1"]),
        ("row", ["6", "LTE", "EUTRANINTERNFREQ", "QRxLevMin / QqualMin / Pmax", "Target-frequency suitability", "UL/PRACH still required for a usable camp", "§5.1.3.4"]),
        ("row", ["7", "LTE", "CELLRESEL", "SIntraSearchCfgInd / SIntraSearch / Q", "Skip intra meas when serving very good", "CFG recommended", "§5.1.3.3"]),
        ("row", ["8", "LTE", "CELLRESEL", "SNonIntraSearchCfgInd / SNonIntraSearch / Q", "Start equal/lower inter-freq search", "Example 10; SIntra > SNonIntra", "§5.4.1.1"]),
        ("row", ["9", "LTE", "EUTRANINTERNFREQ", "ThreshXhigh / ThreshXhighQ", "Higher-priority target qualification", "EutranReselTime must persist", "Tables 5-1/5-2"]),
        ("row", ["10", "LTE", "CELLRESEL", "ThrshServLow / ThrshServLowQ", "Permission to leave for lower priority", "Align with connected A2/A5", "Tables 5-3/5-4"]),
        ("row", ["11", "LTE", "EUTRANINTERNFREQ", "ThreshXlow / ThreshXlowQ", "Lower-priority target qualification", "Target must be actually usable", "Tables 5-3/5-4"]),
        ("row", ["12", "LTE", "CELLRESEL", "Qhyst / TReselEutran", "Serving stickiness / intra timer", "Small timer = ping-pong", "§5.1.3.4"]),
        ("row", ["13", "LTE", "EUTRANINTERNFREQ", "QoffsetFreq / EutranReselTime", "Equal-prio freq offset / inter timer", "Positive offset reduces target rank", "§5.1.3.4"]),
        ("row", ["14", "LTE", "EUTRANINTERFREQNCELL", "CellQoffset", "Per-neighbor idle offset", "Affects which 16 neighbors are listed", "§5.1.3.4"]),
        ("row", ["15", "LTE", "EUTRANINTERNFREQ", "MeasPerformanceDemand", "NORMAL / REDUCED / UNDELIVER", "UNDELIVER blocks idle MLB target", "§5.1.3.3"]),
        ("row", ["16", "LTE", "CELLRESEL", "SpeedDepReselCfgInd + SfMedium/High", "Speed-based timer/hyst scaling", "Not a tool for static layer imbalance", "§5.1.3.6"]),
        ("row", ["17", "LTE", "InterFreqBlkCell", "ApplicationScope", "Idle/connected blacklist scope", "Not a load tool", "§5.1.3.2"]),
        ("row", ["18", "LTE", "RRCCONNSTATETIMER", "T320ForLoadBalance / T320ForOther", "Dedicated-priority lifetime", "SPID/PCC always 180 min", "§5.1.3.1"]),
        ("row", ["19", "LTE", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch / InterFreqMlbSwitch", "Idle MLB enable", "License + MlbTargetInd", "§5.3.2.3"]),
        ("row", ["20", "LTE", "EnhancedMlbAlgoSwitch", "DediPrioManageOnLowLoadSw / EnhSw", "Hold low-load dedicated prio", "T320", "MLB Table 5-10"]),
        ("row", ["21", "LTE", "EUTRANINTERNFREQ", "MlbTargetInd", "Allowed / without idle MLB / without connect MLB", "Coverage HO is separate from MLB targeting", "MLB pp.28, 129"]),
        ("row", ["22", "LTE", "CELLMLB", "InterFreqIdleMlbUeNumThd / idle transfer type", "Idle user-number trigger", "See Intra-RAT MLB sheet", "MLB Table 5-2"]),
        ("row", ["23", "LTE", "EUTRANINTERNFREQ", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Filtered from dedicated-priority delivery", "Document: recommended only for SCC-only frequencies", "§5.1.3.1"]),
        ("row", ["24", "LTE", "GlobalProcSwitch", "CellReselectionOptSwitch", "Preferential LTE frequency delivery in dedicated prio", "Idle MLB path", "§5.3.2.3"]),
        ("space", 8),

        ("section", "SN-11  Final MML Command for activations  (maintain sequence)"),
        ("note", "SN-11 columns A–G are the original template: Parameter Sequence | MO | Activation Value | Conditional Parameter | Remarks | Parameter Description | More Notes. Column H is Chart / Doc Ref. Replace LocalCellId / DlEarfcn. Validate in MAE."),
        ("heads", H_MML),
        ("row", ["0", "—",
                 "LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLALGOSWITCH: LocalCellId=<x>; LST RRCCONNSTATETIMER:;",
                 "Read-only", "Dump current before any MOD", "Baseline", "Keep LST output with the change record"]),
        ("row", ["1", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<prio>, SNonIntraSearchCfgInd=CFG, SIntraSearchCfgInd=CFG, SNonIntraSearch=<val>;",
                 "SIntraSearch > SNonIntraSearch", "Huawei example SNonIntraSearch=10. Capacity/hotspot priority above coverage layer.",
                 "Serving common priority + search start", "One cell/layer per command"]),
        ("row", ["2", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;",
                 "EARFCN exists in the SIB5 plan", "Repeat for every non-serving frequency that must be reselectable.",
                 "SIB5 priority + visibility", "Do not use UNDELIVER on a frequency that idle MLB must use"]),
        ("row", ["3", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, ThreshXhigh=<val>, ThreshXlow=<val>, QoffsetFreq=<val>, EutranReselTime=<val>;",
                 "Matches higher vs lower vs equal-priority design", "No universal dBm in the feature book. Calibrate from MR.",
                 "Reselection qualification and equal-prio offset", "Timers must persist"]),
        ("row", ["4", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<val>;",
                 "Align with connected coverage A2/A5", "Lower-priority fallback permission",
                 "Serving-low", "Do not set so low that a dying serving cell never yields"]),
        ("row", ["5", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=<ALLOWED | ALLOWED_WITHOUT_IDLE_MLB | ALLOWED_WITHOUT_CONNECT_MLB>;",
                 "Verify exact enum on the NE. Coverage NoHoFlag remains PERMIT if coverage HO is required.",
                 "Stops a frequency being an idle and/or connected MLB target without blocking ordinary coverage HO.",
                 "MLB target indication", "pp.28, 129"]),
        ("row", ["6", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "Intra-RAT MLB license; MlbTargetInd already set; SIB5 NORMAL",
                 "Enable idle+connected MLB bits. Leave InterFreqBlindMlbSwitch-0 unless containment is proven.",
                 "Master idle MLB bit", "Do this after steps 1–5"]),
        ("row", ["7", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
                 "Idle MLB ON", "SPID/PCC path stays 180 minutes regardless of this parameter.",
                 "Dedicated-priority lifetime", "§5.1.3.1"]),
        ("row", ["8", "Verify",
                 "LST the MOs above. Wait the next SI modification period. Check L.RRCRel.load.DedicatedPri.LTE.High and L.RRCRel.Lowload.DedicatedPri.LTE.High.",
                 "SI is not instant [IM] §7.1.3", "Do not judge camping in the same 15-min as the change.",
                 "Idle dedicated-priority counters (MLB book)", "Close the sequence"]),
    ]


# ===========================================================================
# CONNECTED — abbreviated structure same pattern, quality rows
# ===========================================================================
def connected_blocks():
    return [
        ("section", "Process chart from document  (right-side chart for SN-1)"),
        ("flow", ["Initiate HO function", "Meas or blind", "Deliver meas config", "UE report A1–A5", "Pick target + admit", "Execute HO", "Fig 4-1"]),
        ("note", "Connected Mode Fig 4-1 §§4.1.1–4.1.8 pp.29–65. Necessary coverage HO preempts unnecessary load/optimization HO (Table 4-5). Full MLB algorithm is not in this book — see Intra-RAT MLB."),
        ("space", 8),
        ("section", "SN list"),
        ("heads", H_SN),
        ("row", ["1", "LTE", "Working Principal", "Measurement-based HO using A1–A5; blind only if containment known", "Meas flags / NRT", "See SN-1", "Fig 4-1"]),
        ("row", ["2", "LTE", "Major highlighted Point", "RSRP preferred; A4>A2; A4 TTT 5120 ms disables FreqPri; example dBm are not design values", "—", "See MAJOR", "Table 4-9, 4-15, §11.4.1.2"]),
        ("row", ["3", "LTE", "Benefit and Limitations", "Coverage protection vs not a load algorithm; gaps steal TTI; unnecessary HO admits all QCIs", "—", "See SN-3", "Tables 4-16/4-17"]),
        ("row", ["4", "LTE", "Selection criteria / Trigger", "A2 coverage; A1/A4 FreqPri; A4/A5 MLB execution; UE filters in MLB book", "Correct A2 family", "See SN-4", "Ch.5 / Ch.11"]),
        ("row", ["5", "LTE", "Activation parameter / Switch", "A1/A2/A4/A5 groups, meas flags, object cap, FreqPri vs MLB switches", "License + NRT", "See SN-5", "Ch.4–5, 11"]),
        ("row", ["6", "LTE", "Prerequisite functions", "NRT, object cap, SMeasure, gap, admission, CA PCC", "—", "See SN-6", "Tables 4-3/4-4"]),
        ("row", ["7", "LTE", "Mutually impacted", "FreqPri reverse MLB pair, A4 vs A2, protect timer, decouple SW", "—", "See SN-7", "p.303"]),
        ("row", ["8", "LTE", "Relation with Other Feature", "Idle, MLB, CA, VoLTE QCI, ES", "—", "See SN-8", "§3"]),
        ("row", ["9", "LTE", "License Requirements", "Basic coverage, FreqPri package, MLB A4, MLB A5, blind", "MAE", "See SN-9", "pp.142–143 MLB"]),
        ("row", ["10", "LTE", "All Parameter List", "Object → filter → A1/A2 → A3/A4/A5 → FreqPri → punish", "Connected order", "See SN-10", "—"]),
        ("row", ["11", "LTE", "Final MML Command", "LST → flags → object cap → coverage → A4 → event type → FreqPri coord → verify", "Sequence", "See SN-11", "MML"]),
        ("space", 8),

        ("section", "SN-1  Working Principal  (connected order)"),
        ("heads", H_STEP),
        ("row", ["1", "LTE", "Handover-function initiation", "Coverage / FreqPri / MLB / service / quality", "Classification: necessary vs unnecessary offload vs unnecessary optimization", "Table 3-1; pp.26–28", "pp.21–28"]),
        ("row", ["2", "LTE", "Choose measurement-based or blind", "Fig 4-2", "Blind only when immediate mobility is required and neighbor contains serving", "Higher access-failure risk", "pp.31–32"]),
        ("row", ["3", "LTE", "Deliver measurement configuration", "EUTRANINTERNFREQ object + report config", "FREQ_MEAS_FLAG; not HO_TRG_FREQ_FORBID; object count limits", "Equal-priority objects may be picked randomly if over cap", "Tables 4-1 to 4-4, p.125"]),
        ("row", ["4", "LTE", "Event A1 — serving good", "A1: Ms−Hys > Thresh for TTT", "Stops coverage meas; can start A1-based FreqPri", "Table 4-8", "pp.41–46"]),
        ("row", ["5", "LTE", "Event A2 — serving poor", "A2: Ms+Hys < Thresh for TTT", "Starts inter-freq coverage meas. Separate families for A3 / A4-A5 / IRAT / blind", "Tables 5-1, 5-3, 5-10", "pp.97–100"]),
        ("row", ["6", "LTE", "Event A3 — neighbor relatively better", "Mn+Ofn+Ocn−Hys > Ms+Ofs+Ocs+A3Offset", "Ofn/Ofs frequency offsets; Ocn/Ocs CIO", "Table 5-16", "p.127"]),
        ("row", ["7", "LTE", "Event A4 — neighbor absolutely good", "Mn+Ofn+Ocn−Hys > Thresh", "Used by MLB and FreqPri (target good enough, need not beat serving)", "Table 5-22; Table 11-5; MLB p.137", "p.145"]),
        ("row", ["8", "LTE", "Event A5 — serving poor AND neighbor good", "Ms+Hys<Th1 AND Mn+Ofn+Ocn−Hys>Th2", "Strongest coverage-protection semantics", "Tables 5-18/5-19", "pp.129–132"]),
        ("row", ["9", "LTE", "Target pick + admission", "Best filtered target; necessary = any QCI; unnecessary = all QCIs", "Inter-eNB unnecessary HO does not immediately try next target after admit fail", "Tables 4-16/4-17", "pp.61–65"]),
        ("row", ["10", "LTE", "Execute HO then punish/retry on fail", "Res/Opt/NonRes punish timers and counts", "Do not treat every prep fail as an RF-threshold problem", "Tables 4-16/4-17", "pp.61–65"]),
        ("space", 8),

        ("section", "SN-2  Major highlighted Point"),
        ("major", "Huawei recommends RSRP as the general trigger quantity. RSRQ fluctuates with load. [CM] Table 4-15 pp.55–56"),
        ("major", "A4 target threshold must be better than the relevant coverage A2, otherwise ping-pong. [CM] Table 5-22 pp.147–148; Table 11-8 pp.315–316"),
        ("major", "InterFreqHoA4TimeToTrig = 5120 ms disables frequency-priority, CQI and service-based inter-frequency HO. [CM] Table 4-9 p.48"),
        ("major", "MML examples such as A1/A2 −85/−87 dBm and A4 −103 dBm are command examples, not design values. [CM] §11.4.1.2 pp.317–318"),
        ("major", "Equal-priority frequencies may be selected randomly for measurement objects — not load balance. [CM] §5.3.1.2 p.125"),
        ("major", "A frequency that is a FreqPri target must not have a reverse MLB-target relationship. Ping-pong warning. [CM] p.303"),
        ("major", "MlbBasedFreqPriHoSwitch lets MLB own heavy-load decisions when specified MLB functions are active. [CM] Table 11-7 p.313"),
        ("space", 8),

        ("section", "SN-3  Benefit and Limitations"),
        ("heads", ["Type", "RAT", "Item", "Document statement", "Condition", "User impact", "Chart / Doc Ref"]),
        ("row", ["Benefit", "LTE", "Coverage protection", "A2→A3/A4/A5 avoids drop", "Necessary HO priority", "Indoor / edge rescue", "Table 4-5"]),
        ("row", ["Benefit", "LTE", "A4 offload gate", "Target need only be good enough", "MLB / FreqPri", "Can leave a still-usable source", "Table 5-22"]),
        ("row", ["Benefit", "LTE", "A5 dual condition", "Serving poor AND target good", "Coverage fallback", "Strongest protection semantics", "Tables 5-18/5-19"]),
        ("row", ["Benefit", "LTE", "FreqPri + A1", "Service on high band, low band kept for coverage", "Fig 11-1/11-2", "Coverage-layer protection intent", "pp.299–300"]),
        ("row", ["Limitation", "LTE", "Not a load algorithm", "Equal static priority ≠ instantaneous load share", "Need MLB book", "Random object pick among equal prio", "p.125"]),
        ("row", ["Limitation", "LTE", "Measurement gaps", "Gap pattern steals DL TTIs", "Old UE / VoLTE", "Prefer A1/A2 gated meas", "Fig 4-10"]),
        ("row", ["Limitation", "LTE", "Admission", "Unnecessary/offload HO must admit ALL QCIs", "VoLTE + data", "Prep fail often admission, not RF", "Tables 4-16/4-17"]),
        ("row", ["Limitation", "LTE", "Blind HO", "No candidate measurement", "Containment required", "Higher access failure", "Fig 4-2"]),
        ("space", 8),

        ("section", "SN-4  Selection criteria / Trigger Condition"),
        ("heads", ["Rule", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration", "Chart / Doc Ref"]),
        ("row", ["Coverage", "LTE", "A2 family then A3/A4/A5", "Serving crosses A2", "CovBasedInterFreqHoMode = IMMEDIATE / SIGNAL / FREQPRIORITY", "Must preempt load HO", "§5.3.1"]),
        ("row", ["Frequency priority", "LTE", "A1 / good serving + A4 target", "Place service on high-priority frequency", "LoadTriggerFreqPriHoSwitch: overlap, load info, neighbor not UE-number MLB triggered, no PCI conflict", "Waiting timer can wait for highest-priority freq", "pp.302–303, 308–309"]),
        ("row", ["MLB HO execution", "LTE", "MlbInterFreqHoEventType A4 or A5", "MLB source trigger already true", "A5 needs non-cosited MLB license", "UE pick is in the MLB book", "MLB Table 6-3"]),
        ("row", ["Stop coverage meas", "LTE", "A1", "Serving recovers", "ReduceInvalidA1A2RptSigSwitch can cut signalling", "Keep A1/A2 hyst consistent", "pp.111–115"]),
        ("row", ["Blind", "LTE", "BlindHoPriority", "Immediate mobility + neighbor contains serving", "No RF measurement of candidate", "Use only with known full overlap", "Table 5-22 p.149"]),
        ("row", ["High mobility", "LTE", "FREQ_PRI_HO_FORBID_SW / MLB_HO_FORBID_SW", "UE detected above ~30 km/h", "No FreqPri/MLB meas", "Near-NLOS can misclassify", "pp.304–305; MLB pp.136–141"]),
        ("space", 8),

        ("section", "SN-5  Activation parameter / Switch   (core table)"),
        ("heads", H_ACT),
        ("row", ["5.1", "LTE", "INTERFREQHOGROUP", "A1/A2 Hyst + TTT + correct A2 family", "Calibrate from MR. Not the MML example dBm.", "Filter coefficient also adds delay", "Tables 5-3, 5-10, 4-9"]),
        ("row", ["5.2", "LTE", "INTERFREQHOGROUP", "A3 offset / hyst / TTT", "Relative HO among similar coverage", "QoffsetFreqConn + CIO", "Table 5-16"]),
        ("row", ["5.3", "LTE", "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp/Rsrq + A4 Hyst/TTT", "A4 better than coverage A2. TTT ≠ 5120 ms if FreqPri required.", "IfMlbThdRsrpOffset / FreqPriHoA4ThldRsrpOffset", "Table 11-5; Table 4-9 p.48"]),
        ("row", ["5.4", "LTE", "INTERFREQHOGROUP", "A5 Thd1 / Thd2 (+ MLB A5 Thd1)", "Serving poor AND target good", "A5 MLB needs extra license", "Tables 5-18/5-19; MLB Table 6-3"]),
        ("row", ["5.5", "LTE", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType", "A4 for co-sited FDD; A5 if licensed non-cosited", "MLB license", "MLB §6.1.1.5.2"]),
        ("row", ["5.6", "LTE", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG / HO_TRG_FREQ_FORBID_MEAS_FLAG", "Select meas; deselect forbid for required HO targets", "Object cap must be enough", "§4.1.4.1.2"]),
        ("row", ["5.7", "LTE", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "≥ number of needed inter-freq objects", "Else random missing frequency", "Tables 4-3/4-4"]),
        ("row", ["5.8", "LTE", "FreqPri switches", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch / ReduceInvalidFreqPriHoSwitch", "ON when MLB should own heavy load", "No reverse MLB target on a FreqPri pair", "Table 11-7; p.303"]),
        ("row", ["5.9", "LTE", "INTRARATHOCOMM", "FreqPriInHoProtectionTimer / FreqPriIFHoWaitingTimer / CovBasedIfHoWaitingTimer", "Non-zero protect after incoming unnecessary HO", "Stops bounce-back", "p.300, 308–309"]),
        ("row", ["5.10", "LTE", "HOMEASCOMM", "SMeasure", "UE may skip inter-freq meas while serving RSRP above SMeasure", "Can silently suppress A4", "§4.1.5 p.55"]),
        ("space", 8),

        ("section", "SN-6  Prerequisite functions"),
        ("heads", H_PRE),
        ("row", ["6.1", "LTE", "EUTRANINTERFREQNCELL", "Symmetric NRT, NoHoFlag, no PCI conflict", "Required", "Looks like a bad A4 if missing", "§4.1.4"]),
        ("row", ["6.2", "LTE", "CELLUEMEASCONTROLCFG", "Object capacity ≥ needed freqs", "Required", "Some frequencies never measured", "Tables 4-3/4-4"]),
        ("row", ["6.3", "LTE", "HOMEASCOMM", "SMeasure compatible with intended A4/FreqPri", "Verify RRC", "A4 never delivered", "§4.1.5"]),
        ("row", ["6.4", "LTE", "Gap pattern", "AutoGapSwitch / GapPatternType", "Acceptable TTI cost", "VoLTE / old UE TP drop", "Fig 4-10"]),
        ("row", ["6.5", "LTE", "Admission / X2", "HoAdmitSwitch / X2RoHoAdmitSwitch", "Healthy", "Unnecessary HO needs all-QCI admit", "Tables 4-16/4-17"]),
        ("row", ["6.6", "LTE", "CA book", "PCC anchoring vs FreqPri", "Do not fight PCC policy", "Ping-pong §11.3.2.3", "p.312"]),
        ("space", 8),

        ("section", "SN-7  Mutually impacted"),
        ("heads", H_IMP),
        ("row", ["7.1", "LTE", "FreqPri target A→B", "MLB target B→A", "Forbidden — ping-pong", "Remove reverse MLB target", "p.303"]),
        ("row", ["7.2", "LTE", "A4 threshold", "Coverage A2", "A4 must be better than A2", "Ping-pong if too close", "Table 5-22"]),
        ("row", ["7.3", "LTE", "Incoming unnecessary HO", "Immediate FreqPri re-meas", "Use FreqPriInHoProtectionTimer", "Bounce-back", "p.300"]),
        ("row", ["7.4", "LTE", "Coverage meas already delivered", "MLB A4 meas", "LOAD_COVERAGE_MEAS_DECOUPLE_SW allows load meas after coverage meas", "MLB meas blocked otherwise", "MLB p.139"]),
        ("row", ["7.5", "LTE", "Large CIO / QoffsetFreqConn", "RF problem", "Can mask overshoot", "Fix RF first", "pp.46–48"]),
        ("row", ["7.6", "LTE", "Virtual-grid smart carrier selection", "FreqPri HO", "Ping-pong if concurrent", "Do not combine casually", "p.312"]),
        ("space", 8),

        ("section", "SN-8  Relation with Other Feature"),
        ("heads", H_IMP),
        ("row", ["8.1", "LTE", "Idle Mode", "Dedicated prio discarded at connect", "Align ThreshXhigh↔A1 and ThrshServLow↔A5", "Next access vs session", "IM §5.1.3.1"]),
        ("row", ["8.2", "LTE", "Intra-RAT MLB", "This book executes A4/A5 HO", "MLB decides who/when. ONLY_STRONGEST_CELL is in MLB book.", "HO engine vs load algorithm", "MLB Table 6-5"]),
        ("row", ["8.3", "LTE", "Carrier Aggregation", "PCC anchoring in CA book", "PCC move ≠ SCC traffic", "§3 p.23; §11.3.2.3", "CA document"]),
        ("row", ["8.4", "LTE", "VoLTE / QCI", "QCI-specific hyst/TTT ; unnecessary HO admits all QCIs", "Never offload if target cannot admit all QCIs", "Tables 4-16/4-17", "QCI"]),
        ("row", ["8.5", "LTE", "Energy saving", "Sleeping cell must not remain FreqPri/MLB target", "Coordinate ES", "MLB function-impact tables", "MLB pp.153–156"]),
        ("space", 8),

        ("section", "SN-9  License Requirements"),
        ("heads", ["Function", "RAT", "Switch / function", "License (verify in MAE)", "If missing", "Notes", "Chart / Doc Ref"]),
        ("row", ["Coverage A1–A5", "LTE", "INTERFREQHOGROUP", "Basic LTE mobility", "Usually included", "Confirm lean package", "—"]),
        ("row", ["Frequency-priority HO", "LTE", "FreqPri Ch.11", "FreqPri / service-based mobility package — verify name", "A1 high-band steering does not run", "Fig 11-1/11-2", "Ch.11"]),
        ("row", ["MLB A4", "LTE", "Intra-RAT MLB", "Intra-RAT Mobility Load Balancing", "No load HO", "See MLB sheet", "MLB"]),
        ("row", ["MLB event A5", "LTE", "MlbInterFreqHoEventType=A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay on A4", "pp.142–143", "MLB Table 6-3"]),
        ("row", ["Blind HO", "LTE", "BlindHoPriority", "Blind option if sold separately", "Blind unavailable", "Higher access risk", "Table 5-22"]),
        ("space", 8),

        ("section", "SN-10  All Parameter List  (sequence / connected order)"),
        ("heads", H_PAR),
        ("row", ["1", "LTE", "EUTRANINTERNFREQ", "DlEarfcn / MeasBandWidth / QoffsetFreqConn", "Measurement object", "NRT must exist", "Table 4-2"]),
        ("row", ["2", "LTE", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG / HO_TRG_FREQ_FORBID_MEAS_FLAG / INTER_FREQ_FILTER_FLAG", "Frequency filtering", "Silent no-HO if wrong", "§4.1.4.1.2"]),
        ("row", ["3", "LTE", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "How many objects can be delivered", "Random drop if over limit", "Tables 4-3/4-4"]),
        ("row", ["4", "LTE", "HOMEASCOMM", "SMeasure", "Skip inter-freq meas when serving strong", "Can suppress A4", "§4.1.5"]),
        ("row", ["5", "LTE", "CELLHOPARACFG", "EutranFilterCoeffRsrp / Rsrq", "L3 filter", "Over-smooth delays rescue", "Table 4-14"]),
        ("row", ["6", "LTE", "ENODEBALGOSWITCH", "AutoGapSwitch / GapPatternType / DedicatedGapPatternType", "Measurement gap", "Steals TTI", "Fig 4-10"]),
        ("row", ["7", "LTE", "INTERFREQHOGROUP", "InterFreqHoA1A2Hyst / TimeToTrig", "A1/A2 stability", "QCI-specific optional", "Table 4-9"]),
        ("row", ["8", "LTE", "INTERFREQHOGROUP", "Coverage A2 threshold families", "A3 vs A4/A5 vs IRAT vs blind", "Wrong family = wrong HO", "Tables 5-3, 5-10"]),
        ("row", ["9", "LTE", "INTERFREQHOGROUP", "A3 offset / hyst / TTT / A3RsrqOffset", "Relative HO", "Ofn/Ocn apply", "Table 5-16"]),
        ("row", ["10", "LTE", "EUTRANINTERFREQNCELL", "CellIndividualOffset", "Connected CIO (Ocn)", "Large CIO masks RF", "pp.46–48"]),
        ("row", ["11", "LTE", "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp/Rsrq + A4 Hyst/TTT", "MLB/FreqPri A4", "TTT 5120 ms disables FreqPri", "Tables 11-5, 4-9"]),
        ("row", ["12", "LTE", "EUTRANINTERNFREQ", "IfMlbThdRsrpOffset / FreqPriHoA4ThldRsrpOffset", "Per-frequency A4 offset", "Plus operator/QCI offset", "Table 11-5; MLB p.137"]),
        ("row", ["13", "LTE", "INTERFREQHOGROUP", "A5 Thd1/Thd2 / Mlb A5 Thd1", "Dual-condition HO", "A5 MLB license", "Tables 5-18/5-19"]),
        ("row", ["14", "LTE", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType", "A4 or A5 for FDD MLB", "Table 6-3", "MLB §6.1.1.5.2"]),
        ("row", ["15", "LTE", "CELLALGOSWITCH", "CovBasedInterFreqHoMode + waiting timer", "Coverage execution timing", "Fig 5-3–5-7", "§5.3.1"]),
        ("row", ["16", "LTE", "INTRARATHOCOMM", "FreqPriIFHoWaitingTimer / FreqPriInHoProtectionTimer", "Wait / incoming protect", "p.300, 308–309", "Ch.11"]),
        ("row", ["17", "LTE", "FreqPri switches", "MlbBasedFreqPriHoSwitch / LoadTrigger / ReduceInvalid", "MLB owns heavy load", "No reverse target", "Table 11-7"]),
        ("row", ["18", "LTE", "CELLOPHOCFG", "HighMobiUeHoForbidSw / FREQ_PRI_HO_FORBID_SW", "No MLB/FreqPri if ~30 km/h", "NLOS misclassify", "pp.304–305"]),
        ("row", ["19", "LTE", "HOMEASCOMM", "HO fail punish timers / counts", "Retry vs penalty", "Prep fail ≠ RF", "Tables 4-16/4-17"]),
        ("row", ["20", "LTE", "EUTRANINTERFREQNCELL", "BlindHoPriority / InterFreqMlbBlindHo", "Blind path", "Containment required", "Table 5-22 p.149"]),
        ("row", ["21", "LTE", "RATFREQPRIORITYGROUP", "CovIFHo RSRP/RSRQ Hyst/TTT", "Per-freq/QCI coverage values", "Do not copy VoLTE to all QCI", "Table 4-9"]),
        ("space", 8),

        ("section", "SN-11  Final MML Command for activations  (maintain sequence)"),
        ("heads", H_MML),
        ("row", ["0", "—",
                 "LST INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST EUTRANINTERFREQNCELL: LocalCellId=<x>; LST CELLHOPARACFG: LocalCellId=<x>; LST HOMEASCOMM:;",
                 "Read-only", "Dump first. Do not paste example −85/−87/−103 dBm as live values.", "Baseline", "§11.4.1.2"]),
        ("row", ["1", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>; FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for required HO targets;",
                 "NRT exists", "Audit flags before any dBm change. Most silent A4 failures are flags/NRT.", "Meas eligibility", "§4.1.4.1.2"]),
        ("row", ["2", "CELLUEMEASCONTROLCFG",
                 "MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<n>, MaxEutranFddMeasFreqNum=<n>;",
                 "n ≥ number of inter-frequency objects this cell must measure", "Otherwise equal-priority frequencies are randomly dropped.", "Object cap", "p.125"]),
        ("row", ["3", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; set coverage A1/A2/A5 from MR (correct A2 family);",
                 "A2 family matches intended HO type (A3 vs A4/A5 vs blind)", "Coverage rescue first.", "Coverage A1/A2/A5", "Tables 5-3, 5-10"]),
        ("row", ["4", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<calib>, InterFreqHoA4Hyst=<hyst>, InterFreqHoA4TimeToTrig=<ttt>;",
                 "A4 better than coverage A2. TTT ≠ 5120 ms if FreqPri/MLB A4 is required.", "Main absolute target gate for MLB/FreqPri.", "A4", "Table 4-9 p.48"]),
        ("row", ["5", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbInterFreqHoEventType=A4;",
                 "A5 only with Intra-LTE Load Balancing for Non-cosited Cells license", "Co-sited FDD typically A4.", "MLB event type", "MLB Table 6-3"]),
        ("row", ["6", "FreqPri related MO",
                 "Enable MlbBasedFreqPriHoSwitch and LoadTriggerFreqPriHoSwitch when MLB is used; keep FreqPriInHoProtectionTimer non-zero; no reverse MLB target on a FreqPri pair;",
                 "MLB feature active; confirm exact MO name in MAE", "Stops FreqPri fighting MLB.", "Coordination", "p.303; Table 11-7"]),
        ("row", ["7", "Verify",
                 "L.HHO.InterFreq.Coverage.Prep/Exec/Succ; L.HHO.InterFreq.FreqPri.*; L.RRC.ReEst.ReconfFail.Att. PrepSucc=ExecAtt/PrepAtt; ExecSucc=ExecSucc/ExecAtt.",
                 "Busy-hour, pair-level", "Table 5-24 p.153; Tables 11-9/11-10 pp.318–319.", "Counters", "Close the sequence"]),
    ]


# ===========================================================================
# MLB
# ===========================================================================
def mlb_blocks():
    return [
        ("section", "Process chart from document  (right-side chart for SN-1)"),
        ("flow", ["Eval load N/C", "Trigger thd+offset", "Admit target", "Select UEs", "A4/A5 HO or idle release", "Penalty / stop", "Fig 3-1"]),
        ("note", "MLB Fig 3-1 p.15. Equalisation vs offload: Fig 4-1 p.16. Idle vs connected transfer: Figs 4-4/4-5 pp.18–19. Load = N/C. Normalised difference = (Load_s−Load_t)/Load_s."),
        ("space", 8),
        ("section", "SN list"),
        ("heads", H_SN),
        ("row", ["1", "LTE", "Working Principal", "Equalise or offload inter-frequency LTE load via HO or idle dedicated prio", "License + NRT + load exchange", "See SN-1", "Fig 3-1, Fig 4-1"]),
        ("row", ["2", "LTE", "Major highlighted Point", "ActiveUe+SE for unequal BW; ONLY_STRONGEST_CELL; 5s+MaxUe≥40 over-transfer; PRB MLB skips CA UEs; 7-day smart thd", "—", "See MAJOR", "Table 5-5, 6-5, p.136"]),
        ("row", ["3", "LTE", "Benefit and Limitations", "True load equalisation vs cannot fix RF/CA scheduler; blind risk; ES interaction", "—", "See SN-3", "§6.1.2.2 p.141"]),
        ("row", ["4", "LTE", "Selection criteria / Trigger", "UE-number / PRB thd; target admit; UE filters; volume; freq pick; A4/A5", "Table 6-2", "See SN-4", "§6.1.1"]),
        ("row", ["5", "LTE", "Activation parameter / Switch", "MlbAlgoSwitch, trigger mode, eval SW, strategy, MlbTargetInd, event type", "License first", "See SN-5", "Table 6-2"]),
        ("row", ["6", "LTE", "Prerequisite functions", "License, X2 load, overlap, HO SR, HW/transport, A4 delivery, SIB5 if idle", "—", "See SN-6", "pp.27–29"]),
        ("row", ["7", "LTE", "Mutually impacted", "Fixed-proportion idle, PRB exclusive modes, FreqPri reverse, ES, CA, smart relearn", "—", "See SN-7", "§5.4.2.2"]),
        ("row", ["8", "LTE", "Relation with Other Feature", "Idle method, Connected HO engine, CA PCC, ES, SON smart thd, inter-RAT separate", "—", "See SN-8", "—"]),
        ("row", ["9", "LTE", "License Requirements", "Intra-RAT MLB; A5 non-cosited; blind; CA transfer; SON thd", "MAE", "See SN-9", "pp.142–143"]),
        ("row", ["10", "LTE", "All Parameter List", "Switch → model → trigger → target → UE pick → event → idle T320", "Connected order", "See SN-10", "Ch.8 → parameter reference"]),
        ("row", ["11", "LTE", "Final MML Command", "LST → MlbTargetInd → eval SW → CELLMLB mode/strategy → thd/volume → CA → master bit → T320 → verify", "Sequence", "See SN-11", "MML"]),
        ("space", 8),

        ("section", "SN-1  Working Principal  (connected order)"),
        ("heads", H_STEP),
        ("row", ["1", "LTE", "Choose equalisation or offload", "Peer load available → equalisation; else offload", "Fig 4-1 p.16", "Equalisation is safer when X2/intra-eNB load exists", "§4.1"]),
        ("row", ["2", "LTE", "Choose load indicator", "UE number vs PRB vs HW/transport Low/Med/High/OverLoad", "Fig 4-2, Fig 4-3", "UE number ≈ short non-GBR; PRB ≈ GBR/sustained", "§4.2"]),
        ("row", ["3", "LTE", "Compute user-number load", "Load = N/C ; diff = (Load_s−Load_t)/Load_s", "ActiveUeBasedLoadEvalSw uses DL-buffer UEs; SpectralEffBasedLoadEvalSw uses PRB, scale, GBR, SE", "SE refresh every 1 min if ≥10 UL-sync UEs", "§5.1.1 pp.24–27"]),
        ("row", ["4", "LTE", "Trigger connected UE-number MLB", "N ≥ InterFreqMlbUeNumThd + MlbUeNumOffset for MlbTrigJudgePeriod", "Stop when N < InterFreqMlbUeNumThd", "Table 6-2: InterFreqMlbSwitch + UE_NUMBER_ONLY + SynchronizedUE", "p.127–128"]),
        ("row", ["5", "LTE", "Admit target cell/frequency", "NoHoFlag, not blacklisted, no PCI conflict, HO SR, meas flags, OverlapInd, MlbTargetInd, Low/Med HW+transport, not punished", "WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB block one mode", "pp.27–29, 129", "§5.1.1.2"]),
        ("row", ["6", "LTE", "Select UEs", "UL-sync, not emergency, SPID/QCI, eMBMS, protect/punish timers, optional ARP/PRB/MCS/SNR", "ONLY_STRONGEST_CELL recommended", "Else immediate coverage HO bounce", "pp.130–140; Table 6-5 p.159"]),
        ("row", ["7", "LTE", "Limit volume and pick frequency", "min(needed delta, hyst, MlbMaxUeNum); FAIR / PRIORITY / LOADPRIORITY", "LoadTransferEnhSw changes multi-target math", "5 s eval + MlbMaxUeNum≥40 over-transfers", "pp.134–137"]),
        ("row", ["8", "LTE", "Execute transfer", "FDD: measurement-based HO A4 or A5; optional blind; idle: RRC release + dedicated prio + T320", "A5 needs extra license", "Redirection generally not used in FDD", "§6.1.1.5.2 p.136"]),
        ("row", ["9", "LTE", "Idle transfer (if enabled)", "Dedicated prio: NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "T320ForLoadBalance", "Affects next session only", "§5.1.1.5 p.33"]),
        ("space", 8),

        ("section", "SN-2  Major highlighted Point"),
        ("major", "Huawei recommends ActiveUeBasedLoadEvalSw when MLB frequencies have different bandwidths, plus SpectralEffBasedLoadEvalSw when SE differs significantly (e.g. >30%). [MLB] Table 5-5 pp.47–48"),
        ("major", "Raw UE-count balancing on unequal bandwidth can reduce DL throughput. [MLB] §6.1.2.2 p.141"),
        ("major", "MlbHoCellSelectStrategy = ONLY_STRONGEST_CELL is recommended. Otherwise the UE is coverage-HO’d back. [MLB] Table 6-5 p.159"),
        ("major", "InterFreqLoadEvalPrd = 5 s together with MlbMaxUeNum ≥ 40 can transfer excessive UEs. [MLB] p.136"),
        ("major", "PRB-usage MLB does not transfer CA UEs, is burst-sensitive, does not guarantee experience fairness, and loses gain when CA penetration ≳ 60%. [MLB] Fig 6-2, §6.5 pp.207–215"),
        ("major", "NCellTrigThldSmartOptAlgoSw: 7 days collect, then calculate, refresh every 7 days. Aggressive first-week seeds: extra CPU/HO and up to 5% TP fluctuation. [MLB] pp.29–31, 162, 224–225"),
        ("major", "Fixed-proportion idle + user-number connected MLB = ping-pong. Adaptive-proportion idle is not recommended without Huawei support. [MLB] §5.4.2.2, §5.5.2.1"),
        ("space", 8),

        ("section", "SN-3  Benefit and Limitations"),
        ("heads", ["Type", "RAT", "Item", "Document statement", "Condition", "User impact", "Chart / Doc Ref"]),
        ("row", ["Benefit", "LTE", "Equalisation", "Uses source and target load to reduce difference", "Load exchange available", "Safer than blind offload", "Fig 4-1"]),
        ("row", ["Benefit", "LTE", "BW / SE aware", "C includes PRB, scale, GBR, measured SE when switches ON", "ActiveUe + SpectralEff", "Needed for 15 vs 20 MHz", "Table 5-5"]),
        ("row", ["Benefit", "LTE", "Idle + connected pair", "Idle steers next access; connected moves now", "T320 vs HO", "Idle has less gap/HO cost", "Figs 4-4/4-5"]),
        ("row", ["Benefit", "LTE", "CA transfer option", "CaUserLoadTransferSw can move CA/PCC UEs in user-number MLB", "Target CA capability condition", "Otherwise CA UEs stay filtered", "pp.129–136, 157"]),
        ("row", ["Limitation", "LTE", "Unequal BW / coverage / UE band", "Gain falls when coverage, capability or PLMN differ", "p.39, 141, 214", "Not a substitute for RF fix", "§6.1.2.2"]),
        ("row", ["Limitation", "LTE", "PRB MLB", "No CA UE; no experience fairness; bursty", "CA ≳ 60% kills gain", "Use as supplement only", "§6.5"]),
        ("row", ["Limitation", "LTE", "Blind offload", "No target load check", "Can overload a ‘low-load’ looking cell", "pp.167, 231", "Offload path"]),
        ("row", ["Limitation", "LTE", "Energy saving", "MLB and carrier shutdown change each other’s targets", "Coordinate ES", "pp.153–156, 219–222", "Function impact"]),
        ("space", 8),

        ("section", "SN-4  Selection criteria / Trigger Condition"),
        ("heads", ["Rule", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration", "Chart / Doc Ref"]),
        ("row", ["Connected UE-number", "LTE", "InterFreqMlbUeNumThd + MlbUeNumOffset", "N ≥ thd+offset for whole MlbTrigJudgePeriod", "Stop N < thd. Mode UE_NUMBER_ONLY + SynchronizedUE", "Use ActiveUe on unequal BW", "p.128; Table 6-2"]),
        ("row", ["Idle UE-number", "LTE", "InterFreqIdleMlbUeNumThd + IdleUE", "Idle-user condition holds", "RRC release + dedicated prio", "Next session only", "Table 5-2"]),
        ("row", ["PRB usage", "LTE", "InterFreqMlbThd / UlThd + LoadOffset + min UE", "PRB ≥ thd+offset", "Does not move CA UEs", "Bursty services can false-trigger", "Table 6-18; §6.5.1"]),
        ("row", ["Target admit", "LTE", "MlbTargetInd, OverlapInd, NoHoFlag, HO SR, HW/transport", "Before using a neighbor as equalisation target", "Low/Med HW+transport; not punished", "Reject ⇒ punish CellPunishPrdNum × eval period", "pp.27–29"]),
        ("row", ["UE filter", "LTE", "CELLMLBUESEL + protect timers", "After trigger and admit", "UL-sync, not emergency, QCI/SPID, optional ARP/PRB/MCS/SNR", "Do not pick edge UEs only for PRB", "pp.130–135"]),
        ("row", ["Volume / freq", "LTE", "MlbMaxUeNum / FreqSelectStrategy", "At execution", "FAIRSTRATEGY / PRIORITYBASED / LOADPRIORITY", "LOADPRIORITY uses above-average load difference", "pp.137, 213"]),
        ("row", ["HO event", "LTE", "MlbInterFreqHoEventType", "After UE selected", "A4 co-sited; A5 licensed non-cosited", "LOAD_COVERAGE_MEAS_DECOUPLE_SW if coverage meas already on", "Table 6-3; p.139"]),
        ("space", 8),

        ("section", "SN-5  Activation parameter / Switch   (core table)"),
        ("heads", H_ACT),
        ("row", ["5.1", "LTE", "CELLALGOSWITCH", "InterFreqMlbSwitch", "Master intra-RAT MLB bit", "License + NRT + MlbTargetInd", "Table 6-2 p.127"]),
        ("row", ["5.2", "LTE", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch", "Idle dedicated-priority transfer", "T320 + SIB5 NORMAL + idle-allowed MlbTargetInd", "§5.1.1.5"]),
        ("row", ["5.3", "LTE", "CELLALGOSWITCH", "InterFreqBlindMlbSwitch", "Blind MLB — leave OFF unless containment proven", "BlindHoPriority", "p.136"]),
        ("row", ["5.4", "LTE", "CELLMLB", "MlbTriggerMode", "UE_NUMBER_ONLY typical for user-experience equalisation", "PRB_ONLY skips CA UEs", "Table 6-2; §6.5"]),
        ("row", ["5.5", "LTE", "CELLMLB", "InterFreqUeTrsfType", "SynchronizedUE (connected). IdleUE for idle. PrbMlbSynchronizedUE for PRB.", "Must match trigger mode", "Table 6-2"]),
        ("row", ["5.6", "LTE", "eval SW", "ActiveUeBasedLoadEvalSw", "ON when frequencies have different bandwidths (Huawei)", "Load model N", "Table 5-5"]),
        ("row", ["5.7", "LTE", "eval SW", "SpectralEffBasedLoadEvalSw", "ON when SE differs significantly e.g. >30% (Huawei)", "≥10 UL-sync UEs for SE refresh", "Table 5-5"]),
        ("row", ["5.8", "LTE", "eval SW", "LoadTransferEnhSw", "ON for multi-target / PRB-diff calculation", "Several candidate frequencies", "p.134"]),
        ("row", ["5.9", "LTE", "eval SW", "CaUserLoadTransferSw", "ON only if CA UEs must be transferable", "Target CA capability ≥ serving on one documented path", "pp.129–136, 157"]),
        ("row", ["5.10", "LTE", "CELLMLB", "MlbHoCellSelectStrategy", "ONLY_STRONGEST_CELL (Huawei recommended)", "A4/A5 meas success", "Table 6-5 p.159"]),
        ("row", ["5.11", "LTE", "CELLMLB", "FreqSelectStrategy", "FAIRSTRATEGY / PRIORITYBASED / LOADPRIORITY", "PRIORITYBASED uses MlbFreqPriority + freq penalty", "pp.137, 213"]),
        ("row", ["5.12", "LTE", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType + MlbTargetInd", "A4 (or A5 if licensed). ALLOWED / WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB", "OverlapInd + NoHoFlag for coverage HO remain separate", "pp.28, 129; Table 6-3"]),
        ("space", 8),

        ("section", "SN-6  Prerequisite functions"),
        ("heads", H_PRE),
        ("row", ["6.1", "LTE", "License", "Intra-RAT MLB (+ A5 license if event=A5)", "Present on eNodeB", "Switch ON but no Load HO counters", "pp.142–143"]),
        ("row", ["6.2", "LTE", "X2 / intra-eNB", "Load exchange for equalisation", "Working", "Without it only offload/blind remains", "§4.1"]),
        ("row", ["6.3", "LTE", "NRT", "OverlapInd, PERMIT_HO, no PCI conflict", "Valid pairs", "Target never admitted", "pp.27–28"]),
        ("row", ["6.4", "LTE", "Connected Mode meas", "A4/A5 actually delivered (flags, object cap, SMeasure, gap)", "See Connected SN-6", "Trigger with 0 meas success", "p.139"]),
        ("row", ["6.5", "LTE", "CELLMLB", "NCellHoSuccRateThld", "Pair HO success ≥ thd", "Failed pair is not a target — fix RF/HO first", "p.27"]),
        ("row", ["6.6", "LTE", "HW / transport", "LowLoad / MediumLoad", "High/OverLoad not a normal equalisation target", "Fig 4-3", "pp.17–18"]),
        ("row", ["6.7", "LTE", "Idle SIB5", "NORMAL if idle MLB ON", "UNDELIVER cannot be idle target", "IM §5.3.2.3", "Idle sheet"]),
        ("row", ["6.8", "LTE", "MAE counters", "≥1×15-min period if smart thd ON", "Required for NCellTrigThldSmartOptAlgoSw", "§§5.1.3.4, 6.5.3.4", "pp.46, 223"]),
        ("space", 8),

        ("section", "SN-7  Mutually impacted"),
        ("heads", H_IMP),
        ("row", ["7.1", "LTE", "User-number connected MLB", "Fixed-proportion idle MLB", "Do not combine — ping-pong", "Keep fixed-proportion OFF", "§5.4.2.2"]),
        ("row", ["7.2", "LTE", "PRB_USAGE MLB", "PRB_VALUATION MLB", "Mutually exclusive in the same mode", "Pick one if PRB used", "pp.216, 245"]),
        ("row", ["7.3", "LTE", "MLB", "FreqPri reverse target", "Ping-pong [CM] p.303", "MlbBasedFreqPriHoSwitch + no reverse pair", "CM Table 11-7"]),
        ("row", ["7.4", "LTE", "MLB", "Carrier shutdown / deep dormancy", "Target set and gain change", "Coordinate ES", "pp.153–156"]),
        ("row", ["7.5", "LTE", "Connected offload active", "Flexible CA combination", "CA avoids a cell under connected offload", "Expected — pp.168, 233", "CA"]),
        ("row", ["7.6", "LTE", "Smart learned thds", "BW change / CA-eval SW / upgrade / board / cell deactivation", "Forced relearn or stale thd", "Do not treat learned thds as daily manual MOD targets", "pp.29–31, 89–91"]),
        ("space", 8),

        ("section", "SN-8  Relation with Other Feature"),
        ("heads", H_IMP),
        ("row", ["8.1", "LTE", "Idle Mode", "Idle MLB method = dedicated prio + T320", "SIB5 NORMAL; MlbTargetInd idle allow/forbid", "Idle sheet SN-5/11", "§5.1.1.5"]),
        ("row", ["8.2", "LTE", "Connected Mode", "HO engine A4/A5, meas, admit, punish", "ONLY_STRONGEST_CELL + A4 typical co-sited", "Connected sheet SN-5", "§6.1.1.5.2"]),
        ("row", ["8.3", "LTE", "Carrier Aggregation", "PCC transfer vs SCell scheduling", "PRB MLB skips CA UEs; user-number needs CaUserLoadTransferSw", "pp.129–136, 157, 207–215", "CA document"]),
        ("row", ["8.4", "LTE", "Energy saving", "Mutual target exclusion", "No MLB onto a sleeping carrier", "pp.153–156", "ES"]),
        ("row", ["8.5", "LTE", "Inter-RAT MLB", "Separate document", "Out of this three-book set", "Do not use this sheet for IRAT", "IRAT MLB book"]),
        ("space", 8),

        ("section", "SN-9  License Requirements"),
        ("heads", ["Function", "RAT", "Switch / function", "License (verify in MAE)", "If missing", "Notes", "Chart / Doc Ref"]),
        ("row", ["Connected/idle equalisation", "LTE", "InterFreqMlbSwitch / IdleMlbSwitch", "Intra-RAT Mobility Load Balancing", "No Load HO / dedicated-pri counters", "Core", "Table 6-2"]),
        ("row", ["MLB event A5", "LTE", "MlbInterFreqHoEventType=A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Stay on A4", "pp.142–143", "Table 6-3"]),
        ("row", ["Blind MLB", "LTE", "InterFreqBlindMlbSwitch", "Blind option if sold separately", "Blind unavailable", "Higher risk without load exchange", "pp.167, 231"]),
        ("row", ["CA user transfer", "LTE", "CaUserLoadTransferSw", "CA + MLB CA-transfer option — verify", "CA UEs filtered", "p.157", "pp.129–136"]),
        ("row", ["Smart n-cell thd", "LTE", "NCellTrigThldSmartOptAlgoSw", "SON/intelligent MLB option + MAE 15-min counters", "No learned pair thd", "7-day learn / 7-day refresh", "pp.29–31"]),
        ("space", 8),

        ("section", "SN-10  All Parameter List  (sequence / connected order)"),
        ("note", "Ch.8 of the MLB book does not print full ranges/defaults — use the version-matched parameter reference (p.299). Conditional Parameter = other-feature relation."),
        ("heads", H_PAR),
        ("row", ["1", "LTE", "CELLALGOSWITCH", "InterFreqMlbSwitch / InterFreqIdleMlbSwitch / InterFreqBlindMlbSwitch", "Master bits", "License", "Table 6-2"]),
        ("row", ["2", "LTE", "CELLMLB", "MlbTriggerMode / PrbLoadCalcMethod / InterFreqUeTrsfType / InterFreqMLBRanShareMode", "Mode", "Must match algorithm", "Table 6-2 / 6-18"]),
        ("row", ["3", "LTE", "eval SW", "ActiveUeBasedLoadEvalSw / SpectralEffBasedLoadEvalSw / LoadTransferEnhSw / CaUserLoadTransferSw", "Load model and CA", "Table 5-5; unequal BW / SE>30% / CA", "pp.25–27, 129–136"]),
        ("row", ["4", "LTE", "CELL / CELLMLB", "CellCapacityScaleFactor / MuMimoPrbStatOptSwitch / MultiRruMode", "Capability scale", "Do not fake bandwidth with scale", "§5.1.1"]),
        ("row", ["5", "LTE", "CELLMLB", "MlbTrigJudgePeriod / InterFreqLoadEvalPrd", "Timing", "5 s + MaxUe≥40 warning", "p.128, p.136"]),
        ("row", ["6", "LTE", "CELLMLB", "InterFreqMlbUeNumThd / MlbUeNumOffset / InterFreqIdleMlbUeNumThd", "UE-number trigger", "Enter = thd+offset; leave = thd", "p.128; Table 5-2"]),
        ("row", ["7", "LTE", "CELLMLB", "InterFreqMlbThd / InterFreqMlbUlThd / LoadOffset / MlbMinUeNumThd(+Offset)", "PRB trigger family", "Supplement only; no CA UE", "§6.5.1"]),
        ("row", ["8", "LTE", "CELLMLB", "LoadDiffThd / InterFreqOffloadOffset / InterFrqUeNumOffloadOffset / InterFIdleUeNumOffloadOfs", "Equalise vs offload slack", "Offload if load exchange missing", "§4.1"]),
        ("row", ["9", "LTE", "CELLMLB", "MlbMaxUeNum / MlbIdleUeNumAdjFactor", "Transfer volume", "Never ≥40 with 5 s eval", "p.136"]),
        ("row", ["10", "LTE", "CELLMLB", "FreqSelectStrategy / MlbFreqPriority / MlbFreqUlPriority", "Freq pick + UL priority for UL PRB MLB", "FAIR / PRIORITY / LOAD", "pp.137, 213"]),
        ("row", ["11", "LTE", "CELLMLB", "MlbHoCellSelectStrategy", "ONLY_STRONGEST_CELL recommended", "Avoid coverage bounce", "Table 6-5 p.159"]),
        ("row", ["12", "LTE", "EUTRANINTERNFREQ", "MlbTargetInd / MlbInterFreqHoEventType / IfMlbThdRsrpOffset", "Target allow + A4/A5 + offset", "WITHOUT_* blocks one MLB mode", "pp.28, 129, 137"]),
        ("row", ["13", "LTE", "EUTRANINTERNFREQ / NCELL", "OverlapInd / NoHoFlag / AggregationAttribute / LoadBalanceNCellScope", "Relation eligibility", "Coverage HO ≠ MLB target", "pp.27–28"]),
        ("row", ["14", "LTE", "CELLMLB", "NCellHoSuccRateThld / CellPunishPrdNum / FreqPunishPrdNum / PunishJudgePrdNum", "Admit / penalty", "Do not lower HO SR thd to force MLB", "p.29"]),
        ("row", ["15", "LTE", "CELLMLBUESEL", "UeSelectArp/Prb/DlMcs/QciPrio + thds / SnrBasedUeSelectionMode", "Who is movable", "Protect emergency / QCI policy", "pp.130–135"]),
        ("row", ["16", "LTE", "CELLMLB", "MlbHoInProtectTimer / MlbUeSelectPunishTimer / MlbHoInProtectMode", "Re-MLB protect", "Ping-pong guard", "pp.130–135"]),
        ("row", ["17", "LTE", "CELLOPHOCFG", "MLB_HO_FORBID_SW", "No MLB meas if UE > ~30 km/h", "NLOS misclassify", "pp.136–141"]),
        ("row", ["18", "LTE", "EUTRANINTERNFREQ", "InterFreqMlbBlindHo / BlindHoPriority", "Blind path", "Containment required", "p.136"]),
        ("row", ["19", "LTE", "EUTRANINTERNFREQ", "InterFreqMlbDlPrbOffset / UlPrbOffset", "Per-freq PRB bias", "PRB mode only", "p.210"]),
        ("row", ["20", "LTE", "CELLPRBVALMLB / MlbQciGroup", "PrbValMlbTrigThd / AdmitThd / FilterFactor / min QoS bit rates", "PRB-evaluation MLB", "Exclusive vs PRB_USAGE", "Table 6-25; §5.8.1"]),
        ("row", ["21", "LTE", "smart", "NCellTrigThldSmartOptAlgoSw + LocalToNCell / NToLocal UE/PRB thds", "7-day learn / 7-day refresh", "Need 15-min counter subscription", "pp.29–31, 89–91"]),
        ("row", ["22", "LTE", "RRCCONNSTATETIMER / EnhancedMlb", "T320ForLoadBalance / DediPrioManageOnLowLoadSw", "Idle dedicated prio life / hold", "Idle MLB", "§5.1.1.5; Table 5-10"]),
        ("row", ["23", "LTE", "QCI / SPID relation", "QCIEUTRANRELATION.MobilityTargetInd / CellQciPara.QciAlgoSwitch", "QCI/SPID target prohibition", "Connected Mode QCI policy", "pp.130–135"]),
        ("space", 8),

        ("section", "SN-11  Final MML Command for activations  (maintain sequence)"),
        ("heads", H_MML),
        ("row", ["0", "—",
                 "LST CELLALGOSWITCH: LocalCellId=<x>; LST CELLMLB: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLMLBUESEL: LocalCellId=<x>; check Intra-RAT MLB license;",
                 "MAE license + version-matched parameter reference (Ch.8 p.299)", "Dump current. Feature book does not list every default/range.", "Baseline", "Attach LST to the change record"]),
        ("row", ["1", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;",
                 "OverlapInd valid; NoHoFlag=PERMIT_HO_ENUM; FREQ_MEAS_FLAG; not forbid-meas. A5 only with non-cosited MLB license.",
                 "Allow this frequency as MLB target. Use WITHOUT_CONNECT_MLB / WITHOUT_IDLE_MLB to block one mode.",
                 "Target + event", "pp.28, 129; Table 6-3"]),
        ("row", ["2", "eval related MO",
                 "Enable ActiveUeBasedLoadEvalSw, SpectralEffBasedLoadEvalSw, LoadTransferEnhSw as required (unequal BW / SE e.g. >30% / multi-target);",
                 "Confirm exact bit/MO names in MAE of the running version", "Huawei Table 5-5. Do this before chasing UE-number thd.", "Load model", "pp.47–48"]),
        ("row", ["3", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL, FreqSelectStrategy=<FAIRSTRATEGY|PRIORITYBASED|LOADPRIORITY>;",
                 "InterFreqMlbSwitch will be ON", "ONLY_STRONGEST_CELL is Huawei-recommended (Table 6-5 p.159).", "Mode + pick strategy", "Table 6-2"]),
        ("row", ["4", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<thd>, MlbUeNumOffset=<ofs>, MlbMaxUeNum=<n>, MlbTrigJudgePeriod=<p>, InterFreqLoadEvalPrd=<prd>;",
                 "Do not use InterFreqLoadEvalPrd=5 s with MlbMaxUeNum≥40 (p.136). Calibrate thd after ActiveUe/SE ON.",
                 "Trigger and volume. Enter = thd+offset; leave = thd.", "UE-number trigger", "p.128, p.136"]),
        ("row", ["5", "CA related",
                 "Enable CaUserLoadTransferSw only if CA UEs must be moved and target CA capability conditions are met;",
                 "CA license + combination plan", "If OFF, CA UEs treating the cell as PCell/SCell are filtered (p.157).", "CA transfer", "pp.129–136"]),
        ("row", ["6", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "Steps 1–4 done. Blind bit remains 0 unless designed. Idle also needs T320 and SIB5 NORMAL.",
                 "Master bits last (or same window after targets).", "Activate MLB", "Table 6-2"]),
        ("row", ["7", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
                 "Idle MLB ON", "SPID/PCC T320 remains 180 min.", "Dedicated-priority lifetime", "§5.1.1.5"]),
        ("row", ["8", "CELLMLBUESEL",
                 "MOD CELLMLBUESEL: LocalCellId=<x>; apply QCI/ARP/emergency protection; avoid selecting edge UEs only for PRB;",
                 "ONLY_STRONGEST_CELL already on CELLMLB", "UE selection policy.", "UE pick", "pp.130–135"]),
        ("row", ["9", "Verify",
                 "L.HHO.InterFreq.Load.Prep/Exec/Succ and UeNumLoad.*; HighLoad.Dur/Num; Load.Meas/MeasSucc; ActiveUser.DL; PCell/SCell; PRB; Thrp.bits/Time.DL; idle DedicatedPri. SON logs: Inter-Frequency Handover Statistics; Inter-Frequency Idle Mode Release Statistics.",
                 "Subscribe 15-min counters if smart thd is used. When both PRB and UE-number MLB are on: PRB HO ≈ Load − UeNumLoad (§6.5.4.3).",
                 "Tables 6-6, 6-21, 6-28 pp.164, 227–228, 248–249. Log contents §§6.1.4.2, 6.5.4.2.",
                 "Counters + SON logs", "Close the sequence"]),
    ]


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    # cover on first sheet
    ws = wb.active
    ws.title = "tmp"
    build_sheet(wb, "Mobility Management", "Mobility Management - Detailed Notes", cover_blocks(), tab=BLUE)
    build_sheet(wb, "Idle Mode Management", "Idle Mode Management - Detailed Notes", idle_blocks(), tab="008000")
    build_sheet(wb, "Connected Mode", "Mobility Management in Connected Mode - Detailed Notes", connected_blocks(), tab="1F4E79")
    build_sheet(wb, "Intra-RAT MLB", "Intra-RAT Mobility Load Balancing - Detailed Notes", mlb_blocks(), tab="2E75B6")
    del wb["tmp"]
    wb.properties.title = "eRAN21.1 Mobility Management Detailed Notes"
    wb.properties.subject = "Idle / Connected / Intra-RAT MLB PDF summary — SN-1 to SN-11"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
