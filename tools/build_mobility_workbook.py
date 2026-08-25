#!/usr/bin/env python3
"""eRAN21.1 4G mobility — step-by-step Excel summary of three Huawei books.

Not the user SN sample. Organised as one synchronised chain, then each
feature in procedure order: Idle → Connected → Intra-RAT MLB → activation.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/4G_LTE_Mobility_Management/4G_LTE_Mobility_Management_eRAN21.1_v3.1.xlsx"
COLS = 6

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
C = Alignment(wrap_text=True, vertical="center", horizontal="center")
T = Alignment(wrap_text=True, vertical="top", horizontal="left")
LI = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)

H_STEP = ["Step", "What happens", "MO / parameter", "Rule in the feature book", "If this is wrong", "Source"]
H_NOTE = ["No.", "Huawei caution", "What it means in the network", "Do / do not", "Related step", "Source"]
H_PAR = ["Order", "MO", "Parameter", "Role in the procedure", "Depends on / couples with", "Source"]
H_MML = ["Order", "MO", "MML command (run in this order)", "Must already be true", "Notes", "Source"]
H_LINK = ["Open this sheet", "Book", "Issue", "This sheet answers", "Read after", "Source"]

W = [10, 22, 52, 30, 32, 20]


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


def setup(ws, footer, tab=BLUE):
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(0.35, 0.35, 0.5, 0.45)
    ws.oddHeader.left.text = "4G LTE Mobility Management  |  Huawei eRAN21.1  |  v3.1"
    ws.oddFooter.left.text = footer
    ws.oddFooter.right.text = "Page &P of &N"
    ws.sheet_properties.tabColor = tab
    ws.freeze_panes = "A3"
    ws.print_title_rows = "1:1"
    ws.sheet_format.defaultRowHeight = 18


def title(ws, r, text):
    merge(ws, r, 1, COLS, text, size=16, bold=True, color=WHITE, bg=BLUE, align=C, h=30)
    return r + 1


def spacer(ws, r, h=8):
    for c in range(1, COLS + 1):
        put(ws, r, c, "", bg=WHITE, h=h)
    return r + 1


def section(ws, r, text):
    merge(ws, r, 1, COLS, "  " + text, size=12, bold=True, color=GREEN, bg=YELLOW, align=LI, h=22)
    return r + 1


def major(ws, r, text):
    put(ws, r, 1, "NOTE", size=9, bold=True, color=WHITE, bg=GREEN, align=C, h=22)
    merge(ws, r, 2, COLS, "  " + text, size=10, bold=True, color=GREEN, bg=MAJOR_BG, align=LI, h=22)
    return r + 1


def note(ws, r, text):
    merge(ws, r, 1, COLS, "  " + text, size=9, bold=False, color=BLUE, bg=WHITE, align=LI, h=20)
    return r + 1


def auto_h(values):
    longest = 1
    for i, v in enumerate(values):
        s = "" if v is None else str(v)
        width = 16 if i == 0 else 40
        longest = max(longest, (len(s) + width - 1) // width, s.count("\n") + 1)
    return min(72, max(22, 14 + longest * 12))


def heads(ws, r, names):
    names = list(names) + [""] * COLS
    for c in range(1, COLS + 1):
        put(ws, r, c, names[c - 1], size=9, bold=True, color=BLACK, bg=HDR, align=C, h=20)
    return r + 1


def rec(ws, r, values):
    values = list(values) + [""] * COLS
    h = auto_h(values[:COLS])
    for c in range(1, COLS + 1):
        put(ws, r, c, values[c - 1], size=10, bg=GREY, align=T, h=h)
    return r + 1


def flow(ws, r, boxes, ref=""):
    """Three process boxes per row: box → box → box."""
    steps = list(boxes)
    box_cols = [1, 3, 5]
    arr_cols = [2, 4]
    i = 0
    while i < len(steps):
        chunk = steps[i:i + 3]
        for c in range(1, COLS + 1):
            put(ws, r, c, "", bg=WHITE, h=28)
        for j, step in enumerate(chunk):
            put(ws, r, box_cols[j], step, size=9, bold=True, color=WHITE, bg=BLUE, align=C, h=28)
            if j < len(chunk) - 1:
                put(ws, r, arr_cols[j], "→", size=16, bold=True, color=GREEN, bg=WHITE, align=C, h=28)
        r += 1
        i += 3
        if i < len(steps):
            for c in range(1, COLS + 1):
                put(ws, r, c, "", bg=WHITE, h=12)
            merge(ws, r, 1, COLS, "↓", size=12, bold=True, color=GREEN, bg=WHITE, align=C, h=12)
            r += 1
    if ref:
        r = note(ws, r, "Document chart: " + ref)
    return r


def mml_heads(ws, r):
    put(ws, r, 1, "SN", size=9, bold=True, color=BLACK, bg=HDR, align=C, h=20)
    merge(ws, r, 2, 3, "MML", size=9, bold=True, color=BLACK, bg=HDR, align=C, h=20)
    put(ws, r, 4, "Purpose", size=9, bold=True, color=BLACK, bg=HDR, align=C, h=20)
    merge(ws, r, 5, 6, "Note", size=9, bold=True, color=BLACK, bg=HDR, align=C, h=20)
    return r + 1


def mml_row(ws, r, sn, mml, purpose, note_txt):
    h = auto_h([str(sn), mml, purpose, note_txt])
    put(ws, r, 1, sn, size=10, bold=True, bg=GREY, align=C, h=h)
    merge(ws, r, 2, 3, mml, size=9, bg=GREY, align=T, h=h)
    put(ws, r, 4, purpose, size=10, bg=GREY, align=T, h=h)
    merge(ws, r, 5, 6, note_txt, size=10, bg=GREY, align=T, h=h)
    return r + 1


def combined_mml(ws, r, rows):
    r = section(ws, r, "Combined MML Command (all Together)")
    r = note(ws, r, "All commands below are in execution order. Replace <x>, <earfcn>, <g> and <val> on the NE. Confirm enum names in MAE. Example dBm in the Connected book are not design values.")
    r = mml_heads(ws, r)
    for i, (mml, purpose, note_txt) in enumerate(rows, 1):
        r = mml_row(ws, r, i, mml, purpose, note_txt)
    return r


# Feature MML in sequence. Purpose uses // as in the attached snap.
MML_IDLE = [
    ("LST CELLRESEL: LocalCellId=<x>;",
     "//Dump serving idle parameters before change",
     "Keep LST output with the change record"),
    ("LST EUTRANINTERNFREQ: LocalCellId=<x>;",
     "//Dump SIB5 inter-frequency list before change",
     "Repeat review for every DlEarfcn"),
    ("LST CELLALGOSWITCH: LocalCellId=<x>;",
     "//Dump idle/MLB switch bits before change",
     "Read-only"),
    ("LST RRCCONNSTATETIMER:;",
     "//Dump T320 before change",
     "SPID/PCC T320 stays 180 min"),
    ("MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<prio>, SIntraSearchCfgInd=CFG, SNonIntraSearchCfgInd=CFG, SIntraSearch=<val>, SNonIntraSearch=<val>;",
     "//Set serving common priority and search start",
     "SIntraSearch > SNonIntraSearch. Huawei example SNonIntraSearch=10. Capacity/hotspot above coverage layer."),
    ("MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;",
     "//Publish this frequency in SIB5 with priority",
     "Repeat for every non-serving frequency that must be reselectable. Do not use UNDELIVER on an idle-MLB target."),
    ("MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, ThreshXhigh=<val>, ThreshXlow=<val>, QoffsetFreq=<val>, EutranReselTime=<val>;",
     "//Set higher/lower/equal-priority reselection qualification",
     "Calibrate from MR. No universal dBm in the feature book."),
    ("MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<val>;",
     "//Permit leave to a lower-priority frequency",
     "Align with connected coverage A2/A5. Do not set so low that a dying serving cell never yields."),
    ("MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=ALLOWED;",
     "//Allow this frequency as idle and/or connected MLB target",
     "Use ALLOWED_WITHOUT_IDLE_MLB or ALLOWED_WITHOUT_CONNECT_MLB to block one mode. Coverage NoHoFlag stays PERMIT if coverage HO is required."),
    ("MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqIdleMlbSwitch-1;",
     "//Turn on intra-LTE idle MLB",
     "Do this after SIB5 NORMAL and MlbTargetInd. Leave InterFreqBlindMlbSwitch-0 unless designed. Confirm bit name in MAE."),
    ("MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
     "//Set lifetime of load-balance dedicated priorities",
     "Idle MLB path only. SPID/PCC remains 180 min."),
    ("LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>;",
     "//Verify idle parameters after MOD",
     "Wait the next SI modification period before judging camping"),
]

MML_CONN = [
    ("LST INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>;",
     "//Dump A1–A5 group before change",
     "Do not paste example −85/−87/−103 dBm from the book"),
    ("LST EUTRANINTERNFREQ: LocalCellId=<x>;",
     "//Dump meas objects and MLB/FreqPri flags",
     "Check FREQ_MEAS_FLAG and HO_TRG_FREQ_FORBID_MEAS_FLAG in MAE"),
    ("LST EUTRANINTERFREQNCELL: LocalCellId=<x>;",
     "//Dump NRT before change",
     "Symmetric neighbour, PERMIT_HO, no PCI conflict"),
    ("LST CELLUEMEASCONTROLCFG: LocalCellId=<x>; LST HOMEASCOMM:;",
     "//Dump object cap and SMeasure",
     "SMeasure can silently hide A4"),
    ("MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbInterFreqHoEventType=A4;",
     "//Set MLB/FreqPri event type on this frequency",
     "Select FREQ_MEAS_FLAG and deselect HO_TRG_FREQ_FORBID_MEAS_FLAG in MAE for required HO targets. A5 only with non-cosited MLB license."),
    ("MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<n>, MaxEutranFddMeasFreqNum=<n>;",
     "//Allow enough inter-frequency measurement objects",
     "n ≥ number of frequencies this cell must measure. Otherwise equal-priority objects drop at random."),
    ("MOD HOMEASCOMM: SMeasure=<val>;",
     "//Allow inter-frequency meas when serving is not extremely strong",
     "Confirm parameter presence on this version. Too high a value suppresses A4."),
    ("MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqHoA1A2Hyst=<hyst>, InterFreqHoA1A2TimeToTrig=<ttt>;",
     "//Set coverage A1/A2 stability",
     "Then set the correct coverage A2 family from MR. Wrong family = wrong HO."),
    ("MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<rsrp>, InterFreqHoA4Hyst=<hyst>, InterFreqHoA4TimeToTrig=<ttt>;",
     "//Set A4 absolute target gate for MLB/FreqPri",
     "A4 must be better than coverage A2. TTT must not be 5120 ms if FreqPri or MLB A4 is required. Calibrate from MR."),
    ("MOD INTRARATHOCOMM: LocalCellId=<x>, FreqPriInHoProtectionTimer=<t>, FreqPriIFHoWaitingTimer=<t>;",
     "//Protect against FreqPri bounce-back after incoming unnecessary HO",
     "Confirm exact MO/parameter names in MAE. No reverse MLB target on a FreqPri pair."),
]

MML_MLB = [
    ("LST CELLALGOSWITCH: LocalCellId=<x>; LST CELLMLB: LocalCellId=<x>; LST CELLMLBUESEL: LocalCellId=<x>;",
     "//Dump MLB switches, trigger and UE-pick before change",
     "Confirm Intra-RAT MLB license. Ch.8 p.299 → parameter reference for defaults."),
    ("MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;",
     "//Allow this frequency as MLB target and use event A4",
     "OverlapInd valid, NoHoFlag=PERMIT_HO. A5 only with Intra-LTE Load Balancing for Non-cosited Cells license."),
    ("MOD CELLMLB: LocalCellId=<x>, ActiveUeBasedLoadEvalSw=ON, SpectralEffBasedLoadEvalSw=ON, LoadTransferEnhSw=ON;",
     "//Turn on BW/SE-aware load model",
     "Huawei: ActiveUe when bandwidths differ; SpectralEff when SE differs a lot (e.g. >30%). Confirm these are CELLMLB fields on this version."),
    ("MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL, FreqSelectStrategy=FAIRSTRATEGY;",
     "//Set connected UE-number equalisation strategy",
     "ONLY_STRONGEST_CELL is Huawei-recommended. FAIRSTRATEGY / PRIORITYBASED / LOADPRIORITY as designed."),
    ("MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<thd>, MlbUeNumOffset=<ofs>, MlbMaxUeNum=<n>, MlbTrigJudgePeriod=<p>, InterFreqLoadEvalPrd=<prd>;",
     "//Set trigger threshold and transfer volume",
     "Enter = thd+offset; leave = thd. Do not use eval period 5 s with MlbMaxUeNum≥40."),
    ("MOD CELLMLBUESEL: LocalCellId=<x>;",
     "//Apply QCI/ARP/emergency UE-pick policy",
     "Do not pick edge UEs only for PRB. Confirm fields in MAE."),
    ("MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
     "//Turn on connected and idle intra-RAT MLB last",
     "Blind bit stays 0 unless containment is proven. Idle also needs T320 and SIB5 NORMAL."),
    ("MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
     "//Set idle dedicated-priority lifetime",
     "Idle MLB path. SPID/PCC remains 180 min."),
    ("LST CELLMLB: LocalCellId=<x>; LST CELLALGOSWITCH: LocalCellId=<x>;",
     "//Verify MLB activation",
     "Then check Load HO / UeNumLoad / DedicatedPri counters and SON inter-frequency logs."),
]

MML_ALL = [
    ("LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST EUTRANINTERFREQNCELL: LocalCellId=<x>; LST INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; LST CELLUEMEASCONTROLCFG: LocalCellId=<x>; LST HOMEASCOMM:; LST CELLALGOSWITCH: LocalCellId=<x>; LST CELLMLB: LocalCellId=<x>; LST CELLMLBUESEL: LocalCellId=<x>; LST RRCCONNSTATETIMER:;",
     "//Dump the full mobility baseline before any MOD",
     "One sequence across Idle + Connected + MLB. Keep LST with the change record."),
] + MML_IDLE[4:9] + MML_CONN[4:] + MML_MLB[1:]


def link_cell(cell, sheet, target="A1"):
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!{target}", display=str(cell.value or ""))
    cell.font = ft(10, True, "0563C1", underline="single")
    cell.alignment = T


def build_sheet(wb, name, title_text, blocks, tab=BLUE):
    ws = wb.create_sheet(name)
    widths(ws, W)
    setup(ws, title_text, tab=tab)
    r = 1
    r = title(ws, r, title_text)
    r = spacer(ws, r, 6)
    for b in blocks:
        kind = b[0]
        if kind == "section":
            r = section(ws, r, b[1])
        elif kind == "major":
            r = major(ws, r, b[1])
        elif kind == "note":
            r = note(ws, r, b[1])
        elif kind == "flow":
            r = flow(ws, r, b[1], b[2] if len(b) > 2 else "")
        elif kind == "heads":
            r = heads(ws, r, b[1])
        elif kind == "row":
            rec(ws, r, b[1])
            if len(b) > 2:
                link_cell(ws.cell(r, 1), b[2])
            r += 1
        elif kind == "space":
            r = spacer(ws, r, b[1] if len(b) > 1 else 8)
        elif kind == "mml":
            r = combined_mml(ws, r, b[1])
    return ws


def overview():
    return [
        ("section", "What this file is"),
        ("note", "Step-by-step summary of three Huawei eRAN21.1 feature-parameter books. One mobility chain. Not a copy of any sample template. Values in MML examples in the books are examples, not live design."),
        ("space", 6),
        ("heads", ["Book", "Document", "Issue / date", "Owns in the chain", "Does not own", "Read as"]),
        ("row", ["Idle Mode Management", "Idle Mode Management Feature Parameter Description", "eRAN21.1 Issue 04, 2026-05-30",
                 "Where the UE camps. Next RRC setup cell. Dedicated priority at release.", "Connected HO events. Load algorithm.", "Sheet 2"]),
        ("row", ["Connected Mode", "Mobility Management in Connected Mode Feature Parameter Description", "eRAN21.1 Issue 08, 2026-06-30",
                 "Measurement and HO engine (A1–A5). Coverage first. Frequency-priority steering.", "Who to move for load. Load = N/C.", "Sheet 3"]),
        ("row", ["Intra-RAT MLB", "Intra-RAT Mobility Load Balancing Feature Parameter Description", "eRAN21.1 Issue 10, 2026-06-30",
                 "When a cell is overloaded. Which UEs. Idle dedicated-priority transfer or connected load HO.", "RF plan. Coverage A2 family design.", "Sheet 4"]),
        ("space", 8),
        ("section", "How the three books fit (one chain)"),
        ("flow", ["Idle: camp / reselect", "RRC connect", "Coverage A2/A5 protect", "FreqPri A1/A4 steer", "MLB load HO / idle release", "Back to idle (T320)"],
         "Idle Fig 5-1  →  Connected Fig 4-1  →  MLB Fig 3-1"),
        ("major", "Idle decides the next access layer. Connected executes measurement and handover. MLB only decides who and when to move for load. Frequency-priority HO is not the MLB algorithm."),
        ("major", "Necessary coverage HO preempts load / optimisation HO. Do not use connected frequency-priority as a substitute for Intra-RAT MLB."),
        ("space", 8),
        ("section", "Open the sheets in this order"),
        ("heads", H_LINK),
        ("row", ["1. End-to-end chain", "All three", "—", "The full sequence from power-on to the next idle camp", "Start here", "This file"], "1. End-to-end chain"),
        ("row", ["2. Idle Mode", "Idle Mode Management", "Issue 04", "Selection, SIB, reselection, dedicated priority", "After the chain", "Idle book"], "2. Idle Mode"),
        ("row", ["3. Connected Mode", "Connected Mode", "Issue 08", "A1–A5, coverage HO, frequency-priority HO", "After Idle", "Connected book"], "3. Connected Mode"),
        ("row", ["4. Intra-RAT MLB", "Intra-RAT MLB", "Issue 10", "Load model, trigger, UE pick, A4/A5 or idle release", "After Connected", "MLB book"], "4. Intra-RAT MLB"),
        ("row", ["5. Activation order", "All three", "—", "One MML sequence across Idle + Connected + MLB", "Last", "MAE / parameter reference"], "5. Activation order"),
        ("space", 8),
        ("section", "What is not in this file"),
        ("row", ["1", "No AI / daily-KPI / change-request agent", "—", "Feature-book summary only", "—", "—"]),
        ("row", ["2", "No operator-specific band priority table", "—", "Huawei says capacity/hotspot above coverage layer. Exact priority numbers are a design, not a book default.", "Idle + MLB", "MLB §5.1.2.1"]),
        ("row", ["3", "No full default/range list", "—", "MLB Ch.8 points to the version-matched parameter reference.", "Activation sheet", "MLB p.299"]),
        ("row", ["4", "MML uses LocalCellId=<x> and DlEarfcn=<earfcn>", "—", "Confirm enum names and syntax on the NE in MAE-Access.", "Last section of every sheet", "MAE"]),
        ("space", 8),
        ("mml", MML_ALL),
    ]


def chain():
    return [
        ("section", "Read this as one procedure"),
        ("note", "Each row is the next thing that happens to a UE. The Book column tells you which feature document owns that step. Open sheets 2–4 for the detail of that step."),
        ("flow", ["Power-on / idle", "Camped cell → RRC", "Protect coverage", "Steer to high band", "Balance load", "Release → next idle"],
         "Synchronised chain of the three books"),
        ("space", 6),
        ("heads", ["Step", "What happens", "Which book", "Trigger", "What the eNodeB / UE does", "Go to"]),
        ("row", ["1", "PLMN selection then cell selection", "Idle", "Power-on, recovery, or no stored cell", "Criterion S: Srxlev > 0 (and Squal if QQualMin is set). This is a floor, not a load knob.", "Sheet 2 step 1–3"]),
        ("row", ["2", "UE camps and reads system information", "Idle", "Suitable cell found", "SIB1 (80 ms). SIB3 = intra + serving priority. SIB5 = inter-frequency priority, thresholds, neighbour list.", "Sheet 2 step 4"]),
        ("row", ["3", "Idle measurement and reselection", "Idle", "Always for higher-priority freqs; equal/lower after search threshold", "Higher: ThreshXhigh. Equal: ranking Rn/Rs. Lower: ThrshServLow and ThreshXlow. Timers must persist.", "Sheet 2 step 5–9"]),
        ("row", ["4", "RRC setup on the camped cell", "Idle → Connected", "Service request", "Dedicated idle priorities are discarded when the UE enters connected. The camped cell is the serving cell of this session.", "Idle §5.1.3.1"]),
        ("row", ["5", "Classify the HO function", "Connected", "Coverage / service / quality / FreqPri / MLB", "Necessary coverage HO has higher priority than unnecessary load or optimisation HO.", "Connected Table 4-5"]),
        ("row", ["6", "Deliver measurement (or blind only if containment is known)", "Connected", "HO function started", "FREQ_MEAS_FLAG, not forbid-meas, object-count limits. Equal-priority objects may be picked at random if over cap — that is not load balance.", "Connected Fig 4-1"]),
        ("row", ["7", "Coverage protection", "Connected", "Serving crosses A2", "A1 stops coverage meas. A3 relative / A4 target good enough / A5 serving poor AND target good. RSRP is the recommended quantity.", "Sheet 3"]),
        ("row", ["8", "Frequency-priority HO (optional)", "Connected Ch.11", "Serving is good enough (A1) and a high-priority frequency is usable (A4)", "Puts service on high band and keeps low band for coverage. A frequency that is a FreqPri target must not have a reverse MLB-target pair.", "Fig 11-1 / 11-2"]),
        ("row", ["9", "MLB evaluates load", "MLB", "Periodic eval", "User-number load = N/C. Difference = (Load_s − Load_t) / Load_s. Use ActiveUe + SpectralEff when bandwidth or SE differ.", "Sheet 4"]),
        ("row", ["10", "MLB trigger and target admit", "MLB", "N ≥ thd+offset for the judge period (or PRB path)", "Target must pass NoHoFlag, overlap, MLB target indication, HO success, Low/Med HW+transport.", "MLB pp.27–29"]),
        ("row", ["11", "MLB picks UEs and volume", "MLB", "After admit", "ONLY_STRONGEST_CELL is recommended. Cap volume (do not use 5 s eval with MlbMaxUeNum ≥ 40).", "Table 6-5, p.136"]),
        ("row", ["12", "Connected load HO executes", "MLB + Connected", "UE selected", "FDD: measurement-based HO, event A4 (or A5 if licensed non-cosited). Cause: Reduce Load in Serving Cell.", "MLB §6.1.1.5.2"]),
        ("row", ["13", "Idle load transfer at release", "MLB + Idle", "Idle MLB trigger", "RRC release with dedicated reselection priority and T320. Class order: NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN.", "MLB §5.1.1.5"]),
        ("row", ["14", "Back to idle / next session", "Idle", "Release, T320 expiry, or next RRC", "Dedicated priorities last until T320, next RRC, or PLMN select. SI change is not instant — wait the next SI modification period.", "Idle §7.1.3"]),
        ("space", 8),
        ("section", "Priority between these steps"),
        ("row", ["A", "Coverage always wins", "Connected", "A2/A5 running", "Do not let MLB or FreqPri pull a UE off a dying serving cell.", "Table 4-5"]),
        ("row", ["B", "FreqPri and MLB must not fight", "Connected + MLB", "Same frequency pair", "No reverse MLB target on a FreqPri pair. MlbBasedFreqPriHoSwitch lets MLB own heavy load.", "Connected p.303"]),
        ("row", ["C", "Idle MLB and connected user-number MLB", "Idle + MLB", "Both ON", "Do not combine fixed-proportion idle MLB with user-number connected MLB (ping-pong).", "MLB §5.4.2.2"]),
        ("row", ["D", "Admission", "Connected", "Unnecessary / offload HO", "Target must admit ALL QCIs. Prep fail is often admission, not RF.", "Tables 4-16 / 4-17"]),
        ("space", 8),
        ("mml", MML_ALL),
    ]


def idle():
    return [
        ("section", "Procedure chart"),
        ("flow", ["PLMN select", "Criterion S", "Camp + read SI", "Measure", "Reselect", "RRC on camped cell"],
         "Idle Mode Management Fig 4-1 and Fig 5-1"),
        ("space", 6),
        ("section", "Step by step"),
        ("heads", H_STEP),
        ("row", ["1", "PLMN selection", "NAS / PLMN", "Happens before cell selection. SIB reselection priority does not control this.", "Wrong PLMN looks like a reselection fault.", "Fig 5-1"]),
        ("row", ["2", "Initial / stored / release-directed cell selection", "CELLSEL ; RRCConnectionRelease redirect", "Last camped, stored info, or strongest cell per band. SIB priority does not fully control first camp after power-up.", "Do not expect SIB priority to fix first access after reset.", "§5.1.2"]),
        ("row", ["3", "Criterion S (suitable cell)", "QRxLevMin, QQualMin, UePowerMax / PMax", "Srxlev = Qrxlevmeas − (Qrxlevmin + offset) − Pcompensation > 0. Squal > 0 only if QQualMin is set.", "Too-high floor = access failure. This is not a load-balance knob.", "§5.1.2"]),
        ("row", ["4", "Camp and read SI", "SIB1 / SIB3 / SIB4 / SIB5", "SIB3: intra + serving common priority. SIB5: inter-freq priority, thresholds, offset, neighbour list. Max 16 non-serving E-UTRAN frequencies.", "Missing SIB5 frequency = UE cannot reselect there.", "Table 7-1"]),
        ("row", ["5", "Intra-frequency measurement", "SIntraSearch / SIntraSearchQ", "UE measures intra unless serving is above the search threshold. Huawei: set CfgInd = CFG.", "Always-on intra meas if not CFG.", "§5.1.3.3"]),
        ("row", ["6", "Inter-frequency measurement", "SNonIntraSearch ; CellReselPriority", "Higher-priority frequency: always measured. Equal/lower: only after search threshold. Search threshold does not stop high-priority search.", "Huawei example SNonIntraSearch = 10; SIntraSearch > SNonIntraSearch.", "§5.1.3.3 ; §5.4.1.1"]),
        ("row", ["7", "Reselection to a higher-priority frequency", "ThreshXhigh / ThreshXhighQ / EutranReselTime", "Camped > 1 s AND target S > high threshold for the timer.", "Too-low ThreshXhigh = premature high-band camp.", "Tables 5-1 / 5-2"]),
        ("row", ["8", "Reselection to equal-priority frequency", "Qhyst, QoffsetFreq, CellQoffset", "Rn = Qmeas,n − Qoffset ; Rs = Qmeas,s + Qhyst ; Rn > Rs for the timer. Ranking is RF, not load.", "Do not use static offset as hourly load control.", "§5.1.3.4"]),
        ("row", ["9", "Reselection to a lower-priority frequency", "ThrshServLow / ThreshXlow", "No higher target AND serving S < serving-low AND target S > X-low, for the timer.", "Too-low serving-low = UE stays on a dying serving cell.", "Tables 5-3 / 5-4"]),
        ("row", ["10", "Dedicated priority at RRC release", "IdleModeMobilityControlInfo ; T320", "Replaces common SIB priorities until T320, next RRC, or PLMN select. Missing serving freq in the dedicated list is treated as lowest.", "Idle MLB / SPID / PCC use this path. SPID/PCC T320 is always 180 min.", "§5.1.3.1"]),
        ("row", ["11", "Next RRC setup", "Camped cell", "UE establishes RRC on the cell where it is then camped. Idle distribution = next access layer.", "A long connected session cannot be moved by idle MLB.", "Fig 5-1"]),
        ("space", 8),
        ("section", "Huawei cautions (read with the steps above)"),
        ("heads", H_NOTE),
        ("row", ["1", "Priority is per frequency, not per cell", "No SIB5 priority ⇒ no reselection to that frequency. Max 16 non-serving E-UTRAN frequencies.", "Set CellReselPriorityCfgInd = CFG on every frequency the UE must be able to reselect.", "Step 4, 7–9", "§5.1.3.1"]),
        ("row", ["2", "Capacity / hotspot above coverage layer", "Huawei recommendation for broadcast reselection priority.", "Do not publish operator-specific numbers as book defaults. Design the ranks; keep coverage layer under capacity layers.", "Step 7", "MLB §5.1.2.1"]),
        ("row", ["3", "UNDELIVER blocks idle MLB", "MeasPerformanceDemand = UNDELIVER removes the frequency from SIB5. It cannot be an idle-MLB target.", "Keep NORMAL on frequencies idle MLB must use.", "Step 4, 10", "§5.3.2.3"]),
        ("row", ["4", "SI is delayed", "Applied in the next SI modification period; otherwise after change-paging or 3 hours. SIB BER must be ≤ 1%.", "Do not judge camping in the same 15-min as the change.", "Step 4", "§7.1.3, §5.3.4"]),
        ("row", ["5", "Do not use the optimisation-table sentence on QRxLevMinOffset", "It conflicts with the Criterion-S formula.", "Treat Criterion S as written in §5.1.2.", "Step 3", "§5.1.2 vs §5.4.1.1.3"]),
        ("row", ["6", "Adaptive-proportion idle balancing is not recommended without Huawei support", "Fixed-proportion idle + user-number connected MLB causes ping-pong.", "Keep that pair off together.", "Step 10", "MLB §5.5.2.1 ; §5.4.2.2"]),
        ("row", ["7", "Neighbour list cap 16", "SIB listed neighbours are truncated.", "Truncation looks like a threshold problem.", "Step 4", "§5.1.3.4"]),
        ("space", 8),
        ("section", "What idle gives / where it stops"),
        ("heads", ["Type", "Item", "Document statement", "Condition", "User impact", "Source"]),
        ("row", ["Gives", "Low signalling vs connected HO", "Idle transfer avoids gap-assisted inter-frequency measurement and connected HO.", "Dedicated priority / idle MLB", "Lower immediate gap loss", "MLB Figs 4-4 / 4-5"]),
        ("row", ["Gives", "Sets next access cell", "Next RRC starts on the camped cell.", "After reselection or dedicated prio", "Long-term layer mix", "Fig 5-1"]),
        ("row", ["Stops", "Not real-time", "Dedicated priorities discarded at connected / T320 / PLMN select.", "Long RRC session", "Cannot rebalance an already-connected UE", "§5.1.3.1"]),
        ("row", ["Stops", "Highest common priority collects most idle UEs", "Equal-priority still ranks by RF, not load.", "No idle MLB", "Need idle MLB for dynamic load", "§5.1.3.5"]),
        ("space", 8),
        ("section", "Parameters in the same order as the procedure"),
        ("heads", H_PAR),
        ("row", ["1", "CELLSEL", "QRxLevMin / QQualMin / offsets / UePowerMax", "Selection floor (Criterion S)", "Not a load knob", "§5.1.2"]),
        ("row", ["2", "CELLRESEL", "CellReselPriority", "SIB3 serving common priority", "Capacity/hotspot above coverage (Huawei note)", "§5.1.3.1"]),
        ("row", ["3", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd + CellReselPriority", "SIB5 target priority", "CFG required or UE will not reselect there", "§5.1.3.1"]),
        ("row", ["4", "CELLRESEL", "SIntraSearchCfgInd / SNonIntraSearchCfgInd + values", "When to start intra / equal-lower inter search", "CFG; example SNonIntraSearch=10; SIntra > SNonIntra", "§5.4.1.1"]),
        ("row", ["5", "EUTRANINTERNFREQ", "ThreshXhigh / ThreshXlow / QoffsetFreq / EutranReselTime", "Higher / lower / equal-prio qualification", "Calibrate from MR. No universal dBm in the book.", "Tables 5-1 to 5-4"]),
        ("row", ["6", "CELLRESEL", "ThrshServLow / Qhyst / TReselEutran", "Permission to leave serving; stickiness", "Align serving-low with connected A2/A5 thinking", "Tables 5-3 / 5-4"]),
        ("row", ["7", "EUTRANINTERNFREQ", "MeasPerformanceDemand", "NORMAL / REDUCED / UNDELIVER", "UNDELIVER blocks idle MLB target", "§5.1.3.3"]),
        ("row", ["8", "EUTRANINTERNFREQ", "MlbTargetInd", "ALLOWED / WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB", "Coverage NoHoFlag is separate", "MLB pp.28, 129"]),
        ("row", ["9", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch (+ InterFreqMlbSwitch)", "Idle MLB enable", "License + SIB5 NORMAL + MlbTargetInd", "§5.3.2.3"]),
        ("row", ["10", "RRCCONNSTATETIMER", "T320ForLoadBalance", "Lifetime of load-balance dedicated priorities", "SPID/PCC path stays 180 min", "§5.1.3.1"]),
        ("space", 8),
        ("section", "How Idle locks to the other two books"),
        ("heads", H_STEP),
        ("row", ["→ Connected", "Dedicated idle prio discarded at RRC connect", "Align ThreshXhigh with A1-escape and ThrshServLow with A5-return", "Do not change idle and coverage A2 on the same day without a design", "Sheet 3", "§5.1.3.1"]),
        ("row", ["→ MLB", "Idle transfer method = dedicated prio + T320", "InterFreqIdleMlbSwitch + idle-allowed MlbTargetInd + SIB5 NORMAL", "Class order includes UTRAN/GERAN — not a dump tool", "Sheet 4", "MLB §5.1.1.5"]),
        ("space", 8),
        ("section", "Idle activation sequence"),
        ("heads", H_MML),
        ("row", ["0", "LST", "Dump CELLRESEL, EUTRANINTERNFREQ, CELLALGOSWITCH, RRCCONNSTATETIMER", "Read-only", "Keep the LST with the change record", "MAE"]),
        ("row", ["1", "CELLRESEL", "Serving priority + CFG search; SIntraSearch > SNonIntraSearch", "SIB3 plan ready", "Huawei example SNonIntraSearch=10. Capacity/hotspot above coverage.", "§5.4.1.1"]),
        ("row", ["2", "EUTRANINTERNFREQ", "Every needed EARFCN: CfgInd=CFG, priority, MeasPerformanceDemand=NORMAL", "Frequency exists in SIB5 plan", "Repeat per non-serving frequency", "§5.1.3.1"]),
        ("row", ["3", "EUTRANINTERNFREQ", "ThreshXhigh / ThreshXlow / QoffsetFreq / EutranReselTime from MR", "Priority ranks already set", "No universal dBm in the feature book", "Tables 5-1 to 5-4"]),
        ("row", ["4", "CELLRESEL", "ThrshServLow aligned with coverage A2/A5 philosophy", "Connected coverage design known", "Do not set so low that a dying cell never yields", "Tables 5-3 / 5-4"]),
        ("row", ["5", "EUTRANINTERNFREQ", "MlbTargetInd as designed (idle allow or forbid)", "Coverage NoHoFlag remains PERMIT if coverage HO is required", "WITHOUT_IDLE_MLB blocks only idle MLB, not coverage HO", "pp.28, 129"]),
        ("row", ["6", "CELLALGOSWITCH then T320", "Idle MLB bit after 1–5. Then T320ForLoadBalance.", "License present; Blind bit stays off unless designed", "Wait next SI modification period before judging camp", "§7.1.3"]),
        ("space", 8),
        ("mml", MML_IDLE),
    ]


def connected():
    return [
        ("section", "Procedure chart"),
        ("flow", ["Start HO function", "Meas or blind", "Deliver meas config", "UE reports A1–A5", "Pick target + admit", "Execute HO"],
         "Connected Mode Fig 4-1 §§4.1.1–4.1.8. Full MLB algorithm is in the MLB book, not here."),
        ("space", 6),
        ("section", "Step by step"),
        ("heads", H_STEP),
        ("row", ["1", "Start the HO function and classify it", "Coverage / FreqPri / MLB / service / quality", "Necessary coverage vs unnecessary offload vs unnecessary optimisation.", "Wrong class = load HO stealing a coverage rescue.", "pp.21–28 ; Table 4-5"]),
        ("row", ["2", "Choose measurement-based or blind", "Fig 4-2 ; BlindHoPriority", "Blind only when mobility is immediate and the neighbour contains the serving coverage.", "Higher access-failure risk.", "pp.31–32"]),
        ("row", ["3", "Deliver measurement configuration", "EUTRANINTERNFREQ + report config", "FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID deselected for required targets; object-count limits.", "Over cap: equal-priority frequencies may drop at random — not load balance.", "Tables 4-1 to 4-4, p.125"]),
        ("row", ["4", "Event A1 — serving becomes good", "A1: Ms − Hys > Thresh for TTT", "Stops coverage measurements. Can start A1-based FreqPri.", "Keep A1/A2 hysteresis consistent.", "Table 4-8"]),
        ("row", ["5", "Event A2 — serving becomes poor", "A2: Ms + Hys < Thresh for TTT", "Starts inter-frequency coverage measurement. Separate A2 families for A3, A4/A5, IRAT, blind.", "Wrong A2 family = wrong HO.", "Tables 5-1, 5-3, 5-10"]),
        ("row", ["6", "Event A3 — neighbour relatively better", "Mn+Ofn+Ocn−Hys > Ms+Ofs+Ocs+A3Offset", "Ofn/Ofs frequency offsets; Ocn/Ocs CIO.", "Large CIO can mask RF overshoot.", "Table 5-16"]),
        ("row", ["7", "Event A4 — neighbour absolutely good", "Mn+Ofn+Ocn−Hys > Thresh", "Used by MLB and FreqPri: target need only be good enough, not better than serving.", "A4 must be better than coverage A2 or the UE ping-pongs.", "Table 5-22 ; Table 11-5"]),
        ("row", ["8", "Event A5 — serving poor AND neighbour good", "Ms+Hys < Th1 AND Mn+Ofn+Ocn−Hys > Th2", "Strongest coverage-protection semantics. MLB A5 needs extra non-cosited license.", "Use A4 for co-sited FDD MLB unless that license exists.", "Tables 5-18 / 5-19"]),
        ("row", ["9", "Pick target and admit", "Best filtered target", "Necessary HO: any QCI. Unnecessary / offload: ALL QCIs. Inter-eNB unnecessary HO does not immediately try the next target after admit fail.", "Prep fail is often admission, not RF.", "Tables 4-16 / 4-17"]),
        ("row", ["10", "Execute, then punish / retry on fail", "Res / Opt / NonRes punish timers", "Do not treat every prep fail as an RF-threshold problem.", "Fix X2/admission before tightening A4.", "pp.61–65"]),
        ("space", 8),
        ("section", "Huawei cautions"),
        ("heads", H_NOTE),
        ("row", ["1", "Use RSRP as the general trigger quantity", "RSRQ moves with scheduler load and can oscillate.", "Keep RSRP unless there is a documented RSRQ design.", "Steps 4–8", "Table 4-15"]),
        ("row", ["2", "A4 must be better than the relevant coverage A2", "Otherwise coverage and load HO fight.", "Set A4 after coverage A2, from MR, not from MML examples.", "Step 7", "Table 5-22 ; Table 11-8"]),
        ("row", ["3", "A4 TTT = 5120 ms disables FreqPri, CQI and service-based inter-frequency HO", "Treat 5120 ms as off, not as slow.", "Use a real TTT if FreqPri or MLB A4 is required.", "Step 7", "Table 4-9 p.48"]),
        ("row", ["4", "Example A1/A2 −85/−87 dBm and A4 −103 dBm are command examples, not design values", "Copying them live is a design error.", "Calibrate from MR.", "Steps 4–8", "§11.4.1.2"]),
        ("row", ["5", "Equal-priority measurement objects may be selected randomly", "That is not load balance.", "Raise MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum to cover every needed frequency.", "Step 3", "§5.3.1.2 p.125"]),
        ("row", ["6", "FreqPri target A→B must not have reverse MLB target B→A", "Ping-pong warning.", "Remove the reverse pair. Use MlbBasedFreqPriHoSwitch when MLB should own heavy load.", "Step 1, 8", "p.303 ; Table 11-7"]),
        ("row", ["7", "SMeasure can silently suppress A4", "UE may skip inter-freq meas while serving RSRP is above SMeasure.", "Check HOMEASCOMM if A4 never arrives.", "Step 3", "§4.1.5"]),
        ("space", 8),
        ("section", "What connected mobility gives / where it stops"),
        ("heads", ["Type", "Item", "Document statement", "Condition", "User impact", "Source"]),
        ("row", ["Gives", "Coverage rescue", "A2 → A3/A4/A5 avoids drop", "Necessary HO priority", "Indoor / edge", "Table 4-5"]),
        ("row", ["Gives", "A4 offload gate", "Target need only be good enough", "MLB / FreqPri", "Can leave a still-usable source", "Table 5-22"]),
        ("row", ["Gives", "FreqPri + A1", "Service on high band; low band kept for coverage", "Fig 11-1 / 11-2", "Coverage-layer protection intent", "pp.299–300"]),
        ("row", ["Stops", "Not a load algorithm", "Equal static priority ≠ instantaneous load share", "Need MLB book", "Random object pick among equal prio", "p.125"]),
        ("row", ["Stops", "Measurement gaps steal DL TTIs", "Gap pattern cost", "Old UE / VoLTE", "Prefer A1/A2 gated meas", "Fig 4-10"]),
        ("space", 8),
        ("section", "Parameters in procedure order"),
        ("heads", H_PAR),
        ("row", ["1", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG / HO_TRG_FREQ_FORBID_MEAS_FLAG", "Whether the frequency is measured / allowed as HO target", "Silent no-HO if wrong", "§4.1.4.1.2"]),
        ("row", ["2", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "How many objects can be delivered", "Random drop if over limit", "Tables 4-3 / 4-4"]),
        ("row", ["3", "HOMEASCOMM", "SMeasure", "Skip inter-freq meas when serving is strong", "Can hide A4", "§4.1.5"]),
        ("row", ["4", "INTERFREQHOGROUP", "A1/A2 hyst + TTT + correct A2 family", "Coverage start/stop", "Wrong family = wrong HO", "Tables 5-3, 5-10"]),
        ("row", ["5", "INTERFREQHOGROUP", "A3 / A4 / A5 thresholds, hyst, TTT", "Relative vs absolute vs dual-condition HO", "A4 TTT ≠ 5120 ms if FreqPri/MLB A4 needed", "Tables 5-16, 5-22, 5-18"]),
        ("row", ["6", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType + IfMlbThdRsrpOffset + FreqPriHoA4ThldRsrpOffset", "Which event MLB/FreqPri uses, per-freq A4 offset", "A5 needs non-cosited MLB license", "MLB Table 6-3"]),
        ("row", ["7", "FreqPri switches", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch / FreqPriInHoProtectionTimer", "Who owns heavy load; bounce-back guard", "No reverse MLB pair", "Table 11-7 ; p.300"]),
        ("space", 8),
        ("section", "How Connected locks to the other two books"),
        ("heads", H_STEP),
        ("row", ["← Idle", "Dedicated prio already discarded", "Idle only set the serving cell of this session", "Align idle thresholds with A1/A5", "Sheet 2", "Idle §5.1.3.1"]),
        ("row", ["→ MLB", "This book is the HO engine; MLB is the load brain", "ONLY_STRONGEST_CELL and A4 live in the MLB book", "LOAD_COVERAGE_MEAS_DECOUPLE_SW if coverage meas already on", "Sheet 4", "MLB Table 6-5 ; p.139"]),
        ("space", 8),
        ("section", "Connected activation sequence"),
        ("heads", H_MML),
        ("row", ["0", "LST", "Dump INTERFREQHOGROUP, EUTRANINTERNFREQ, NCELL, CELLHOPARACFG, HOMEASCOMM", "Read-only", "Do not paste example −85/−87/−103 dBm", "§11.4.1.2"]),
        ("row", ["1", "EUTRANINTERNFREQ", "Meas flag on; forbid-meas off for required HO targets", "NRT exists", "Most silent A4 failures are flags/NRT, not dBm", "§4.1.4.1.2"]),
        ("row", ["2", "CELLUEMEASCONTROLCFG", "Object cap ≥ number of inter-frequency objects this cell must measure", "Frequency plan known", "Otherwise equal-prio freqs drop at random", "p.125"]),
        ("row", ["3", "INTERFREQHOGROUP", "Coverage A1/A2/A5 from MR, correct A2 family", "Coverage design first", "Rescue before offload", "Tables 5-3, 5-10"]),
        ("row", ["4", "INTERFREQHOGROUP", "A4 from MR; better than coverage A2; TTT not 5120 ms if FreqPri/MLB A4 is required", "Step 3 done", "Main absolute target gate", "Table 4-9"]),
        ("row", ["5", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType = A4 (A5 only with non-cosited MLB license)", "MLB will be used", "Co-sited FDD typically A4", "MLB Table 6-3"]),
        ("row", ["6", "FreqPri coordination", "MlbBasedFreqPriHoSwitch / LoadTrigger / non-zero incoming protect timer; no reverse MLB pair", "Confirm exact MO name in MAE", "Stops FreqPri fighting MLB", "p.303"]),
        ("space", 8),
        ("mml", MML_CONN),
    ]


def mlb():
    return [
        ("section", "Procedure chart"),
        ("flow", ["Eval load N/C", "Trigger thd + offset", "Admit target", "Select UEs", "A4/A5 HO or idle release", "Penalty / stop"],
         "MLB Fig 3-1. Equalisation vs offload: Fig 4-1. Idle vs connected transfer: Figs 4-4 / 4-5. Load = N/C. Difference = (Load_s − Load_t) / Load_s."),
        ("space", 6),
        ("section", "Step by step"),
        ("heads", H_STEP),
        ("row", ["1", "Choose equalisation or offload", "Peer load available or not", "Equalisation when source and target load are known. Offload if not.", "Equalisation is safer when X2 / intra-eNB load exists.", "Fig 4-1 §4.1"]),
        ("row", ["2", "Choose the load indicator", "UE number vs PRB vs HW/transport", "UE number ≈ short non-GBR. PRB ≈ GBR / sustained. HW/transport = Low/Med/High/OverLoad.", "PRB MLB does not move CA UEs and is burst-sensitive.", "Fig 4-2, Fig 4-3"]),
        ("row", ["3", "Compute user-number load", "Load = N/C", "ActiveUeBasedLoadEvalSw uses DL-buffer UEs. SpectralEffBasedLoadEvalSw uses PRB, scale, GBR, SE (refresh 1 min if ≥10 UL-sync UEs).", "Raw UE-count on unequal bandwidth can reduce DL throughput.", "§5.1.1 ; Table 5-5 ; §6.1.2.2"]),
        ("row", ["4", "Trigger connected UE-number MLB", "InterFreqMlbUeNumThd + MlbUeNumOffset + MlbTrigJudgePeriod", "Enter when N ≥ thd+offset for the whole judge period. Leave when N < thd.", "Mode UE_NUMBER_ONLY + SynchronizedUE.", "Table 6-2 ; p.128"]),
        ("row", ["5", "Admit the target cell / frequency", "MlbTargetInd, OverlapInd, NoHoFlag, HO SR, HW/transport", "WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB block one mode. Low/Med HW+transport. Not punished.", "Failed pair is not a target — fix RF/HO first (NCellHoSuccRateThld).", "pp.27–29"]),
        ("row", ["6", "Select UEs", "CELLMLBUESEL + protect timers", "UL-sync, not emergency, SPID/QCI, optional ARP/PRB/MCS/SNR. ONLY_STRONGEST_CELL recommended.", "Otherwise the UE is coverage-HO’d back.", "Table 6-5 p.159"]),
        ("row", ["7", "Limit volume and pick frequency", "MlbMaxUeNum ; FAIR / PRIORITY / LOADPRIORITY", "min(needed delta, hyst, max UE). LoadTransferEnhSw changes multi-target math.", "5 s eval + MlbMaxUeNum ≥ 40 over-transfers.", "pp.134–137"]),
        ("row", ["8", "Execute connected transfer", "A4 or A5 HO (FDD); optional blind", "Redirection generally not used in FDD. A5 needs extra license.", "Blind only with known containment.", "§6.1.1.5.2"]),
        ("row", ["9", "Execute idle transfer (if enabled)", "RRC release + dedicated prio + T320", "Class order NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN.", "Affects the next session only.", "§5.1.1.5"]),
        ("row", ["10", "Stop / punish", "CellPunishPrdNum × eval period", "Target reject (no radio resource) punishes the pair.", "Do not lower HO-success threshold to force MLB.", "p.29"]),
        ("space", 8),
        ("section", "Huawei cautions"),
        ("heads", H_NOTE),
        ("row", ["1", "ActiveUe when bandwidths differ; SpectralEff when SE differs a lot (e.g. >30%)", "Otherwise N/C is unfair across 15 vs 20 MHz.", "Turn load-model switches on before chasing the UE-number threshold.", "Step 3", "Table 5-5"]),
        ("row", ["2", "ONLY_STRONGEST_CELL", "Recommended cell-pick strategy.", "Without it, coverage HO bounces the UE back.", "Step 6", "Table 6-5"]),
        ("row", ["3", "Do not use InterFreqLoadEvalPrd = 5 s with MlbMaxUeNum ≥ 40", "Over-transfer.", "Lower volume or lengthen eval.", "Step 7", "p.136"]),
        ("row", ["4", "PRB-usage MLB skips CA UEs; gain dies when CA penetration ≳ 60%", "No experience fairness; bursty.", "Use as a supplement, not the main user-experience equaliser.", "Step 2", "Fig 6-2 ; §6.5"]),
        ("row", ["5", "Smart neighbour thds: 7 days collect, then refresh every 7 days", "Aggressive first-week seeds: extra CPU/HO and up to 5% TP swing.", "Need 15-min counter subscription. Do not treat learned thds as daily manual MOD targets.", "Step 4", "pp.29–31, 162"]),
        ("row", ["6", "Fixed-proportion idle + user-number connected MLB = ping-pong", "Adaptive-proportion idle not recommended without Huawei support.", "Keep that pair off together.", "Step 9", "§5.4.2.2 ; §5.5.2.1"]),
        ("row", ["7", "MLB and energy saving change each other’s targets", "A sleeping cell must not remain an MLB / high idle target.", "Coordinate ES.", "Step 5", "pp.153–156"]),
        ("space", 8),
        ("section", "What MLB gives / where it stops"),
        ("heads", ["Type", "Item", "Document statement", "Condition", "User impact", "Source"]),
        ("row", ["Gives", "True load equalisation", "Uses source and target load to reduce the difference", "Load exchange available", "Safer than blind offload", "Fig 4-1"]),
        ("row", ["Gives", "Idle + connected pair", "Idle steers next access; connected moves now", "T320 vs HO", "Idle has less gap/HO cost", "Figs 4-4 / 4-5"]),
        ("row", ["Stops", "Cannot fix RF, coverage, or band-capability mismatch", "Gain falls when coverage, UE band or PLMN differ", "p.39, 141, 214", "Not a substitute for RF", "§6.1.2.2"]),
        ("row", ["Stops", "Blind offload has no target-load check", "Can overload a cell that only looks low-load", "pp.167, 231", "Leave blind off unless designed", "Offload path"]),
        ("space", 8),
        ("section", "Parameters in procedure order"),
        ("heads", H_PAR),
        ("row", ["1", "CELLALGOSWITCH", "InterFreqMlbSwitch / IdleMlbSwitch / BlindMlbSwitch", "Master bits", "License first. Blind off unless designed.", "Table 6-2"]),
        ("row", ["2", "CELLMLB", "MlbTriggerMode / InterFreqUeTrsfType", "UE_NUMBER_ONLY + SynchronizedUE typical for experience equalisation", "Must match. PRB_ONLY skips CA UEs.", "Table 6-2"]),
        ("row", ["3", "eval SW", "ActiveUeBasedLoadEvalSw / SpectralEffBasedLoadEvalSw / LoadTransferEnhSw / CaUserLoadTransferSw", "Load model and CA", "Table 5-5 ; CA pp.129–136", "pp.25–27"]),
        ("row", ["4", "CELLMLB", "InterFreqMlbUeNumThd / Offset / Idle thd / MaxUeNum / eval periods", "When to enter/leave and how many UEs", "Enter = thd+offset; leave = thd. Never 5 s + MaxUe≥40.", "p.128, p.136"]),
        ("row", ["5", "CELLMLB", "MlbHoCellSelectStrategy / FreqSelectStrategy", "ONLY_STRONGEST_CELL ; FAIR / PRIORITY / LOAD", "Avoid coverage bounce", "Table 6-5"]),
        ("row", ["6", "EUTRANINTERNFREQ", "MlbTargetInd / MlbInterFreqHoEventType / IfMlbThdRsrpOffset", "Who may be a target; A4/A5; offset", "WITHOUT_* blocks one MLB mode", "pp.28, 129"]),
        ("row", ["7", "CELLMLBUESEL + punish / protect timers", "Who is movable ; re-MLB guard", "Protect emergency / QCI policy", "Do not pick edge UEs only for PRB", "pp.130–135"]),
        ("row", ["8", "RRCCONNSTATETIMER / EnhancedMlb", "T320ForLoadBalance / DediPrioManageOnLowLoadSw", "Idle dedicated-prio life / hold", "Idle MLB path", "§5.1.1.5 ; Table 5-10"]),
        ("space", 8),
        ("section", "How MLB locks to the other two books"),
        ("heads", H_STEP),
        ("row", ["← Idle", "Idle method = dedicated prio + T320 + SIB5 NORMAL", "MlbTargetInd can forbid idle targeting without blocking coverage HO", "UNDELIVER cannot be an idle target", "Sheet 2", "§5.1.1.5"]),
        ("row", ["← Connected", "HO engine A4/A5, meas flags, object cap, SMeasure, admission", "Trigger with 0 meas success = Connected sheet, not a thd problem", "ONLY_STRONGEST_CELL + A4 typical co-sited", "Sheet 3", "§6.1.1.5.2"]),
        ("space", 8),
        ("section", "MLB activation sequence"),
        ("heads", H_MML),
        ("row", ["0", "LST + license", "Dump CELLALGOSWITCH, CELLMLB, EUTRANINTERNFREQ, CELLMLBUESEL. Confirm Intra-RAT MLB license.", "MAE + parameter reference (Ch.8 p.299)", "Book does not list every default", "p.299"]),
        ("row", ["1", "EUTRANINTERNFREQ", "MlbTargetInd=ALLOWED, event A4, overlap and PERMIT_HO already valid", "Connected meas flags already correct", "WITHOUT_CONNECT_MLB / WITHOUT_IDLE_MLB to block one mode", "pp.28, 129"]),
        ("row", ["2", "Load model", "ActiveUe / SpectralEff / LoadTransferEnh as required (unequal BW / SE e.g. >30% / multi-target)", "Confirm bit names in MAE", "Do this before chasing UE-number thd", "Table 5-5"]),
        ("row", ["3", "CELLMLB", "UE_NUMBER_ONLY, SynchronizedUE, ONLY_STRONGEST_CELL, frequency strategy", "Targets already allowed", "Table 6-5 recommended pick", "Table 6-2"]),
        ("row", ["4", "CELLMLB", "UE-number thd/offset, MaxUeNum, judge/eval periods from live load — not 5 s + MaxUe≥40", "Load model ON", "Enter = thd+offset; leave = thd", "p.128, p.136"]),
        ("row", ["5", "CA / UE pick", "CaUserLoadTransferSw only if CA UEs must move. Then CELLMLBUESEL policy.", "CA plan", "If OFF, CA UEs stay filtered", "p.157"]),
        ("row", ["6", "Master bits last", "InterFreqMlbSwitch-1 and IdleMlbSwitch-1 if idle transfer is required. Blind stays 0.", "Steps 1–4 done. Idle also needs T320 and SIB5 NORMAL.", "Activate after targets exist", "Table 6-2"]),
        ("row", ["7", "Verify", "Load HO counters (Load vs UeNumLoad), HighLoad Dur/Num, meas success, DL active users, throughput, idle DedicatedPri. SON: Inter-Frequency Handover Statistics ; Idle Mode Release Statistics.", "15-min counters if smart thd is used", "When both PRB and UE-number MLB are on: PRB HO ≈ Load − UeNumLoad", "Tables 6-6, 6-21 ; §§6.1.4.2, 6.5.4.2"]),
        ("space", 8),
        ("mml", MML_MLB),
    ]


def activate():
    return [
        ("section", "One sequence across all three books"),
        ("note", "Do Idle camping first (next access layer), then Connected coverage and A4 (HO engine), then MLB load brain last. LST before every MOD. Placeholders: LocalCellId=<x>, DlEarfcn=<earfcn>. Confirm enums in MAE. Example dBm in the Connected book are not design values."),
        ("flow", ["LST baseline", "Idle SIB / priority / search", "Connected flags + coverage", "A4 gate from MR", "MLB target + load model", "MLB thd + master bit + T320"],
         "Activation order — not the sample-file MML"),
        ("space", 6),
        ("heads", ["Step", "Book", "MO", "What to do", "Stop if this is missing", "Source"]),
        ("row", ["0", "All", "LST *", "Dump current CELLRESEL, EUTRANINTERNFREQ, NCELL, INTERFREQHOGROUP, CELLHOPARACFG, HOMEASCOMM, CELLALGOSWITCH, CELLMLB, CELLMLBUESEL, RRCCONNSTATETIMER. Check Intra-RAT MLB license.", "No baseline = no rollback", "MAE"]),
        ("row", ["1", "Idle", "CELLRESEL", "Serving common priority. CFG search. SIntraSearch > SNonIntraSearch (example SNonIntraSearch=10). Capacity/hotspot ranks above coverage layer.", "SIB3 not planned", "Idle §5.4.1.1 ; MLB §5.1.2.1"]),
        ("row", ["2", "Idle", "EUTRANINTERNFREQ", "Every needed EARFCN: CfgInd=CFG, priority, NORMAL (not UNDELIVER), reselection thresholds from MR, MlbTargetInd as designed.", "Frequency missing from SIB5 plan", "Idle §5.1.3.1"]),
        ("row", ["3", "Idle", "CELLRESEL", "ThrshServLow aligned with later coverage A2/A5 thinking.", "Coverage A2 not yet known — set after step 5 or use current A2", "Idle Tables 5-3 / 5-4"]),
        ("row", ["4", "Connected", "EUTRANINTERNFREQ + NCELL", "Meas flag on; forbid-meas off for HO targets; symmetric NRT; NoHoFlag PERMIT; no PCI conflict.", "Most silent A4 failures start here", "Connected §4.1.4"]),
        ("row", ["5", "Connected", "CELLUEMEASCONTROLCFG + HOMEASCOMM", "Object cap ≥ needed freqs. SMeasure must not hide A4.", "Equal-prio random drop / silent A4", "p.125 ; §4.1.5"]),
        ("row", ["6", "Connected", "INTERFREQHOGROUP", "Coverage A1/A2/A5 from MR, correct A2 family. Then A4 better than that A2. A4 TTT ≠ 5120 ms if FreqPri or MLB A4 is required.", "Do not paste −85/−87/−103 dBm from the book", "§11.4.1.2 ; Table 4-9"]),
        ("row", ["7", "Connected", "FreqPri coordination", "If FreqPri is used with MLB: MlbBasedFreqPriHoSwitch / LoadTrigger as designed; incoming protect timer non-zero; no reverse MLB pair.", "Ping-pong A↔B", "p.303"]),
        ("row", ["8", "MLB", "EUTRANINTERNFREQ", "MlbTargetInd and MlbInterFreqHoEventType=A4 (A5 only with non-cosited MLB license).", "Overlap / PERMIT_HO / meas already done in steps 4–6", "MLB pp.28, 129 ; Table 6-3"]),
        ("row", ["9", "MLB", "Load-model switches", "ActiveUe if bandwidths differ. SpectralEff if SE differs a lot (e.g. >30%). LoadTransferEnh if multi-target.", "Do this before the UE-number threshold", "Table 5-5"]),
        ("row", ["10", "MLB", "CELLMLB", "UE_NUMBER_ONLY, SynchronizedUE, ONLY_STRONGEST_CELL, frequency strategy, thd/offset/MaxUeNum/periods. Never 5 s eval + MaxUe ≥ 40.", "Load model from step 9", "Table 6-2, 6-5 ; p.136"]),
        ("row", ["11", "MLB + Idle", "CELLALGOSWITCH + T320", "Master MLB bits last. Idle bit only if SIB5 NORMAL and T320 set. Blind bit stays 0 unless containment is proven.", "License + steps 1–10", "Table 6-2"]),
        ("row", ["12", "All", "Verify after SI delay", "Idle: wait next SI modification period; dedicated-pri counters. Connected: Coverage / FreqPri HO SR. MLB: Load HO, meas success, HighLoad, DL TP, SON handover / idle-release logs.", "Same 15-min as the change is too early for idle", "Idle §7.1.3 ; MLB Tables 6-6, 6-21"]),
        ("space", 8),
        ("section", "License reminders (confirm exact names in MAE)"),
        ("heads", ["No.", "License / package", "What it enables", "If missing", "Where used", "Source"]),
        ("row", ["1", "Basic idle + coverage A1–A5", "Usually in the LTE eNodeB package", "If missing, nothing else in this file will look right", "Sheets 2–3", "—"]),
        ("row", ["2", "Intra-RAT Mobility Load Balancing", "Connected and idle MLB bits", "Switch ON but Load HO / DedicatedPri counters stay 0", "Sheet 4", "MLB Table 6-2"]),
        ("row", ["3", "Intra-LTE Load Balancing for Non-cosited Cells", "Required for MLB event A5", "Stay on A4 without it", "Sheet 3–4", "MLB pp.142–143"]),
        ("row", ["4", "FreqPri / service-based mobility package", "A1 high-band steering", "Verify the sold name", "Sheet 3", "Ch.11"]),
        ("row", ["5", "Blind / CA-transfer / SON smart-thd options", "Sold separately in some packages", "Leave off if the license is not present", "Sheet 4", "MLB"]),
        ("space", 8),
        ("major", "SymbolShutdownSwitch in the old CSV sample is Symbol Power Saving. It is not a mobility command and is not used here."),
        ("space", 8),
        ("mml", MML_ALL),
    ]


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    wb.active.title = "tmp"
    build_sheet(wb, "Read me", "4G LTE Mobility Management — eRAN21.1 v3.1  |  How to read this file", overview(), BLUE)
    build_sheet(wb, "1. End-to-end chain", "1. End-to-end chain  |  Idle → Connected → MLB → Idle", chain(), "1F4E79")
    build_sheet(wb, "2. Idle Mode", "2. Idle Mode Management  |  eRAN21.1 Issue 04  |  Step by step", idle(), "008000")
    build_sheet(wb, "3. Connected Mode", "3. Mobility Management in Connected Mode  |  eRAN21.1 Issue 08  |  Step by step", connected(), "2E75B6")
    build_sheet(wb, "4. Intra-RAT MLB", "4. Intra-RAT Mobility Load Balancing  |  eRAN21.1 Issue 10  |  Step by step", mlb(), "C65911")
    build_sheet(wb, "5. Activation order", "5. Activation order  |  One sequence across the three books", activate(), BLUE)
    del wb["tmp"]
    wb.properties.title = "4G LTE Mobility Management eRAN21.1 v3.1"
    wb.properties.subject = "Step-by-step summary with Combined MML Command on every sheet"
    wb.properties.version = "3.1"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
