#!/usr/bin/env python3
"""Exclusive Huawei eRAN21.1 Mobility Management PDF summary.

Format (attached Excel sample):
  Title:   dark blue #005596, white bold, centered A-F
  Section: yellow #FFFF00, bold green #008000
  Header:  light blue #DDEBF7, black bold
  Data:    light grey #F2F2F2, black
  Six columns, spacer rows, Calibri, gridlines

Content: PDF summary only. No AI agent. No daily KPI engine.
Sources:
  Idle Mode Management, eRAN21.1, Issue 04
  Mobility Management in Connected Mode, eRAN21.1, Issue 08
  Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/docs/4G_LTE_Mobility_Management/Mobility_Management_eRAN21.1_Workbook.xlsx"

TITLE_BG, TITLE_FG = "005596", "FFFFFF"
SECTION_BG, SECTION_FG = "FFFF00", "008000"
THEAD_BG, THEAD_FG = "DDEBF7", "000000"
DATA_BG, DATA_FG = "F2F2F2", "000000"
WHITE, GRID = "FFFFFF", "B4B4B4"

thin = Border(
    left=Side(style="thin", color=GRID),
    right=Side(style="thin", color=GRID),
    top=Side(style="thin", color=GRID),
    bottom=Side(style="thin", color=GRID),
)
left = Alignment(wrap_text=True, vertical="center", horizontal="left")
center = Alignment(wrap_text=True, vertical="center", horizontal="center")
top = Alignment(wrap_text=True, vertical="top", horizontal="left")


def fill(c):
    return PatternFill("solid", fgColor=c)


def fnt(size=10, bold=False, color=DATA_FG):
    return Font(name="Calibri", size=size, bold=bold, color=color)


def put(ws, r, c, val, size=10, bold=False, color=DATA_FG, bg=DATA_BG, align=None, h=None):
    x = ws.cell(r, c, val)
    x.font = fnt(size, bold, color)
    x.fill = fill(bg)
    x.alignment = align or left
    x.border = thin
    if h:
        ws.row_dimensions[r].height = h


def merge_put(ws, r, c1, c2, val, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    put(ws, r, c1, val, **kw)
    bg = kw.get("bg", DATA_BG)
    color = kw.get("color", DATA_FG)
    size = kw.get("size", 10)
    bold = kw.get("bold", False)
    for c in range(c1 + 1, c2 + 1):
        y = ws.cell(r, c)
        y.fill = fill(bg)
        y.border = thin
        y.font = fnt(size, bold, color)


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
    ws.oddFooter.right.text = "Huawei eRAN21.1 PDF summary  |  Page &P of &N"
    ws.sheet_properties.tabColor = TITLE_BG
    ws.freeze_panes = "A3"


def title(ws, r, cols, text):
    merge_put(ws, r, 1, cols, text, size=16, bold=True, color=TITLE_FG, bg=TITLE_BG, align=center, h=28)
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
        put(ws, r, c, name, size=10, bold=True, color=THEAD_FG, bg=THEAD_BG, align=center, h=22)
    return r + 1


def row(ws, r, values, h=30):
    for c, v in enumerate(values, 1):
        put(ws, r, c, v, size=10, color=DATA_FG, bg=DATA_BG, align=top, h=h)
    return r + 1


def sheet(wb, name, title_text, sections):
    ws = wb.create_sheet(name)
    cols = 6
    widths(ws, [24, 12, 30, 38, 36, 56])
    setup(ws, title_text)
    r = 1
    r = title(ws, r, cols, title_text)
    for sec in sections:
        r = spacer(ws, r, cols)
        r = section(ws, r, cols, sec["name"])
        r = headers(ws, r, sec["headers"])
        for rec in sec["rows"]:
            r = row(ws, r, rec, h=sec.get("h", 30))
        r = spacer(ws, r, cols)
    return ws


# ---------------------------------------------------------------------------
# Cover — Mobility Management
# ---------------------------------------------------------------------------
def cover_sections():
    return [
        {
            "name": "Section 1: Feature Introduction",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 34,
            "rows": [
                ["Scope", "LTE FDD/TDD", "Mobility Management", "Idle + Connected + Intra-RAT MLB", "eRAN21.1 feature-parameter books",
                 "This workbook is an exclusive feature-parameter summary of the three Huawei PDFs only."],
                ["Feature 1", "LTE FDD/TDD", "Idle Mode Management", "Cell selection, reselection, dedicated priority, SI", "Issue 04, 2026-05-30",
                 "Controls camping and the cell used for the next RRC setup. Chart: PLMN → Criterion S → Camp → SIB3/SIB5 → Reselect → RRC. [IM] Fig 4-1, Fig 5-1."],
                ["Feature 2", "LTE FDD/TDD", "Mobility Management in Connected Mode", "A1/A2/A3/A4/A5, coverage HO, frequency-priority HO", "Issue 08, 2026-06-30",
                 "Executes connected HO. Necessary coverage HO has higher priority than unnecessary load/optimization HO. Load algorithm itself is in the MLB book. [CM] Table 3-1, Fig 4-1."],
                ["Feature 3", "LTE FDD/TDD", "Intra-RAT Mobility Load Balancing", "Idle dedicated-priority transfer + connected load HO", "Issue 10, 2026-06-30",
                 "Equalises or offloads inter-frequency LTE load. Chart: Eval load → Trigger → Admit target → Select UE → HO or idle release. [MLB] Fig 3-1, Fig 4-1."],
                ["MAJOR note", "LTE FDD/TDD", "How the three books fit", "Idle = next access layer; Connected = measurement/HO engine; MLB = who/when to move for load", "Use in this sequence",
                 "Do not treat connected frequency-priority HO as a substitute for the MLB algorithm. [CM] §3; [MLB] §3–§4."],
                ["MAJOR note", "LTE FDD/TDD", "Sample MML in the original CSV template", "ENodeBAlgoSwitch.SymbolShutdownSwitch", "Not a mobility parameter",
                 "That example belongs to Symbol Power Saving. It is not used in this summary."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (when each feature runs)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration"],
            "h": 30,
            "rows": [
                ["Idle reselection", "LTE FDD/TDD", "SIB3 / SIB5 priorities and thresholds", "After camp; higher-priority freqs always measured", "ThreshXhigh / ThrshServLow / ThreshXlow / Qoffset", "SI change applies in the next SI modification period, not instantly. [IM] §7.1.3"],
                ["Coverage HO", "LTE FDD/TDD", "A2 start, A1 stop, A3/A4/A5 report", "Serving quality crosses A2", "Separate A2 families for A3 vs A4/A5 vs blind vs IRAT", "Necessary HO preempts load HO. [CM] Table 4-5"],
                ["Frequency-priority HO", "LTE FDD/TDD", "A1 / A4 + FreqPri switches", "Serving good enough to place service on a high-priority frequency", "MlbBasedFreqPriHoSwitch can block FreqPri when MLB is active", "Designed to keep low band for coverage. [CM] Fig 11-1/11-2"],
                ["Connected MLB", "LTE FDD/TDD", "CELLMLB trigger mode UE-number or PRB", "N ≥ InterFreqMlbUeNumThd + Offset for MlbTrigJudgePeriod, or PRB ≥ thd + offset", "Load = N/C when user-number mode is used", "FDD normally uses measurement-based HO, not redirection. [MLB] §6.1.1"],
                ["Idle MLB", "LTE FDD/TDD", "RRCConnectionRelease + IdleModeMobilityControlInfo", "Idle-user load crosses InterFreqIdleMlbUeNumThd", "Dedicated priorities live until T320 / next RRC / PLMN select", "Affects the next session, not an already-connected long call. [IM] §5.1.3.1; [MLB] §5.1.1.5"],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (document reading order)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 28,
            "rows": [
                ["Read Feature 1", "LTE FDD/TDD", "Idle Mode Management sheet", "Summarise selection, priority, thresholds, dedicated priority", "SIB3/SIB5/T320", "Camping hierarchy and idle MLB interface"],
                ["Read Feature 2", "LTE FDD/TDD", "Connected Mode sheet", "Summarise events, meas, coverage vs FreqPri", "A1–A5, flags, object cap", "HO execution engine"],
                ["Read Feature 3", "LTE FDD/TDD", "Intra-RAT MLB sheet", "Summarise load model, trigger, target, UE pick, MML", "CELLMLB / MlbTargetInd", "Load transfer algorithm"],
                ["Validate on NE", "LTE FDD/TDD", "MAE-Access / parameter reference", "Confirm enum, default, license, MML syntax", "Version-matched spreadsheet (MLB Ch.8 points there)", "Printed feature books do not list every default/range"],
            ],
        },
    ]


# ---------------------------------------------------------------------------
# Feature 1 Idle
# ---------------------------------------------------------------------------
def idle_sections():
    return [
        {
            "name": "Section 1: Feature Introduction  (SN-1 Working Principal / SN-2 Major highlighted Point)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 38,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD/TDD", "Idle mode process",
                 "PLMN select → cell select (Criterion S) → camp → read SI → measure → reselect",
                 "RRC_IDLE until the next RRC connection",
                 "CHART (doc): PLMN → S → Camp → SIB3/SIB5 → Reselect → RRC on camped cell. [IM] Fig 4-1 §4 printed p.6; Fig 5-1 §5.1 pp.13–14."],
                ["Criterion S", "LTE FDD/TDD", "CELLSEL / CELLRESEL / EUTRANINTERNFREQ",
                 "Srxlev = Qrxlevmeas − (Qrxlevmin + offset) − Pcompensation > 0; Squal if QQualMin configured",
                 "If QQualMin is 0/absent, selection is RSRP-only",
                 "Suitability floor, not a load-balance knob. [IM] §5.1.2 pp.15–16."],
                ["Common priority", "LTE FDD/TDD", "CELLRESEL.CellReselPriority ; EUTRANINTERNFREQ.CellReselPriority",
                 "SIB3 serving priority; SIB5 non-serving priority when CfgInd=CFG",
                 "Larger value = higher priority; no priority ⇒ no reselection to that frequency",
                 "Priority is frequency-level, not per-cell. Max 16 non-serving E-UTRAN frequencies. [IM] §5.1.3.1 pp.17–18."],
                ["Measurement rule", "LTE FDD/TDD", "SIntraSearch / SNonIntraSearch / MeasPerformanceDemand",
                 "Higher-priority frequencies are always measured. Equal/lower only after search threshold",
                 "NORMAL / REDUCED / UNDELIVER",
                 "UNDELIVER removes the frequency from SIB5 and it cannot be an idle-MLB target. [IM] §5.1.3.3; §5.3.2.3."],
                ["SN-2 Major highlighted Point", "LTE FDD/TDD", "Document cautions",
                 "Initial selection is not priority balancing; SI apply delay; SIB neighbor max 16; SIB BER ≤1%",
                 "Huawei recommends SNonIntraSearchCfgInd=CFG and SIntraSearch > SNonIntraSearch; example SNonIntraSearch=10",
                 "Updated SI applies in the next SI modification period; otherwise reread after change-paging or 3 hours. [IM] §5.4.1.1; §7.1.3; [MLB] §5.1.2.1. Optimization-table text on QRxLevMinOffset conflicts with the Criterion-S formula — do not use that table sentence. [IM] §5.1.2 vs §5.4.1.1.3."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 Benefit and Limitations / SN-4 Selection criteria)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration"],
            "h": 34,
            "rows": [
                ["Benefit — idle transfer", "LTE FDD/TDD", "Dedicated priority in RRCConnectionRelease",
                 "Idle MLB / SPID / operator / PCC anchoring at release",
                 "Avoids gap-assisted inter-frequency measurement and connected HO",
                 "Lower immediate UX impact than connected HO. [MLB] §4.3 Figs 4-4/4-5 pp.18–19."],
                ["Limitation", "LTE FDD/TDD", "Dedicated priority lifetime",
                 "Discarded when UE enters connected, does PLMN select, or T320 expires",
                 "T320ForLoadBalance for MLB release; SPID/PCC always 180 min",
                 "Cannot rebalance a UE already in a long RRC session. [IM] §5.1.3.1 pp.18–22."],
                ["Limitation", "LTE FDD/TDD", "Adaptive-proportion idle balancing",
                 "Do not enable without Huawei engineering support",
                 "Huawei explicitly not recommended [MLB] §5.5.2.1 p.80",
                 "Fixed-proportion idle + user-number connected MLB causes ping-pong. [MLB] §5.4.2.2."],
                ["Higher-priority reselection", "LTE FDD/TDD", "ThreshXhigh / ThreshXhighQ + EutranReselTime",
                 "Camped >1 s AND target S > high threshold for the reselection timer",
                 "Always-on measurement of higher-priority frequencies",
                 "[IM] Tables 5-1/5-2 pp.29–32."],
                ["Lower-priority reselection", "LTE FDD/TDD", "ThrshServLow / ThreshXlow (+ Q forms)",
                 "No qualifying higher-priority target AND serving S < serving-low AND target S > X-low",
                 "Must persist for the reselection timer",
                 "[IM] Tables 5-3/5-4 pp.33–35."],
                ["Equal-priority ranking", "LTE FDD/TDD", "Qhyst / QoffsetFreq / CellQoffset",
                 "Rn > Rs for the reselection period; camped >1 s",
                 "Rn = Qmeas,n − (QoffsetFreq + CellQoffset); Rs = Qmeas,s + Qhyst",
                 "Larger positive target offset reduces reselection probability. [IM] §5.1.3.4 pp.27–28."],
                ["Idle MLB trigger", "LTE FDD/TDD", "InterFreqIdleMlbUeNumThd / IdleUE transfer type",
                 "Idle-user number condition holds through the judge period",
                 "Low-load target frequencies get higher dedicated reselection priority",
                 "Class order: NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN. [MLB] §5.1.1.5 p.33."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5 Activation parameter / Switch)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 32,
            "rows": [
                ["Broadcast priority", "LTE FDD/TDD", "CELLRESEL / EUTRANINTERNFREQ",
                 "Set CellReselPriority; set CellReselPriorityCfgInd=CFG on non-serving freqs",
                 "Without CFG the UE does not reselect to that frequency",
                 "Huawei: capacity/hotspot frequencies above the low-band coverage layer in broadcast priority. [MLB] §5.1.2.1 pp.38–39."],
                ["Search thresholds", "LTE FDD/TDD", "CELLRESEL",
                 "SIntraSearchCfgInd=CFG; SNonIntraSearchCfgInd=CFG",
                 "SIntraSearch > SNonIntraSearch; example SNonIntraSearch=10",
                 "CFG avoids continuous equal/lower-priority measurement and extra battery use. [IM] §5.4.1.1."],
                ["Meas performance", "LTE FDD/TDD", "EUTRANINTERNFREQ.MeasPerformanceDemand",
                 "NORMAL for main capacity / intensive-coverage frequencies",
                 "REDUCED / UNDELIVER restrict delivery",
                 "UNDELIVER must not be used as a congestion tool if idle MLB needs that frequency. [IM] §5.1.3.3; §5.3.2.3."],
                ["Idle MLB switch", "LTE FDD/TDD", "CELLALGOSWITCH.MlbAlgoSwitch",
                 "InterFreqIdleMlbSwitch (typically with InterFreqMlbSwitch)",
                 "Target frequency MlbTargetInd must allow idle MLB",
                 "ALLOWED_WITHOUT_IDLE_MLB blocks idle equalisation/offload to that frequency. [MLB] pp.28, 129."],
                ["Dedicated-priority hold", "LTE FDD/TDD", "EnhancedMlbAlgoSwitch",
                 "DediPrioManageOnLowLoadSw ; optional EnhSw for all RRC-release causes",
                 "T320ForLoadBalance",
                 "Prevents released UEs returning to higher-load frequencies. [MLB] §5.3.1 Table 5-10 pp.62–68."],
                ["T320", "LTE FDD/TDD", "RRCCONNSTATETIMER",
                 "T320ForLoadBalance for load-balance release; T320ForOther for other causes",
                 "SPID/PCC anchoring T320 is always 180 minutes",
                 "[IM] §5.1.3.1."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["SIB5 frequency list", "LTE FDD/TDD", "EUTRANINTERNFREQ", "All required non-serving LTE frequencies present", "Max 16 non-serving E-UTRAN freqs", "[IM] §5.1.3.1"],
                ["Neighbor broadcast", "LTE FDD/TDD", "EUTRANINTERFREQNCELL", "SIB4/SIB5 listed neighbors", "At most 16 listed neighbors per frequency", "Non-zero offsets affect which 16 are chosen. [IM] §5.1.3.4"],
                ["Blacklist", "LTE FDD/TDD", "InterFreqBlkCell.ApplicationScope", "Invalid cells only", "65535 = idle + connected; only 16 idle blacklists delivered", "[IM] §5.1.3.2"],
                ["SI quality", "LTE FDD/TDD", "SIB broadcast", "BER ≤ 1%", "Required for reselection", "[IM] §5.3.4"],
                ["UE capability", "LTE FDD/TDD", "incMonEUTRA / band support", "Legacy UEs may see only NORMAL and ≤8 dedicated freqs", "Enhanced UEs up to 16", "[IM] §5.1.3.1 / §5.1.3.3"],
                ["Forbidden combo", "LTE FDD/TDD", "Fixed-proportion idle + user-number connected MLB", "Do not combine", "Ping-pong warning", "[MLB] §5.4.2.2"],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 28,
            "rows": [
                ["Idle vs connected MLB", "LTE FDD/TDD", "Dedicated prio discarded at RRC connect", "Connected A1–A5 / MLB then own the session", "Coordinate, do not fight", "[IM] §5.1.3.1; [CM] §3"],
                ["Idle vs CA", "LTE FDD/TDD", "PCC anchoring + CaUserLoadTransferSw", "Changes whether CA-capable UEs are idle-steered", "PCC ≠ SCC load", "[MLB] Fig 5-3 pp.34–36"],
                ["RSRQ thresholds", "LTE FDD/TDD", "Squal / ThreshX*Q", "RSRQ moves with cell load", "Can oscillate", "Huawei generally prefers RSRP for connected trigger; same caution applies. [CM] Table 4-15"],
                ["Energy saving", "LTE FDD/TDD", "Carrier shutdown / dormancy", "Sleeping cell must not remain a high idle-priority target", "Mutual impact tables in MLB book", "[MLB] pp.153–156"],
                ["Redirection", "LTE FDD/TDD", "Load-based RRC redirect", "Not location/RF measured; QoS/priority based", "More disruptive than HO", "[IM] §6.1.1 pp.69–70"],
                ["GERAN/UTRAN", "LTE FDD/TDD", "Dedicated class order", "NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "Idle MLB class", "[MLB] §5.1.1.5 p.33"],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["Basic idle selection / SIB reselection", "LTE FDD/TDD", "CELLSEL / CELLRESEL / SIB3/SIB5", "Basic LTE eNodeB", "Normally included", "Confirm on the live license file."],
                ["Idle MLB / dedicated load priority", "LTE FDD/TDD", "InterFreqIdleMlbSwitch", "Intra-RAT Mobility Load Balancing family (confirm exact name)", "Same family as connected MLB", "If missing, switches may appear ON but idle dedicated-pri counters stay 0."],
                ["Blind idle / blind MLB", "LTE FDD/TDD", "InterFreqBlindMlbSwitch", "Blind MLB option if sold separately", "Documented as higher risk", "[MLB]"],
                ["DediPrioManageOnLowLoad(Enh)", "LTE FDD/TDD", "Enhanced MLB idle-priority management", "Verify enhanced MLB option", "Holds low-load dedicated priority", "[MLB] Table 5-10"],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence / connected order)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document / Huawei note", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["1", "LTE FDD/TDD", "CELLSEL", "QRxLevMin / QQualMin / offsets / UePowerMax", "Criterion S floors", "[IM] §5.1.2"],
                ["2", "LTE FDD/TDD", "CELLRESEL", "QRxLevMin / QQualMin / PMax", "Serving/intra reselection suitability", "[IM] §5.1.3.4"],
                ["3", "LTE FDD/TDD", "CELLRESEL", "CellReselPriority", "SIB3 serving common priority", "[IM] §5.1.3.1"],
                ["4", "LTE FDD/TDD", "EUTRANINTERNFREQ", "CellReselPriorityCfgInd / CellReselPriority", "SIB5 target priority; CfgInd=CFG required", "[IM] §5.1.3.1"],
                ["5", "LTE FDD/TDD", "EUTRANINTERNFREQ", "QRxLevMin / QqualMin / Pmax", "Target-frequency suitability", "[IM] §5.1.3.4"],
                ["6", "LTE FDD/TDD", "CELLRESEL", "SIntraSearchCfgInd / SIntraSearch / SIntraSearchQ", "Skip intra meas when serving very good; CFG recommended", "[IM] §5.1.3.3"],
                ["7", "LTE FDD/TDD", "CELLRESEL", "SNonIntraSearchCfgInd / SNonIntraSearch / Q", "Equal/lower inter-freq search start; example 10", "[IM] §5.4.1.1; [MLB] §5.1.2.1"],
                ["8", "LTE FDD/TDD", "EUTRANINTERNFREQ", "ThreshXhigh / ThreshXhighQ", "Higher-priority target qualification", "[IM] Tables 5-1/5-2"],
                ["9", "LTE FDD/TDD", "CELLRESEL", "ThrshServLow / ThrshServLowQ", "Permission to leave for lower priority", "[IM] Tables 5-3/5-4"],
                ["10", "LTE FDD/TDD", "EUTRANINTERNFREQ", "ThreshXlow / ThreshXlowQ", "Lower-priority target qualification", "[IM] Tables 5-3/5-4"],
                ["11", "LTE FDD/TDD", "CELLRESEL", "Qhyst / TReselEutran", "Serving stickiness / intra timer", "[IM] §5.1.3.4"],
                ["12", "LTE FDD/TDD", "EUTRANINTERNFREQ", "QoffsetFreq / EutranReselTime", "Equal-prio freq offset / inter timer", "[IM] §5.1.3.4"],
                ["13", "LTE FDD/TDD", "EUTRANINTERFREQNCELL", "CellQoffset", "Per-neighbor idle offset", "[IM] §5.1.3.4"],
                ["14", "LTE FDD/TDD", "EUTRANINTERNFREQ", "MeasPerformanceDemand", "NORMAL / REDUCED / UNDELIVER", "[IM] §5.1.3.3"],
                ["15", "LTE FDD/TDD", "CELLRESEL", "SpeedDepReselCfgInd + SfMedium/High", "Speed-based scaling", "[IM] §5.1.3.6"],
                ["16", "LTE FDD/TDD", "RRCCONNSTATETIMER", "T320ForLoadBalance / T320ForOther", "Dedicated-priority lifetime", "[IM] §5.1.3.1"],
                ["17", "LTE FDD/TDD", "CELLALGOSWITCH", "InterFreqIdleMlbSwitch / InterFreqMlbSwitch", "Idle MLB enable", "[IM] §5.3.2.3"],
                ["18", "LTE FDD/TDD", "EnhancedMlbAlgoSwitch", "DediPrioManageOnLowLoadSw / EnhSw", "Hold low-load dedicated prio", "[MLB] Table 5-10"],
                ["19", "LTE FDD/TDD", "EUTRANINTERNFREQ", "MlbTargetInd", "ALLOWED / ALLOWED_WITHOUT_IDLE_MLB / ALLOWED_WITHOUT_CONNECT_MLB", "[MLB] pp.28, 129"],
                ["20", "LTE FDD/TDD", "CELLMLB", "InterFreqIdleMlbUeNumThd / idle transfer type", "Idle user-number trigger", "[MLB] Table 5-2"],
                ["21", "LTE FDD/TDD", "GlobalProcSwitch", "CellReselectionOptSwitch", "Preferential LTE freq delivery in dedicated prio", "[IM] §5.3.2.3"],
                ["22", "LTE FDD/TDD", "EUTRANINTERNFREQ", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Filtered from dedicated-priority delivery; recommended only for SCC-only freqs", "[IM] §5.1.3.1"],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 36,
            "rows": [
                ["0", "LTE FDD/TDD", "—",
                 "LST CELLRESEL: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLALGOSWITCH: LocalCellId=<x>; LST RRCCONNSTATETIMER:;",
                 "Read-only", "Always dump current configuration. LocalCellId / DlEarfcn are placeholders. Validate syntax in MAE of the running eRAN version."],
                ["1", "LTE FDD/TDD", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<prio>, SNonIntraSearchCfgInd=CFG, SIntraSearchCfgInd=CFG, SNonIntraSearch=<val>;",
                 "SIntraSearch > SNonIntraSearch", "Document example SNonIntraSearch=10. Priority: capacity/hotspot above coverage layer. [IM] §5.4.1.1; [MLB] §5.1.2.1"],
                ["2", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;",
                 "Frequency exists in SIB5 plan", "Repeat per non-serving frequency that must be reselectable. Do not set UNDELIVER on a frequency that idle MLB must use."],
                ["3", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, ThreshXhigh=<val>, ThreshXlow=<val>, QoffsetFreq=<val>, EutranReselTime=<val>;",
                 "Matches reselection direction (higher vs lower priority)", "Calibrate from live MR. Feature book does not give a universal dBm design value."],
                ["4", "LTE FDD/TDD", "CELLRESEL",
                 "MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<val>;",
                 "Align with connected coverage A2/A5 philosophy", "Lower-priority fallback permission."],
                ["5", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=<ALLOWED | ALLOWED_WITHOUT_IDLE_MLB | ALLOWED_WITHOUT_CONNECT_MLB>;",
                 "Verify exact enum on the NE", "Use WITHOUT_IDLE_MLB / WITHOUT_CONNECT_MLB to stop a frequency being an MLB target while keeping ordinary coverage mobility."],
                ["6", "LTE FDD/TDD", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "Intra-RAT MLB license; target indications already set", "Do not enable InterFreqBlindMlbSwitch in the same step unless containment is proven."],
                ["7", "LTE FDD/TDD", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
                 "Idle MLB ON", "SPID/PCC path stays 180 min regardless of this parameter."],
                ["8", "LTE FDD/TDD", "Verify",
                 "LST the MOs above; wait next SI modification period; check L.RRCRel.load.DedicatedPri.LTE.High and L.RRCRel.Lowload.DedicatedPri.LTE.High",
                 "SI is not instant [IM] §7.1.3", "Idle dedicated-priority counters are in the MLB book."],
            ],
        },
    ]


# ---------------------------------------------------------------------------
# Feature 2 Connected
# ---------------------------------------------------------------------------
def connected_sections():
    return [
        {
            "name": "Section 1: Feature Introduction  (SN-1 / SN-2)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 38,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD/TDD", "Connected-mode mobility",
                 "Initiate → meas or blind → deliver meas → report → pick target → HO → punish/retry",
                 "Necessary coverage HO > unnecessary load/optimization HO",
                 "CHART (doc): Fig 4-1 §§4.1.1–4.1.8 pp.29–65. This book is the HO engine. The full load algorithm is in Intra-RAT MLB. [CM] Table 3-1 pp.21–23."],
                ["Event A1", "LTE FDD/TDD", "Serving becomes good",
                 "Enter: Ms − Hys > Thresh for TTT", "Stops coverage measurements; can start A1-based FreqPri in same-coverage multi-band",
                 "[CM] Table 4-8 pp.41–46; Table 5-1 pp.97–98."],
                ["Event A2", "LTE FDD/TDD", "Serving becomes poor",
                 "Enter: Ms + Hys < Thresh for TTT", "Starts inter-frequency coverage measurements",
                 "Separate A2 families for A3-based, A4/A5-based, IRAT and blind HO. [CM] Tables 5-3, 5-10."],
                ["Event A3 / A4 / A5", "LTE FDD/TDD", "Target decision",
                 "A3 relative; A4 absolute target good; A5 serving poor AND target good",
                 "A4 used by MLB/FreqPri as 'target good enough'",
                 "A3: Mn+Ofn+Ocn−Hys > Ms+Ofs+Ocs+Off. [CM] Table 5-16 p.127; Table 5-22 p.145; Tables 5-18/5-19."],
                ["SN-2 Major highlighted Point", "LTE FDD/TDD", "Document cautions",
                 "RSRP recommended; A4 must be better than coverage A2; A4 TTT=5120 ms disables FreqPri/CQI/service IFHO",
                 "Example MML −85/−87/−103 dBm are command examples, not design values",
                 "Equal-priority frequencies may be selected randomly for measurement objects. [CM] §5.3.1.2 p.125; Table 4-9 p.48; Table 4-15; §11.4.1.2 pp.317–318."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 / SN-4)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration"],
            "h": 32,
            "rows": [
                ["Coverage HO", "LTE FDD/TDD", "A2 → A3/A4/A5 ; CovBasedInterFreqHoMode",
                 "Serving crosses the applicable A2 family",
                 "IMMEDIATE / BASEDONSIGNALSTRENGTH / BASEDONFREQPRIORITY + waiting timer",
                 "Necessary HO: admit any QCI. [CM] §5.3.1.1–5.3.1.3; Tables 4-16/4-17."],
                ["Frequency-priority HO", "LTE FDD/TDD", "A4 + FreqPriA4 offset + FreqPriIFHoWaitingTimer",
                 "Place service on high band while keeping low band for coverage",
                 "MlbBasedFreqPriHoSwitch blocks FreqPri when specified MLB functions are active",
                 "LoadTriggerFreqPriHoSwitch needs overlap, obtainable load, neighbor not in UE-number MLB trigger, no PCI conflict. [CM] Figs 11-1/11-2; pp.301–303."],
                ["MLB HO (executed here)", "LTE FDD/TDD", "MlbInterFreqHoEventType A4 or A5",
                 "MLB source trigger already true (Feature 3)",
                 "A5 requires Intra-LTE Load Balancing for Non-cosited Cells license",
                 "FDD normally measurement-based HO. [MLB] §6.1.1.5.2 pp.137–139; Table 6-3."],
                ["Blind HO", "LTE FDD/TDD", "BlindHoPriority / InterFreqMlbBlindHo",
                 "Immediate mobility required and neighbor coverage contains serving",
                 "No candidate measurement",
                 "Higher access-failure risk. Use only when containment is known. [CM] Fig 4-2; Table 5-22 p.149."],
                ["Benefit / limitation", "LTE FDD/TDD", "Connected mobility",
                 "Always",
                 "Protects coverage; A4 allows offload without beating serving",
                 "Not itself a load-share algorithm. Meas gaps steal TTIs. Unnecessary (offload) HO must admit ALL QCIs. [CM] Tables 4-16/4-17."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 32,
            "rows": [
                ["Coverage meas", "LTE FDD/TDD", "INTERFREQHOGROUP",
                 "Configure A1/A2 hyst, TTT and the correct A2 family",
                 "Filter coefficient + TTT both add delay",
                 "Wrong A2 family ⇒ wrong HO behaviour. [CM] Tables 5-3, 5-10, 4-9."],
                ["A4 / FreqPri / MLB gate", "LTE FDD/TDD", "INTERFREQHOGROUP / EUTRANINTERNFREQ",
                 "Set InterFreqLoadBasedHoA4ThdRsrp/Rsrq, IfMlbThdRsrpOffset, FreqPriHoA4ThldRsrpOffset, A4 hyst/TTT",
                 "A4 better than coverage A2; do not set A4 TTT to 5120 ms if FreqPri is required",
                 "[CM] Table 5-22 pp.147–148; Table 11-5; Table 4-9 p.48."],
                ["A5 coverage", "LTE FDD/TDD", "INTERFREQHOGROUP",
                 "Configure A5 Thd1 (serving poor) and Thd2 (target good)",
                 "Dual condition",
                 "Strongest 'serving must be poor' semantics. [CM] Tables 5-18/5-19."],
                ["Meas eligibility", "LTE FDD/TDD", "EUTRANINTERNFREQ / CELLUEMEASCONTROLCFG",
                 "FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for needed targets",
                 "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum must cover required objects",
                 "SMeasure can silently skip inter-freq meas. [CM] §4.1.4.1.2; §4.1.5; Tables 4-3/4-4."],
                ["FreqPri vs MLB", "LTE FDD/TDD", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch / ReduceInvalidFreqPriHoSwitch",
                 "Enable so MLB owns heavy-load decisions",
                 "A FreqPri target must not have a reverse MLB-target relationship",
                 "Ping-pong warning. [CM] p.303; Table 11-7 p.313."],
                ["Protect / punish", "LTE FDD/TDD", "INTRARATHOCOMM / HOMEASCOMM / CELLOPHOCFG",
                 "FreqPriInHoProtectionTimer; HO fail punish timers; optional high-mobility forbid ~30 km/h",
                 "NLOS can misclassify high-mobility",
                 "[CM] p.300; Tables 4-16/4-17; [MLB] pp.136–141."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["Neighbor relation", "LTE FDD/TDD", "EUTRANINTERFREQNCELL", "NoHoFlag, CIO, PCI uniqueness", "Required both directions", "Missing NRT looks like a bad threshold."],
                ["Measurement objects", "LTE FDD/TDD", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "Must be ≥ number of needed inter-freq objects", "[CM] Tables 4-3/4-4"],
                ["SMeasure", "LTE FDD/TDD", "HOMEASCOMM", "Skip inter-freq meas while serving RSRP above SMeasure", "Verify actual RRC config", "[CM] §4.1.5 p.55"],
                ["Measurement gap", "LTE FDD/TDD", "AutoGapSwitch / GapPatternType / DedicatedGapPatternType", "Gap steals DL TTIs", "Prefer A1/A2 gated meas where supported", "[CM] Fig 4-10"],
                ["Admission", "LTE FDD/TDD", "HoAdmitSwitch / X2RoHoAdmitSwitch", "Necessary = any QCI; unnecessary = all QCIs", "Prep fail is often admission", "[CM] Tables 4-16/4-17"],
                ["CA / PCC", "LTE FDD/TDD", "Carrier Aggregation book", "PCC anchoring not fully specified in this CM book", "Misaligned PCC + FreqPri can ping-pong", "[CM] §3 p.23; §11.3.2.3 p.312"],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["FreqPri vs MLB reverse pair", "LTE FDD/TDD", "Target A→B and MLB B→A", "Forbidden combination", "Ping-pong", "[CM] p.303"],
                ["A4 vs coverage A2", "LTE FDD/TDD", "Threshold separation", "A4 must be better than coverage A2", "Ping-pong if too close", "[CM] Table 5-22; Table 11-8"],
                ["Incoming unnecessary HO", "LTE FDD/TDD", "FreqPriInHoProtectionTimer", "Blocks immediate FreqPri re-meas", "Bounce-back guard", "[CM] p.300"],
                ["LOAD_COVERAGE_MEAS_DECOUPLE_SW", "LTE FDD/TDD", "MLB book", "Allows load meas after coverage meas already delivered", "If coverage meas blocks MLB A4", "[MLB] p.139"],
                ["Virtual-grid smart carrier selection", "LTE FDD/TDD", "Concurrent with FreqPri", "Can cause ping-pong", "Do not combine casually", "[CM] p.312"],
                ["Related: Idle Mode", "LTE FDD/TDD", "Dedicated idle prio discarded at connect", "Align idle ThreshXhigh with A1-escape; ThrshServLow with A5-return", "Next access vs session", "[IM]/[CM]"],
                ["Related: Intra-RAT MLB", "LTE FDD/TDD", "This feature executes A4/A5 HO", "MLB decides who/when", "ONLY_STRONGEST_CELL recommended in MLB book", "[MLB] Table 6-5 p.159"],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["Coverage A1–A5 HO", "LTE FDD/TDD", "INTERFREQHOGROUP", "Basic LTE mobility", "Normally included", "Confirm lean packages."],
                ["Frequency-priority HO", "LTE FDD/TDD", "FreqPri function", "Frequency-priority / service-based mobility package — verify exact name", "If missing, A1-based high-band steering does not run", "[CM] Ch.11"],
                ["MLB A4 (co-sited)", "LTE FDD/TDD", "Intra-RAT MLB", "Intra-RAT Mobility Load Balancing", "See Feature 3", "[MLB]"],
                ["MLB event A5", "LTE FDD/TDD", "MlbInterFreqHoEventType=A5", "Intra-LTE Load Balancing for Non-cosited Cells", "Required only if event type is A5", "[MLB] pp.142–143"],
                ["Blind HO", "LTE FDD/TDD", "BlindHoPriority", "Blind handover option if sold separately", "Higher access risk", "[CM] Table 5-22"],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document / Huawei note", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["1", "LTE FDD/TDD", "EUTRANINTERNFREQ", "DlEarfcn / MeasBandWidth / QoffsetFreqConn", "Measurement object", "[CM] Table 4-2"],
                ["2", "LTE FDD/TDD", "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG / HO_TRG_FREQ_FORBID_MEAS_FLAG / INTER_FREQ_FILTER_FLAG", "Frequency filtering", "[CM] §4.1.4.1.2"],
                ["3", "LTE FDD/TDD", "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "Delivered object count", "[CM] Tables 4-3/4-4"],
                ["4", "LTE FDD/TDD", "HOMEASCOMM", "SMeasure", "Skip inter-freq meas when serving strong", "[CM] §4.1.5"],
                ["5", "LTE FDD/TDD", "CELLHOPARACFG", "EutranFilterCoeffRsrp / Rsrq", "L3 filter", "[CM] Table 4-14"],
                ["6", "LTE FDD/TDD", "ENODEBALGOSWITCH", "AutoGapSwitch / GapPatternType / DedicatedGapPatternType", "Measurement gap", "[CM] Fig 4-10"],
                ["7", "LTE FDD/TDD", "INTERFREQHOGROUP", "InterFreqHoA1A2Hyst / TimeToTrig", "A1/A2 stability", "[CM] Table 4-9"],
                ["8", "LTE FDD/TDD", "INTERFREQHOGROUP", "Coverage A2 threshold families", "A3 vs A4/A5 vs IRAT vs blind", "[CM] Tables 5-3, 5-10"],
                ["9", "LTE FDD/TDD", "INTERFREQHOGROUP", "InterFreqHoA3Offset / A3 Hyst / TTT / A3RsrqOffset", "Relative HO", "[CM] Table 5-16"],
                ["10", "LTE FDD/TDD", "EUTRANINTERFREQNCELL", "CellIndividualOffset", "Connected CIO (Ocn)", "[CM] pp.46–48"],
                ["11", "LTE FDD/TDD", "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp/Rsrq + A4 Hyst/TTT", "MLB/FreqPri A4; TTT 5120 ms disables FreqPri", "[CM] Tables 11-5, 4-9"],
                ["12", "LTE FDD/TDD", "EUTRANINTERNFREQ", "IfMlbThdRsrpOffset / FreqPriHoA4ThldRsrpOffset", "Per-frequency A4 offset", "[MLB] p.137; [CM] Table 11-5"],
                ["13", "LTE FDD/TDD", "INTERFREQHOGROUP", "A5 Thd1/Thd2 + Mlb A5 Thd1", "Dual-condition HO", "[CM] Tables 5-18/5-19; [MLB] Table 6-3"],
                ["14", "LTE FDD/TDD", "EUTRANINTERNFREQ", "MlbInterFreqHoEventType", "A4 or A5 for FDD MLB", "[MLB] §6.1.1.5.2"],
                ["15", "LTE FDD/TDD", "CELLALGOSWITCH", "CovBasedInterFreqHoMode + CovBasedIfHoWaitingTimer", "Coverage execution timing", "[CM] §5.3.1"],
                ["16", "LTE FDD/TDD", "INTRARATHOCOMM", "FreqPriIFHoWaitingTimer / FreqPriInHoProtectionTimer", "Wait / incoming protect", "[CM] pp.300, 308–309"],
                ["17", "LTE FDD/TDD", "FreqPri switches", "MlbBasedFreqPriHoSwitch / LoadTriggerFreqPriHoSwitch / ReduceInvalidFreqPriHoSwitch", "MLB owns heavy load", "[CM] Table 11-7"],
                ["18", "LTE FDD/TDD", "CELLOPHOCFG", "HighMobiUeHoForbidSw / FREQ_PRI_HO_FORBID_SW", "No MLB/FreqPri meas above ~30 km/h", "[CM] pp.304–305; [MLB] pp.136–141"],
                ["19", "LTE FDD/TDD", "HOMEASCOMM", "Res/Opt/NonRes HO fail punish timers and counts", "Retry vs penalty", "[CM] Tables 4-16/4-17"],
                ["20", "LTE FDD/TDD", "EUTRANINTERFREQNCELL", "BlindHoPriority / InterFreqMlbBlindHo", "Blind path", "[CM] Table 5-22 p.149; [MLB] p.136"],
                ["21", "LTE FDD/TDD", "RATFREQPRIORITYGROUP", "CovIFHo RSRP/RSRQ Hyst/TTT", "Per-freq/QCI coverage values", "[CM] Table 4-9"],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 34,
            "rows": [
                ["0", "LTE FDD/TDD", "—",
                 "LST INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST EUTRANINTERFREQNCELL: LocalCellId=<x>; LST CELLHOPARACFG: LocalCellId=<x>; LST HOMEASCOMM:;",
                 "Read-only", "Dump first. Do not paste example −85/−87/−103 dBm as live design values. [CM] §11.4.1.2"],
                ["1", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>; (FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for required HO targets);",
                 "NRT exists", "Audit flags/NRT before any threshold change."],
                ["2", "LTE FDD/TDD", "CELLUEMEASCONTROLCFG",
                 "MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<n>, MaxEutranFddMeasFreqNum=<n>;",
                 "n ≥ number of inter-frequency objects the cell must measure", "Equal-priority objects may otherwise be randomly dropped. [CM] p.125"],
                ["3", "LTE FDD/TDD", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>; set coverage A1/A2/A5 from MR, not from MML examples;",
                 "Correct A2 family for the intended HO type", "Coverage rescue configuration."],
                ["4", "LTE FDD/TDD", "INTERFREQHOGROUP",
                 "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<calib>, InterFreqHoA4Hyst=<hyst>, InterFreqHoA4TimeToTrig=<ttt>;",
                 "A4 better than coverage A2; TTT ≠ 5120 ms if FreqPri/MLB A4 required", "[CM] Table 4-9 p.48; Table 5-22"],
                ["5", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbInterFreqHoEventType=A4;",
                 "A5 only with non-cosited MLB license", "[MLB] Table 6-3 pp.142–143"],
                ["6", "LTE FDD/TDD", "FreqPri / ENODEBALGOSWITCH related MO",
                 "Enable MlbBasedFreqPriHoSwitch and LoadTriggerFreqPriHoSwitch when MLB is used; keep FreqPriInHoProtectionTimer non-zero;",
                 "No reverse MLB target on a FreqPri pair", "Confirm exact MO name in MAE. [CM] p.303; Table 11-7"],
                ["7", "LTE FDD/TDD", "Verify",
                 "L.HHO.InterFreq.Coverage.*; L.HHO.InterFreq.FreqPri.*; L.RRC.ReEst.ReconfFail.Att. PrepSucc=ExecAtt/PrepAtt; ExecSucc=ExecSucc/ExecAtt.",
                 "Busy-hour pair-level", "Counters [CM] Table 5-24 p.153; Tables 11-9/11-10 pp.318–319."],
            ],
        },
    ]


# ---------------------------------------------------------------------------
# Feature 3 MLB
# ---------------------------------------------------------------------------
def mlb_sections():
    return [
        {
            "name": "Section 1: Feature Introduction  (SN-1 / SN-2)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Current Value / Status", "Notes / Rationale"],
            "h": 38,
            "rows": [
                ["SN-1 Working Principal", "LTE FDD/TDD", "Intra-RAT MLB",
                 "Coordinate load among overlapping inter-frequency LTE cells; transfer by connected HO or idle dedicated-priority release",
                 "Equalisation uses peer load; offload can run without full target load",
                 "CHART (doc): Fig 3-1 p.15; Fig 4-1 equalisation vs offload p.16. Load (user-number) = N/C ; normalised difference (Load_s−Load_t)/Load_s. [MLB] §3–§4; §5.1.1 pp.24–27."],
                ["Load indicators", "LTE FDD/TDD", "PRB vs UE number vs HW/transport",
                 "PRB/service satisfaction ~ GBR; UE count ~ short non-GBR; HW/transport = Low/Medium/High/OverLoad",
                 "ActiveUeBasedLoadEvalSw uses UEs with buffered DL data; SpectralEffBasedLoadEvalSw includes PRB, scale, GBR, measured SE",
                 "Huawei recommends ActiveUe when MLB frequencies have different bandwidths, and SE eval when SE differs significantly (e.g. >30%). [MLB] Fig 4-2/4-3; Table 5-5 pp.47–48."],
                ["Modes", "LTE FDD/TDD", "CELLMLB.MlbTriggerMode",
                 "UE_NUMBER_ONLY / PRB_ONLY / PRB_OR_UE_NUMBER",
                 "Idle uses dedicated priority + T320; connected uses A4/A5 HO (FDD)",
                 "PRB MLB does not transfer CA UEs and does not guarantee experience fairness; gain falls when CA penetration ≳60%. [MLB] §6.5 pp.207–215."],
                ["SN-2 Major highlighted Point", "LTE FDD/TDD", "Document cautions",
                 "Raw UE-count on unequal BW can reduce DL throughput; ONLY_STRONGEST_CELL recommended; 5 s eval + MlbMaxUeNum≥40 over-transfers",
                 "Smart neighbor thds: collect 7 days, refresh every 7 days",
                 "First-week aggressive seeds: extra CPU/HO and up to 5% TP fluctuation. [MLB] §6.1.2.2 p.141; Table 6-5 p.159; p.136; pp.29–31, 162, 224–225."],
            ],
        },
        {
            "name": "Section 2: Triggering Conditions  (SN-3 / SN-4)",
            "headers": ["Feature Part", "RAT", "MO Name / Check Item", "When it starts", "Parameter Detail", "User Experience Consideration"],
            "h": 32,
            "rows": [
                ["Connected UE-number", "LTE FDD/TDD", "InterFreqMlbUeNumThd + MlbUeNumOffset + MlbTrigJudgePeriod",
                 "N ≥ Thd+Offset throughout the judge period; stop when N < Thd",
                 "Requires InterFreqMlbSwitch, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE",
                 "[MLB] Table 6-2 p.127; §6.1.1.1 p.128."],
                ["Idle UE-number", "LTE FDD/TDD", "InterFreqIdleMlbUeNumThd + IdleUE",
                 "Analogous idle-user condition",
                 "RRC release with IdleModeMobilityControlInfo",
                 "[MLB] Table 5-2; §5.1.1.5."],
                ["PRB usage", "LTE FDD/TDD", "InterFreqMlbThd / UlThd + LoadOffset + min UE",
                 "PRB ≥ thd + offset and minimum UE condition",
                 "PrbLoadCalcMethod=PRB_USAGE; InterFreqUeTrsfType=PrbMlbSynchronizedUE",
                 "Bursty, no CA UE transfer. [MLB] Table 6-18; §6.5.1."],
                ["Target admit", "LTE FDD/TDD", "NoHoFlag, blacklist, PCI, HO SR, meas flags, OverlapInd, MlbTargetInd, HW/transport",
                 "Before a cell is used as equalisation target",
                 "Low/Med HW+transport; not in penalty; HO success ≥ NCellHoSuccRateThld",
                 "Target reject 'no radio resource' ⇒ punish CellPunishPrdNum × InterFreqLoadEvalPrd. [MLB] pp.27–29, 129."],
                ["UE selection", "LTE FDD/TDD", "CELLMLBUESEL + protect/punish timers",
                 "After source trigger and target admit",
                 "UL-sync, not emergency; ARP/PRB/MCS/QCI/SNR options; high-mobility forbid optional",
                 "ONLY_STRONGEST_CELL avoids immediate coverage HO bounce. [MLB] pp.130–141; Table 6-5 p.159."],
                ["Transfer volume / freq pick", "LTE FDD/TDD", "MlbMaxUeNum / FreqSelectStrategy",
                 "At execution",
                 "FAIRSTRATEGY / PRIORITYBASED / LOADPRIORITY; LoadTransferEnhSw changes multi-target math",
                 "InterFreqLoadEvalPrd=5 s AND MlbMaxUeNum≥40 can transfer excessive UEs. [MLB] pp.134–137."],
            ],
        },
        {
            "name": "Section 3: eNodeB Actions  (SN-5)",
            "headers": ["Action Area", "RAT", "Feature Part", "eNodeB Action", "Parameter / Condition", "Operational Meaning"],
            "h": 30,
            "rows": [
                ["Master switch", "LTE FDD/TDD", "CELLALGOSWITCH.MlbAlgoSwitch",
                 "InterFreqMlbSwitch-1; InterFreqIdleMlbSwitch-1 if idle transfer required; Blind-0 unless designed",
                 "License + NRT + MlbTargetInd",
                 "Table 6-2 p.127."],
                ["Trigger mode", "LTE FDD/TDD", "CELLMLB",
                 "MlbTriggerMode=UE_NUMBER_ONLY (typical document recommendation for user-experience oriented equalisation)",
                 "InterFreqUeTrsfType must match (SynchronizedUE / IdleUE / PrbMlbSynchronizedUE)",
                 "PRB_ONLY is supplementary. [MLB] Table 5-5; §6.5."],
                ["Load model", "LTE FDD/TDD", "ActiveUeBasedLoadEvalSw / SpectralEffBasedLoadEvalSw / LoadTransferEnhSw",
                 "Turn on as recommended for unequal bandwidth / SE / multi-target",
                 "SE refresh every minute when ≥10 UL-sync UEs",
                 "[MLB] §5.1.1 pp.25–27; Table 5-5."],
                ["CA transfer", "LTE FDD/TDD", "CaUserLoadTransferSw",
                 "Required if CA UEs must be moved in user-number MLB",
                 "One path requires target CA air-interface capability ≥ serving",
                 "If OFF, CA UEs treating the cell as PCell/SCell are filtered (2CC / FDD+TDD notes). [MLB] pp.129–136, 157."],
                ["Cell / freq pick", "LTE FDD/TDD", "CELLMLB",
                 "MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL; FreqSelectStrategy as designed",
                 "LOADPRIORITY uses above-average load difference",
                 "[MLB] Table 6-5 p.159; pp.137, 213."],
                ["Event / target", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MlbInterFreqHoEventType=A4 (co-sited) or A5 (licensed non-cosited); set MlbTargetInd per frequency",
                 "ALLOWED_WITHOUT_CONNECT_MLB / WITHOUT_IDLE_MLB block one mode",
                 "Ordinary coverage HO is separate from MLB targeting. [MLB] pp.28, 129; Table 6-3."],
            ],
        },
        {
            "name": "Section 4: Prerequisite Functions  (SN-6)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 24,
            "rows": [
                ["License", "LTE FDD/TDD", "Intra-RAT MLB (+ A5 license if used)", "Present on the eNodeB", "Required", "A5: Intra-LTE Load Balancing for Non-cosited Cells [MLB] pp.142–143"],
                ["Load exchange", "LTE FDD/TDD", "X2 / intra-eNB", "Needed for equalisation", "Without it only offload/blind remains", "[MLB] §4.1"],
                ["NRT / overlap / PCI", "LTE FDD/TDD", "OverlapInd, NoHoFlag, PCI uniqueness", "PERMIT_HO; valid overlap", "Required", "[MLB] pp.27–28"],
                ["Meas delivery", "LTE FDD/TDD", "Feature 2 flags / object cap / SMeasure / gap", "A4/A5 must actually be delivered", "Else trigger with 0 meas success", "[CM] + [MLB] p.139"],
                ["HO success floor", "LTE FDD/TDD", "NCellHoSuccRateThld", "Pair must pass", "Failed pair is not a target", "[MLB] p.27"],
                ["HW / transport", "LTE FDD/TDD", "LowLoad / MediumLoad", "High/OverLoad not a normal equalisation target", "Fig 4-3", "[MLB] pp.17–18"],
                ["Smart thd counters", "LTE FDD/TDD", "MAE 15-min counter subscription", "If NCellTrigThldSmartOptAlgoSw ON", "At least one 15-min period", "[MLB] §§5.1.3.4, 6.5.3.4 pp.46, 223"],
                ["Idle SIB", "LTE FDD/TDD", "SIB5 NORMAL", "If idle MLB ON", "UNDELIVER freqs cannot be idle targets", "[IM] §5.3.2.3"],
            ],
        },
        {
            "name": "Section 5: Mutually Impacted and Related Features  (SN-7 / SN-8)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 26,
            "rows": [
                ["User-number connected + fixed-proportion idle", "LTE FDD/TDD", "Idle proportion feature", "Do not combine", "Ping-pong", "[MLB] §5.4.2.2 pp.72, 74–75"],
                ["PRB_USAGE vs PRB_VALUATION", "LTE FDD/TDD", "Same mode", "Mutually exclusive", "Pick one if PRB used", "[MLB] pp.216, 245"],
                ["MLB vs FreqPri", "LTE FDD/TDD", "Connected Mode Ch.11", "MlbBasedFreqPriHoSwitch; no reverse target", "Ping-pong if both fight", "[CM] p.303; Table 11-7"],
                ["MLB vs energy saving", "LTE FDD/TDD", "Carrier shutdown / deep dormancy", "Changes target selection and gain", "Coordinate", "[MLB] pp.153–156, 219–222"],
                ["Flexible CA", "LTE FDD/TDD", "CA serving-cell combination", "Avoids a cell where connected offload is active", "Expected interaction", "[MLB] pp.168, 233"],
                ["Smart learned thresholds", "LTE FDD/TDD", "NCellTrigThldSmartOptAlgoSw", "Relearn after BW/CA-eval/upgrade/board/cell-deactivation change", "Can go stale", "[MLB] pp.29–31, 89–91"],
                ["Related: Idle / Connected / CA / IRAT MLB", "LTE FDD/TDD", "F1 idle method; F2 HO engine; CA PCC; inter-RAT is a separate book", "A4/A5 + ONLY_STRONGEST_CELL", "Inter-RAT MLB is out of this document set", "[MLB]/[CM]/[IM]"],
            ],
        },
        {
            "name": "Section 6: License Requirements  (SN-9)",
            "headers": ["Topic", "RAT", "Feature Part", "Parameter ID / Check Item", "Document value / status", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["Connected/idle equalisation", "LTE FDD/TDD", "InterFreqMlbSwitch / IdleMlbSwitch", "Intra-RAT Mobility Load Balancing", "Core license — confirm name on NE", "No Load HO counters if missing"],
                ["MLB event A5", "LTE FDD/TDD", "MlbInterFreqHoEventType=A5", "Intra-LTE Load Balancing for Non-cosited Cells", "pp.142–143", "Stay on A4 if license absent"],
                ["Blind MLB", "LTE FDD/TDD", "InterFreqBlindMlbSwitch", "Blind MLB option if sold separately", "Higher risk without load exchange", "[MLB] pp.167, 231"],
                ["CA user transfer", "LTE FDD/TDD", "CaUserLoadTransferSw", "CA license + MLB CA-transfer option — verify", "Else CA UEs filtered", "p.157"],
                ["Smart n-cell threshold", "LTE FDD/TDD", "NCellTrigThldSmartOptAlgoSw", "SON/intelligent MLB option + MAE counters", "7-day learn / 7-day refresh", "pp.29–31"],
            ],
        },
        {
            "name": "Section 7: All Parameter List  (SN-10, sequence)",
            "headers": ["Topic", "RAT", "MO Name / Check Item", "Parameter ID / Check Item", "Document / Huawei note", "Notes / Rationale"],
            "h": 22,
            "rows": [
                ["1", "LTE FDD/TDD", "CELLALGOSWITCH", "InterFreqMlbSwitch / InterFreqIdleMlbSwitch / InterFreqBlindMlbSwitch", "Master bits", "Table 6-2"],
                ["2", "LTE FDD/TDD", "CELLMLB", "MlbTriggerMode / PrbLoadCalcMethod / InterFreqUeTrsfType / InterFreqMLBRanShareMode", "Mode", "Table 6-2 / 6-18"],
                ["3", "LTE FDD/TDD", "eval switches", "ActiveUeBasedLoadEvalSw / SpectralEffBasedLoadEvalSw / LoadTransferEnhSw / CaUserLoadTransferSw", "Load model / CA", "Table 5-5; pp.129–136"],
                ["4", "LTE FDD/TDD", "CELL / CELLMLB", "CellCapacityScaleFactor / MuMimoPrbStatOptSwitch / MultiRruMode", "Capability scale", "§5.1.1"],
                ["5", "LTE FDD/TDD", "CELLMLB", "MlbTrigJudgePeriod / InterFreqLoadEvalPrd", "Timing; 5 s + large MaxUeNum warning", "p.128, p.136"],
                ["6", "LTE FDD/TDD", "CELLMLB", "InterFreqMlbUeNumThd / MlbUeNumOffset / InterFreqIdleMlbUeNumThd", "UE-number trigger", "p.128; Table 5-2"],
                ["7", "LTE FDD/TDD", "CELLMLB", "InterFreqMlbThd / InterFreqMlbUlThd / LoadOffset / MlbMinUeNumThd / Offset", "PRB trigger family", "§6.5.1"],
                ["8", "LTE FDD/TDD", "CELLMLB", "LoadDiffThd / InterFreqOffloadOffset / InterFrqUeNumOffloadOffset / InterFIdleUeNumOffloadOfs", "Equalise vs offload slack", "§4.1"],
                ["9", "LTE FDD/TDD", "CELLMLB", "MlbMaxUeNum / MlbIdleUeNumAdjFactor", "Transfer volume", "p.136"],
                ["10", "LTE FDD/TDD", "CELLMLB", "FreqSelectStrategy / MlbFreqPriority / MlbFreqUlPriority", "FAIR / PRIORITY / LOADPRIORITY", "pp.137, 213"],
                ["11", "LTE FDD/TDD", "CELLMLB", "MlbHoCellSelectStrategy", "ONLY_STRONGEST_CELL recommended", "Table 6-5 p.159"],
                ["12", "LTE FDD/TDD", "EUTRANINTERNFREQ", "MlbTargetInd / MlbInterFreqHoEventType / IfMlbThdRsrpOffset", "Target + A4/A5", "pp.28, 129, 137"],
                ["13", "LTE FDD/TDD", "EUTRANINTERNFREQ / NCELL", "OverlapInd / NoHoFlag / AggregationAttribute / LoadBalanceNCellScope", "Relation eligibility", "pp.27–28"],
                ["14", "LTE FDD/TDD", "CELLMLB", "NCellHoSuccRateThld / CellPunishPrdNum / FreqPunishPrdNum / PunishJudgePrdNum", "Admit / penalty", "p.29"],
                ["15", "LTE FDD/TDD", "CELLMLBUESEL", "UeSelectArp/Prb/DlMcs/QciPrio + thds / SnrBasedUeSelectionMode", "UE pick", "pp.130–135"],
                ["16", "LTE FDD/TDD", "CELLMLB", "MlbHoInProtectTimer / MlbUeSelectPunishTimer / MlbHoInProtectMode", "Re-MLB protect", "pp.130–135"],
                ["17", "LTE FDD/TDD", "CELLOPHOCFG", "MLB_HO_FORBID_SW", "No MLB meas if UE > ~30 km/h", "pp.136–141"],
                ["18", "LTE FDD/TDD", "EUTRANINTERNFREQ", "InterFreqMlbBlindHo / BlindHoPriority", "Blind path", "p.136"],
                ["19", "LTE FDD/TDD", "EUTRANINTERNFREQ", "InterFreqMlbDlPrbOffset / UlPrbOffset", "Per-freq PRB bias", "p.210"],
                ["20", "LTE FDD/TDD", "CELLPRBVALMLB / MlbQciGroup", "PrbValMlbTrigThd / AdmitThd / FilterFactor / min QoS bit rates", "PRB-evaluation MLB", "Table 6-25; §5.8.1"],
                ["21", "LTE FDD/TDD", "smart", "NCellTrigThldSmartOptAlgoSw + LocalToNCell / NToLocal UE/PRB thds", "7-day learn / 7-day refresh", "pp.29–31, 89–91"],
                ["22", "LTE FDD/TDD", "RRCCONNSTATETIMER / EnhancedMlb", "T320ForLoadBalance / DediPrioManageOnLowLoadSw", "Idle dedicated prio", "§5.1.1.5; Table 5-10"],
            ],
        },
        {
            "name": "Section 8: Final MML Command for activations  (SN-11, maintain sequence)",
            "headers": ["Parameter Sequence", "RAT", "MO", "Activation Value", "Conditional Parameter", "Remarks / Parameter Description / More Notes"],
            "h": 36,
            "rows": [
                ["0", "LTE FDD/TDD", "—",
                 "LST CELLALGOSWITCH: LocalCellId=<x>; LST CELLMLB: LocalCellId=<x>; LST EUTRANINTERNFREQ: LocalCellId=<x>; LST CELLMLBUESEL: LocalCellId=<x>; check Intra-RAT MLB license;",
                 "MAE license + version-matched parameter reference", "Feature-book Ch.8 does not list full defaults/ranges (p.299)."],
                ["1", "LTE FDD/TDD", "EUTRANINTERNFREQ",
                 "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;",
                 "OverlapInd valid; NoHoFlag=PERMIT_HO_ENUM; FREQ_MEAS_FLAG; not forbid-meas", "Use A5 only with the non-cosited MLB license. Block a frequency from MLB with WITHOUT_CONNECT_MLB / WITHOUT_IDLE_MLB."],
                ["2", "LTE FDD/TDD", "eval MO / CELLALGOSWITCH related",
                 "Enable ActiveUeBasedLoadEvalSw, SpectralEffBasedLoadEvalSw, LoadTransferEnhSw as required by bandwidth/SE/multi-target;",
                 "Different BW and/or SE difference e.g. >30%", "Confirm exact bit/MO names in MAE. [MLB] Table 5-5"],
                ["3", "LTE FDD/TDD", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL;",
                 "InterFreqMlbSwitch will be ON", "ONLY_STRONGEST_CELL is Huawei-recommended. [MLB] Table 6-5 p.159"],
                ["4", "LTE FDD/TDD", "CELLMLB",
                 "MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<thd>, MlbUeNumOffset=<ofs>, MlbMaxUeNum=<n>, MlbTrigJudgePeriod=<p>, InterFreqLoadEvalPrd=<prd>, FreqSelectStrategy=<FAIRSTRATEGY|PRIORITYBASED|LOADPRIORITY>;",
                 "Do not use InterFreqLoadEvalPrd=5s with MlbMaxUeNum≥40", "Calibrate thd after ActiveUe/SE switches. [MLB] p.136"],
                ["5", "LTE FDD/TDD", "CA related",
                 "Enable CaUserLoadTransferSw only when CA UEs must be transferable and target CA capability conditions are met;",
                 "CA license + combination plan", "If OFF, CA UEs are filtered. [MLB] pp.129–136, 157"],
                ["6", "LTE FDD/TDD", "CELLALGOSWITCH",
                 "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;",
                 "Steps 1–4 done; Blind bit remains 0 unless designed", "Idle path also needs T320 and SIB5 NORMAL."],
                ["7", "LTE FDD/TDD", "RRCCONNSTATETIMER",
                 "MOD RRCCONNSTATETIMER: T320ForLoadBalance=<T320>;",
                 "Idle MLB ON", "[MLB] §5.1.1.5"],
                ["8", "LTE FDD/TDD", "Verify",
                 "L.HHO.InterFreq.Load.* and UeNumLoad.*; L.InterFreq.HighLoad.Dur/Num; L.InterFreq.Load.Meas/MeasSucc; L.Traffic.ActiveUser.DL.Avg; PCell/SCell users; PRB; Thrp.bits/Time.DL; idle DedicatedPri counters. SON logs: Inter-Frequency Handover Statistics; Inter-Frequency Idle Mode Release Statistics.",
                 "Subscribe 15-min counters if smart thd is used", "Tables 6-6, 6-21, 6-28 pp.164, 227–228, 248–249; log contents §§6.1.4.2, 6.5.4.2."],
            ],
        },
    ]


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    # Cover uses first sheet
    ws = wb.active
    ws.title = "Mobility Management"
    cols = 6
    widths(ws, [24, 12, 30, 38, 36, 56])
    setup(ws, "Mobility Management")
    r = 1
    r = title(ws, r, cols, "Mobility Management - Detailed Notes")
    for sec in cover_sections():
        r = spacer(ws, r, cols)
        r = section(ws, r, cols, sec["name"])
        r = headers(ws, r, sec["headers"])
        for rec in sec["rows"]:
            r = row(ws, r, rec, h=sec.get("h", 30))
        r = spacer(ws, r, cols)

    sheet(wb, "Idle Mode Management", "Idle Mode Management - Detailed Notes", idle_sections())
    sheet(wb, "Connected Mode", "Mobility Management in Connected Mode - Detailed Notes", connected_sections())
    sheet(wb, "Intra-RAT MLB", "Intra-RAT Mobility Load Balancing - Detailed Notes", mlb_sections())

    wb.properties.title = "eRAN21.1 Mobility Management PDF summary"
    wb.properties.creator = "Feature-parameter summary"
    wb.properties.subject = "Idle Mode / Connected Mode / Intra-RAT MLB — exclusive PDF summary"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
