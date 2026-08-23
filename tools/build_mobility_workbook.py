#!/usr/bin/env python3
"""Build Huawei eRAN21.1 Mobility Management Excel workbook.

Filled from:
  - Idle Mode Management, eRAN21.1, Issue 04
  - Mobility Management in Connected Mode, eRAN21.1, Issue 08
  - Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10

Template structure requested by user (11 items per feature).
Proposed values are Robi capacity-layer engineering recommendations,
not official Huawei default values. Validate in MAE before live change.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.page import PageMargins
from openpyxl.workbook.child import INVALID_TITLE_REGEX
from copy import copy

OUT = "/workspace/docs/4G_LTE_Mobility_Management/Mobility_Management_eRAN21.1_Workbook.xlsx"

# Colors
NAVY = "0D2B4A"
HUAWEI_RED = "C7000B"
SECTION = "1F4E79"
THEAD = "2E75B6"
SUBHEAD = "5B9BD5"
LIGHT = "D6EAF8"
WHITE = "FFFFFF"
IVORY = "FFF8E7"
GREEN = "C6EFCE"
GREEN_DK = "006100"
AMBER = "FCE4D6"
AMBER_DK = "C65911"
YELLOW = "FFF2CC"
YELLOW_DK = "806000"
GREY = "F2F2F2"
GREY_DK = "7F7F7F"
LILAC = "E2D5F1"
PINK = "F8CBAD"
TEAL = "D5F5E3"
REF = "D6EAF8"

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
thick_bottom = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="medium", color="0D2B4A"),
)

font_white_b = Font(name="Calibri", size=16, bold=True, color=WHITE)
font_white = Font(name="Calibri", size=11, bold=True, color=WHITE)
font_white_s = Font(name="Calibri", size=10, bold=True, color=WHITE)
font_navy_b = Font(name="Calibri", size=13, bold=True, color=NAVY)
font_navy = Font(name="Calibri", size=11, bold=True, color=NAVY)
font_body = Font(name="Calibri", size=10, color="1A1A1A")
font_small = Font(name="Calibri", size=9, italic=True, color="4A4A4A")
font_hdr = Font(name="Calibri", size=9, bold=True, color=WHITE)
font_green = Font(name="Calibri", size=10, bold=True, color=GREEN_DK)
font_warn = Font(name="Calibri", size=10, bold=True, color=AMBER_DK)
font_title = Font(name="Calibri", size=22, bold=True, color=WHITE)
font_sub = Font(name="Calibri", size=12, bold=True, color=WHITE)

wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_c = Alignment(wrap_text=True, vertical="center", horizontal="center")
top_left = Alignment(wrap_text=True, vertical="top", horizontal="left")


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def cell(ws, r, c, value, font=None, fill_hex=None, align=None, border=True, row_h=None):
    x = ws.cell(r, c, value)
    x.font = font or font_body
    x.alignment = align or wrap
    if fill_hex:
        x.fill = fill(fill_hex)
    if border:
        x.border = thin
    if row_h:
        ws.row_dimensions[r].height = row_h
    return x


def banner(ws, row, cols, text, color=NAVY, font=None, height=28):
    merge(ws, row, 1, row, cols)
    cell(ws, row, 1, text, font=font or font_white_b, fill_hex=color, align=Alignment(vertical="center", horizontal="left", indent=1), row_h=height)
    for c in range(2, cols + 1):
        ws.cell(row, c).fill = fill(color)
        ws.cell(row, c).border = thin
    return row + 1


def section(ws, row, cols, sn, title, color=SECTION):
    merge(ws, row, 1, row, cols)
    cell(ws, row, 1, f"  SN-{sn}   {title}", font=font_white, fill_hex=color, align=Alignment(vertical="center", horizontal="left"), row_h=22)
    for c in range(2, cols + 1):
        ws.cell(row, c).fill = fill(color)
        ws.cell(row, c).border = thin
    return row + 1


def note(ws, row, cols, text, color=IVORY, height=36):
    merge(ws, row, 1, row, cols)
    cell(ws, row, 1, text, font=font_small, fill_hex=color, align=top_left, row_h=height)
    for c in range(2, cols + 1):
        ws.cell(row, c).fill = fill(color)
        ws.cell(row, c).border = thin
    return row + 1


def para(ws, row, cols, text, color=WHITE, height=48):
    merge(ws, row, 1, row, cols)
    cell(ws, row, 1, text, font=font_body, fill_hex=color, align=top_left, row_h=height)
    for c in range(2, cols + 1):
        ws.cell(row, c).fill = fill(color)
        ws.cell(row, c).border = thin
    return row + 1


def bullets(ws, row, cols, items, color=WHITE, row_h=18):
    for item in items:
        merge(ws, row, 1, row, cols)
        cell(ws, row, 1, f"  •  {item}", font=font_body, fill_hex=color, align=wrap, row_h=row_h)
        for c in range(2, cols + 1):
            ws.cell(row, c).fill = fill(color)
            ws.cell(row, c).border = thin
        row += 1
    return row


def table(ws, row, headers, rows, header_color=THEAD, col_fills=None, min_h=28):
    ncols = len(headers)
    for c, h in enumerate(headers, 1):
        cell(ws, row, c, h, font=font_hdr, fill_hex=header_color, align=wrap_c, row_h=22)
    row += 1
    for i, rec in enumerate(rows):
        bg = GREY if i % 2 else WHITE
        h = max(min_h, 16 + 8 * max((str(v).count("\n") + 1) for v in rec))
        for c, v in enumerate(rec, 1):
            use = bg
            fnt = font_body
            if col_fills and c in col_fills:
                use = col_fills[c]
                if col_fills[c] == GREEN:
                    fnt = font_green
                elif col_fills[c] == AMBER:
                    fnt = font_warn
                elif col_fills[c] == YELLOW:
                    fnt = Font(name="Calibri", size=10, bold=True, color=YELLOW_DK)
            cell(ws, row, c, v, font=fnt, fill_hex=use, align=top_left, row_h=h)
        row += 1
    return row


def flow_row(ws, row, boxes, colors, cols=10):
    """Simple process-flow using merged colored boxes."""
    n = len(boxes)
    # 10 columns: each box uses 2 cols, arrows in remaining conceptually inside text
    span = max(1, cols // n)
    c = 1
    for i, (box, col) in enumerate(zip(boxes, colors)):
        end = min(cols, c + span - 1)
        if i < n - 1:
            # leave last of span for arrow look by putting arrow in text
            text = f"{box}   →" if end < cols else box
        else:
            text = box
        if c <= end:
            merge(ws, row, c, row, end)
            cell(ws, row, c, text, font=font_white_s, fill_hex=col, align=wrap_c, row_h=28)
            for k in range(c + 1, end + 1):
                ws.cell(row, k).fill = fill(col)
                ws.cell(row, k).border = thin
        c = end + 1
    ws.row_dimensions[row].height = 30
    return row + 1


def apply_sheet_setup(ws, title, landscape=True):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.oddHeader.left.text = "Robi Axiata PLC  |  4G RNO  |  Mobility Management"
    ws.oddFooter.left.text = title
    ws.oddFooter.right.text = "Huawei eRAN21.1  |  Page &P of &N"
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = None
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def add_legend_row(ws, row, cols):
    merge(ws, row, 1, row, cols)
    cell(
        ws,
        row,
        1,
        "  COLOR LEGEND:  Green = Robi proposed / recommended value   |   Amber = risk / do-not-change casually   |   Yellow = L900 coverage-protection   |   Blue = document reference   |   Proposed values are engineering recommendations and must be calibrated on live KPI before CR execution.",
        font=font_small,
        fill_hex=GREY,
        row_h=28,
    )
    for c in range(2, cols + 1):
        ws.cell(row, c).fill = fill(GREY)
        ws.cell(row, c).border = thin
    return row + 1


# ---------------------------------------------------------------------------
# COVER
# ---------------------------------------------------------------------------
def build_cover(wb):
    ws = wb.active
    ws.title = "00_Cover"
    cols = 10
    set_widths(ws, [18, 18, 18, 18, 16, 16, 16, 16, 16, 22])
    apply_sheet_setup(ws, "Cover")

    r = 1
    r = banner(ws, r, cols, "  ROBI AXIATA PLC   ·   RADIO NETWORK OPTIMIZATION", HUAWEI_RED, font_title, 36)
    r = banner(ws, r, cols, "  Mobility Management  |  Huawei LTE eRAN21.1  |  Feature–Parameter Workbook", NAVY, font_sub, 24)
    r = note(
        ws,
        r,
        cols,
        "  Prepared for: Mohammad Selim, Senior RNO Engineer (2G/3G/4G/5G)  ·  Vendor: Huawei  ·  Scope: 4G intra-LTE mobility and MLB  ·  Version: 1.0  ·  Date: 23 Aug 2026",
        LIGHT,
        22,
    )
    r += 1

    r = section(ws, r, cols, "0", "Purpose of this workbook", HUAWEI_RED)
    r = para(
        ws,
        r,
        cols,
        "This workbook converts the three Huawei eRAN21.1 Mobility Management feature documents into an RNO-usable Excel format. "
        "It follows the requested 11-item sequence for each feature: Working Principle → Highlighted Points → Benefit/Limitation → "
        "Selection/Trigger → Activation Switch → Prerequisite → Mutual Impact → Relation with Other Features → License → "
        "Full Parameter List → Final MML. It is the knowledge base for a daily AI agent that flags capacity-layer DL user-throughput "
        "gaps greater than 2 Mbps and prepares a guarded change request.",
        WHITE,
        56,
    )
    r += 1

    r = section(ws, r, cols, "0", "Source documents (must stay attached as reference)", SECTION)
    r = table(
        ws,
        r,
        ["SN", "Document", "Issue", "Date", "Role in this workbook", "PDF short name"],
        [
            ["1", "eRAN Idle Mode Management — Feature Parameter Description", "Issue 04", "2026-05-30", "Camping, reselection, dedicated priority, idle MLB interface", "Idle_Mode_Management_eRAN21.1"],
            ["2", "eRAN Mobility Management in Connected Mode — Feature Parameter Description", "Issue 08", "2026-06-30", "A1/A2/A3/A4/A5, coverage HO, frequency-priority HO, measurement", "Mobility_Management_in_Connected_Mode_eRAN21.1"],
            ["3", "eRAN Intra-RAT Mobility Load Balancing — Feature Parameter Description", "Issue 10", "2026-06-30", "Connected/idle MLB algorithms, CA transfer, counters, intelligent threshold", "Intra-RAT_Mobility_Load_Balancing_eRAN21.1"],
        ],
        col_fills={5: REF},
    )
    r += 1

    r = section(ws, r, cols, "0", "Robi 4G layer architecture used throughout this book", SECTION)
    r = table(
        ws,
        r,
        ["Layer", "Bandwidth", "PRB", "Nominal share of capacity pool", "Intended role", "MLB role", "Idle priority intent", "Connected mobility intent"],
        [
            ["L2600 C1", "20 MHz", "100", "17.4%", "Primary capacity", "Equalization source/target", "Highest, equal among C1–C4", "A4 MLB / A3 if relative quality needed"],
            ["L2600 C2", "20 MHz", "100", "17.4%", "Primary capacity", "Equalization source/target", "Highest, equal among C1–C4", "A4 MLB / A3 if relative quality needed"],
            ["L2600 C3", "20 MHz", "100", "17.4%", "Primary capacity", "Equalization source/target", "Highest, equal among C1–C4", "A4 MLB / A3 if relative quality needed"],
            ["L2600 C4", "20 MHz", "100", "17.4%", "Primary capacity", "Equalization source/target", "Highest, equal among C1–C4", "A4 MLB / A3 if relative quality needed"],
            ["L1800", "20 MHz", "100", "17.4%", "Broad capacity + mobility anchor", "Equalization source/target", "High, next to L2600", "MLB A4 + coverage rescue"],
            ["L2100", "15 MHz", "75", "13.0%", "Capacity (smaller BW)", "Equalization source/target, SE-normalized", "High or one step below L1800", "MLB A4; do not treat as 20 MHz"],
            ["L900", "5 MHz", "25", "NOT in capacity pool", "Indoor / deep coverage only", "NOT a routine MLB target", "Lowest LTE priority", "A5/coverage in; A1/FreqPri out"],
        ],
        col_fills={2: GREEN, 6: YELLOW},
    )
    r = note(
        ws,
        r,
        cols,
        "  Nominal share = PRB / 575. Do not use this as a live traffic target. Correct by spectral efficiency, MIMO, CA PCC/SCC role, UE band support, GBR, interference, transport and availability. L900 25 PRB is excluded from the capacity-balancing pool.",
        YELLOW,
        32,
    )
    r += 1

    r = section(ws, r, cols, "0", "Daily optimization objective", SECTION)
    r = para(
        ws,
        r,
        cols,
        "Reduce DL user-throughput gap among capacity layers (L1800 / L2100 / L2600 C1–C4) from >10 Mbps toward 1–2 Mbps. "
        "A gap > 2 Mbps is an INVESTIGATION TRIGGER, not automatic proof that MLB or A1–A4 must change. "
        "The agent must first classify: load imbalance vs RF/interference vs CA/PCC vs UE capability vs mobility failure vs hardware/transport vs genuine L900 indoor dependence.",
        WHITE,
        48,
    )
    r += 1

    r = section(ws, r, cols, "0", "Workbook map (read in this sequence)", SECTION)
    r = flow_row(ws, r, ["00 Cover", "01 Idle Mode", "02 Connected Mode", "03 Intra-RAT MLB", "04 AI / CR Logic"], [NAVY, "1B4F72", "117A65", HUAWEI_RED, "6C3483"], cols)
    r = table(
        ws,
        r,
        ["Sheet", "Feature", "What you get", "Use when"],
        [
            ["00_Cover", "Introduction", "Layer policy, documents, color legend, safety rules", "First read / hand-over to another engineer"],
            ["01_Idle_Mode", "Feature 1 — Idle Mode Management", "11-item sequence + idle parameter + MML", "Camping / next-RRC-layer imbalance"],
            ["02_Connected_Mode", "Feature 2 — Connected Mode Mobility", "11-item sequence + A1–A5 + FreqPri + MML", "Coverage HO, A2 late, ping-pong, L900 return"],
            ["03_IntraRAT_MLB", "Feature 3 — Intra-RAT MLB", "11-item sequence + MLB algorithm + MML", "Active-user / PRB imbalance, throughput gap"],
            ["04_AI_CR_Logic", "Combined AI agent", "RCA tree, approval rule, CR template, rollback", "Daily KPI monitoring when gap > 2 Mbps"],
            ["05_Parameter_Index", "Cross-feature index", "All key parameters in one filterable table", "Search MO / parameter quickly"],
        ],
    )
    r += 1

    r = section(ws, r, cols, "0", "MAJOR safety notes (highlighted)", HUAWEI_RED)
    r = bullets(
        ws,
        r,
        cols,
        [
            "Do not force equal raw UE count or equal PRB% across unequal bandwidths. Use Load = N / C with ActiveUeBasedLoadEvalSw + SpectralEffBasedLoadEvalSw.",
            "L900 is a coverage safety layer. Permit coverage HO to L900. Prohibit routine connected/idle MLB targeting toward L900.",
            "Coverage mobility always preempts load/frequency-priority mobility. Never weaken L900 return protection only to raise average DL throughput.",
            "PRB-based MLB does not transfer CA UEs and does not guarantee user-experience fairness. Do not use it as the only algorithm on a high-CA L2600 network.",
            "The sample MML row in the original CSV (ENodeBAlgoSwitch / SymbolShutdownSwitch) is a POWER-SAVING example and is NOT used in this workbook.",
            "Huawei native intelligent MLB thresholds learn for 7 days and refresh every 7 days. An external AI must not overwrite them daily.",
            "Change only one parameter family per cluster per trial. Auto-rollback if drop, VoLTE, re-establishment, HO success or L900 indoor KPI degrades.",
            "All MML in this book uses placeholder LocalCellId / DlEarfcn / GroupId. Substitute site-specific IDs and validate syntax in MAE-Access / MML help of the running eRAN version.",
        ],
        AMBER,
        20,
    )
    r += 1
    r = add_legend_row(ws, r, cols)
    return ws


# ---------------------------------------------------------------------------
# FEATURE 1 IDLE
# ---------------------------------------------------------------------------
def build_idle(wb):
    ws = wb.create_sheet("01_Idle_Mode")
    cols = 10
    set_widths(ws, [8, 28, 22, 18, 18, 18, 22, 28, 22, 36])
    apply_sheet_setup(ws, "Feature 1 — Idle Mode Management")

    r = 1
    r = banner(ws, r, cols, "  FEATURE 1   ·   Idle Mode Management", HUAWEI_RED, font_title, 32)
    r = banner(ws, r, cols, "  Source: eRAN Idle Mode Management Feature Parameter Description, eRAN21.1, Issue 04 (2026-05-30)   ·   Related: MLB §5 idle transfer, Connected Mode frequency-priority", NAVY, font_white_s, 20)
    r = add_legend_row(ws, r, cols)
    r += 1

    # SN1
    r = section(ws, r, cols, "1", "Working Principle")
    r = para(
        ws,
        r,
        cols,
        "An LTE UE in RRC_IDLE performs PLMN selection, cell selection, system-information reception, measurement, reselection and paging. "
        "After PLMN selection the UE searches a suitable cell, camps, reads SI, measures serving/neighbor cells, reselections by frequency priority + thresholds + timers, "
        "and on service initiation establishes RRC on the cell where it is then camped. Idle distribution therefore decides the NEXT access / PCC layer. "
        "Ref: [IM] Fig 4-1 §4 printed p.6; Fig 5-1 §5.1 pp.13–14.",
        WHITE,
        56,
    )
    r = note(ws, r, cols, "  CHART / SEQUENCE (right-side process requested in template) — idle chain from power-on to next RRC:", LIGHT, 18)
    r = flow_row(ws, r, ["PLMN select", "Cell select (S)", "Camp + read SIB3/5", "Measure / reselect", "RRC on camped cell"], ["1B4F72", "1A5276", "117A65", "B9770E", HUAWEI_RED], cols)
    r = table(
        ws,
        r,
        ["Step", "State / action", "What the UE uses", "Key MO / parameter family", "Robi meaning", "Doc ref"],
        [
            ["A", "Initial / stored / release-directed selection", "Last camped cell, RRC release frequency, or strongest cell per scanned band", "CellSel.*, RRCConnectionRelease", "SIB priority does NOT fully control first camp after power-up", "[IM] §5.1.2 pp.15–16"],
            ["B", "Suitability Criterion S", "Srxlev>0 and (if configured) Squal>0", "QRxLevMin, QQualMin, PMax / UePowerMax", "Admission floor — not a load knob", "[IM] §5.1.2 pp.15–16"],
            ["C", "Common priority", "SIB3 serving priority, SIB5 inter-frequency priority", "CellResel.CellReselPriority, EutranInterNFreq.CellReselPriority", "L2600 highest equal; L900 lowest", "[IM] §5.1.3.1 pp.17–18"],
            ["D", "Measurement rules", "Higher priority always measured; equal/lower only after search threshold", "SIntraSearch, SNonIntraSearch, MeasPerformanceDemand", "L900 UE continuously looks for L2600/L1800/L2100", "[IM] §5.1.3.3 pp.23–25"],
            ["E", "Higher-priority reselection", "Target S > ThreshXhigh / ThreshXhighQ for EutranReselTime", "ThreshXhigh, ThreshXhighQ", "L900 → capacity only if high-band is truly good", "[IM] Tables 5-1/5-2 pp.29–32"],
            ["F", "Lower-priority reselection", "Serving < ThrshServLow AND target > ThreshXlow", "ThrshServLow, ThreshXlow", "Capacity → L900 indoor fallback", "[IM] Tables 5-3/5-4 pp.33–35"],
            ["G", "Equal-priority ranking", "Rn = Qmeas,n − Qoffset ; Rs = Qmeas,s + Qhyst", "Qhyst, QoffsetFreq, CellQoffset", "How four L2600 carriers compete", "[IM] §5.1.3.4 pp.27–28"],
            ["H", "Dedicated priority / idle MLB", "IdleModeMobilityControlInfo in RRCConnectionRelease; lives until T320 / next RRC", "T320ForLoadBalance, InterFreqIdleMlbSwitch", "Steers NEXT session, not a long connected session", "[IM] §5.1.3.1 pp.18–22; [MLB] §5.1.1.5"],
        ],
        col_fills={5: GREEN},
    )
    r += 1

    # SN2
    r = section(ws, r, cols, "2", "Major Highlighted Points")
    r = bullets(
        ws,
        r,
        cols,
        [
            "Priorities are FREQUENCY-level, not per-cell. No SIB5 priority ⇒ no reselection to that frequency. Max 16 non-serving E-UTRAN frequencies in SIB5. [IM] §5.1.3.1",
            "Initial cell selection is not frequency-priority balancing. Evaluate camping after reselection has had time to run, not only immediate attach. [IM] §5.1.2",
            "Higher-priority frequencies are measured regardless of serving quality. Thresholds do not stop L900 UEs from searching L2600; they only decide whether L2600 may WIN. [IM] §5.1.3.3",
            "QRxLevMin / QQualMin are suitability floors. Do not use them as daily MLB knobs. Wrong tightening of L900 floors causes indoor access failure.",
            "Dedicated priorities REPLACE common SIB priorities and are discarded when the UE enters connected mode, does PLMN selection, or T320 expires.",
            "Updated SIB is applied in the next SI modification period. UEs otherwise reread SI after change-paging or after 3 hours. Do not judge an idle change in the same 15-min period. [IM] §7.1.3",
            "Huawei recommends SNonIntraSearchCfgInd=CFG and SIntraSearch > SNonIntraSearch; example SNonIntraSearch=10 to reduce return ping-pong. [IM] §5.4.1.1; [MLB] §5.1.2.1",
            "MeasPerformanceDemand=UNDELIVER removes a frequency from SIB5 and blocks idle MLB targeting. Never use UNDELIVER as congestion control on a capacity carrier. [IM] §5.3.2.3",
            "Document inconsistency: optimization table text on QRxLevMinOffset conflicts with the Criterion-S formula. Do not use that table sentence as a layer-balance rule. [IM] §5.1.2 vs §5.4.1.1.3",
        ],
        IVORY,
        20,
    )
    r += 1

    # SN3
    r = section(ws, r, cols, "3", "Benefit and Limitations")
    r = table(
        ws,
        r,
        ["Type", "Item", "Detail", "Impact on Robi throughput-gap work", "Doc / engineering"],
        [
            ["Benefit", "Low signalling cost", "Idle transfer avoids gap-assisted inter-frequency measurement and connected HO", "Preferred first step for predictable busy-hour camping bias", "[MLB] §4.3 pp.18–19"],
            ["Benefit", "Sets next PCC/access layer", "RRC starts on the camped cell", "Reduces L900-started sessions when high-band is usable", "Engineering"],
            ["Benefit", "Dedicated priority is temporary", "T320-bounded; does not permanently rewrite SIB", "Safer than changing broadcast priority every day", "[IM] §5.1.3.1"],
            ["Limitation", "Not real-time", "Cannot rebalance a UE already in a long RRC session", "Must pair with connected MLB / CA", "Engineering"],
            ["Limitation", "Static common priority packs the highest usable tier", "Equal L2600 priority still ranks by RF, not by load", "Need idle MLB + connected MLB for C1–C4 fairness", "[IM] §5.1.3.5"],
            ["Limitation", "SI delay", "Change is not instant; BER of SI broadcast must be ≤1%", "Wait SI modification + camping turnover before KPI judge", "[IM] §5.3.4, §7.1.3"],
            ["Limitation", "UE capability split", "Legacy UEs may see only NORMAL frequencies; max 8 dedicated freqs without enhanced meas", "Audit camping by UE band class", "[IM] §5.1.3.1 / §5.1.3.3"],
            ["Limitation", "Load-based redirection is blind to location quality", "Selects by QoS/priority, not measured RF", "Use only with verified co-coverage", "[IM] §6.1.1 pp.69–70"],
            ["Limitation", "Adaptive-proportion idle balancing", "Huawei explicitly does not recommend without Huawei engineering support", "Do not enable in the AI agent", "[MLB] §5.5.2.1 p.80"],
        ],
        col_fills={1: REF},
    )
    r += 1

    # SN4
    r = section(ws, r, cols, "4", "Selection Criteria / Trigger Condition")
    r = para(ws, r, cols, "Idle mode has no 'UE selection' like connected MLB. The UE applies 3GPP Criterion S + reselection rules. eNodeB 'selects' only when it writes dedicated priorities at RRC release (idle MLB / SPID / PCC / operator).", WHITE, 36)
    r = table(
        ws,
        r,
        ["Rule", "Formula / condition", "Parameters", "Robi application", "Pass / fail meaning"],
        [
            ["Criterion S (RSRP)", "Srxlev = Qrxlevmeas − (Qrxlevmin + offset) − Pcompensation  > 0", "QRxLevMin, QRxLevMinOffset, UePowerMax / PMax", "Keep L900 floors coverage-safe; do not raise to push users off L900", "Fail = cell not suitable"],
            ["Criterion S (RSRQ)", "Squal = Qqualmeas − (Qqualmin + offset) > 0  (if QQualMin configured ≠ 0/absent)", "QQualMin, QQualMinOffset", "Optional interference gate; RSRQ moves with load — introduce only after live check", "If QQualMin absent, RSRQ not used"],
            ["Higher-priority target", "Camped >1 s AND target S > ThreshXhigh[/Q] for EutranReselTime", "ThreshXhigh, ThreshXhighQ, EutranReselTime", "L900→L2600/L1800/L2100 only if target is robust", "Protects indoor users from premature high-band camp"],
            ["Lower-priority target", "No higher target AND serving S < ThrshServLow[/Q] AND target S > ThreshXlow[/Q]", "ThrshServLow, ThreshXlow", "Capacity→L900 indoor rescue", "Too-low ThrshServLow = sticky poor high-band"],
            ["Equal-priority rank", "Rn > Rs for reselection time; Rn = Qmeas,n − (QoffsetFreq + CellQoffset)", "Qhyst, QoffsetFreq, CellQoffset", "Four L2600 carriers start equal; small offset only for stable RF asymmetry", "Hourly load must NOT be chased by static offset"],
            ["Inter-freq search start", "Equal/lower measured when Srxlev ≤ SNonIntraSearchP or Squal ≤ SNonIntraSearchQ", "SNonIntraSearch, SNonIntraSearchQ", "Does not suppress higher-priority search from L900", "CFG recommended to avoid always-on equal/lower search"],
            ["Idle MLB release trigger", "Source idle-user load ≥ InterFreqIdleMlbUeNumThd (+offset) and target admitted", "InterFreqIdleMlbSwitch, InterFreqIdleMlbUeNumThd", "Dynamic dedicated priority toward low-load capacity freq", "See Feature 3"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    # SN5
    r = section(ws, r, cols, "5", "Activation Parameter / Switch  —  Core settings in table")
    r = note(ws, r, cols, "  Conditional Parameter = what must be ON/OFF for the main switch to take effect. Green = Robi proposed starting policy (calibrate). Amber = do not enable without extra design.", IVORY, 22)
    r = table(
        ws,
        r,
        ["SN", "MO", "Parameter / Switch", "Description", "Proposed value", "Conditional parameter", "Notes", "Layer apply", "Doc ref"],
        [
            ["5.1", "CellResel", "CellReselPriority", "Serving-frequency common reselection priority in SIB3. Larger = higher.", "L2600=7; L1800=6; L2100=5 or 6; L900=2", "Must be broadcast; UE uses dedicated list instead if present", "Do not give four L2600 carriers different common priorities", "All", "[IM] §5.1.3.1"],
            ["5.2", "EutranInterNFreq", "CellReselPriorityCfgInd", "Whether SIB5 carries priority for this non-serving frequency", "CFG", "Required or UE will not reselect to that frequency", "Configure for all 6 paired capacity freqs + L900", "All inter-freq", "[IM] §5.1.3.1"],
            ["5.3", "EutranInterNFreq", "CellReselPriority", "SIB5 priority of the target frequency", "Same hierarchy as 5.1", "CellReselPriorityCfgInd=CFG", "Highest equal across L2600 C1–C4", "Target freq", "[IM] §5.1.3.1"],
            ["5.4", "EutranInterNFreq", "MeasPerformanceDemand", "NORMAL / REDUCED / UNDELIVER visibility and meas performance", "NORMAL on all principal LTE layers", "UNDELIVER ⇒ not in SIB5 and not idle-MLB target", "Never UNDELIVER a capacity carrier to 'save load'", "All", "[IM] §5.1.3.3; §5.3.2.3"],
            ["5.5", "CellResel", "SNonIntraSearchCfgInd + SNonIntraSearch", "Start equal/lower-priority inter-freq search", "CFG ; example 10 (Huawei idle-MLB note)", "Set SIntraSearch > SNonIntraSearch", "Reduces ping-pong return; does not stop higher-priority search", "Serving", "[IM] §5.4.1.1; [MLB] §5.1.2.1"],
            ["5.6", "CellAlgoSwitch.MlbAlgoSwitch", "InterFreqIdleMlbSwitch", "Enables intra-LTE idle-mode load equalization / dedicated-priority steering", "ON on capacity layers after connected MLB design is ready", "InterFreqMlbSwitch usually also needed; target MlbTargetInd must allow idle MLB", "Do not combine with fixed-proportion idle balancing (ping-pong warning)", "Capacity", "[IM] §5.3.2.3; [MLB] §5.4.2.2"],
            ["5.7", "EnhancedMlbAlgoSwitch", "DediPrioManageOnLowLoadSw (+ EnhSw)", "Keeps released UEs from returning to higher-load frequencies", "ON (evaluate EnhSw for all RRC-release causes)", "Idle MLB / dedicated-priority path", "Helps hold the steered camping layer for T320", "Capacity", "[MLB] §5.3.1 Table 5-10"],
            ["5.8", "RrcConnStateTimer", "T320ForLoadBalance", "Lifetime of load-balance dedicated priorities", "Start MIN30 or MIN60; do not jump to 180 min on day 1", "Used for MLB release; SPID/PCC anchoring uses 180 min always", "Too long = stale steering after load flips", "eNodeB", "[IM] §5.1.3.1"],
            ["5.9", "GlobalProcSwitch.ProtocolMsgOptSwitch", "CellReselectionOptSwitch", "Improves preferential delivery of LTE frequencies in dedicated priority", "ON when idle MLB is ON", "Dedicated-priority path", "Huawei recommendation with idle equalization", "eNodeB", "[IM] §5.3.2.3"],
        ],
        col_fills={5: GREEN, 6: AMBER},
    )
    r += 1

    # SN6
    r = section(ws, r, cols, "6", "Prerequisite Functions")
    r = table(
        ws,
        r,
        ["SN", "Prerequisite", "MO / check", "Proposed / required", "If missing, what fails", "Notes"],
        [
            ["6.1", "Complete SIB5 inter-frequency set", "EutranInterNFreq for L900/1800/2100/2600 C1–C4", "All principal freqs present, NORMAL", "UE cannot reselect to missing freq", "Max 16 non-serving E-UTRAN freqs"],
            ["6.2", "Neighbor list not truncated", "EutranInterFreqNCell ; SIB4/SIB5 ≤16 listed neighbors per frequency", "Cosited sector neighbors all listed", "Looks like a 'threshold problem' but is NRT/SIB truncation", "Non-zero offsets affect which 16 are chosen"],
            ["6.3", "Blacklist not hiding a valid layer", "InterFreqBlkCell.ApplicationScope", "Do not blacklist a capacity layer for load reasons", "65535 = idle+connected; only 16 idle blacklists delivered", "Blacklist = invalid target, not MLB"],
            ["6.4", "SI broadcast quality", "SIB BER ≤ 1%", "Healthy", "Failed SIB5 decode looks like failed layer balance", "[IM] §5.3.4"],
            ["6.5", "UE band capability inventory", "UE capability / SPID", "Know % of UEs without L2600 or L2100", "Idle priority cannot move incapable UEs", "Segment KPI by band class"],
            ["6.6", "X2 / intra-eNB neighbor relation for later connected/idle MLB", "EutranInterFreqNCell.NoHoFlag, OverlapInd", "PERMIT and correct overlap on capacity pairs", "Idle MLB target admission fails", "See Feature 3"],
            ["6.7", "Do not enable adaptive-proportion idle MLB", "Fixed/adaptive proportion features", "OFF unless Huawei on-site support", "Huawei not-recommended; ping-pong with user-number MLB", "[MLB] §5.5.2.1"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    # SN7
    r = section(ws, r, cols, "7", "Mutually Impacted Parameters / Features")
    r = table(
        ws,
        r,
        ["This idle control", "Conflicts / couples with", "What goes wrong", "Proposed coordination", "Conditional action"],
        [
            ["High L2600 common priority", "Connected MLB moving PCC to L1800, then release back to L2600", "Cyclic steering / ping-pong", "Idle sets initial anchor; connected MLB owns the session; dedicated priority only if target still low-load", "Do not change priority + offset + idle MLB + connected MLB the same day"],
            ["Fixed-proportion idle balancing", "User-number connected MLB", "Huawei ping-pong warning", "Keep fixed-proportion OFF", "Mutually avoid"],
            ["Aggressive ThreshXhigh (too low)", "L900 indoor UL / VoLTE", "Premature high-band camp, setup fail, low TP", "Require robust high-band S before leaving L900", "Amber — coverage"],
            ["Aggressive ThrshServLow (too low)", "Coverage HO A2/A5", "UE stays on dying L2600 too long in idle then accesses poorly", "Align idle serving-low with connected A2 philosophy", "Coordinate with Feature 2"],
            ["Negative QoffsetFreq toward one L2600", "CA PCC anchoring + FreqPri HO", "One carrier permanently hoards idle UEs", "Offsets only for stable RF, not hourly load", "Prefer idle MLB dedicated priority"],
            ["T320 long", "Fast traffic tide", "UEs keep old low-load target after that carrier becomes busy", "30–60 min start", "Review weekly"],
            ["RSRQ-based reselection", "Scheduler load (RSRQ falls when cell is busy)", "Oscillation: busy → RSRQ bad → leave → RSRQ recovers → return", "RSRP primary; RSRQ only after distribution check", "Same warning as connected mode"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    # SN8
    r = section(ws, r, cols, "8", "Relation with Other Features")
    r = table(
        ws,
        r,
        ["Related feature", "How it relates", "What to check together", "Robi rule"],
        [
            ["Intra-RAT MLB (Feature 3)", "Idle MLB is the idle transfer method of MLB (dedicated priority + T320)", "InterFreqIdleMlbSwitch, MlbTargetInd idle allow/forbid, T320", "Idle MLB ON among capacity layers; FORBID idle MLB target = L900"],
            ["Connected Mode (Feature 2)", "Dedicated idle priorities are discarded at RRC connect; A1–A5 / FreqPri take over", "A1 L900-escape vs ThreshXhigh; A5 L900-return vs ThrshServLow", "Idle = next access; connected = session safety + MLB"],
            ["Carrier Aggregation", "PCC anchoring and CaUserLoadTransferSw change whether CA-capable UEs are idle-steered; PCC ≠ SCC load", "PCell users, SCell users, active CC", "Do not call a layer 'unloaded' if it is a busy SCell"],
            ["SPID / operator / RAN sharing", "Dedicated priorities can come from SPID or operator policy, not only MLB", "SPID maps, T320=180 min for SPID/PCC", "Do not let SPID fight MLB daily"],
            ["Redirection", "RRC release to a frequency; load-based redirection is not RF-measured", "Use only with full overlap", "Prefer measurement-based connected HO"],
            ["Energy saving / carrier shutdown", "A sleeping L2600 must not remain highest idle priority target", "ES state vs SIB5 list", "Coordinate with ES team before MLB/idle changes"],
            ["GERAN / UTRAN reselection", "MLB dedicated class order: NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "2G/3G priority below L900", "Do not dump LTE overflow onto 2G"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    # SN9
    r = section(ws, r, cols, "9", "License Requirements")
    r = table(
        ws,
        r,
        ["Function", "License / entitlement (verify in MAE-Access / license file)", "If license missing", "Robi action"],
        [
            ["Basic idle selection / SIB3/SIB5 reselection", "Basic LTE eNodeB — normally no extra feature license", "N/A", "Always available"],
            ["User-number / PRB idle MLB, dedicated-priority load release", "Intra-RAT Mobility Load Balancing (same family as connected MLB) — confirm exact license name on site", "Idle MLB MML accepted but feature inactive / counters stay 0", "Check license before enabling InterFreqIdleMlbSwitch"],
            ["InterFreqBlindMlbSwitch idle/blind path", "Intra-RAT MLB + blind MLB option if sold separately", "Blind path unavailable", "Keep blind OFF unless full containment proven"],
            ["DediPrioManageOnLowLoad(Enh)", "Enhanced MLB / idle-priority management option — verify", "Released UEs bounce back to high-load freq", "Enable only after license confirmed"],
            ["Enhanced measurement / 16 dedicated frequencies", "UE capability incMonEUTRA, not only eNB license", "Legacy UEs get ≤8 NORMAL freqs", "Keep all primary layers NORMAL"],
        ],
        col_fills={2: YELLOW},
    )
    r = note(ws, r, cols, "  Exact Huawei license IDs are in the matching eRAN21.1 License Documentation / MAE license center. This sheet lists function-to-license mapping from the feature books; confirm on the live eNodeB before CR.", REF, 28)
    r += 1

    # SN10
    r = section(ws, r, cols, "10", "All Parameter List  —  Feature 1  (sequence / connected order)")
    r = note(ws, r, cols, "  Order = engineering sequence: suitability → serving broadcast → inter-freq broadcast → measurement → reselection timers → dedicated/idle MLB. Add more columns anytime. Conditional Parameter here = relation with other feature as requested.", IVORY, 22)
    idle_params = [
        ["1", "CellSel", "QRxLevMin", "Minimum RSRP for cell selection Criterion S", "Keep current coverage-safe value; do not use as MLB", "Pcompensation / UePowerMax", "Not a daily load knob", "Selection", "All", "[IM] §5.1.2", "Amber"],
        ["2", "CellSel", "QQualMin", "Minimum RSRQ for selection; 0/absent ⇒ RSRP-only", "Keep absent/0 until RSRQ plan approved", "QQualMinOffset", "RSRQ moves with load", "Selection", "All", "[IM] §5.1.2", ""],
        ["3", "Cell", "UePowerMax", "Used in Pcompensation for Srxlev", "Do not change for layer balance", "UE max Tx power", "UL constraint", "Selection", "All", "[IM] §5.1.2", "Amber"],
        ["4", "CellResel", "QRxLevMin / QQualMin / PMax", "Reselection suitability of serving/intra", "Coverage-safe; L900 must remain campable indoors", "Same as selection philosophy", "Protect L900 indoor", "Resel suitability", "Especially L900", "[IM] §5.1.3.4", "Yellow"],
        ["5", "CellResel", "CellReselPriority", "SIB3 serving common priority", "L2600=7; L1800=6; L2100=5/6; L900=2", "SIB3 broadcast", "Equal L2600 C1–C4", "Priority", "Serving", "[IM] §5.1.3.1", "Green"],
        ["6", "EutranInterNFreq", "CellReselPriorityCfgInd", "SIB5 priority present", "CFG", "Without CFG no reselection to that freq", "Mandatory for 6+1 layers", "Priority", "Inter-freq", "[IM] §5.1.3.1", "Green"],
        ["7", "EutranInterNFreq", "CellReselPriority", "SIB5 target priority", "Mirror serving hierarchy", "CfgInd=CFG", "Do not stack C1>C2>C3>C4", "Priority", "Inter-freq", "[IM] §5.1.3.1", "Green"],
        ["8", "EutranInterNFreq", "QRxLevMin / QqualMin / Pmax", "Target-frequency suitability", "High-band slightly stricter than L900 if UL-limited; never block L900", "UL / PRACH KPI", "DL-only qualification is dangerous on L2600 indoor", "Target S", "Target", "[IM] §5.1.3.4", "Amber"],
        ["9", "CellResel", "SIntraSearchCfgInd / SIntraSearch / SIntraSearchQ", "Skip intra-freq meas when serving is very good", "CFG; SIntraSearch > SNonIntraSearch", "Battery vs intra HO readiness", "Huawei recommends CFG", "Meas", "Serving", "[IM] §5.1.3.3 / §5.4.1.1", "Green"],
        ["10", "CellResel", "SNonIntraSearchCfgInd / SNonIntraSearch / Q", "Start equal/lower inter-freq search", "CFG; example 10", "Does not stop higher-priority meas", "MLB doc suggests small value to cut ping-pong", "Meas", "Serving", "[IM] §5.4.1.1; [MLB] §5.1.2.1", "Green"],
        ["11", "EutranInterNFreq", "ThreshXhigh / ThreshXhighQ", "Higher-priority target must exceed this", "Calibrate from MR: L900→high-band only if target robust (start from current, ±1–2 dB max per trial)", "EutranReselTime", "Too low = indoor L2600 camp", "Higher-prio reselection", "L900 source", "[IM] Tables 5-1/5-2", "Yellow"],
        ["12", "CellResel", "ThrshServLow / ThrshServLowQ", "Serving must be this poor before lower-prio reselection", "Keep high-band→L900 possible before access collapse", "A2/A5 connected alignment", "Too low = sticky dying L2600", "Lower-prio reselection", "Capacity serving", "[IM] Tables 5-3/5-4", "Yellow"],
        ["13", "EutranInterNFreq", "ThreshXlow / ThreshXlowQ", "Lower-priority (L900) target quality floor", "L900 must be actually usable", "ThrshServLow", "Do not open L900 fallback to a weak L900", "Lower-prio reselection", "L900 target", "[IM] Tables 5-3/5-4", "Yellow"],
        ["14", "CellResel", "Qhyst", "Serving stickiness in equal-priority rank", "Keep default unless ping-pong", "QoffsetFreq", "Large Qhyst slows needed reselection", "Equal-prio rank", "Serving", "[IM] §5.1.3.4", ""],
        ["15", "EutranInterNFreq", "QoffsetFreq", "Frequency offset in Rn", "0 among L2600 C1–C4 at start", "CellQoffset", "Static bias only for proven RF asymmetry", "Equal-prio rank", "L2600", "[IM] §5.1.3.4", "Green"],
        ["16", "EutranInterFreqNCell", "CellQoffset", "Per-neighbor idle offset", "0 unless overshoot/NRT issue", "SIB neighbor list (max 16)", "Not a daily load tool", "Equal-prio rank", "Per cell", "[IM] §5.1.3.4", "Amber"],
        ["17", "CellResel", "TReselEutran", "Intra-freq reselection timer", "Keep default urban value", "Speed-scaling", "Small = ping-pong", "Timer", "Serving", "[IM] §5.4.1.1.3", ""],
        ["18", "EutranInterNFreq", "EutranReselTime", "Inter-freq reselection persistence", "Keep / slightly longer than intra if ping-pong", "ThreshXhigh/low", "Must persist for the timer", "Timer", "Inter-freq", "[IM] §5.1.3.4", ""],
        ["19", "EutranInterNFreq", "MeasPerformanceDemand", "NORMAL / REDUCED / UNDELIVER", "NORMAL", "Idle MLB target eligibility", "UNDELIVER forbidden on capacity", "Meas delivery", "All", "[IM] §5.1.3.3", "Green"],
        ["20", "CellResel", "SpeedDepReselCfgInd + SfMedium/High", "Speed-based timer/hyst scaling", "Do not use to solve static layer imbalance", "NCellChangeMedium/High, TEvaluation", "High-speed corridors only", "Speed", "Highway clusters", "[IM] §5.1.3.6", "Amber"],
        ["21", "InterFreqBlkCell", "ApplicationScope", "Idle/connected blacklist scope", "Do not load-blacklist a good layer", "Max 16 idle blacklists delivered", "65535 = both modes", "Blacklist", "Invalid cells only", "[IM] §5.1.3.2", "Amber"],
        ["22", "RrcConnStateTimer", "T320ForLoadBalance", "Dedicated-priority lifetime for MLB release", "MIN30 or MIN60 to start", "InterFreqIdleMlbSwitch", "SPID/PCC always 180 min", "Dedicated prio", "eNodeB", "[IM] §5.1.3.1", "Green"],
        ["23", "RrcConnStateTimer", "T320ForOther", "T320 for other release causes", "Keep current unless design says else", "DediPrioManageOnLowLoadEnhSw", "EnhSw can apply to all releases", "Dedicated prio", "eNodeB", "[MLB] Table 5-10", ""],
        ["24", "CellAlgoSwitch", "InterFreqIdleMlbSwitch", "Idle MLB enable bit", "ON capacity; see Feature 3", "InterFreqMlbSwitch, MlbTargetInd", "Not with fixed-proportion", "Idle MLB", "Capacity", "[IM] §5.3.2.3", "Green"],
        ["25", "EnhancedMlbAlgoSwitch", "DediPrioManageOnLowLoadSw", "Hold low-load dedicated priority", "ON with idle MLB", "T320", "Reduces bounce-back", "Idle MLB", "Capacity", "[MLB] §5.3.1", "Green"],
        ["26", "EutranInterNFreq", "MlbTargetInd", "Whether frequency may be idle and/or connected MLB target", "Capacity: ALLOWED; toward L900: WITHOUT_IDLE_MLB and WITHOUT_CONNECT_MLB (verify enum on site)", "OverlapInd, NoHoFlag", "Coverage HO to L900 must remain", "Idle+Conn MLB", "Per frequency", "[MLB] pp.28, 129", "Yellow"],
        ["27", "CellMLB", "InterFreqIdleMlbUeNumThd / IdleUE transfer type", "Idle user-number trigger", "Calibrate after 7-day baseline; do not copy 20 MHz threshold onto L2100/L900", "Active/SE eval switches", "See Feature 3", "Idle MLB trigger", "Capacity", "[MLB] Table 5-2", "Green"],
        ["28", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Filter frequency from dedicated-priority delivery", "Use only for true SCC-only frequencies", "Enhanced meas capability", "Huawei: recommended only for SCC-only", "Filter", "SCC-only", "[IM] §5.1.3.1", "Amber"],
    ]
    r = table(
        ws,
        r,
        ["Seq", "MO", "Parameter", "Description", "Proposed value", "Conditional parameter / other-feature relation", "Notes", "Function group", "Apply on", "Doc ref", "Flag"],
        idle_params,
        col_fills={5: GREEN},
        min_h=26,
    )
    r += 1

    # SN11
    r = section(ws, r, cols, "11", "Final MML Command for Activations  —  maintain sequence")
    r = note(
        ws,
        r,
        cols,
        "  SEQUENCE RULE: audit/list first → set common priority/SIB → set search/reselection thresholds → allow idle MLB target on capacity only → enable idle MLB switches → set T320. "
        "Replace LocalCellId / DlEarfcn with live values. Commands are eRAN-style MOD/LST. Validate on a lab or single cluster. The original CSV SymbolShutdownSwitch row is deleted on purpose.",
        AMBER,
        40,
    )
    r = table(
        ws,
        r,
        ["Seq", "MO", "Activation value / MML skeleton", "Conditional parameter", "Remarks", "Parameter description", "More notes"],
        [
            ["0", "—", "LST CELLRESEL: LocalCellId=<x>;  LST EUTRANINTERNFREQ: LocalCellId=<x>;  LST CELLALGOSWITCH: LocalCellId=<x>;  LST RRCCONNSTATETIMER:;", "Read-only", "Always dump current before change", "Baseline", "Keep LST output with the CR"],
            ["1", "CELLRESEL", "MOD CELLRESEL: LocalCellId=<x>, CellReselPriority=<7|6|5|2>, SNonIntraSearchCfgInd=CFG, SNonIntraSearch=10, SIntraSearchCfgInd=CFG;", "SIntraSearch > SNonIntraSearch", "Priority by layer: L2600=7, L1800=6, L2100=5/6, L900=2", "Serving common priority + search start", "One layer per command; do not broadcast a new hierarchy on all sites the same night"],
            ["2", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>, CellReselPriorityCfgInd=CFG, CellReselPriority=<prio>, MeasPerformanceDemand=NORMAL;", "EARFCN exists; neighbor cells listed", "Repeat for every paired capacity + L900 frequency", "SIB5 priority + visibility", "UNDELIVER forbidden on capacity"],
            ["3", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<high-band>, ThreshXhigh=<calib>, ThreshXlow=<calib>, QoffsetFreq=0, EutranReselTime=<current>;", "Idle Feature 1 SN-4 rules", "L900 source: ThreshXhigh protects indoor; L2600 mutual: QoffsetFreq=0", "Higher/lower qualification and equal-prio offset", "±1–2 dB per trial max"],
            ["4", "CELLRESEL", "MOD CELLRESEL: LocalCellId=<x>, ThrshServLow=<calib>;", "Align with connected A2/A5 philosophy", "Capacity cells: allow L900 fallback before access collapse", "Serving-low for lower-prio reselection", "Yellow — L900 protection"],
            ["5", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;  (also WITHOUT_IDLE_MLB if enum available — verify on NE)", "Coverage HO to L900 must stay allowed (NoHoFlag=PERMIT for coverage path)", "L900 is not a capacity MLB target", "Target eligibility", "Verify exact enum in eRAN21.1 parameter reference"],
            ["6", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L1800|L2100|L2600>, MlbTargetInd=ALLOWED;", "OverlapInd valid; NoHoFlag=PERMIT_HO_ENUM", "Capacity layers may be idle+connected MLB targets", "Target eligibility", "From L900 toward capacity is allowed"],
            ["7", "CELLALGOSWITCH", "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;", "License present; Feature 3 connected settings designed", "Turn idle MLB on only after target indications are correct", "Idle+connected MLB master bits", "Do not enable InterFreqBlindMlbSwitch in the same step"],
            ["8", "ENHANCED MLB / CELLALGOSWITCH", "MOD CELLALGOSWITCH / related EnhancedMlb MO: DediPrioManageOnLowLoadSw=ON;  (EnhSw only after test)", "T320 set", "Holds steered camping", "Dedicated priority hold", "Exact MO name: confirm in MAE of running version"],
            ["9", "RRCCONNSTATETIMER", "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;", "Idle MLB ON", "Start conservative", "Dedicated-priority lifetime", "Increase only if bounce-back is proven"],
            ["10", "CELLMLB", "MOD CELLMLB: LocalCellId=<x>, InterFreqUeTrsfType=IdleUE, InterFreqIdleMlbUeNumThd=<calib>;   (if this site uses a dedicated idle trigger set — some networks keep connected+idle types on the same MO family)", "User-number mode; do not mix fixed-proportion", "See Feature 3 for full CELLMLB", "Idle trigger", "If both idle and connected types exist, follow live MO help"],
            ["11", "Verification", "LST CELLRESEL / EUTRANINTERNFREQ / CELLALGOSWITCH / RRCCONNSTATETIMER / CELLMLB;  monitor L.RRCRel.load.DedicatedPri.LTE.High and camping/RRC-setup by layer for ≥1 SI modification + busy hour", "SI modification period elapsed", "Do not judge in the same quarter-hour", "KPI / counter close-loop", "Counters: [MLB] idle dedicated-priority counters"],
        ],
        col_fills={3: GREEN, 4: AMBER},
        min_h=36,
    )
    r += 1
    r = note(ws, r, cols, "  Chart reminder: Idle Mode does not equalize four L2600 carriers by common priority. Common priority = coverage intent. Dedicated priority = release-time load. Connected MLB/CA = session load. Ref [IM] conclusion + [MLB] §5.1.1.5.", REF, 28)
    return ws


# ---------------------------------------------------------------------------
# FEATURE 2 CONNECTED
# ---------------------------------------------------------------------------
def build_connected(wb):
    ws = wb.create_sheet("02_Connected_Mode")
    cols = 10
    set_widths(ws, [8, 28, 22, 18, 18, 18, 24, 28, 22, 36])
    apply_sheet_setup(ws, "Feature 2 — Connected Mode Mobility")

    r = 1
    r = banner(ws, r, cols, "  FEATURE 2   ·   Mobility Management in Connected Mode", HUAWEI_RED, font_title, 32)
    r = banner(ws, r, cols, "  Source: eRAN Mobility Management in Connected Mode, eRAN21.1, Issue 08 (2026-06-30)   ·   Load-HO algorithm details are in the MLB book, not fully in this book", NAVY, font_white_s, 20)
    r = add_legend_row(ws, r, cols)
    r += 1

    r = section(ws, r, cols, "1", "Working Principle")
    r = para(
        ws,
        r,
        cols,
        "Connected-mode mobility keeps an RRC_CONNECTED UE on a usable cell. The common flow is: initiation decision → measurement-based or blind → deliver measurement configuration → "
        "receive report → choose target/policy → execute HO → retry/penalty. Huawei classifies: (1) necessary/coverage HO to avoid drop; (2) unnecessary/offload HO with cause "
        "'Reduce Load in Serving Cell'; (3) unnecessary/optimization HO including frequency-priority HO. This document does NOT contain the full MLB algorithm — see Feature 3. "
        "Ref: [CM] Table 3-1 pp.21–23; Fig 4-1 §§4.1.1–4.1.8 pp.29–65; Handover Function Classification pp.26–28.",
        WHITE,
        64,
    )
    r = note(ws, r, cols, "  CHART / STATE MODEL (right-side process):", LIGHT, 16)
    r = flow_row(ws, r, ["RRC connected", "A2 / MLB / FreqPri start meas", "A3/A4/A5 report", "Admit + HO", "A1 stop / protect timer"], ["1B4F72", "B9770E", "117A65", HUAWEI_RED, "6C3483"], cols)
    r = table(
        ws,
        r,
        ["Event", "Enter condition (concept)", "Leave / stop", "Robi use", "Doc"],
        [
            ["A1 serving good", "Ms − Hys > Thresh for TTT", "Ms + Hys < Thresh", "Stop coverage meas; L900→capacity gate when high-band overlap is implied", "[CM] Table 4-8 pp.41–46"],
            ["A2 serving poor", "Ms + Hys < Thresh for TTT", "Ms − Hys > Thresh", "Start inter-freq coverage meas (L2600 indoor → measure L1800/L2100/L900)", "[CM] Table 5-1 pp.97–98"],
            ["A3 relative better", "Mn+Ofn+Ocn−Hys > Ms+Ofs+Ocs+Off", "reverse + Hys", "Among similar overlapping capacity carriers", "[CM] Table 5-16 p.127"],
            ["A4 absolute target good", "Mn+Ofn+Ocn−Hys > Thresh", "reverse + Hys", "MLB / FreqPri: target is 'good enough' without beating serving", "[CM] Table 5-22 p.145; [MLB] pp.137–139"],
            ["A5 serving poor AND target good", "Ms+Hys<Th1 AND Mn+Ofn+Ocn−Hys>Th2", "either side fails", "Protected capacity→L900 fallback", "[CM] Tables 5-18/5-19 pp.129–132"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "2", "Major Highlighted Points")
    r = bullets(
        ws,
        r,
        cols,
        [
            "Necessary coverage HO has higher measurement/preemption priority than unnecessary optimization HO. Throughput balancing must never block indoor rescue. [CM] Table 4-5 pp.38–39",
            "Blind HO/redirection has higher access-failure risk and should be used only when immediate mobility is required and neighbor contains the serving cell. [CM] Fig 4-2 pp.31–32; Table 5-22 p.149",
            "Huawei recommends RSRP as the general trigger quantity; RSRQ fluctuates with load. Do not substitute RSRQ for a real MLB load metric. [CM] Table 4-15 pp.55–56",
            "Frequency-priority HO is designed to put service on high bands and keep low band for continuous coverage. [CM] Fig 11-1/11-2 pp.299–300",
            "MlbBasedFreqPriHoSwitch blocks FreqPri HO when specified MLB functions are active, so MLB owns heavy-load decisions. [CM] pp.301–302; Table 11-7 p.313",
            "A frequency that is a FreqPri target must not have a reverse MLB-target relationship back to the source — ping-pong warning. [CM] p.303",
            "Equal-priority frequencies may be selected randomly for measurement objects — four L2600 carriers with the same static priority will NOT load-balance. [CM] §5.3.1.2 p.125",
            "InterFreqHoA4TimeToTrig=5120 ms disables frequency-priority, CQI and service-based inter-freq HO. [CM] Table 4-9 p.48",
            "A4 target threshold must be higher (better) than the relevant coverage A2, else ping-pong. [CM] Table 5-22 pp.147–148; Table 11-8 pp.315–316",
            "MML examples such as A1/A2 −85/−87 dBm and A4 −103 dBm are COMMAND EXAMPLES, not Robi design values. [CM] §11.4.1.2 pp.317–318",
            "A 7-carrier site may need 6 inter-freq objects. If MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum is smaller, some carriers are never measured. [CM] Tables 4-3/4-4",
        ],
        IVORY,
        20,
    )
    r += 1

    r = section(ws, r, cols, "3", "Benefit and Limitations")
    r = table(
        ws,
        r,
        ["Type", "Item", "Detail", "Robi impact"],
        [
            ["Benefit", "Coverage protection", "A2→A3/A4/A5 saves indoor / edge users", "Hard safeguard vs throughput-only optimization"],
            ["Benefit", "A4 for offload", "Moves UE to a 'good enough' low-load carrier without requiring it to beat serving", "Main event for L1800/L2100 ↔ L2600 MLB"],
            ["Benefit", "A5 for L900", "Requires BOTH serving poor AND L900 good", "Best L900-protection semantics"],
            ["Benefit", "FreqPri + A1", "Can release L900 users to high band when overlap is strong", "Stops L900 becoming a capacity layer"],
            ["Limitation", "Not a load algorithm", "Equal priority ≠ instantaneous load share", "Need Feature 3 MLB"],
            ["Limitation", "Measurement gaps", "Inter-freq meas steals DL TTIs on older UEs / VoLTE", "Track TP before/after broad meas; prefer A1/A2 gated meas"],
            ["Limitation", "Admission rules", "Unnecessary/offload HO must admit ALL QCIs; necessary HO needs any QCI", "Prep fail is often admission, not RF threshold"],
            ["Limitation", "CA/PCC", "PCC move ≠ SCC traffic move; CA details are in the CA book", "Check PCell vs SCell before blaming mobility"],
            ["Limitation", "High-speed / NLOS", "FREQ_PRI_HO_FORBID_SW ~30 km/h; NLOS can misclassify", "Do not use FreqPri to chase highway users"],
        ],
        col_fills={1: REF},
    )
    r += 1

    r = section(ws, r, cols, "4", "Selection Criteria / Trigger Condition")
    r = table(
        ws,
        r,
        ["Path", "Trigger", "Target selection", "UE / cell filters", "Robi rule"],
        [
            ["Coverage", "A2 (family depends on A3 vs A4/A5 vs IRAT vs blind)", "A3 relative, A4 absolute, or A5 dual; CovBasedInterFreqHoMode = IMMEDIATE / SIGNAL / FREQPRIORITY", "NRT, meas flags, object limit, SMeasure", "Must preempt MLB"],
            ["Frequency priority", "A1-based or configured FreqPri when serving is good enough; A4 qualifies target", "Highest-priority freq that passes A4; can wait FreqPriIFHoWaitingTimer", "LoadTriggerFreqPriHoSwitch: overlap, load obtainable, neighbor not in UE-number MLB trigger, no PCI conflict", "L900→capacity only; never FreqPri toward L900 as capacity"],
            ["MLB (executed here, decided in Feature 3)", "Source load trigger in CELLMLB", "A4 or A5 per MlbInterFreqHoEventType; ONLY_STRONGEST_CELL recommended", "UL-sync, not emergency, QCI/SPID, not high-mobility if forbid SW ON, protect timers", "Capacity pool only"],
            ["Blind", "Immediate mobility required + BlindHoPriority", "Configured priority, no candidate meas", "Neighbor coverage must contain serving", "OFF for daily throughput work"],
            ["Stop meas", "A1 serving recovered", "Stop coverage meas", "ReduceInvalidA1A2RptSigSwitch can cut signalling", "Keep A1/A2 hysteresis consistent"],
        ],
        col_fills={5: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "5", "Activation Parameter / Switch  —  Core settings")
    r = table(
        ws,
        r,
        ["SN", "MO", "Parameter / Switch", "Description", "Proposed value", "Conditional parameter", "Notes", "Doc"],
        [
            ["5.1", "InterFreqHoGroup", "A1/A2 hyst + TTT", "Stability of coverage meas start/stop", "Longer TTT than coverage-emergency; do not copy example −85/−87 as design", "Filter coefficient", "Separate A2 families exist for A3 vs A4/A5 vs blind", "[CM] Tables 5-3, 5-10, 4-9"],
            ["5.2", "InterFreqHoGroup", "InterFreqHoA3Offset + A3 Hyst/TTT", "Relative HO among similar layers", "Small offset; use for L2600↔L2600 or L1800↔L2100 if overlap similar", "QoffsetFreqConn, CIO", "Large CIO hides RF problems", "[CM] Table 5-16"],
            ["5.3", "InterFreqHoGroup", "InterFreqLoadBasedHoA4ThdRsrp + IfMlbThdRsrpOffset + FreqPriHoA4ThldRsrpOffset", "Absolute target quality for MLB/FreqPri", "Calibrate so target is usable AND A4 > coverage A2 (better). Not −103 by default.", "A4 Hyst/TTT; license if A5 MLB", "Main capacity-layer MLB event", "[CM] Table 11-5; [MLB] pp.137–139"],
            ["5.4", "InterFreqHoGroup", "A5 Thd1/Thd2 (coverage) / Mlb A5 Thd1", "Dual condition HO", "Use for capacity→L900 coverage fallback", "Intra-LTE Load Balancing for Non-cosited Cells license if MLB A5", "Strongest L900 protection", "[CM] Tables 5-18/5-19; [MLB] Table 6-3"],
            ["5.5", "CellAlgoSwitch", "CovBasedInterFreqHoMode", "When to execute coverage HO after report", "Keep current unless waiting-timer design is written", "CovBasedIfHoWaitingTimer", "IMMEDIATE vs SIGNAL vs FREQPRIORITY", "[CM] §5.3.1.1–5.3.1.3"],
            ["5.6", "ENodeBAlgoSwitch / related", "MlbBasedFreqPriHoSwitch", "FreqPri yields to MLB in heavy load", "ON wherever connected MLB is ON", "MLB feature active", "Stops FreqPri fighting MLB", "[CM] Table 11-7 p.313"],
            ["5.7", "related FreqPri", "LoadTriggerFreqPriHoSwitch + ReduceInvalidFreqPriHoSwitch", "FreqPri only if neighbor not already MLB-heavy", "ON after MLB live", "Overlap, load info, no PCI conflict", "Reduces invalid FreqPri", "[CM] pp.302–303"],
            ["5.8", "CellOpHoCfg / IntraRatHoComm", "MLB_HO_FORBID_SW / FREQ_PRI_HO_FORBID_SW", "No MLB/FreqPri meas for UE > ~30 km/h", "ON in urban high-mobility cells if HO fail high", "NLOS can misclassify stationary UE", "Do not use as indoor tool", "[CM] pp.304–305; [MLB] pp.136–141"],
            ["5.9", "EutranInterNFreq", "FREQ_MEAS_FLAG / HO_TRG_FREQ_FORBID_MEAS_FLAG", "Allow or forbid measuring a frequency as HO target", "Capacity: measure; do not forbid L900 measurement (coverage)", "Object-number limits", "Audit before touching thresholds", "[CM] §4.1.4.1.2"],
            ["5.10", "InterFreqHoGroup", "InterFreqHoA4TimeToTrig", "A4 TTT; 5120 ms disables FreqPri/CQI/service IFHO", "Do NOT set 5120 ms if FreqPri/MLB A4 is required", "A4 hyst", "Common silent killer of FreqPri", "[CM] Table 4-9 p.48"],
        ],
        col_fills={5: GREEN, 6: AMBER},
    )
    r += 1

    r = section(ws, r, cols, "6", "Prerequisite Functions")
    r = table(
        ws,
        r,
        ["SN", "Prerequisite", "Check", "Proposed", "If missing"],
        [
            ["6.1", "Symmetric NRT + no PCI conflict", "EutranInterFreqNCell both directions", "All capacity pairs + L900 coverage relations", "Looks like bad A4"],
            ["6.2", "Measurement object capacity", "MaxNonIntraMeasObjNum, MaxEutranFddMeasFreqNum ≥ number of needed inter-freq", "≥6 from a 7-layer site", "Random missing L2600 carrier"],
            ["6.3", "SMeasure not silently suppressing meas", "HoMeasComm.SMeasure vs intended A4/FreqPri", "Verify actual RRC config", "A4 never delivered"],
            ["6.4", "Gap pattern acceptable", "AutoGapSwitch, GapPatternType", "Prefer A1/A2 gated meas", "TP drop on VoLTE/old UE"],
            ["6.5", "Admission / X2", "HoAdmitSwitch, X2RoHoAdmitSwitch, transport state", "Healthy", "PrepAtt high, ExecAtt low"],
            ["6.6", "CA/PCC policy known", "CA book + PCC anchoring", "Do not FreqPri against PCC anchor", "Ping-pong [CM] §11.3.2.3"],
            ["6.7", "Penalty timers understood", "ResHoPreFailPunishTimer, OptHoPreFailPunishTimer, NonResHoPreFailPunishTimes", "Do not answer all fail types by lowering A4", "Wrong RCA"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "7", "Mutually Impacted")
    r = table(
        ws,
        r,
        ["Control A", "Control B", "Risk", "Coordination"],
        [
            ["FreqPri target A→B", "MLB target B→A", "Documented ping-pong", "No reverse MLB target on a FreqPri pair [CM] p.303"],
            ["FreqPri HO", "Connected MLB", "Both move the same UE", "MlbBasedFreqPriHoSwitch ON"],
            ["Incoming unnecessary HO", "Immediate FreqPri re-meas", "Bounce back", "FreqPriInHoProtectionTimer [CM] p.300"],
            ["A4 too close to A2", "Coverage meas start", "Ping-pong", "Keep A4 better than coverage A2"],
            ["Large CIO / QoffsetFreqConn", "RF problem", "Masks overshoot, edge TP fall", "Fix RF first"],
            ["Virtual-grid smart carrier selection", "FreqPri", "Ping-pong [CM] p.312", "Do not enable both casually"],
            ["NSA / CA PCC anchoring", "FreqPri / MLB", "PCC ping-pong", "Read CA + NSA notes before CR"],
            ["LOAD_COVERAGE_MEAS_DECOUPLE_SW", "Already-delivered coverage meas", "Allows load meas after coverage meas [MLB] p.139", "Enable if coverage meas blocks MLB A4"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "8", "Relation with Other Features")
    r = table(
        ws,
        r,
        ["Feature", "Relation", "Robi rule"],
        [
            ["Idle Mode (F1)", "Idle sets access layer; connected discards dedicated idle priority", "Align ThreshXhigh with A1-escape; ThrshServLow with A5-return"],
            ["Intra-RAT MLB (F3)", "MLB uses A4/A5, HO mode, UE filters, protect timers in this feature", "This feature = measurement/HO engine; MLB = who/when"],
            ["Carrier Aggregation", "PCC anchoring documented outside this book", "Judge aggregate CA throughput, not only PCell TP"],
            ["VoLTE / QCI", "QCI-specific hyst/TTT and mobility-target-ind", "Never offload QCI1 if target cannot admit all QCIs"],
            ["Inter-RAT MLB / IRAT HO", "Separate A2 family and IRAT documents", "Out of this 4G capacity-layer scope"],
            ["Energy saving", "Sleeping cell must not stay a FreqPri/MLB target", "Coordinate ES"],
        ],
        col_fills={3: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "9", "License Requirements")
    r = table(
        ws,
        r,
        ["Function", "License (verify on NE)", "Note"],
        [
            ["Basic coverage intra/inter-freq HO A1–A5", "Basic LTE mobility — normally included", "Still confirm if operator uses lean license packages"],
            ["Frequency-priority HO", "Frequency Priority Based Handover / Service-based mobility package — verify exact name", "If missing, A1-escape from L900 will not run"],
            ["MLB A4 on co-sited cells", "Intra-RAT Mobility Load Balancing", "See Feature 3"],
            ["MLB event A5", "Intra-LTE Load Balancing for Non-cosited Cells [MLB] pp.142–143", "Needed only if MlbInterFreqHoEventType=A5"],
            ["Blind HO", "Blind handover option if sold separately", "Keep OFF for this project unless overlap proven"],
        ],
        col_fills={2: YELLOW},
    )
    r += 1

    r = section(ws, r, cols, "10", "All Parameter List  —  Feature 2  (sequence)")
    conn_params = [
        ["1", "EutranInterNFreq", "DlEarfcn / MeasBandWidth", "Inter-freq measurement object", "All 6 non-serving layers configured", "NRT exists", "Object limit", "Meas object", "[CM] Table 4-2"],
        ["2", "EutranInterNFreq", "FREQ_MEAS_FLAG", "Allow frequency in meas", "Selected on all needed layers", "HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for coverage/MLB targets", "Audit first", "Meas filter", "[CM] §4.1.4.1.2"],
        ["3", "EutranInterNFreq", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Forbid as HO target meas", "Deselected for L900 coverage and capacity MLB", "SCC-only exception", "Silent no-HO", "Meas filter", "[CM] §4.1.4.1.2"],
        ["4", "CellUeMeasControlCfg", "MaxNonIntraMeasObjNum / MaxEutranFddMeasFreqNum", "How many inter-freq objects can be delivered", "≥ needed pairs (7-layer ⇒ 6)", "Equal-priority random pick if over limit", "Four L2600 at same prio can be unevenly exposed", "Object cap", "[CM] Tables 4-3/4-4"],
        ["5", "HoMeasComm", "SMeasure", "Skip inter-freq meas while serving RSRP above this", "Verify vs intended FreqPri/MLB", "A1/A2", "Can silently suppress A4", "Meas gate", "[CM] §4.1.5 p.55"],
        ["6", "CellHoParaCfg", "EutranFilterCoeffRsrp / Rsrq", "L3 filter", "Keep current; do not loosen L900 rescue too much", "TTT", "Over-smooth delays indoor rescue", "Filter", "[CM] Table 4-14"],
        ["7", "ENodeBAlgoSwitch", "AutoGapSwitch / GapPatternType / DedicatedGapPatternType", "Measurement gap", "Prefer A1/A2 gated meas", "VoLTE / old UE TP", "Gap steals TTI", "Gap", "[CM] Fig 4-10"],
        ["8", "InterFreqHoGroup", "InterFreqHoA1A2Hyst / TimeToTrig", "A1/A2 stability", "Calibrate; not −85/−87 blindly", "A2 family per HO type", "QCI-specific optional", "A1/A2", "[CM] Table 4-9"],
        ["9", "InterFreqHoGroup", "Coverage A2 thresholds (A3-based / A4A5-based / blind families)", "Start coverage meas", "L2600 A2 early enough for indoor; L900 A2 not hyper-aggressive", "A1 stop", "Wrong family = wrong behavior", "A2", "[CM] Tables 5-3, 5-10"],
        ["10", "InterFreqHoGroup", "InterFreqHoA3Offset / A3RsrqOffset / A3 Hyst / TTT", "Relative HO", "Small; similar-coverage capacity only", "QoffsetFreqConn, CIO", "Not for L900 capacity offload", "A3", "[CM] Table 5-16"],
        ["11", "EutranInterNFreq", "QoffsetFreqConn", "Connected frequency offset (Ofn)", "0 unless stable RF bias", "A3/A4/A5", "Positive on target makes HO harder", "Offset", "[CM] pp.46–48"],
        ["12", "EutranInterFreqNCell", "CellIndividualOffset", "Connected CIO (Ocn)", "0 globally toward L900", "NRT", "Positive CIO to L900 fills 5 MHz", "CIO", "[CM] pp.46–48"],
        ["13", "InterFreqHoGroup", "InterFreqLoadBasedHoA4ThdRsrp / Rsrq", "MLB/FreqPri A4 absolute target", "Calibrate; A4 better than coverage A2", "IfMlbThdRsrpOffset, FreqPriHoA4ThldRsrpOffset", "Main capacity MLB event", "A4", "[CM] Table 11-5; [MLB] p.137"],
        ["14", "EutranInterNFreq", "IfMlbThdRsrpOffset / FreqPriHoA4ThldRsrpOffset", "Per-frequency A4 offset", "Use to differentiate L2100 15 MHz vs 20 MHz if needed", "A4 base thd", "Prefer SE-based MLB over harsh A4 bias", "A4 offset", "[MLB] p.137"],
        ["15", "InterFreqHoGroup", "InterFreqHoA4Hyst / InterFreqHoA4TimeToTrig", "A4 stability", "Longer than coverage TTT; NEVER 5120 ms if FreqPri needed", "5120 ms disables FreqPri/CQI/service IFHO", "Silent disable risk", "A4", "[CM] Table 4-9 p.48"],
        ["16", "InterFreqHoGroup", "Coverage A5 Thd1/Thd2 + Hyst/TTT", "Serving poor AND target good", "Capacity→L900 fallback", "L900 suitability", "Best L900 semantics", "A5", "[CM] Tables 5-18/5-19"],
        ["17", "InterFreqHoGroup", "MlbInterFreqHoA5Thd1Rsrp/Rsrq", "MLB A5 serving-poor half", "Only if event type = A5", "Non-cosited MLB license", "Co-sited capacity: prefer A4", "MLB A5", "[MLB] Table 6-3"],
        ["18", "EutranInterNFreq", "MlbInterFreqHoEventType", "A4 or A5 for FDD MLB HO", "A4 on co-sited capacity", "A5 needs extra license", "A5 if source must be poor first", "MLB event", "[MLB] §6.1.1.5.2"],
        ["19", "CellAlgoSwitch", "CovBasedInterFreqHoMode", "Immediate / signal-strength / freq-priority wait", "Keep unless designed", "CovBasedIfHoWaitingTimer", "Fig 5-3–5-7", "Coverage exec", "[CM] §5.3.1"],
        ["20", "IntraRatHoComm", "CovBasedIfHoWaitingTimer / FreqPriIFHoWaitingTimer / FreqPriInHoProtectionTimer", "Wait / protect", "Keep incoming-protect ON with FreqPri", "FreqPri vs MLB", "Stops bounce", "Timers", "[CM] pp.300, 308–309"],
        ["21", "ENodeB / FreqPri switches", "MlbBasedFreqPriHoSwitch", "MLB owns heavy load", "ON if MLB ON", "MLB license", "Required coordination", "FreqPri vs MLB", "[CM] Table 11-7"],
        ["22", "FreqPri", "LoadTriggerFreqPriHoSwitch / ReduceInvalidFreqPriHoSwitch", "FreqPri only to healthy low-load overlap neighbor", "ON after MLB", "No PCI conflict; neighbor not UE-number MLB triggered", "Invalid FreqPri reduction", "FreqPri", "[CM] pp.302–303"],
        ["23", "CellOpHoCfg", "HighMobiUeHoForbidSw = MLB_HO_FORBID_SW", "No MLB meas if UE >~30 km/h", "ON where high-speed HO fail", "NLOS misclassify", "Urban highway optional", "Safeguard", "[MLB] pp.136–141"],
        ["24", "related", "FREQ_PRI_HO_FORBID_SW", "No FreqPri if high mobility", "Same policy as 23", "NLOS", "Indoor: leave OFF unless proven", "Safeguard", "[CM] pp.304–305"],
        ["25", "HoMeasComm", "Res/Opt/NonRes HO fail punish timers & counts", "Retry vs penalty", "Do not lower A4 after admission fail", "Separate prep vs exec vs RLF", "RCA first", "Penalty", "[CM] Tables 4-16/4-17"],
        ["26", "EutranInterFreqNCell", "BlindHoPriority / InterFreqMlbBlindHo", "Blind path", "OFF / no priority unless full containment", "Neighbor contains serving", "Higher access fail", "Blind", "[CM] Table 5-22 p.149"],
        ["27", "CellAlgoSwitch", "LOAD_COVERAGE_MEAS_DECOUPLE_SW", "Allow load meas after coverage meas already delivered", "ON if coverage meas blocks MLB", "MLB A4", "Documented in MLB book", "Decouple", "[MLB] p.139"],
        ["28", "RatFreqPriorityGroup", "CovIFHo RSRP/RSRQ Hyst/TTT", "Per-freq/QCI coverage values", "Use for VoLTE if needed", "QCI", "Do not globally copy to data QCI", "QCI", "[CM] Table 4-9"],
    ]
    r = table(
        ws,
        r,
        ["Seq", "MO", "Parameter", "Description", "Proposed value", "Conditional / other feature", "Notes", "Group", "Doc ref"],
        conn_params,
        col_fills={5: GREEN},
        min_h=26,
    )
    r += 1

    r = section(ws, r, cols, "11", "Final MML Command for Activations  —  maintain sequence")
    r = note(
        ws,
        r,
        cols,
        "  SEQUENCE: LST/audit NRT+meas flags+object limits → fix flags → set A1/A2/A5 coverage safety (L900) → set A4 MLB/FreqPri quality gate → FreqPri vs MLB coordination switches → forbid reverse MLB target → verification counters. "
        "Do not put example −85/−87/−103 dBm into live MML without MR calibration.",
        AMBER,
        36,
    )
    r = table(
        ws,
        r,
        ["Seq", "MO", "Activation value / MML skeleton", "Conditional parameter", "Remarks", "Parameter description", "More notes"],
        [
            ["0", "—", "LST INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>;  LST EUTRANINTERNFREQ: LocalCellId=<x>;  LST EUTRANINTERFREQNCELL: LocalCellId=<x>;  LST CELLHOPARACFG: LocalCellId=<x>;  LST HOMEASCOMM:;", "Read-only", "Dump first", "Baseline", "Attach to CR"],
            ["1", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<earfcn>,  (ensure FREQ_MEAS_FLAG selected; HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for L900 and all capacity freqs);", "NRT exists; object cap sufficient", "Audit before any threshold CR", "Meas eligibility", "Most 'A4 not working' cases are flags/NRT"],
            ["2", "CELLUEMEASCONTROLCFG", "MOD CELLUEMEASCONTROLCFG: LocalCellId=<x>, MaxNonIntraMeasObjNum=<≥6>, MaxEutranFddMeasFreqNum=<≥6>;", "7-layer site", "Otherwise some L2600 never measured", "Object cap", "Equal-prio random exposure"],
            ["3", "INTERFREQHOGROUP", "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>,  (set coverage A2 / A1 / A5 Thd for L2600→L900 rescue — calibrate from MR, do not paste example values);", "A2 family matches A5 coverage design", "Protect indoor first", "Coverage safety", "Yellow — L900"],
            ["4", "INTERFREQHOGROUP", "MOD INTERFREQHOGROUP: LocalCellId=<x>, InterFreqHoGroupId=<g>, InterFreqLoadBasedHoA4ThdRsrp=<calib>, InterFreqHoA4Hyst=<2..4>, InterFreqHoA4TimeToTrig=MS320;", "A4 better than coverage A2; TTT ≠ 5120 ms", "Capacity-layer MLB/FreqPri gate", "A4 absolute target", "MS320 is a starting TTT, not a mandate"],
            ["5", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<capa>, MlbInterFreqHoEventType=A4, IfMlbThdRsrpOffset=0;", "Co-sited overlap; A5 license not required", "A4 for capacity MLB", "MLB event type", "Use A5 only for non-cosited / extra license"],
            ["6", "EUTRANINTERFREQNCELL", "MOD EUTRANINTERFREQNCELL: LocalCellId=<x>, DlEarfcn=<L900>, CellId=<l900>, CellIndividualOffset=0;", "Do not use +CIO to 'help' HO success into L900", "Avoid filling 5 MHz", "CIO", "Yellow"],
            ["7", "ENODEBALGOSWITCH / FreqPri MO", "Enable MlbBasedFreqPriHoSwitch-1 and LoadTriggerFreqPriHoSwitch-1; set FreqPriInHoProtectionTimer to current-or-default non-zero;", "Connected MLB will be ON (Feature 3)", "FreqPri yields in heavy load; no immediate bounce", "Coordination switches", "Confirm exact MO in MAE"],
            ["8", "CELLOPHOCFG", "MOD CELLOPHOCFG: LocalCellId=<x>, HighMobiUeHoForbidSw=MLB_HO_FORBID_SW-1;   (only on proven high-speed cells)", "NLOS false-positive risk", "Optional", "High-mobility MLB forbid", "Not default for indoor cluster"],
            ["9", "Verification", "Monitor L.HHO.InterFreq.Coverage.*, L.HHO.InterFreq.FreqPri.*, L.RRC.ReEst.ReconfFail.Att, ping-pong, drop, VoLTE. PrepSucc=ExecAtt/PrepAtt; ExecSucc=ExecSucc/ExecAtt.", "15-min + busy hour", "Judge pair-level, not only source cell", "Counters [CM] Table 5-24, 11-9/11-10", "Close-loop"],
        ],
        col_fills={3: GREEN, 4: AMBER},
        min_h=34,
    )
    return ws


# ---------------------------------------------------------------------------
# FEATURE 3 MLB
# ---------------------------------------------------------------------------
def build_mlb(wb):
    ws = wb.create_sheet("03_IntraRAT_MLB")
    cols = 10
    set_widths(ws, [8, 28, 24, 18, 18, 20, 24, 26, 20, 36])
    apply_sheet_setup(ws, "Feature 3 — Intra-RAT MLB")

    r = 1
    r = banner(ws, r, cols, "  FEATURE 3   ·   Intra-RAT Mobility Load Balancing", HUAWEI_RED, font_title, 32)
    r = banner(ws, r, cols, "  Source: eRAN Intra-RAT Mobility Load Balancing, eRAN21.1, Issue 10 (2026-06-30)   ·   Execution events live in Connected Mode book", NAVY, font_white_s, 20)
    r = add_legend_row(ws, r, cols)
    r += 1

    r = section(ws, r, cols, "1", "Working Principle")
    r = para(
        ws,
        r,
        cols,
        "Intra-RAT MLB coordinates load among overlapping inter-frequency (and inter-duplex) LTE cells and transfers load by connected handover or idle reselection (dedicated priority). "
        "Equalization uses peer load information and reduces source–target load difference. Offload protects a cell even if target load is incomplete. "
        "User-number model: Load = N / C ; normalized difference = (Load_s − Load_t) / Load_s. "
        "Huawei recommends ActiveUeBasedLoadEvalSw when bandwidths differ, plus SpectralEffBasedLoadEvalSw when SE differs (e.g. >30%). "
        "Ref: [MLB] §3 Fig 3-1 p.15; §4.1–4.3 Figs 4-1–4-5 pp.16–19; §5.1.1 pp.24–27; Table 5-5 pp.47–48.",
        WHITE,
        64,
    )
    r = note(ws, r, cols, "  CHART / MLB LOOP:", LIGHT, 16)
    r = flow_row(ws, r, ["Eval load N/C", "Trigger if N≥Thd+Offs", "Admit target", "Select UEs", "HO or idle release"], ["1B4F72", "B9770E", "117A65", "6C3483", HUAWEI_RED], cols)
    r = table(
        ws,
        r,
        ["Mode", "What is equalized", "Transfer", "CA UEs", "Robi role"],
        [
            ["User-number connected (UE_NUMBER_ONLY + SynchronizedUE)", "Active/UL-sync users vs cell capability C", "Measurement-based HO (A4/A5) or blind HO", "Only if CaUserLoadTransferSw + conditions", "PRIMARY algorithm for DL TP fairness"],
            ["User-number idle (IdleUE)", "Idle users / future sessions", "RRC release + dedicated priority + T320", "Depends on PCC anchoring + CaUserLoadTransferSw", "SECONDARY — next RRC layer"],
            ["PRB_USAGE connected (PRB_ONLY + PrbMlbSynchronizedUE)", "PRB usage or unused-PRB (BW-aware if LoadTransferEnhSw off)", "HO of heavy users", "NOT transferred", "SUPPLEMENT only; weak if CA > ~60%"],
            ["PRB evaluation (GBR + min non-GBR ± control)", "QoS/GBR pressure", "HO", "See live help", "Only if GBR/VoLTE saturation is the problem"],
            ["Offload / blind", "Protect source vs threshold without full target load", "Blind HO/redirect", "Risky", "Not for co-sited Huawei capacity pool"],
        ],
        col_fills={5: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "2", "Major Highlighted Points")
    r = bullets(
        ws,
        r,
        cols,
        [
            "Do not treat 15 MHz L2100 or 5 MHz L900 as equal to 20 MHz in raw UE-count MLB. Raw UE balancing can reduce DL throughput. [MLB] §6.1.2.2 p.141",
            "ActiveUeBasedLoadEvalSw uses UEs with buffered DL data (non-zero TTI samples). SpectralEffBasedLoadEvalSw refreshes SE every minute when ≥10 UL-sync UEs. [MLB] §5.1.1 pp.25–27",
            "Trigger (connected user-number): N ≥ InterFreqMlbUeNumThd + MlbUeNumOffset for whole MlbTrigJudgePeriod; stop when N < InterFreqMlbUeNumThd. [MLB] §6.1.1.1 p.128",
            "Target needs: NoHoFlag permit, not blacklisted, no PCI conflict, HO success ≥ NCellHoSuccRateThld, FREQ_MEAS_FLAG, not HO_TRG_FREQ_FORBID, valid OverlapInd, MlbTargetInd, Low/Medium HW+transport, not in penalty. [MLB] pp.27–29, 129",
            "MlbTargetInd ALLOWED_WITHOUT_IDLE_MLB / ALLOWED_WITHOUT_CONNECT_MLB can block one mode. Use this to protect L900. [MLB] pp.28, 129",
            "MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL is Huawei-recommended. Else MLB HO to non-strongest cell is immediately coverage-HO'd back. [MLB] Table 6-5 p.159",
            "InterFreqLoadEvalPrd=5 s AND MlbMaxUeNum≥40 can over-transfer. [MLB] p.136",
            "PRB MLB: burst-sensitive, no UE-experience fairness, no CA UE transfer, gain collapses when CA penetration ≳60%. [MLB] Fig 6-2, §6.5 pp.207–215",
            "NCellTrigThldSmartOptAlgoSw learns 7 days then refreshes every 7 days. Do not overwrite daily. First week aggressive seeds: extra CPU/HO and up to 5% TP fluctuation. [MLB] pp.29–31, 89–91, 162, 224–225",
            "Fixed-proportion idle + user-number connected MLB = ping-pong. Adaptive-proportion idle = not recommended. [MLB] §5.4.2.2, §5.5.2.1",
            "Hardware/transport load states: Low / Medium / High / OverLoad. High/OverLoad target is not a valid equalisation target. [MLB] Fig 4-3 pp.17–18",
        ],
        IVORY,
        20,
    )
    r += 1

    r = section(ws, r, cols, "3", "Benefit and Limitations")
    r = table(
        ws,
        r,
        ["Type", "Item", "Detail", "Robi impact"],
        [
            ["Benefit", "True load equalisation", "Uses peer load, not only RF", "Right tool for 10 Mbps TP gap if cause is load"],
            ["Benefit", "BW/SE aware when switches ON", "C includes PRB, scale factor, GBR, SE", "Essential for 15 vs 20 MHz"],
            ["Benefit", "Idle + connected pair", "Idle reduces future HO; connected fixes now", "Use both, not only A3 CIO"],
            ["Benefit", "CA transfer option", "CaUserLoadTransferSw can move PCC", "Needed on high-CA L2600"],
            ["Limitation", "Cannot fix RF / CA scheduler / transport", "Low TP + low PRB is not an MLB case", "RCA first"],
            ["Limitation", "Coverage mismatch / UE band / PLMN", "Gain disappears", "[MLB] pp.39, 141, 214"],
            ["Limitation", "Blind offload", "Can overload a 'low-load' looking cell", "[MLB] pp.167, 231"],
            ["Limitation", "Energy-saving interaction", "MLB and carrier shutdown change each other's targets", "[MLB] pp.153–156, 219–222"],
            ["Limitation", "L900 erosion", "Not quantified in Huawei book — engineering risk", "Hard guardrail"],
        ],
        col_fills={1: REF},
    )
    r += 1

    r = section(ws, r, cols, "4", "Selection Criteria / Trigger Condition")
    r = table(
        ws,
        r,
        ["Stage", "Condition", "Parameters", "Robi rule"],
        [
            ["Source trigger (UE-number connected)", "N ≥ InterFreqMlbUeNumThd + MlbUeNumOffset for MlbTrigJudgePeriod", "InterFreqMlbUeNumThd, MlbUeNumOffset, MlbTrigJudgePeriod, ActiveUeBasedLoadEvalSw", "Use ACTIVE UEs, not raw UL-sync, on mixed BW"],
            ["Source stop", "N < InterFreqMlbUeNumThd", "Hysteresis via offset", "Avoid chatter"],
            ["Source trigger (PRB)", "PRB ≥ InterFreqMlbThd + LoadOffset AND min UE condition", "InterFreqMlbThd, InterFreqMlbUlThd, MlbMinUeNumThd, per-freq PRB offsets", "Supplement only"],
            ["Target admit (equalisation)", "Low/Med HW+transport, load difference, not punished, HO SR ≥ thd, MlbTargetInd, overlap, PCI OK", "NCellHoSuccRateThld, CellPunishPrdNum, LoadDiffThd", "L900 not admitted as routine target"],
            ["Target reject penalty", "Target rejects 'Reduce Load in Serving Cell' for no radio resource → punish CellPunishPrdNum × InterFreqLoadEvalPrd", "CellPunishPrdNum", "Do not immediately retarget L900"],
            ["UE filter", "UL-sync, not emergency, SPID/QCI, eMBMS, protect timer, failed-transfer punish, optional ARP/PRB/MCS/QCI/SNR", "CellMlbUeSel.*, MlbHoInProtectTimer, MlbUeSelectPunishTimer", "Do not pick indoor-edge UEs just because they eat PRB"],
            ["UE count to move", "min(needed delta, source hyst, MlbMaxUeNum); optional MLB_UE_SEL_OPT_SW", "MlbMaxUeNum, LoadTransferEnhSw", "Keep MlbMaxUeNum conservative; never ≥40 with 5 s eval"],
            ["Freq pick", "FAIRSTRATEGY / PRIORITYBASED / LOADPRIORITY", "FreqSelectStrategy, MlbFreqPriority", "LOADPRIORITY among 4×L2600 if X2 load OK"],
            ["HO event", "A4 (co-sited) or A5 (license)", "MlbInterFreqHoEventType, A4 thd + offset", "A4 for Robi co-sited capacity"],
            ["Idle MLB", "Analogous idle UE thd; dedicated priority class: NG-RAN > E-UTRAN low-load > E-UTRAN high-load > UTRAN > GERAN", "InterFreqIdleMlbUeNumThd, T320", "See Feature 1"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "5", "Activation Parameter / Switch  —  Core settings")
    r = table(
        ws,
        r,
        ["SN", "MO", "Parameter / Switch", "Description", "Proposed value", "Conditional parameter", "Notes", "Doc"],
        [
            ["5.1", "CellAlgoSwitch.MlbAlgoSwitch", "InterFreqMlbSwitch", "Master intra-freq-pair MLB", "ON on capacity cells", "License; NRT; MlbTargetInd", "Core bit [MLB] Table 6-2 p.127", "[MLB] Table 6-2"],
            ["5.2", "CellAlgoSwitch.MlbAlgoSwitch", "InterFreqIdleMlbSwitch", "Idle transfer", "ON after target policy set", "T320, dedicated prio", "Not with fixed-proportion", "[IM]/[MLB]"],
            ["5.3", "CellAlgoSwitch.MlbAlgoSwitch", "InterFreqBlindMlbSwitch", "Blind MLB", "OFF", "Full containment + BlindHoPriority", "Amber — do not use for TP campaign", "[MLB]"],
            ["5.4", "CellMLB", "MlbTriggerMode", "UE_NUMBER_ONLY / PRB_ONLY / PRB_OR_UE_NUMBER", "UE_NUMBER_ONLY", "PRB_ONLY excludes CA UEs", "Primary = user number", "[MLB] Table 6-2"],
            ["5.5", "CellMLB", "InterFreqUeTrsfType", "SynchronizedUE / IdleUE / PrbMlbSynchronizedUE", "SynchronizedUE (+ IdleUE if idle MLB ON — follow live MO help if separate)", "Trigger mode match", "Must match algorithm", "[MLB] Table 6-2"],
            ["5.6", "related eval SW", "ActiveUeBasedLoadEvalSw", "Load uses active DL-buffer UEs", "ON", "Different BW layers", "Huawei recommended for unequal BW", "[MLB] Table 5-5"],
            ["5.7", "related eval SW", "SpectralEffBasedLoadEvalSw", "C uses PRB, scale, GBR, measured SE", "ON", "Needs ≥10 UL-sync UEs for SE refresh / min", "If SE differs >~30%", "[MLB] Table 5-5"],
            ["5.8", "related", "LoadTransferEnhSw", "Better multi-target / PRB-diff calculation", "ON", "Multi L2600 targets", "Recommended with 4 carriers", "[MLB] §6.1.1.5.1"],
            ["5.9", "related", "CaUserLoadTransferSw", "Allow CA UE / PCC transfer", "ON after PCC/SCC audit", "Target CA capability ≥ serving on one path", "Else CA UEs filtered (2CC/FDD+TDD notes p.157)", "[MLB] pp.129–136"],
            ["5.10", "CellMLB", "MlbHoCellSelectStrategy", "Which measured cell to use", "ONLY_STRONGEST_CELL", "A4 meas success", "Huawei recommended p.159", "[MLB] Table 6-5"],
            ["5.11", "CellMLB", "FreqSelectStrategy", "How to pick frequency among candidates", "LOADPRIORITY for 4×L2600", "Reliable load exchange", "PRIORITYBASED if CA/PCC architecture requires a preferred PCC", "[MLB] pp.137, 213"],
            ["5.12", "EutranInterNFreq", "MlbInterFreqHoEventType", "A4 or A5", "A4", "A5 license if used", "Co-sited capacity", "[MLB] Table 6-3"],
            ["5.13", "EutranInterNFreq", "MlbTargetInd", "May this frequency be MLB target", "Capacity ALLOWED; L900 WITHOUT connect+idle MLB", "OverlapInd, NoHoFlag", "Coverage HO to L900 stays", "[MLB] pp.28, 129"],
            ["5.14", "related smart", "NCellTrigThldSmartOptAlgoSw", "Neighbor-specific learned thds", "Optional ON; do not seed aggressively week 1", "≥1×15-min counter subscribed in MAE", "7-day learn / 7-day refresh", "[MLB] pp.29–31"],
        ],
        col_fills={5: GREEN, 6: AMBER},
    )
    r += 1

    r = section(ws, r, cols, "6", "Prerequisite Functions")
    r = table(
        ws,
        r,
        ["SN", "Prerequisite", "Proposed / required", "If missing"],
        [
            ["6.1", "Intra-RAT MLB license + (if A5) Non-cosited MLB license", "Present on every eNB in the cluster", "Switch ON but no transfer"],
            ["6.2", "X2 / intra-eNB load exchange on Huawei co-sited sectors", "Working; else only offload/blind", "Unsafe equalisation"],
            ["6.3", "Correct OverlapInd and PERMIT_HO on capacity pairs", "Set", "Target never admitted"],
            ["6.4", "NCellHoSuccRateThld pass on pair", "Fix coverage HO first if pair already unhealthy", "Target filtered"],
            ["6.5", "HW/transport not High/OverLoad on intended target", "Alarms clear", "Target rejected / punished"],
            ["6.6", "Feature 2 A4 meas actually delivered (flags, object cap, SMeasure, gap)", "See Feature 2 SN-6", "Load trigger with 0 meas success"],
            ["6.7", "Idle Feature 1 SIB5/NORMAL if idle MLB ON", "See Feature 1", "Dedicated prio cannot include the freq"],
            ["6.8", "MAE 15-min counter subscription if smart threshold ON", "At least one 15-min period", "[MLB] §§5.1.3.4, 6.5.3.4"],
            ["6.9", "CA combination / PCC policy known before CaUserLoadTransferSw", "PCell/SCell baseline", "Wrong PCC move"],
        ],
        col_fills={3: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "7", "Mutually Impacted")
    r = table(
        ws,
        r,
        ["MLB item", "Other item", "Risk", "Rule"],
        [
            ["User-number connected MLB", "Fixed-proportion idle MLB", "Ping-pong [MLB] §5.4.2.2", "Fixed-proportion OFF"],
            ["User-number connected", "PRB_ONLY as sole mode", "CA UEs stuck; fairness not guaranteed", "UE_NUMBER_ONLY primary"],
            ["PRB_USAGE vs PRB_VALUATION", "Same mode mutually exclusive [MLB] pp.216, 245", "Cannot both", "Pick one if supplement needed"],
            ["MLB", "FreqPri (Feature 2)", "Opposite moves", "MlbBasedFreqPriHoSwitch; no reverse target"],
            ["MLB", "Carrier shutdown / deep dormancy", "Target disappears / gain loss [MLB] pp.153–156", "ES coordination"],
            ["MLB", "Flexible CA", "CA avoids a cell under connected offload [MLB] pp.168, 233", "Expected — do not 'fix' by forcing CA onto overloaded cell"],
            ["Smart learned thd", "BW change / board SW / upgrade / CA-eval SW change", "Forced relearn or stale thd", "Freeze AI CRs during relearn week"],
            ["Aggressive MlbMaxUeNum + 5 s eval", "CPU / signalling / TP", "Overshoot", "Conservative max UE"],
            ["MLB toward L900", "Indoor VoLTE / 5 MHz saturation", "Coverage layer becomes capacity", "Forbidden"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "8", "Relation with Other Features")
    r = table(
        ws,
        r,
        ["Feature", "Relation", "Robi rule"],
        [
            ["Idle Mode (F1)", "Idle MLB is MLB's idle method", "Dedicated prio + T320; L900 not idle target"],
            ["Connected Mode (F2)", "HO execution, A4/A5, meas, admit, punish", "ONLY_STRONGEST_CELL + A4"],
            ["Carrier Aggregation", "PCC transfer vs SCell scheduling; PRB MLB skips CA UEs", "Read PCell+SCell before CR"],
            ["Energy saving", "Mutual target exclusion", "No MLB onto sleeping carrier"],
            ["Admission control / X2 HO admit", "Unnecessary HO must admit all QCIs [CM]", "Prep fail ≠ A4 too high"],
            ["Intelligent n-cell threshold SON", "Learns 7 days", "AI monitors, does not daily overwrite"],
            ["Inter-RAT MLB", "Out of this book", "Do not dump 4G overflow to 3G/2G for this TP KPI"],
        ],
        col_fills={3: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "9", "License Requirements")
    r = table(
        ws,
        r,
        ["Function", "License (verify in MAE / license file)", "If missing"],
        [
            ["InterFreqMlbSwitch user-number / PRB connected equalisation", "Intra-RAT Mobility Load Balancing", "No load HO counters"],
            ["Idle MLB / dedicated-priority load release", "Same Intra-RAT MLB family (confirm)", "No L.RRCRel.load.DedicatedPri.*"],
            ["MlbInterFreqHoEventType = A5", "Intra-LTE Load Balancing for Non-cosited Cells [MLB] pp.142–143", "Stay on A4"],
            ["InterFreqBlindMlbSwitch / blind HO", "Blind MLB option if sold separately", "Keep OFF"],
            ["CaUserLoadTransferSw / flexible CA interaction", "CA license + MLB CA-transfer option — verify", "CA UEs stay filtered"],
            ["NCellTrigThldSmartOptAlgoSw", "Intelligent / SON MLB option + MAE counter subscription", "No learned neighbor thd"],
        ],
        col_fills={2: YELLOW},
    )
    r += 1

    r = section(ws, r, cols, "10", "All Parameter List  —  Feature 3  (sequence)")
    mlb_params = [
        ["1", "CellAlgoSwitch", "InterFreqMlbSwitch", "Master connected/idle MLB family bit", "ON capacity", "License, NRT", "Table 6-2", "Activate", "[MLB] p.127"],
        ["2", "CellAlgoSwitch", "InterFreqIdleMlbSwitch", "Idle MLB", "ON capacity", "T320, SIB5", "After target policy", "Activate", "[MLB]/[IM]"],
        ["3", "CellAlgoSwitch", "InterFreqBlindMlbSwitch", "Blind MLB", "OFF", "Containment", "Not for this campaign", "Activate", "[MLB]"],
        ["4", "CellMLB", "MlbTriggerMode", "UE_NUMBER_ONLY / PRB_ONLY / PRB_OR_UE_NUMBER", "UE_NUMBER_ONLY", "CA penetration", "Primary", "Trigger", "[MLB] Table 6-2"],
        ["5", "CellMLB", "InterFreqUeTrsfType", "SynchronizedUE / IdleUE / PrbMlb…", "SynchronizedUE", "Mode match", "IdleUE additional if idle ON", "Trigger", "[MLB] Table 6-2"],
        ["6", "CellMLB", "PrbLoadCalcMethod", "PRB_USAGE vs valuation", "Leave default unless PRB supplement designed", "Mutually exclusive vs valuation", "Not primary", "PRB", "[MLB] Table 6-18"],
        ["7", "eval SW", "ActiveUeBasedLoadEvalSw", "Active DL-buffer UEs", "ON", "Unequal BW", "Recommended", "Load model", "[MLB] Table 5-5"],
        ["8", "eval SW", "SpectralEffBasedLoadEvalSw", "SE-aware C", "ON", "≥10 UL-sync for refresh", "Recommended", "Load model", "[MLB] Table 5-5"],
        ["9", "CellMLB / Cell", "CellCapacityScaleFactor", "Manual C scale", "1 unless known HW/MIMO difference", "SE switch", "Do not fake L2100 to 20 MHz", "Load model", "[MLB] §5.1.1"],
        ["10", "eval SW", "LoadTransferEnhSw", "Enhanced multi-target / PRB-diff", "ON", "4×L2600", "Recommended", "Transfer vol", "[MLB] p.134"],
        ["11", "eval SW", "CaUserLoadTransferSw", "CA/PCC transfer", "ON after audit", "Target CA C ≥ serving (one path)", "Else CA filtered", "CA", "[MLB] pp.129–136"],
        ["12", "CellMLB", "MlbTrigJudgePeriod", "How long condition must hold", "Start conservative (do not use fastest if chatter)", "InterFreqLoadEvalPrd", "Stability", "Timing", "[MLB] p.128"],
        ["13", "CellMLB", "InterFreqLoadEvalPrd", "Load-exchange / eval period", "Do not use 5 s with large MlbMaxUeNum", "MlbMaxUeNum", "Over-transfer warning p.136", "Timing", "[MLB] p.136"],
        ["14", "CellMLB", "InterFreqMlbUeNumThd", "Connected UE-number enter/leave thd (leave = thd; enter = thd+offset)", "Calibrate per layer after SE switch ON; do not copy 20 MHz thd to L2100", "Active/SE switches", "Key daily CR candidate", "Trigger", "[MLB] p.128"],
        ["15", "CellMLB", "MlbUeNumOffset", "Enter hysteresis", "Keep >0", "Thd", "Stops on/off chatter", "Trigger", "[MLB] p.128"],
        ["16", "CellMLB", "InterFreqIdleMlbUeNumThd / InterFIdleUeNumOffloadOfs", "Idle trigger / offload offset", "Calibrate separately", "Idle switch", "Feature 1", "Idle", "[MLB] Table 5-2"],
        ["17", "CellMLB", "MlbMaxUeNum", "Max UEs per eval", "Conservative (e.g. start well below 40; site-calibrate)", "Eval period", "Never ≥40 with 5 s", "Volume", "[MLB] p.136"],
        ["18", "CellMLB", "LoadDiffThd / InterFreqOffloadOffset / InterFrqUeNumOffloadOffset", "How different target must be / offload slack", "Keep equalisation (not blind offload) in Huawei pool", "Load exchange", "Offload if X2 missing only", "Admit", "[MLB]"],
        ["19", "CellMLB", "InterFreqMlbThd / InterFreqMlbUlThd / LoadOffset / MlbMinUeNumThd", "PRB trigger family", "Use only if supplement designed", "CA exclusion", "Not primary", "PRB", "[MLB] §6.5.1"],
        ["20", "EutranInterNFreq", "InterFreqMlbDlPrbOffset / UlPrbOffset", "Per-freq PRB bias", "0 at start", "PRB mode", "Static bias last resort", "PRB", "[MLB] p.210"],
        ["21", "CellMLB", "FreqSelectStrategy / MlbFreqPriority / MlbFreqUlPriority", "Freq pick", "LOADPRIORITY", "Load exchange", "PRIORITYBASED only for PCC design", "Freq pick", "[MLB] p.137"],
        ["22", "CellMLB", "MlbHoCellSelectStrategy", "Cell pick among measured", "ONLY_STRONGEST_CELL", "A4", "Mandatory recommended", "Cell pick", "[MLB] Table 6-5"],
        ["23", "EutranInterNFreq", "MlbTargetInd / OverlapInd", "Target eligibility + overlap", "Capacity ALLOWED + overlap; L900 no connect/idle MLB", "NoHoFlag", "Yellow L900", "Target", "[MLB] pp.28, 129"],
        ["24", "EutranInterFreqNCell", "NoHoFlag", "Permit HO", "PERMIT on capacity and L900 coverage path", "Blacklist", "MLB exclusion ≠ coverage block", "NRT", "[MLB] p.27"],
        ["25", "CellMLB", "NCellHoSuccRateThld", "Min pair HO success to stay a target", "Keep; fix RF if pair fails it", "Coverage HO KPI", "Do not lower to force MLB", "Admit", "[MLB] p.27"],
        ["26", "CellMLB", "CellPunishPrdNum / FreqPunishPrdNum / PunishJudgePrdNum", "Target/freq penalty", "Keep default unless stuck", "Eval period", "Rejected target cool-down", "Penalty", "[MLB] p.29"],
        ["27", "CellMlbUeSel", "UeSelectArpPrio / PrbPrio / DlMcsPrio / QciPrio + thds", "Who is movable", "Prefer good-radio moderate-load UEs; protect QCI1", "SPID, eMBMS, emergency", "No edge-PRB hunting", "UE pick", "[MLB] pp.130–135"],
        ["28", "EutranInterNFreq", "SnrBasedUeSelectionMode", "SNR-aware UE pick", "ON if available and RF variance high", "Meas", "Avoid moving cell-edge to L2600", "UE pick", "[MLB] p.130"],
        ["29", "CellMLB", "MlbHoInProtectTimer / MlbUeSelectPunishTimer", "No immediate re-MLB", "Keep non-zero", "FreqPri protect timer", "Ping-pong guard", "Protect", "[MLB]"],
        ["30", "EutranInterNFreq", "MlbInterFreqHoEventType + IfMlbThdRsrpOffset", "A4/A5 + offset", "A4, offset 0 start", "Feature 2 A4 thd", "Co-sited", "Event", "[MLB] p.137"],
        ["31", "CellMLB", "InterFreqMlbBlindHo + BlindHoPriority", "Blind path", "OFF", "Containment", "Amber", "Blind", "[MLB] p.136"],
        ["32", "CellOpHoCfg", "MLB_HO_FORBID_SW", "No MLB for high-mobility UE", "Optional high-speed cells", "NLOS", "Feature 2", "Safeguard", "[MLB] pp.136–141"],
        ["33", "smart", "NCellTrigThldSmartOptAlgoSw + LocalToNCellMlbUeNumThld / NToLocal… / LocalToNCellMlbPrbThld", "Learned pair thds", "Monitor, do not daily MOD", "7-day window; 15-min counters", "AI read-only", "SON", "[MLB] pp.29–31"],
        ["34", "CellPrbValMlb", "PrbValMlbTrigThd / AdmitThd / FilterFactor", "PRB-evaluation MLB", "OFF unless GBR problem", "Exclusive vs PRB_USAGE", "Not TP-fairness tool", "PRB val", "[MLB] Table 6-25"],
        ["35", "RrcConnStateTimer", "T320ForLoadBalance", "Idle dedicated prio life", "MIN30/60", "Idle MLB", "Feature 1", "Idle", "[MLB] §5.1.1.5"],
        ["36", "EnhancedMlb", "DediPrioManageOnLowLoadSw", "Hold low-load prio", "ON with idle MLB", "T320", "Feature 1", "Idle", "[MLB] Table 5-10"],
    ]
    r = table(
        ws,
        r,
        ["Seq", "MO", "Parameter", "Description", "Proposed value", "Conditional / other feature", "Notes", "Group", "Doc ref"],
        mlb_params,
        col_fills={5: GREEN},
        min_h=24,
    )
    r += 1

    r = section(ws, r, cols, "11", "Final MML Command for Activations  —  maintain sequence")
    r = note(
        ws,
        r,
        cols,
        "  SEQUENCE: license+LST → target eligibility (allow capacity, block L900 MLB) → load-eval switches → CELLMLB trigger/volume/strategy → CA transfer → idle bits/T320 → A4 event (Feature 2) → verify counters. "
        "One family per cluster. Original CSV SymbolShutdownSwitch example is not part of MLB.",
        AMBER,
        36,
    )
    r = table(
        ws,
        r,
        ["Seq", "MO", "Activation value / MML skeleton", "Conditional parameter", "Remarks", "Parameter description", "More notes"],
        [
            ["0", "—", "LST CELLALGOSWITCH: LocalCellId=<x>;  LST CELLMLB: LocalCellId=<x>;  LST EUTRANINTERNFREQ: LocalCellId=<x>;  LST CELLMLBUESEL: LocalCellId=<x>;  check license Intra-RAT MLB;", "License + MAE", "Baseline", "Dump", "Attach to CR"],
            ["1", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L1800|L2100|L2600>, MlbTargetInd=ALLOWED, MlbInterFreqHoEventType=A4;", "OverlapInd valid; NoHoFlag=PERMIT; FREQ_MEAS_FLAG; not forbid-meas", "Capacity may be targets", "Target + event", "From L900 toward capacity also ALLOWED"],
            ["2", "EUTRANINTERNFREQ", "MOD EUTRANINTERNFREQ: LocalCellId=<x>, DlEarfcn=<L900>, MlbTargetInd=ALLOWED_WITHOUT_CONNECT_MLB;  (add WITHOUT_IDLE_MLB if enum exists — verify)", "Coverage HO to L900 remains PERMIT", "L900 is coverage, not capacity target", "L900 protect", "Yellow — verify exact enum in parameter reference"],
            ["3", "CELLALGOSWITCH / eval MO", "Turn ON ActiveUeBasedLoadEvalSw, SpectralEffBasedLoadEvalSw, LoadTransferEnhSw;", "Unequal BW (L2100 15 MHz)", "Do this BEFORE chasing UE-number thd", "Load model", "Exact bit names: confirm MAE"],
            ["4", "CELLMLB", "MOD CELLMLB: LocalCellId=<x>, MlbTriggerMode=UE_NUMBER_ONLY, InterFreqUeTrsfType=SynchronizedUE, FreqSelectStrategy=LOADPRIORITY, MlbHoCellSelectStrategy=ONLY_STRONGEST_CELL;", "InterFreqMlbSwitch will be ON", "Core connected equalisation", "Mode + pick strategy", "Huawei ONLY_STRONGEST_CELL"],
            ["5", "CELLMLB", "MOD CELLMLB: LocalCellId=<x>, InterFreqMlbUeNumThd=<calib>, MlbUeNumOffset=<hyst>, MlbMaxUeNum=<conservative>, MlbTrigJudgePeriod=<stable>, InterFreqLoadEvalPrd=<not 5s-if-maxUE-large>;", "Active+SE eval already ON", "Do not copy 20 MHz thd to L2100", "Trigger + volume", "Key guarded CR fields"],
            ["6", "CA / CELLALGOSWITCH", "Enable CaUserLoadTransferSw only after PCell/SCell/active-CC baseline;", "Target CA capability path", "High CA on L2600", "CA transfer", "PRB MLB still will not move CA UEs"],
            ["7", "CELLALGOSWITCH", "MOD CELLALGOSWITCH: LocalCellId=<x>, MlbAlgoSwitch=InterFreqMlbSwitch-1&InterFreqIdleMlbSwitch-1;", "Steps 1–5 done; idle SIB5 NORMAL; L900 not idle target", "Master ON last (or in same window after targets)", "Master bits", "Blind bit stays 0"],
            ["8", "RRCCONNSTATETIMER", "MOD RRCCONNSTATETIMER: T320ForLoadBalance=MIN30;", "Idle MLB ON", "Start conservative", "Dedicated prio life", "Feature 1"],
            ["9", "CELLMLBUESEL", "MOD CELLMLBUESEL: LocalCellId=<x>,  (protect QCI1 / emergency; avoid aggressive UeSelectPrbPrio on edge MCS);", "VoLTE cells", "Do not move weak indoor to L2600", "UE pick", "ONLY_STRONGEST already on CELLMLB"],
            ["10", "Verification", "L.HHO.InterFreq.Load.* and UeNumLoad.*; L.InterFreq.HighLoad.Dur.Cell; L.InterFreq.Load.Meas(Succ); L.Traffic.ActiveUser.DL.Avg; L.Traffic.User.PCell/SCell.DL.Avg; L.ChMeas.PRB.DL.Used.Avg; L.Thrp.bits.DL / L.Thrp.Time.DL; idle DedicatedPri counters. SON logs: Inter-Frequency Handover Statistics + Idle Mode Release Statistics.", "15-min subscribed", "Judge sector-cluster, not only source", "Counters [MLB] Tables 6-6, 6-21, 6-28 pp.164, 227–228, 248–249", "Prep/Exec/Succ formulas in sheet 04"],
        ],
        col_fills={3: GREEN, 4: AMBER},
        min_h=36,
    )
    r += 1
    r = note(ws, r, cols, "  Expected SON log contents: [MLB] §§6.1.4.2 and 6.5.4.2 pp.163, 226–227. A successful HO that only shifts congestion or degrades the moved UE is NOT a successful optimisation.", REF, 28)
    return ws


# ---------------------------------------------------------------------------
# AI / CR
# ---------------------------------------------------------------------------
def build_ai(wb):
    ws = wb.create_sheet("04_AI_CR_Logic")
    cols = 10
    set_widths(ws, [8, 22, 22, 18, 16, 16, 16, 18, 20, 36])
    apply_sheet_setup(ws, "AI agent and change request")

    r = 1
    r = banner(ws, r, cols, "  COMBINED AI AGENT  ·  Daily KPI  →  RCA  →  Guarded CR", HUAWEI_RED, font_title, 32)
    r = banner(ws, r, cols, "  Trigger: capacity-layer DL user-throughput gap > 2 Mbps   ·   Objective: reduce gap toward 1–2 Mbps without harming L900 / VoLTE / mobility", NAVY, font_white_s, 20)
    r = add_legend_row(ws, r, cols)
    r += 1

    r = section(ws, r, cols, "A", "End-to-end agent sequence")
    r = flow_row(ws, r, ["Import KPI+params", "Group co-sited sector", "Gap >2 Mbps?", "RCA class", "CR or RCA-only"], [NAVY, "1B4F72", "B9770E", "6C3483", HUAWEI_RED], cols)
    r = bullets(
        ws,
        r,
        cols,
        [
            "Group by site + sector: L1800, L2100, L2600 C1–C4 (capacity). Keep L900 in the same group only as a safeguard, not as a fairness peer.",
            "Use busy-hour and ≥7 days when possible. Reject intervals with outage, missing counters, topology change, or sleeping cell.",
            "Effective capacity C_i ≈ PRB_i × SE_i × MIMO_i × Availability_i − GBR_i − control. TargetUsers_i = TotalCapacityUsers × C_i / ΣC.",
            "Primary score = active-UE share vs C share; PRB%, TP (mean/median/P5), high-load duration, CA PCell/SCell as guardrails.",
        ],
        WHITE,
        20,
    )
    r += 1

    r = section(ws, r, cols, "B", "RCA classes  (gap > 2 Mbps is not a CR by itself)")
    r = table(
        ws,
        r,
        ["Class", "Typical observation", "MLB/A3/A4 change?", "What to do instead / additionally"],
        [
            ["1 Load / PCell user imbalance", "Low TP + high PRB + high active UEs on source; target spare C + good overlap", "YES — Feature 3 candidate", "Active-UE MLB, LOADPRIORITY, modest thd/offset"],
            ["2 Heavy-user PRB imbalance", "Few UEs, very high PRB, non-CA dominated", "YES — PRB supplement only", "Never as sole algorithm if CA high"],
            ["3 CA / SCell problem", "PCell looks bad but SCell traffic healthy, or SCell never activates", "NO (first)", "CA combination / activation / scheduler; then maybe PCC policy"],
            ["4 RF / interference", "Low TP + low CQI/SINR, PRB not high", "NO", "RF, PCI, overshoot, external interference"],
            ["5 UE capability", "L2600 empty because UEs have no band", "NO", "Device mix; do not punish L1800"],
            ["6 Mobility / meas failure", "High PrepAtt, low Exec, meas succ low", "NO until flags/NRT/A4 delivery fixed", "Feature 2 audit"],
            ["7 HW / transport / alarm", "Low TP, resources look free, alarms or HighLoad transport", "NO — target illegal", "Clear alarm; do not MLB onto it"],
            ["8 L900 indoor dependence", "Users only survive on L900; high-band RSRP poor", "NO forced offload", "RF / indoor; A1-escape only where overlap is strong"],
        ],
        col_fills={3: AMBER},
    )
    r += 1

    r = section(ws, r, cols, "C", "Approval rule for generating a Change Request")
    r = para(
        ws,
        r,
        cols,
        "IF throughput gap > 2 Mbps AND condition is sustained (not 1×15-min) AND source normalized load is high AND target has spare effective capacity "
        "AND target RF/overlap is adequate AND target HO success healthy AND no HW/transport/RF alarm AND drop / VoLTE / re-est / L900 indoor guardrails pass "
        "AND sample size sufficient THEN generate a BOUNDED CR (one parameter family). ELSE RCA only.",
        GREEN,
        48,
    )
    r += 1

    r = section(ws, r, cols, "D", "Which CR family to pick (least invasive first)")
    r = table(
        ws,
        r,
        ["Order", "If you see", "CR family (one only)", "Sheet"],
        [
            ["1", "Missing NRT / meas flag / object cap / L900 set as MLB target by mistake", "Eligibility / flag / MlbTargetInd fix", "F2 + F3"],
            ["2", "Raw UE MLB on 15 vs 20 MHz, SE ignored", "Turn ON ActiveUe + SpectralEff eval", "F3 SN-5"],
            ["3", "Good overlap, proven load delta, healthy HO", "Modest InterFreqMlbUeNumThd / offset / MlbMaxUeNum", "F3 SN-11 step 5"],
            ["4", "Four L2600 unequal but RF similar, load exchange OK", "FreqSelectStrategy=LOADPRIORITY; equal idle priority", "F3 + F1"],
            ["5", "CA UEs stuck on busy PCC, target CA capable", "CaUserLoadTransferSw after audit", "F3"],
            ["6", "Released UEs bounce to busy layer", "Idle MLB + T320 + DediPrioManageOnLowLoadSw", "F1"],
            ["7", "L900 occupancy high AND high-band MR is good", "A1/FreqPri escape / ThreshXhigh (not weaker L900 floors)", "F1+F2"],
            ["8", "Ping-pong", "Increase A4 vs A2 separation, TTT/hyst, kill reverse target, ONLY_STRONGEST_CELL", "F2+F3"],
            ["9", "Drops / late indoor", "Earlier A2/A5 rescue — NOT more offload", "F2"],
        ],
        col_fills={3: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "E", "Change Request template  (copy one row per CR)")
    r = table(
        ws,
        r,
        ["Field", "What to fill", "Example", "Mandatory?"],
        [
            ["CR ID", "Date-cluster-seq", "20260823-DHA-SEC2-01", "Yes"],
            ["Site / sector / cells", "All 6 capacity + L900 CGI", "", "Yes"],
            ["Busy-hour window", "Local BH, ≥7d if possible", "", "Yes"],
            ["TP max / min / gap Mbps", "Capacity layers only", "18.4 / 6.1 / 12.3", "Yes"],
            ["RCA class (1–8)", "From SN-B", "1", "Yes"],
            ["Source / target EARFCN", "Capacity only", "L1800 → L2600C2", "Yes"],
            ["Active UE, PRB%, SE, PCell/SCell", "Source and target", "", "Yes"],
            ["Target RSRP/RSRQ/CQI evidence", "MR or HO meas succ", "", "Yes"],
            ["Load / Coverage / FreqPri HO succ", "Prep/Exec/Succ", "", "Yes"],
            ["Alarms / availability", "None on target", "", "Yes"],
            ["MO + parameter + old → new", "One family", "CELLMLB InterFreqMlbUeNumThd 30→24", "Yes"],
            ["MML sequence from F1/F2/F3 SN-11", "Paste validated MML", "", "Yes"],
            ["L900 impact assessment", "Users/PRB/VoLTE/access", "No inbound MLB; coverage HO unchanged", "Yes"],
            ["CA impact", "PCell/SCell/active CC", "", "Yes"],
            ["Expected UE move", "Order-of-magnitude", "< MlbMaxUeNum per period", "Yes"],
            ["Rollback trigger", "Numeric", "Drop +0.1pp or HO exec −2pp or L900 PRB +10pp", "Yes"],
            ["Approver", "RNO + change board", "", "Yes"],
        ],
        col_fills={4: GREEN},
    )
    r += 1

    r = section(ws, r, cols, "F", "Hard rollback / block list")
    r = bullets(
        ws,
        r,
        cols,
        [
            "Block CR if target alarmed, High/OverLoad transport/HW, PCI conflict, pair HO success below NCellHoSuccRateThld, or sample too small.",
            "Rollback if L900 active users/PRB rise abnormally, L900 TP or VoLTE worsen, MLB Prep/Exec success fall, ping-pong/drop/re-est rise, CA active-CC fall, or high-load just moves without TP gain.",
            "Do not daily MOD LocalToNCellMlbUeNumThld / smart learned thresholds.",
            "Do not change trigger + event thd + target scope + transfer volume in one night.",
            "Evaluate at sector-cluster level. Moving congestion is not an optimisation win.",
        ],
        AMBER,
        20,
    )
    r += 1

    r = section(ws, r, cols, "G", "KPI / counter pack the agent must ingest")
    r = table(
        ws,
        r,
        ["Group", "Counter / KPI", "Use", "Source book"],
        [
            ["Throughput", "L.Thrp.bits.DL / L.Thrp.Time.DL  (+ median/P5 if available)", "Gap detection", "[MLB]"],
            ["Load", "L.ChMeas.PRB.DL.Used.Avg, UL PRB", "Guardrail / PRB RCA", "[MLB]"],
            ["Users", "L.Traffic.User.Ulsync.Avg, L.Traffic.ActiveUser.DL.Avg", "N in Load=N/C", "[MLB]"],
            ["CA", "L.Traffic.User.PCell.DL.Avg, L.Traffic.User.SCell.DL.Avg", "Class 3", "[MLB]"],
            ["MLB HO", "L.HHO.InterFreq.Load.Prep/Exec/ExecSucc (+ UeNumLoad + InterFddTdd + NCell)", "MLB health", "[MLB] Table 6-6"],
            ["MLB state", "L.InterFreq.HighLoad.Dur/Num.Cell, Load.Meas / MeasSucc", "Trigger vs meas", "[MLB]"],
            ["Coverage HO", "L.HHO.InterFreq.Coverage.Prep/Exec/Succ", "Safeguard", "[CM] Table 5-24"],
            ["FreqPri HO", "L.HHO.InterFreq.FreqPri.*", "L900 escape / high-band steer", "[CM] Tables 11-9/11-10"],
            ["Idle steer", "L.RRCRel.load.DedicatedPri.LTE.High, L.RRCRel.Lowload.DedicatedPri.LTE.High", "Idle MLB working?", "[MLB]"],
            ["Safety", "Drop, RRC re-est (L.RRC.ReEst.ReconfFail.Att), VoLTE, accessibility, availability", "Hard guardrail", "[CM]/ops"],
        ],
        col_fills={4: REF},
    )
    r += 1
    r = note(
        ws,
        r,
        cols,
        "  Formula when both PRB and UE-number MLB are on: PRB-based HO ≈ total Load counters − UeNumLoad counters. [MLB] §6.5.4.3 pp.227–228.  "
        "PrepSucc = ExecAtt/PrepAtt ; ExecSucc = ExecSucc/ExecAtt ; E2E = ExecSucc/PrepAtt.",
        REF,
        32,
    )
    return ws


# ---------------------------------------------------------------------------
# INDEX
# ---------------------------------------------------------------------------
def build_index(wb):
    ws = wb.create_sheet("05_Parameter_Index")
    cols = 10
    set_widths(ws, [8, 14, 22, 28, 36, 28, 18, 14, 14, 22])
    apply_sheet_setup(ws, "Parameter index")

    r = 1
    r = banner(ws, r, cols, "  CROSS-FEATURE PARAMETER INDEX  (filter this sheet in Excel)", NAVY, font_title, 28)
    r = add_legend_row(ws, r, cols)

    headers = ["SN", "Feature", "MO", "Parameter", "Proposed / policy", "Conditional", "Group", "L900 flag", "CR-able daily?", "Doc"]
    rows = [
        ["1", "F1 Idle", "CellResel", "CellReselPriority", "L2600=7 L1800=6 L2100=5/6 L900=2", "SIB3", "Priority", "Lowest", "Rare", "[IM]"],
        ["2", "F1 Idle", "EutranInterNFreq", "CellReselPriority + CfgInd", "CFG + same hierarchy", "SIB5", "Priority", "Lowest toward L900", "Rare", "[IM]"],
        ["3", "F1 Idle", "EutranInterNFreq", "MeasPerformanceDemand", "NORMAL", "Idle MLB target", "Meas", "NORMAL (keep visible)", "No as congestion tool", "[IM]"],
        ["4", "F1 Idle", "CellResel", "SNonIntraSearch", "CFG; example 10", "SIntraSearch > this", "Meas", "", "Slow", "[IM]/[MLB]"],
        ["5", "F1 Idle", "EutranInterNFreq", "ThreshXhigh", "Calibrate L900→high-band", "Timer", "Resel", "YES protect", "Guarded", "[IM]"],
        ["6", "F1 Idle", "CellResel", "ThrshServLow", "Allow fallback before collapse", "A2/A5 align", "Resel", "YES", "Guarded", "[IM]"],
        ["7", "F1 Idle", "EutranInterNFreq", "QoffsetFreq", "0 among L2600", "Equal prio", "Rank", "", "Only stable RF", "[IM]"],
        ["8", "F1 Idle", "RrcConnStateTimer", "T320ForLoadBalance", "MIN30/60", "Idle MLB", "Dedicated", "", "Yes small steps", "[IM]"],
        ["9", "F1 Idle", "CellAlgoSwitch", "InterFreqIdleMlbSwitch", "ON capacity", "License + target ind", "Idle MLB", "OFF toward L900", "After design", "[MLB]"],
        ["10", "F2 Conn", "InterFreqHoGroup", "A1/A2/A5 coverage set", "Calibrate; not example dBm", "Meas flags", "Coverage", "A5 in / A1 out", "Safety first", "[CM]"],
        ["11", "F2 Conn", "InterFreqHoGroup", "InterFreqLoadBasedHoA4ThdRsrp", "Calibrate; A4 better than A2", "TTT ≠ 5120 ms", "MLB/FreqPri", "", "Guarded", "[CM]/[MLB]"],
        ["12", "F2 Conn", "EutranInterNFreq", "MlbInterFreqHoEventType", "A4", "A5 license", "MLB event", "Not MLB A4 toward L900", "Rare", "[MLB]"],
        ["13", "F2 Conn", "EutranInterFreqNCell", "CellIndividualOffset", "0 toward L900", "NRT", "CIO", "YES", "No global +CIO", "[CM]"],
        ["14", "F2 Conn", "FreqPri SW", "MlbBasedFreqPriHoSwitch", "ON if MLB ON", "MLB", "Coord", "", "Once", "[CM]"],
        ["15", "F2 Conn", "CellUeMeasControlCfg", "MaxNonIntraMeasObjNum", "≥6 on 7-layer", "4×L2600", "Object cap", "Keep L900 measurable", "If missing carriers", "[CM]"],
        ["16", "F3 MLB", "CellAlgoSwitch", "InterFreqMlbSwitch", "ON capacity", "License", "Master", "Not to fill L900", "Once", "[MLB]"],
        ["17", "F3 MLB", "CellMLB", "MlbTriggerMode", "UE_NUMBER_ONLY", "CA", "Mode", "", "Rare", "[MLB]"],
        ["18", "F3 MLB", "eval", "ActiveUeBasedLoadEvalSw", "ON", "Unequal BW", "Model", "", "Once then leave", "[MLB]"],
        ["19", "F3 MLB", "eval", "SpectralEffBasedLoadEvalSw", "ON", "SE refresh", "Model", "", "Once then leave", "[MLB]"],
        ["20", "F3 MLB", "CellMLB", "InterFreqMlbUeNumThd + Offset", "Calibrate per layer", "Active+SE ON", "Trigger", "Do not include L900 in pool", "MAIN daily CR", "[MLB]"],
        ["21", "F3 MLB", "CellMLB", "MlbMaxUeNum", "Conservative; not ≥40 with 5s", "Eval prd", "Volume", "", "Guarded", "[MLB]"],
        ["22", "F3 MLB", "CellMLB", "MlbHoCellSelectStrategy", "ONLY_STRONGEST_CELL", "A4", "Pick", "", "Set once", "[MLB]"],
        ["23", "F3 MLB", "CellMLB", "FreqSelectStrategy", "LOADPRIORITY", "Load exchange", "Freq pick", "", "Set once", "[MLB]"],
        ["24", "F3 MLB", "EutranInterNFreq", "MlbTargetInd", "Capacity ALLOWED; L900 without connect/idle MLB", "Overlap", "Target", "YES", "Set once + audit", "[MLB]"],
        ["25", "F3 MLB", "eval", "CaUserLoadTransferSw", "ON after CA audit", "CA license", "CA", "", "After baseline", "[MLB]"],
        ["26", "F3 MLB", "smart", "NCellTrigThldSmartOptAlgoSw / learned thds", "Monitor only daily", "7-day learn", "SON", "", "NO daily overwrite", "[MLB]"],
        ["27", "F3 MLB", "CellAlgoSwitch", "InterFreqBlindMlbSwitch", "OFF", "Containment", "Blind", "", "No", "[MLB]"],
        ["28", "All", "EutranInterNFreq", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "Deselected for L900 + capacity", "Coverage", "Filter", "Must stay measurable", "Audit", "[CM]"],
    ]
    start = r
    r = table(ws, r, headers, rows, col_fills={5: GREEN, 8: YELLOW}, min_h=22)
    ws.auto_filter.ref = f"A{start}:J{r-1}"
    ws.freeze_panes = f"A{start+1}"
    r += 1
    r = note(ws, r, cols, "  Use Excel AutoFilter on row headers. 'CR-able daily?' = whether the AI agent may propose this as a same-week change. Learned SON thresholds are monitor-only.", REF, 24)
    return ws


def style_tabs(wb):
    colors = {
        "00_Cover": "0D2B4A",
        "01_Idle_Mode": "1B4F72",
        "02_Connected_Mode": "117A65",
        "03_IntraRAT_MLB": "C7000B",
        "04_AI_CR_Logic": "6C3483",
        "05_Parameter_Index": "1F4E79",
    }
    for name, col in colors.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = col


def main():
    import os

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    build_cover(wb)
    build_idle(wb)
    build_connected(wb)
    build_mlb(wb)
    build_ai(wb)
    build_index(wb)
    style_tabs(wb)
    wb.properties.title = "Robi 4G Mobility Management eRAN21.1"
    wb.properties.creator = "RNO workbook generator for Mohammad Selim / Robi Axiata PLC"
    wb.properties.subject = "Idle Mode, Connected Mode, Intra-RAT MLB, AI CR logic"
    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
