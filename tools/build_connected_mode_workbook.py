#!/usr/bin/env python3
"""Connected Mode eRAN21.1 — one Word-like sheet per feature chapter."""

import os
import sys
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cm_docstyle import DocSheet, NAVY

VER = "v2.2"
OUT = "/workspace/docs/4G_LTE_Mobility_Management/Connected_Mode_eRAN21.1_Feature_Sheets_v2.2.xlsx"
DOC = "Mobility Management in Connected Mode Feature Parameter Description"
ISSUE = "Huawei eRAN21.1 Issue 08 (2026-06-30)"


def R(feat, sec):
    """Reference column: Feature ID + document name + issue + section."""
    return f"{feat}  ·  {DOC}  ·  eRAN21.1 Issue 08  ·  {sec}"

# Sheet names (Excel limit 31 characters)
S = {
    "toc": "00 Contents",
    "ov": "03 Overview",
    "b": "04 Basic Functions",
    "c": "05 Coverage HO",
    "s": "06 Service HO",
    "d": "07 Distance HO",
    "u": "08 UL-Quality HO",
    "q": "09 CQI Inter-Freq HO",
    "r": "10 Service-Request HO",
    "f": "11 Freq-Priority HO",
    "p": "12 Speed-based HO",
    "m": "13 UTRAN Multi-PLMN",
    "g": "14 GERAN Multi-PLMN",
}

ORDER = ["toc", "ov", "b", "c", "s", "d", "u", "q", "r", "f", "p", "m", "g"]


def nav_for(key):
    i = ORDER.index(key)
    prev_k = ORDER[i - 1] if i > 0 else None
    next_k = ORDER[i + 1] if i + 1 < len(ORDER) else None
    items = [("Contents", S["toc"])]
    items.append((f"← {S[prev_k]}" if prev_k else "", S[prev_k] if prev_k else None))
    items.append((f"{S[next_k]} →" if next_k else "", S[next_k] if next_k else None))
    if key != "b":
        items.append(("Ch.4 Basic (must)", S["b"]))
    return items


def start(wb, key, chapter, name, meta, tab=NAVY):
    ws = wb.create_sheet(S[key])
    ws.sheet_properties.tabColor = tab
    d = DocSheet(ws, f"{chapter} {name}")
    d.banner(f"{DOC}   ·   {ISSUE}   ·   {VER}")
    d.title(chapter, name)
    d.meta(meta)
    d.nav(nav_for(key))
    d.space(8)
    return d


def sheet_toc_fix(wb):
    ws = wb.create_sheet(S["toc"])
    d = DocSheet(ws, f"{DOC}  |  Contents")
    d.banner(f"{DOC}   ·   {ISSUE}   ·   {VER}")
    d.title("Contents", "How to read this file")
    d.meta("Version v2.2. One sheet = one feature chapter. Every feature sheet uses boxed Common / Overview and numbered Principle. Chapter 4 is basic for every later feature.")
    d.nav([("This page", None), (S["ov"] + " →", S["ov"]), ("Ch.4 Basic (must)", S["b"])])
    d.h1("What this file is")
    d.para(f"{DOC}. {ISSUE}.")
    d.para("The document describes several connected-mode handover features. Chapter 4 is the common engine (measurement, events A1–A5 / B1–B2, admission, retry). Every later feature re-uses Chapter 4 and only changes how the handover is started and which event / target is used.")
    d.info_box("How to read every feature sheet", [
        "One sheet per feature. Hyperlinks jump to the related feature in one click.",
        "Gridlines are off. Introduction → Common for sub-group / Overview (boxed) → Principle (numbered, boxed) → each sub-group → Combined summary → Parameter list.",
        "Chapter 5 Common for sub-group (document §5.1) is shared by 5.2–5.6. Chapters 6, 7 and 8 use the same Common-for-sub-group box style.",
        "This file is version v2.2.",
        "Parameter list: Value = value only (blue). Command-example dBm in the book are not design values.",
        "FDD Feature IDs from §2.3. TDD uses the TD* equivalent unless the document says FDD only.",
        "LBFD-131111 FDD↔TDD is covered inside 5.3, 6.2 and 7.2 (inter-duplex = inter-frequency).",
    ])
    d.h1("Read in this sequence")
    heads = ["Ch.", "Feature in the document", "Page", "Feature ID / switch", "Open"]
    from cm_docstyle import ft, fl, L, NAVY, HDR_BG, LINE, LINK, TEXT, COLS
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.hyperlink import Hyperlink
    r = d.r
    ws = d.ws
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=COLS)
    for c, h in enumerate(heads, 1):
        cell = ws.cell(r, c, h)
        cell.font = ft(9, True, NAVY)
        cell.fill = fl(HDR_BG)
        cell.alignment = L
        cell.border = Border(bottom=Side(style="thin", color=NAVY))
    d.r += 1
    rows = [
        ("3", "Overview", "21", "—", S["ov"]),
        ("4", "Basic Functions (must for all later features)", "29", "LBFD-002018", S["b"]),
        ("5", "Coverage-based Handover", "97", "LBFD-00201801 / 00201802 · LOFD-001019 / 001020 / 001078", S["c"]),
        ("6", "Service-based Handover", "189", "LBFD-00201805 · LOFD-171207 · LOFD-001043 / 001046", S["s"]),
        ("7", "Distance-based Handover", "225", "LBFD-00201804 · LOFD-001072 / 001073", S["d"]),
        ("8", "UL-Quality-based Handover", "245", "UlQualityInterFreqHoSwitch / UlQualityInterRATHoSwitch", S["u"]),
        ("9", "CQI-based Inter-Frequency Handover (FDD)", "273", "FDD. Disabled if A4 TTT = 5120 ms", S["q"]),
        ("10", "Service-Request-based Inter-Frequency Handover", "281", "ServiceReqInterFreqHoSwitch", S["r"]),
        ("11", "Frequency-Priority-based Inter-Frequency Handover", "299", "FreqPriorIFHOSwitch", S["f"]),
        ("12", "Speed-based Inter-Frequency Handover (FDD)", "320", "Requires Ch.5 inter-frequency coverage HO", S["p"]),
        ("13", "Separate Mobility Policies to UTRAN for Multi PLMN (FDD)", "330", "LOFD-070216", S["m"]),
        ("14", "Separate Mobility Policies to GERAN for Multi PLMN", "337", "LOFD-111204", S["g"]),
    ]
    for ch, name, page, ids, sheet in rows:
        r = d.r
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=COLS)
        ws.row_dimensions[r].height = 22
        for c, v in enumerate([ch, name, page, ids, "Open sheet →"], 1):
            cell = ws.cell(r, c, v)
            cell.font = ft(11, c == 1, TEXT)
            cell.alignment = L
            cell.border = Border(bottom=Side(style="hair", color=LINE))
        link = ws.cell(r, 5)
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{sheet}'!A1", display="Open sheet →")
        link.font = ft(11, True, LINK, underline="single")
        d.r += 1
    d.h1("Connection between features")
    d.para("Idle camping is not in this document. This book starts after RRC connect. Coverage (Ch.5) is the necessary rescue. Service / distance / UL-quality / CQI / service-request / frequency-priority / speed are unnecessary or specialised handovers that still use the Chapter 4 engine. Multi-PLMN (Ch.13–14) only changes the UTRAN/GERAN policy at target decision. Frequency-priority must not form a reverse pair with MLB on the same frequency.")
    return d


def sheet_overview(wb):
    d = start(
        wb, "ov", "3", "Overview of Mobility Management in Connected Mode",
        "Document page 21. This chapter is the map. It does not switch a feature on.",
    )
    d.h1("Introduction")
    d.info_box("What this chapter is", [
        "Connected-mode mobility keeps a UE in service after RRC setup.",
        "The eNodeB classifies the handover, delivers measurement (or uses blind), waits for an event, picks a target, admits it, then executes.",
        "Chapter 4 is that engine. Chapters 5–12 are different reasons to start the engine. Chapters 13–14 change inter-RAT policy per PLMN.",
        "This chapter has no switch. It is the map only.",
    ])
    d.h1("Principle")
    d.info_box("Principle", [
        "Necessary coverage handover has higher priority than unnecessary load or optimisation handover.",
        "If the serving cell can no longer carry the UE, coverage (Ch.5) runs first.",
        "A4 for most unnecessary inter-frequency handovers: the neighbour only needs to be good enough, not better than serving.",
        "Keep that A4 a higher RSRP requirement than coverage A2, or the UE returns immediately.",
    ])
    d.callout("CORE", "Core setting", "There is no switch in Chapter 3. Configure Chapter 4 first, then the feature chapter you need.")
    d.h1("How later features attach")
    d.pair_boxes(
        ("Necessary / specialised starters", [
            "Ch.5 Coverage — UE at cell edge. Intra-frequency A3. Inter-frequency A2 then A3/A4/A5 or blind. IRAT B1/B2 or blind.",
            "Ch.7 Distance — TA says the UE is beyond the planned cell. Event A4 / B1.",
            "Ch.8 UL-quality — uplink MCS and IBLER are poor. Event A4, then blind if A4 never comes.",
            "Ch.12 Speed (FDD) — high-speed UE to the coverage layer. Requires Ch.5.3.",
        ]),
        ("Unnecessary / policy starters", [
            "Ch.6 Service — QCI should live on another frequency or RAT. Event A4 (LTE) or B1 (IRAT).",
            "Ch.9 CQI (FDD) — downlink CQI is poor while RSRP may still look usable. Event A4. Off if A4 TTT = 5120 ms.",
            "Ch.10 Service-request — a new high-priority QCI arrives. Event A4.",
            "Ch.11 Frequency-priority — serving is good (A1); high-priority frequency (A4).",
            "Ch.13–14 Multi-PLMN — UTRAN/GERAN capability per operator at target decision. No new event.",
        ]),
    )
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "One engine (Ch.4). Many starters (Ch.5–12). Coverage always wins.",
        "A4 for most unnecessary inter-frequency handovers.",
        "Keep that A4 better than coverage A2 or the UE returns immediately.",
        "Open Chapter 4 next. There is no activation parameter in Chapter 3.",
    ])
    d.h1("Parameter list")
    d.para("Chapter 3 has no activation parameter. Open Chapter 4.")
    d.param_heads()
    d.param_row(
        1, "—", "—", "—",
        "Overview only. All switches start in Chapter 4 and the feature chapter.",
        "—", "Overview",
        "Chapter 3 has no activation parameter.",
        R("—", "§3 Overview"),
        link_sheet=S["b"],
    )
    return d


def sheet_basic(wb):
    d = start(
        wb, "b", "4", "Basic Functions of Mobility Management in Connected Mode",
        "Document page 29. Feature ID LBFD-002018 (TDD: TDLBFD-002018). Must be configured for every later feature.",
    )
    d.h1("Introduction — what this chapter is")
    d.info_box("Common for all later features", [
        "This is not one handover type. It is the common procedure used by coverage, service, distance, UL-quality, CQI, service-request, frequency-priority and speed.",
        "4.1.1 Overall process (document Fig 4-1).",
        "4.1.2 Handover function initiation — necessary vs unnecessary.",
        "4.1.3 Processing mode — measurement-based or blind.",
        "4.1.4 Measurement configuration — object, event, gap, SMeasure.",
        "4.1.5 Measurement reporting.",
        "4.1.6 Target cell / frequency decision and admission.",
        "4.1.7 Execution.",
        "4.1.8 Retry and penalty.",
        "If flags, neighbour lists, object capacity, events or admission are wrong here, every later feature looks broken.",
    ])
    d.h1("Principle")
    d.info_box("Principle  ·  Fig 4-1", [
        "Start HO function.",
        "Choose measurement-based or blind.",
        "Deliver measurement configuration.",
        "UE reports A1–A5 (or B1/B2).",
        "Pick target and admit.",
        "Execute.",
        "Punish / retry on fail.",
    ])
    d.callout("CORE", "Core setting for the whole book", [
        "Neighbour frequency: FREQ_MEAS_FLAG selected. HO_TRG_FREQ_FORBID_MEAS_FLAG deselected for required HO targets.",
        "Object capacity ≥ the number of inter-frequency objects this cell must measure. If over the cap, equal-priority frequencies may be picked at random. That is not load balance.",
        "Trigger quantity: RSRP recommended. RSRQ moves with scheduler load.",
        "InterFreqHoA4TimeToTrig must not be 5120 ms if frequency-priority, CQI or service-based inter-frequency HO is required (eRAN21.1 Table 4-9). 5120 ms means off, not slow.",
    ])

    d.h2("4.1.2  Necessary vs unnecessary")
    d.pair_boxes(
        ("Necessary", [
            "Serving cannot carry the UE.",
            "Cause: Handover desirable for radio reasons (intra-RAT) or Time Critical Handover (IRAT).",
            "Coverage-based HO is necessary.",
            "Admits any QCI.",
            "After admit fail, intra-eNodeB necessary HO tries the next candidate.",
            "Coverage preempts unnecessary features (Ch.6–12).",
        ]),
        ("Unnecessary", [
            "Serving can still carry the UE.",
            "Offload (MLB) or optimisation (service, frequency-priority, CQI, service-request).",
            "Target must admit ALL QCIs.",
            "Inter-eNodeB unnecessary HO does not immediately try the next target after admit fail.",
            "A4 TTT = 5120 ms disables FreqPri, CQI and service-based IFHO (eRAN21.1 Table 4-9).",
        ]),
    )

    d.h2("4.1.4  Events and offset calculation")
    d.para("Events only say signal quality. The feature chapter decides which event is used.")
    d.callout("CALC", "Calculation  ·  entering condition must hold for TimeToTrig", [
        "A1  Ms − Hys > Thresh     serving becomes good (stops coverage measurement; can start FreqPri)",
        "A2  Ms + Hys < Thresh     serving becomes poor (starts inter-frequency / IRAT measurement)",
        "A3  Mn + Ofn + Ocn − Hys > Ms + Ofs + Ocs + Off     neighbour relatively better",
        "A4  Mn + Ofn + Ocn − Hys > Thresh     neighbour absolutely good (MLB / FreqPri / service / CQI / distance / UL-quality gate)",
        "A5  Ms + Hys < Th1  AND  Mn + Ofn + Ocn − Hys > Th2     serving poor AND neighbour good",
        "B1  Mn + Ofn − Hys > Thresh     IRAT neighbour absolutely good",
        "B2  Ms + Hys < Th1  AND  Mn + Ofn − Hys > Th2     IRAT coverage pair",
        "Ofn / Ofs = frequency offset. Ocn / Ocs = CellIndividualOffset (CIO). Off = A3 offset. Hys = hysteresis of that event.",
        "QCI/operator/SPID offsets are added to the base threshold, then clamped: MAX(−140, MIN(−43, result)) for RSRP.",
    ])
    d.callout("CALC", "Example with values  ·  for understanding only  ·  not a live design  ·  not the book MML −85/−87/−103 dBm", [
        "One teaching set so the arithmetic is visible. Calibrate real cells from MR. Later chapters re-use the same idea.",
        "Assume: Ms = −95 dBm (serving RSRP), Mn = −90 dBm (neighbour RSRP), Hys = 2 dB, Ofn = Ofs = 0, Ocn = Ocs = 0, Off = 2 dB, TimeToTrig = 320 ms.",
        "A1  Thresh = −100 dBm.   Ms − Hys = −95 − 2 = −97.   −97 > −100 → ENTER. Serving is good. Coverage measurement can stop.",
        "     If Ms = −105 dBm: −105 − 2 = −107.   −107 > −100 → NO. A1 does not enter.",
        "A2  Thresh = −100 dBm.   Ms + Hys = −95 + 2 = −93.   −93 < −100 → NO. Serving is not poor. Inter-frequency measurement does not start.",
        "     If Ms = −105 dBm: −105 + 2 = −103.   −103 < −100 → ENTER. Start inter-frequency / IRAT measurement.",
        "TimeToTrig. The entering condition must stay true for 320 ms. If it is true for 200 ms and then Ms recovers, the timer resets and no report is sent.",
    ])
    d.callout("CALC", "Same teaching set  ·  A3 / A4 / A5  ·  still not a design", [
        "A3  Left = Mn + Ofn + Ocn − Hys = −90 − 2 = −92.   Right = Ms + Ofs + Ocs + Off = −95 + 2 = −93.   −92 > −93 → ENTER. Neighbour is relatively better.",
        "     If Off = 6 dB: Right = −89.   −92 > −89 → NO. Larger A3 offset makes intra-frequency HO harder.",
        "     If CIO Ocn = +3 dB: Left = −90 + 3 − 2 = −89.   −89 > −93 → still ENTER, and the neighbour looks 3 dB better.",
        "A4  Thresh = −105 dBm.   Mn + Ofn + Ocn − Hys = −92.   −92 > −105 → ENTER. Neighbour is absolutely good enough (need not beat serving).",
        "     If Mn = −110 dBm: −110 − 2 = −112.   −112 > −105 → NO. Target is not good enough.",
        "A5  Th1 = −110 dBm, Th2 = −105 dBm.   Ms + Hys = −93 < −110? NO. Serving is not poor, so A5 does not enter even though the neighbour is good.",
        "     If Ms = −115 dBm: −115 + 2 = −113 < −110 YES, and −92 > −105 YES → ENTER. Serving poor AND neighbour good.",
    ])
    d.callout("CALC", "Same teaching set  ·  B1 / B2, offset clamp, A4 vs coverage A2", [
        "B1  IRAT Mn = −92 (configured IRAT quantity), Ofn = 0, Hys = 2, Thresh = −100.   −92 − 2 = −94.   −94 > −100 → ENTER.",
        "B2  Serving A2-style Th1 = −110, IRAT B1-style Th2 = −100.   Ms = −115: −115 + 2 = −113 < −110 YES, and −94 > −100 YES → ENTER.",
        "     If Ms = −95: −93 < −110? NO. B2 does not enter even if the IRAT neighbour is good.",
        "Offset then clamp (A2 RSRP). Base A3InterFreqHoA2ThdRsrp = −100 dBm. SPID factor = +6 dB. Result = −94. MAX(−140, MIN(−43, −94)) = −94 dBm. Used A2 = −94 dBm.",
        "     If SPID factor = +70 dB: −100 + 70 = −30 → clamp to −43 dBm (cannot go above −43).",
        "     If SPID factor = −50 dB: −100 − 50 = −150 → clamp to −140 dBm (cannot go below −140).",
        "A4 vs coverage A2 (ping-pong). ‘A4 better than A2’ means a higher RSRP requirement: A4_thd > A2_thd (for example A4 = −105, A2 = −110).",
        "     Safe pair: A2 = −110 dBm, A4 = −105 dBm, Mn = −102 dBm. A4: −102 − 2 = −104 > −105 → HO. After HO, serving = −102. A2: −102 + 2 = −100 < −110? NO. Coverage measurement does not start.",
        "     Unsafe pair: A2 = −100 dBm, A4 = −110 dBm, Mn = −107 dBm. A4: −107 − 2 = −109 > −110 → HO. After HO, serving = −107. A2: −107 + 2 = −105 < −100 YES. Coverage measurement starts at once — ping-pong.",
    ])
    d.para("A1/A2 Hys and TTT: InterFreqHoGroup.InterFreqHoA1A2Hyst / InterFreqHoA1A2TimeToTrig. A3 Off: IntraFreqHoA3Offset or InterFreqHoA3Offset. A4 Hys/TTT: InterFreqHoA4Hyst / InterFreqHoA4TimeToTrig.")

    d.h2("4.1.4  Measurement objects and SMeasure")
    d.info_box("Measurement objects and SMeasure", [
        "The eNodeB filters blacklisted cells, NoHoFlag neighbours, and forbidden TA/LA.",
        "Frequencies are picked in descending priority. If the highest priority maps to several frequencies, the pick among them is random when over the object cap.",
        "SMeasure (HOMEASCOMM): if serving RSRP is above SMeasure, the UE may skip intra/inter/IRAT measurement. A too-high SMeasure silently kills A4.",
        "Measurement gap steals DL TTIs. AutoGap / gap pattern must be acceptable for VoLTE and old UEs.",
        "Algorithm measurement priority: voice/CSFB > necessary coverage > unnecessary single-UE HO > ANR sampling. MEAS_OBJ_PREEMPT_SW allows a higher-priority algorithm to take the gap.",
    ])

    d.h2("4.1.6–4.1.8  Admit, execute, punish")
    d.info_box("Admit, execute, punish", [
        "Necessary: any QCI may be admitted.",
        "Unnecessary offload: all QCIs must be admitted.",
        "Prep fail is often admission or X2, not RF.",
        "Penalty timers after fail: do not treat every prep fail as an A4 problem.",
        "One event set serves every later feature. QCI-specific Hys/TTT is possible.",
        "Object-cap random drop is not MLB. RSRQ as trigger oscillates with load. Blind mode has higher access-failure risk.",
    ])
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Fix flags, NRT, object cap, SMeasure and A2 family before any dBm.",
        "RSRP trigger.",
        "A4 better than coverage A2 (higher RSRP requirement).",
        "A4 TTT ≠ 5120 ms if Ch.6 / Ch.9 / Ch.11 inter-frequency A4 is required.",
        "Then open the feature sheet.",
    ])

    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG", "selected", "Frequency must be measured. Silent no-HO if off.", "Yes", "4.1.4",
         "Whether the UE is configured to measure this E-UTRAN frequency in connected mode.", R("LBFD-002018", "§4.1.4")),
        (2, "EUTRANINTERNFREQ", "HO_TRG_FREQ_FORBID_MEAS_FLAG", "deselected", "Must be off for required HO targets.", "Yes", "4.1.4",
         "If selected, this frequency is forbidden as a handover target even if measured.", R("LBFD-002018", "§4.1.4")),
        (3, "EUTRANINTERFREQNCELL", "NoHoFlag", "PERMIT_HO", "Neighbour allowed as HO target.", "Yes", "4.1.4",
         "Per-neighbour flag allowing or prohibiting handover to that cell.", R("LBFD-002018", "§4.1.4")),
        (4, "CELLUEMEASCONTROLCFG", "MaxNonIntraMeasObjNum", "≥ needed objects", "Over cap: equal-priority objects drop at random.", "Yes", "4.1.4",
         "Maximum number of non-intra-frequency measurement objects the eNodeB may deliver to a UE.", R("LBFD-002018", "Table 4-3")),
        (5, "CELLUEMEASCONTROLCFG", "MaxEutranFddMeasFreqNum", "≥ needed FDD freqs", "Same random-drop risk for FDD.", "Yes", "4.1.4",
         "Maximum number of neighbouring E-UTRAN FDD frequencies that can be delivered for measurement.", R("LBFD-002018", "Table 4-3")),
        (6, "HOMEASCOMM", "SMeasure", "—", "Calibrate from MR. Must not hide intended A4.", "Tune", "4.1.5",
         "Serving-cell RSRP above this value allows the UE to skip intra/inter/IRAT measurement.", R("LBFD-002018", "§4.1.5")),
        (7, "INTRARATHOCOMM", "InterFreqHoA1A2TrigQuan", "RSRP", "Recommended trigger quantity.", "Tune", "4.1.4",
         "Quantity (RSRP / RSRQ / BOTH) used to trigger inter-frequency events A1 and A2.", R("LBFD-002018", "§4.1.4")),
        (8, "INTRARATHOCOMM", "InterFreqHoA4TrigQuan", "RSRP", "Recommended. Reporting SAME_AS_TRIG_QUAN.", "Tune", "4.1.4",
         "Quantity used to trigger inter-frequency event A4.", R("LBFD-002018", "§4.1.4")),
        (9, "INTERFREQHOGROUP", "InterFreqHoA1A2Hyst", "—", "Hys in A1/A2 formulas. Keep A1/A2 pair consistent.", "Tune", "Table 4-8",
         "Hysteresis Hys applied to entering and leaving conditions of events A1 and A2.", R("LBFD-002018", "Table 4-8")),
        (10, "INTERFREQHOGROUP", "InterFreqHoA1A2TimeToTrig", "—", "TTT for A1/A2. QCI-specific optional.", "Tune", "Table 4-8",
         "Duration TimeToTrig that A1/A2 entering or leaving condition must hold.", R("LBFD-002018", "Table 4-8")),
        (11, "INTERFREQHOGROUP", "InterFreqHoA4Hyst", "—", "Hys in A4: Mn+Ofn+Ocn−Hys > Thresh.", "Tune", "Table 4-8",
         "Hysteresis Hys in the event A4 formula.", R("LBFD-002018", "Table 4-8")),
        (12, "INTERFREQHOGROUP", "InterFreqHoA4TimeToTrig", "not 5120 ms", "5120 ms disables FreqPri, CQI and service-based IFHO (eRAN21.1 Table 4-9).", "Yes", "Table 4-9",
         "Duration TimeToTrig for event A4. 5120 ms is treated as disable for FreqPri / CQI / service IFHO.", R("LBFD-002018", "Table 4-9")),
        (13, "INTRAFREQHOGROUP", "IntraFreqHoA3Offset", "—", "Off in intra-frequency A3 formula.", "Tune", "4.1.4",
         "Offset Off added on the serving side of intra-frequency event A3.", R("LBFD-002018", "§4.1.4")),
        (14, "INTERFREQHOGROUP", "InterFreqHoA3Offset", "—", "Off in inter-frequency A3 formula.", "Tune", "4.1.4",
         "Offset Off added on the serving side of inter-frequency event A3.", R("LBFD-002018", "§4.1.4")),
        (15, "EUTRANINTERFREQNCELL", "CellIndividualOffset", "—", "Ocn. Large CIO can mask RF overshoot.", "Tune", "4.1.4",
         "Cell-specific offset Ocn (CIO) for the neighbouring cell in A3/A4/A5.", R("LBFD-002018", "§4.1.4")),
        (16, "EUTRANINTERNFREQ", "QoffsetFreq", "—", "Ofn / connected frequency offset.", "Tune", "4.1.4",
         "Frequency-specific offset Ofn for the neighbouring frequency.", R("LBFD-002018", "§4.1.4")),
        (17, "CELLHOPARACFG", "EutranFilterCoeffRsrp", "—", "L3 RSRP filter. Over-smooth delays coverage rescue.", "Tune", "4.1.4",
         "Layer-3 filter coefficient applied to RSRP before event evaluation.", R("LBFD-002018", "§4.1.4")),
        (18, "CELLALGOSWITCH", "MEAS_OBJ_PREEMPT_SW", "as designed", "Higher-priority algorithm can take limited UE gap capability.", "Tune", "Table 4-4",
         "Allows a higher-priority measurement algorithm to preempt gap resources of a limited-capability UE.", R("LBFD-002018", "Table 4-4")),
        (19, "ENODEBALGOSWITCH", "AutoGapSwitch", "as designed", "Gap pattern cost on DL TTI / VoLTE.", "Tune", "4.1.4",
         "Controls automatic delivery of measurement gap patterns for inter-frequency / IRAT measurement.", R("LBFD-002018", "§4.1.4")),
        (20, "CELLQCIPARA", "QciPriorityForHo", "see default map", "FDD default: QCI5=1, QCI1=2, QCI3=3, QCI2=4, QCI4=5, QCI6=6, QCI7=7, QCI8=8, QCI9=9. Smaller = higher priority.", "Tune", "Table 4-10",
         "Priority used when several QCIs run together so the eNodeB knows which QCI’s HO parameters to deliver.", R("LBFD-002018", "Table 4-10")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_coverage(wb):
    d = start(
        wb, "c", "5", "Coverage-based Handover",
        "Document page 97. §5.1 Common for sub-group is shared by 5.2–5.6. Then each sub-group. Necessary handover. Always configure Chapter 4 first.",
    )
    d.h1("Introduction — sub-groups first")
    d.info_box("Types by target  ·  every type shares §5.1 Common for sub-group", [
        "Intra-frequency handover — measurement-based only   ·   LBFD-00201801   ·   §5.2",
        "Inter-frequency handover — measurement-based + preferential blind + emergency blind   ·   LBFD-00201802   ·   §5.3   ·   also LBFD-131111 FDD↔TDD",
        "E-UTRAN to UTRAN handover — measurement-based + blind   ·   LOFD-001019   ·   §5.4",
        "E-UTRAN to GERAN handover — measurement-based + blind   ·   LOFD-001020   ·   §5.5",
        "E-UTRAN to UTRAN CS/PS steering — which UTRAN layer after IRAT A2   ·   LOFD-001078   ·   §5.6",
    ], "A coverage-based handover is triggered when a UE moves to the cell edge. The book classifies handovers by target.")

    d.h1("5.1  Common for sub-group")
    d.para("Read this before 5.2–5.6. It is the document §5.1 (page 97). Shared by every coverage sub-group. Each item is in its own box.")

    d.info_box("5.1.1  Introduction to handover functions", [
        "Coverage HO starts at the cell edge. The serving cell can no longer carry the UE. It is a necessary handover and preempts Chapters 6–12.",
        "Intra-frequency: the UE already measures the same frequency. The eNodeB delivers event A3 after RRC. No A2 is required.",
        "Inter-frequency and IRAT: the UE cannot measure other frequencies / RATs all the time. Event A2 (serving poor) starts measurement. Event A1 (serving good) stops it.",
        "Blind is used when measurement is not needed (preferential) or not possible in time (emergency).",
        "FDD↔TDD is treated as inter-frequency (LBFD-131111). Same engine as §5.3.",
    ])

    d.pair_boxes(
        ("5.1.2  Measurement-based handover functions", [
            "Intra-frequency (§5.2): measurement-based only. No blind. No initiation-decision phase. Handover runs when any neighbour meets A3.",
            "Inter-frequency (§5.3): A2 starts gap-assisted measurement. The target event is A3, A4 or A5 from EUTRANINTERNFREQ.InterFreqHoEventType. A1 stops measurement if serving recovers.",
            "E-UTRAN to UTRAN (§5.4): IRAT A2 starts. Target B1 (neighbour absolutely good) or B2 (serving poor AND neighbour good). Measurement path: UtranPsHoSwitch. A1 stops.",
            "E-UTRAN to GERAN (§5.5): same A2/A1 idea unless a GERAN A2 offset is set. Target B1/B2 on GERAN. Typical path: GeranRedirectSwitch.",
            "The UE reports the event. The eNodeB then picks the target, admits it, and executes (Chapter 4 engine).",
        ]),
        ("5.1.3  Blind handover functions", [
            "Preferential blind (blind handover). Switch: CELLHOPARACFG IfCoverPreBlindHoSwitch. Use only if the neighbouring cell / frequency fully contains the source. The measurement A2 starts handover without waiting for A3/A4/A5. If A1 arrives before completion, blind HO stops.",
            "Emergency / emergent blind (blind redirection). Switch: CELLHOPARACFG EmcInterFreqBlindHoSwitch. After access the eNodeB checks for a suitable blind neighbour or frequency and delivers a separate blind A2 (CellHoParaCfg.BlindHoA1A2ThdRsrp / Rsrq). When that A2 is reported, the eNodeB redirects the UE to avoid drop. Blind A1 stops it.",
            "If measurement A2 threshold ≤ blind A2, the eNodeB delivers only blind A2.",
            "If no blind neighbouring cell or connected-mode frequency priority is configured, blind A2 is not delivered.",
            "QCI-1: VolteRedirectSwitch allows blind redirect for voice. Release cause to the MME is User Inactivity. Blind IRAT is not the normal path for QCI-1.",
        ], "The book has two blind types. Both apply to inter-frequency (§5.3) and, with the IRAT switches, to UTRAN/GERAN."),
    )

    d.pair_boxes(
        ("5.1.4  Event A2 involved in coverage-based handover", [
            "A3-based inter-frequency A2: A3InterFreqHoA2ThdRsrp / Rsrq. Add operator / QCI / SPID offset, then clamp. Used when InterFreqHoEventType = EventA3.",
            "A4/A5-based inter-frequency A2: InterFreqHoA2ThdRSRP / RSRQ. Used when InterFreqHoEventType = EventA4 or EventA5.",
            "IRAT A2: InterRatHoA2ThdRsrp / Rsrq. UTRAN vs GERAN can take a further A2 offset.",
            "Blind A2: CellHoParaCfg.BlindHoA1A2ThdRsrp / Rsrq. Separate from measurement A2.",
            "A1 of the same family stops measurement or blind. Keep A1 a few dB above the paired A2.",
            "A4/A5 target threshold must be a higher RSRP requirement than this coverage A2 (A4_thd > A2_thd), or the UE returns immediately.",
            "ReduceInvalidA1A2RptSigSwitch: deliver A2 first at RRC setup; deliver A1 only after A2, to cut extra signalling.",
        ], "Several A2 families exist. Mixing them is the usual reason coverage IFHO does not start or ping-pongs. InterFreqHoEventType picks the family."),
        ("5.1.5  Principles for selecting UTRAN or GERAN", [
            "IRAT A2 still starts the function. This section only chooses which RAT is measured / used after A2.",
            "SrvccRatSteeringSwitch ON: for QCI-1 / SRVCC, measure only the highest-priority RAT for voice.",
            "PsRatSteeringSwitch ON: for data (non-QCI-1), measure only the highest-priority RAT for PS.",
            "RatLayerSwitch is legacy. Do not use it on this version.",
            "With the same RF, a smaller B1 TimeToTrig toward UTRAN than GERAN makes UTRAN more likely.",
            "CS/PS steering (§5.6) is a further filter on which UTRAN frequency (CsPriority / PsPriority). It is not a different A2.",
            "Coverage IRAT TTT for offload-oriented IRAT must be < 3 s, or the 3 s measurement stop kills the report.",
        ]),
    )

    d.info_box("5.1.6  Other points that belong with this Common block", [
        "Chapter 4 flags, NRT, object cap and SMeasure must already be correct. Coverage cannot start if the frequency is not a measurement object.",
        "Coverage is necessary HO: any QCI may be admitted. Unnecessary features (Ch.6–12) wait.",
        "Book MML examples (A1/A2 −85/−87 dBm, A4 −103 dBm) are command examples, not design values.",
        "Align idle ThrshServLow with this A2/A5 thinking, but idle is a different document.",
    ])
    d.callout("CALC", "A2 families — do not mix them  ·  example not a design", [
        "A3-based IFHO A2:  A3InterFreqHoA2ThdRsrp  +  operator/QCI offset (eNBCnOpQciRsvdPara)   when InterFreqHoEventType = EventA3",
        "A4/A5-based IFHO A2:  InterFreqHoA2ThdRSRP / RSRQ   when InterFreqHoEventType = EventA4 or EventA5",
        "IRAT A2:  InterRatHoA2ThdRsrp / Rsrq. Further split UTRAN vs GERAN if an A2 offset is set per RAT.",
        "Blind A2:  CellHoParaCfg.BlindHoA1A2ThdRsrp / Rsrq",
        "If measurement A2 threshold ≤ blind A2, the eNodeB delivers only blind A2.",
        "Example: A3-based A2 base = −110 dBm, operator/QCI offset = +4 dB → −106 dBm. Clamp MAX(−140, MIN(−43, −106)) = −106 dBm. That −106 dBm is the A2 the UE uses.",
        "Example ping-pong check: coverage A2 = −110 dBm, so A4 must be higher than −110 dBm (for example A4 = −105 dBm). Unsafe: A2 = −100 and A4 = −110 lets a neighbour at −107 dBm HO, then A2 fires on the new cell (−107 + 2 = −105 < −100).",
    ])

    d.h2("5.2  Coverage-based intra-frequency handover   ·   LBFD-00201801")
    d.info_box("Principle", [
        "Turn ENODEBALGOSWITCH HoAlgoSwitch IntraFreqCoverHoSwitch = ON.",
        "After RRC setup the eNodeB delivers intra-frequency A3. There is no initiation-decision phase.",
        "Handover runs when any neighbour meets A3 for IntraFreqHoA3TimeToTrig.",
        "A3 enter: Mn + Ofn + Ocn − Hys > Ms + Ofs + Ocs + IntraFreqHoA3Offset.",
        "Trigger quantity: IntraFreqHoA3TrigQuan = RSRP (default).",
        "This does not move the UE to another band. Chapter 4 NRT and CIO must be valid. License: none.",
    ], "Overview of this sub-group: measurement-based only. No blind. No A2.")
    d.callout("CORE", "Core setting", "ENODEBALGOSWITCH HoAlgoSwitch IntraFreqCoverHoSwitch = ON. IntraFreqHoA3TrigQuan = RSRP (default).")
    d.callout("CALC", "Calculation + example (not a design)", [
        "A3 enter: Mn + Ofn + Ocn − Hys > Ms + Ofs + Ocs + IntraFreqHoA3Offset, true for IntraFreqHoA3TimeToTrig.",
        "Example: Ms = −95 dBm, Mn = −90 dBm, Hys = 2 dB, all offsets 0 except Off = 2 dB.",
        "Left = −90 − 2 = −92.  Right = −95 + 2 = −93.  −92 > −93 → ENTER. Neighbour is 3 dB stronger, offset asks for 2 dB, hysteresis 2 dB, so A3 just passes.",
        "If IntraFreqHoA3Offset = 6 dB: Right = −89.  −92 > −89 → NO. Increase Off to stop early intra-frequency HO.",
    ])
    d.two_col(
        "Advantage",
        ["Cuts intra-frequency interference and drop on a contiguous layer.", "No extra license."],
        "Limitation",
        ["TDD massive-MIMO beamforming can inflate neighbour RSRP and cause early A3.", "Does not move the UE to another band."],
    )
    d.callout("CONDITION", "Conditions", "Chapter 4 NRT and CIO must be valid. License: none.")

    d.h2("5.3  Coverage-based inter-frequency handover   ·   LBFD-00201802")
    d.info_box("Overview of types in §5.3", [
        "Measurement-based inter-frequency HO. A2 starts gap measurement. Target A3 / A4 / A5 from InterFreqHoEventType. A1 stops if serving recovers. Switch: CELLHOPARACFG InterFreqCoverHoSwitch.",
        "Preferential blind (blind handover). Same measurement A2, but the eNodeB hands over without A3/A4/A5. Use only if the target fully contains the source. Switch: IfCoverPreBlindHoSwitch.",
        "Emergency blind (blind redirection). Separate worse A2 (BlindHoA1A2Thd). Policy is redirection when quality is already too bad to finish measurement. Switch: EmcInterFreqBlindHoSwitch.",
        "Event A2 involved: do not mix A3-based A2 with A4/A5-based A2. See §5.1.4 in Common for sub-group.",
        "FDD↔TDD (LBFD-131111) is the same inter-frequency engine. UTRAN/GERAN selection principles in §5.1.5 apply only when IRAT is also on — they do not replace 5.3.",
    ], "This sub-group uses the shared Common for sub-group (§5.1). The three types inside 5.3 are listed first.")
    d.info_box("Principle", [
        "Configure Chapter 4 first (FREQ_MEAS_FLAG, NoHoFlag, object cap, SMeasure).",
        "Turn CELLHOPARACFG CellHoAlgoSwitch InterFreqCoverHoSwitch = ON.",
        "Set EUTRANINTERNFREQ InterFreqHoEventType = EventA3 or EventA4 or EventA5. This picks both the target event and the A2 family.",
        "Start: A2  Ms + Hys < A2_thd  for TimeToTrig. Stop: A1  Ms − Hys > A1_thd.",
        "Target A3 = neighbour relatively better. Target A4 = neighbour absolutely good. Target A5 = serving poor AND neighbour good.",
        "If IfCoverPreBlindHoSwitch = ON and the neighbour fully contains the source: measurement A2 starts preferential blind HO instead of measurement.",
        "If EmcInterFreqBlindHoSwitch = ON: a worse blind A2 starts emergency redirection. If measurement A2 ≤ blind A2, only blind A2 is delivered.",
        "Optional: ReduceInvalidA1A2RptSigSwitch = ON so A2 is delivered first and A1 only after A2.",
        "Keep A4/A5 a higher RSRP requirement than this coverage A2, or the UE ping-pongs back.",
        "FDD↔TDD uses this same procedure (LBFD-131111).",
    ])
    d.callout("CORE", "Core setting", [
        "CELLHOPARACFG CellHoAlgoSwitch InterFreqCoverHoSwitch = ON.",
        "IfCoverPreBlindHoSwitch = ON only if the target fully contains the source.",
        "EmcInterFreqBlindHoSwitch = ON for emergency redirection.",
        "EUTRANINTERNFREQ InterFreqHoEventType = EventA3 or EventA4 or EventA5 (this picks the A2 family).",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Start meas: A2  Ms + Hys < A2_thd  for TTT. Stop: A1  Ms − Hys > A1_thd.",
        "Target A3: relative. Target A4: Mn+Ofn+Ocn−Hys > InterFreqHoA4ThdRSRP. Target A5: serving < Th1 AND neighbour > Th2.",
        "Preferential blind: same A2 as measurement IFHO, but HO instead of meas. Emergency blind: worse A2, policy = redirection.",
        "Example A2 start: A2_thd = −110 dBm, Hys = 2 dB, Ms = −115 dBm.  −115 + 2 = −113 < −110 → start inter-frequency measurement.",
        "Example A1 stop: A1_thd = −104 dBm (a few dB above A2 −110), Ms = −95 dBm.  −95 − 2 = −97 > −104 → stop measurement, serving recovered. If Ms is still −108: −108 − 2 = −110 > −104? NO — measurement continues.",
        "Example A4 target: A4 = −105 dBm, Mn = −90 dBm, Hys = 2.  −92 > −105 → HO. Keep A4 (−105) higher than A2 (−110) so a just-good target is not immediately A2-poor.",
    ])
    d.two_col(
        "Advantage",
        ["Rescues the UE onto another LTE layer.", "Blind covers UEs that cannot measure in time."],
        "Limitation",
        ["Wrong A2 family = wrong HO.", "Blind has higher access fail.", "Gaps steal DL TTIs."],
    )
    d.callout("CONDITION", "Conditions", "Ch.4 flags and object cap. A4/A5 thd better than this A2. Book MML examples −85/−87 dBm are not design values.")

    d.h2("5.4  Coverage-based inter-RAT handover to UTRAN   ·   LOFD-001019")
    d.info_box("Principle", [
        "Start: IRAT A2. Stop: A1 of the same family.",
        "Target: B1 (Mn + Ofn − Hys > B1_thd) or B2 (serving A2-style Th1 AND neighbour B1-style Th2).",
        "Measurement HO: UtranPsHoSwitch. Blind / redirect: UtranRedirectSwitch.",
        "Quantity: InterRatHoA1A2TrigQuan = RSRP recommended.",
        "Apply §5.1.5 if both UTRAN and GERAN are possible (SrvccRatSteeringSwitch / PsRatSteeringSwitch).",
        "IRAT offload TTT must be < 3 s or the 3 s meas stop kills the report.",
    ], "Overview of this sub-group: measurement-based (B1/B2) and blind redirect. Shared A2 rules: §5.1.4. RAT pick: §5.1.5.")
    d.callout("CORE", "Core setting", "UtranPsHoSwitch and/or UtranRedirectSwitch = ON. InterRatHoA1A2TrigQuan = RSRP recommended.")
    d.callout("CALC", "Calculation + example (not a design)", [
        "B1: Mn + Ofn − Hys > B1_thd. Example: IRAT Mn = −92, Ofn = 0, Hys = 2, B1_thd = −100.  −94 > −100 → ENTER.",
        "B2: serving A2-style Th1 AND B1-style Th2. Example Th1 = −110, Th2 = −100, Ms = −115, Mn = −92, Hys = 2.",
        "     Serving: −115 + 2 = −113 < −110 YES. Neighbour: −94 > −100 YES → ENTER. If Ms = −95, serving −93 < −110? NO — B2 does not enter.",
        "IRAT offload TTT must be < 3 s or the 3 s meas stop kills the report.",
    ])
    d.two_col("Advantage", ["Last rescue toward 3G."], "Limitation", ["IRAT TTT > 3 s blocks offload-oriented IRAT because meas is stopped at 3 s."])

    d.h2("5.5  Coverage-based inter-RAT handover to GERAN   ·   LOFD-001020")
    d.info_box("Principle", [
        "Turn GeranRedirectSwitch = ON.",
        "Start on IRAT A2. Stop on A1. Target B1/B2 on GERAN.",
        "A smaller B1 TTT toward UTRAN than GERAN makes UTRAN more likely given the same RF (§5.1.5).",
        "Blind IRAT is not the normal path for QCI-1.",
    ], "Overview of this sub-group: measurement-based (B1/B2) and redirect. Same A2/A1 pair as UTRAN unless a GERAN A2 offset is set.")
    d.callout("CORE", "Core setting", "GeranRedirectSwitch = ON. Smaller B1 TTT toward UTRAN than GERAN makes UTRAN more likely given the same RF.")
    d.two_col("Advantage", ["GSM as last coverage."], "Limitation", ["Blind IRAT is not for QCI-1."])

    d.h2("5.6  Coverage-based E-UTRAN to UTRAN CS/PS steering   ·   LOFD-001078")
    d.info_box("Principle", [
        "Requires §5.4 coverage IRAT to UTRAN first.",
        "Turn CELLALGOSWITCH FreqLayerSwitch UtranFreqLayerMeasSwitch and/or UtranFreqLayerBlindSwitch = ON.",
        "Set UTRANNFREQ CsPriority / PsPriority. Priority_0 = do not use that frequency for that service.",
        "Voice vs data RAT pick is still §5.1.5 (SrvccRatSteeringSwitch / PsRatSteeringSwitch). RatLayerSwitch is legacy.",
    ], "Overview of this sub-group: not a new A2. After IRAT A2, only the UTRAN frequency with the highest CS or PS priority is measured / used.")
    d.callout("CORE", "Core setting", "CELLALGOSWITCH FreqLayerSwitch UtranFreqLayerMeasSwitch and/or UtranFreqLayerBlindSwitch = ON. UTRANNFREQ CsPriority / PsPriority set. Priority_0 = do not use that frequency for that service.")
    d.callout("CONDITION", "Conditions", "Needs §5.4 coverage IRAT to UTRAN. Voice vs data RAT steering: SrvccRatSteeringSwitch / PsRatSteeringSwitch. RatLayerSwitch is legacy — not recommended.")

    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Read §5.1 Common for sub-group first. It supports every sub-group.",
        "Intra-frequency A3 is always-on coverage. No A2.",
        "Inter-frequency: pick one A2 family from the event type, then A3 or A4/A5, optional preferential / emergency blind.",
        "IRAT: separate A2, then B1/B2. §5.1.5 picks UTRAN vs GERAN. §5.6 only chooses which UTRAN layer.",
        "Coverage is necessary: it preempts Ch.6–12.",
    ])

    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "ENODEBALGOSWITCH", "IntraFreqCoverHoSwitch", "ON", "§5.2 activation.", "Yes", "5.2 Intra-freq",
         "Enables coverage-based intra-frequency handover (event A3).", R("LBFD-00201801", "§5.2")),
        (2, "INTRARATHOCOMM", "IntraFreqHoA3TrigQuan", "RSRP", "Default. A3 trigger quantity.", "Tune", "5.2 Intra-freq",
         "Measurement quantity that triggers intra-frequency event A3.", R("LBFD-00201801", "§5.2.4")),
        (3, "INTRAFREQHOGROUP", "IntraFreqHoA3Offset", "—", "Off in intra-frequency A3.", "Tune", "5.2 Intra-freq",
         "Offset Off in the intra-frequency A3 formula (Mn+Ofn+Ocn−Hys > Ms+Ofs+Ocs+Off).", R("LBFD-00201801", "§5.2 / §4.1.4")),
        (4, "CELLHOPARACFG", "InterFreqCoverHoSwitch", "ON", "§5.3 activation.", "Yes", "5.3 Inter-freq",
         "Enables coverage-based inter-frequency handover (measurement-based and related blind options).", R("LBFD-00201802", "§5.3")),
        (5, "CELLHOPARACFG", "IfCoverPreBlindHoSwitch", "ON if contained", "Preferential blind. Target must fully cover source.", "Tune", "5.3 Inter-freq",
         "When ON, an inter-frequency A2 starts preferential blind HO instead of measurement-based IFHO.", R("LBFD-00201802", "§5.3.1")),
        (6, "CELLHOPARACFG", "EmcInterFreqBlindHoSwitch", "ON if needed", "Emergency inter-frequency redirection.", "Tune", "5.3 Inter-freq",
         "Enables emergency blind redirection when serving quality is too poor to complete measurement.", R("LBFD-00201802", "§5.3.1")),
        (7, "EUTRANINTERNFREQ", "InterFreqHoEventType", "EventA3 or A4 or A5", "Selects which A2 family and which target event.", "Yes", "5.3 Inter-freq",
         "Chooses the target event (A3 / A4 / A5) and therefore which coverage A2 family is delivered.", R("LBFD-00201802", "§5.1.4 / §5.3")),
        (8, "INTERFREQHOGROUP", "A3InterFreqHoA2ThdRsrp", "—", "A2 when event type is A3. Add operator/QCI offset then clamp.", "Tune", "5.3 Inter-freq",
         "Event A2 RSRP threshold used when the inter-frequency target event is A3. Effective thd = this + operator/QCI offset.", R("LBFD-00201802", "Table 5-10")),
        (9, "INTERFREQHOGROUP", "InterFreqHoA2ThdRSRP", "—", "A2 when event type is A4 or A5. Book examples are not design values.", "Tune", "5.3 Inter-freq",
         "Event A2 RSRP threshold used when the inter-frequency target event is A4 or A5.", R("LBFD-00201802", "Table 5-10")),
        (10, "INTERFREQHOGROUP", "InterFreqHoA4ThdRSRP", "—", "A4 target. Must be better than coverage A2.", "Tune", "5.3 Inter-freq",
         "Absolute neighbour RSRP threshold for coverage inter-frequency event A4.", R("LBFD-00201802", "§5.3")),
        (11, "CELLHOPARACFG", "BlindHoA1A2ThdRsrp", "—", "Blind A2. If meas A2 ≤ this, only blind A2 is delivered.", "Tune", "5.3 Blind",
         "Event A2 RSRP threshold that starts coverage blind inter-frequency / IRAT handling.", R("LBFD-00201802", "Table 5-3")),
        (12, "CELLALGOSWITCH", "ReduceInvalidA1A2RptSigSwitch", "ON recommended", "Deliver A2 first, A1 after A2.", "Tune", "5.3 Inter-freq",
         "Delivers A2 measurement first at RRC setup and A1 only after A2, to reduce extra A1/A2 signalling.", R("LBFD-00201802", "§5.3.1")),
        (13, "ENODEBALGOSWITCH", "UtranPsHoSwitch", "ON", "§5.4 measurement HO to UTRAN.", "Yes", "5.4 UTRAN",
         "Enables measurement-based PS handover from E-UTRAN to UTRAN.", R("LOFD-001019", "§5.4")),
        (14, "ENODEBALGOSWITCH", "UtranRedirectSwitch", "ON", "§5.4 blind / redirect to UTRAN.", "Yes", "5.4 UTRAN",
         "Enables redirection (blind path) from E-UTRAN to UTRAN.", R("LOFD-001019", "§5.4")),
        (15, "INTERRATHOCOMM", "InterRatHoA1A2TrigQuan", "RSRP", "IRAT A1/A2 quantity.", "Tune", "5.4 / 5.5",
         "Quantity used to trigger inter-RAT events A1 and A2.", R("LOFD-001019", "§5.1.4")),
        (16, "INTERRATHOCOMMGROUP", "InterRatHoA2ThdRsrp", "—", "IRAT coverage A2.", "Tune", "5.4 / 5.5",
         "Serving-cell A2 RSRP threshold that starts coverage inter-RAT measurement.", R("LOFD-001019", "Table 5-3")),
        (17, "ENODEBALGOSWITCH", "GeranRedirectSwitch", "ON", "§5.5 GERAN.", "Yes", "5.5 GERAN",
         "Enables coverage-based handover / redirection from E-UTRAN to GERAN.", R("LOFD-001020", "§5.5")),
        (18, "CELLALGOSWITCH", "UtranFreqLayerMeasSwitch", "ON", "§5.6 CS/PS steering measurement.", "Yes", "5.6 CS/PS",
         "Enables measurement-based CS/PS frequency-layer steering toward UTRAN.", R("LOFD-001078", "§5.6")),
        (19, "CELLALGOSWITCH", "UtranFreqLayerBlindSwitch", "ON", "§5.6 CS/PS steering blind.", "Yes", "5.6 CS/PS",
         "Enables blind CS/PS frequency-layer steering toward UTRAN.", R("LOFD-001078", "§5.6")),
        (20, "UTRANNFREQ", "CsPriority / PsPriority", "Priority_16 or plan", "Priority_0 excludes that frequency from that service.", "Tune", "5.6 CS/PS",
         "Priority of a neighbouring UTRAN frequency for CS services versus PS services. Priority_0 = not used for that service.", R("LOFD-001078", "§5.6")),
        (21, "ENODEBALGOSWITCH", "SrvccRatSteeringSwitch", "ON if SRVCC", "Voice RAT pick after IRAT A2. RatLayerSwitch is legacy.", "Tune", "5.1.5",
         "After IRAT A2, restricts measurement to the highest-priority RAT for QCI-1 / SRVCC.", R("LOFD-001019", "Table 5-4")),
        (22, "ENODEBALGOSWITCH", "PsRatSteeringSwitch", "ON if data IRAT", "Data RAT pick after IRAT A2.", "Tune", "5.1.5",
         "After IRAT A2, restricts measurement to the highest-priority RAT for data (non-QCI-1) services.", R("LOFD-001019", "Table 5-4")),
    ]
    for row in rows:
        d.param_row(*row)
    return d

def sheet_service(wb):
    d = start(
        wb, "s", "6", "Service-based Handover",
        "Document page 189. Unnecessary HO. Sub-groups 6.2–6.4. Uses Chapter 4 A4 / B1. Requires Chapter 4; coverage A2 must not pull the UE back.",
    )
    d.h1("Introduction — sub-groups first")
    d.info_box("Types by target  ·  every type shares this Common for sub-group", [
        "Inter-frequency (including FDD↔TDD)   ·   LBFD-00201805 and LOFD-171207 enhancement   ·   §6.2",
        "E-UTRAN to UTRAN   ·   LOFD-001043   ·   §6.3",
        "E-UTRAN to GERAN   ·   LOFD-001046   ·   §6.4",
        "Measurement-based only. Not coverage rescue. Not blind.",
    ], "This function allows services with different QCIs to be carried by different frequencies.")
    d.h1("6.1  Common for sub-group")
    d.pair_boxes(
        ("What is common", [
            "Start when the highest-priority QCI on the UE is allowed to leave the serving frequency (PERMIT_HO / MUST_HO) and the serving EARFCN is not in that QCI’s target group.",
            "Stop if no A4/B1 for 3 s.",
            "Bind every QCI you intend to move (CNOPERATORQCIPARA → group).",
            "Chapter 4 object cap and A4 TTT ≠ 5120 ms.",
            "A4 threshold must be a higher RSRP requirement than coverage A2, or coverage measurement starts immediately after arrival.",
        ]),
        ("LTE vs IRAT", [
            "§6.2 LTE path = event A4. Two switches: eNodeB ServiceBasedInterFreqHoSwitch AND cell SrvBasedInterFreqHoSw.",
            "FDD: ServiceIfHoCfgGroup + ServiceIfDlEarfcnGrp. Multi-freq ON: several EARFCNs in priority order; OFF: only index 0.",
            "§6.3 UTRAN path = event B1. UtranServiceHoSwitch. InterRatHoState = MUST_HO or PERMIT_HO.",
            "§6.4 GERAN path = event B1. GeranServiceHoSwitch. Same SERVICEIRHOCFGGROUP bind.",
            "TDD: mutually exclusive with service-request HO. FDD: not exclusive the same way.",
            "If MLB HoAdmitSwitch is ON in a cell already in MLB, service HO prep can fail. Document recommends it deselected.",
        ]),
    )
    d.h1("Principle")
    d.info_box("Principle", [
        "Service HO is QCI steering, not edge rescue.",
        "LTE: A4. IRAT: B1. IRAT offload TTT must be < 3 s.",
        "Unnecessary HO: target must admit all QCIs.",
        "Keep A4 above coverage A2. Do not combine with 5120 ms A4 TTT.",
    ])

    d.h2("6.2  Service-based inter-frequency handover")
    d.info_box("Principle", [
        "Target event A4. Threshold = InterFreqLoadBasedHoA4ThdRSRP + QCI/operator offset.",
        "Enter: Mn + Ofn + Ocn − Hys > that threshold, for InterFreqHoA4TimeToTrig.",
        "HoAlgoSwitch ServiceBasedInterFreqHoSwitch = ON AND CELLALGOSWITCH SrvBasedInterFreqHoSw = ON.",
        "CNOPERATORQCIPARA binds QCI to ServiceIfHoCfgGroup. InterFreqHoState = PERMIT_HO.",
        "ServiceBasedMultiFreqHoSwitch ON: CA-incapable UEs pick high priority + high BW + light load. CA-capable UEs are not service-HO’d.",
        "ServBasedHoBackSwitch allows return to the source frequency on the next service HO.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Target event A4. Threshold = InterFreqLoadBasedHoA4ThdRSRP + QCI/operator offset (same A4 pool as FreqPri in many versions).",
        "Enter: Mn + Ofn + Ocn − Hys > that threshold, for InterFreqHoA4TimeToTrig.",
        "Example: base A4 = −105 dBm, QCI offset = +3 dB → used A4 = −102 dBm. Mn = −90, Hys = 2.  −92 > −102 → ENTER.",
        "If coverage A2 is −110 dBm, used A4 −102 dBm is higher than A2, so the UE is not coverage-measured out of the new cell at once.",
        "ServiceBasedMultiFreqHoSwitch ON: for CA-incapable UEs, pick high priority + high BW + light load. CA-capable UEs are then not service-HO’d (use CA smart selection instead).",
        "Light-load example: serving MlbUeNumThd = 40, offset = 5. Neighbour UL-sync UEs = 30.  30 < 40+5 → treated as light load.",
        "ServBasedHoBackSwitch allows return to the source frequency on the next service HO.",
    ])
    d.two_col(
        "Advantage",
        ["Puts a QCI on the band designed for it.", "FDD can steer one QCI to several EARFCNs (enhancement LOFD-171207)."],
        "Limitation",
        ["Unnecessary HO: target must admit all QCIs.", "Ping-pong if A4 is not better than coverage A2.", "MLB admit conflict if HoAdmitSwitch is ON."],
    )

    d.h2("6.3  Service-based IRAT to UTRAN")
    d.info_box("Principle", [
        "HoAlgoSwitch UtranServiceHoSwitch = ON.",
        "SERVICEIRHOCFGGROUP InterRatHoState = MUST_HO or PERMIT_HO. Bind QCI on CNOPERATORQCIPARA.",
        "Target B1. IRAT offload TTT must be < 3 s or the 3 s meas stop kills the report.",
        "Coverage IRAT A2 still owns true edge. Do not use this as coverage rescue.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Target B1. Example: IRAT Mn = −92, Hys = 2, B1_thd = −100.  −94 > −100 → ENTER.",
        "IRAT offload TTT must be < 3 s or the 3 s meas stop kills the report.",
    ])
    d.two_col("Advantage", ["Force a QCI (often voice) to 3G."], "Limitation", ["Coverage IRAT A2 still owns true edge. Do not use this as coverage rescue."])

    d.h2("6.4  Service-based IRAT to GERAN")
    d.info_box("Principle", [
        "HoAlgoSwitch GeranServiceHoSwitch = ON.",
        "Same SERVICEIRHOCFGGROUP bind as 6.3.",
        "Same 3 s IRAT meas stop. Not a coverage substitute.",
    ])
    d.two_col("Advantage", ["GSM for a chosen QCI."], "Limitation", ["Same 3 s IRAT meas stop. Not a coverage substitute."])

    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Service HO is QCI steering, not edge rescue.",
        "LTE path = A4. IRAT path = B1.",
        "Two switches on LTE (eNodeB + cell). Bind every QCI you intend to move.",
        "Keep A4 above coverage A2. Do not combine with 5120 ms A4 TTT.",
    ])

    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "ENODEBALGOSWITCH", "ServiceBasedInterFreqHoSwitch", "ON", "eNodeB master for §6.2.", "Yes", "6.2 Inter-freq",
         "eNodeB-level switch that enables service-based inter-frequency handover.", R("LBFD-00201805", "§6.2")),
        (2, "CELLALGOSWITCH", "SrvBasedInterFreqHoSw", "ON", "Cell allow for §6.2. Both this and row 1 are required.", "Yes", "6.2 Inter-freq",
         "Cell-level switch that allows service-based inter-frequency handover in this cell. Both eNodeB and cell bits are required.", R("LBFD-00201805", "§6.2")),
        (3, "SERVICEIFHOCFGGROUP", "InterFreqHoState", "PERMIT_HO", "QCI is allowed to leave serving frequency.", "Yes", "6.2 Inter-freq",
         "Whether a QCI bound to this group is permitted to leave the serving frequency by service-based HO.", R("LBFD-00201805", "§6.2")),
        (4, "CNOPERATORQCIPARA", "ServiceIfHoCfgGroupId", "group id", "Binds QCI to the group.", "Yes", "6.2 Inter-freq",
         "Links a QCI of an operator to a ServiceIfHoCfgGroup (target-frequency policy).", R("LBFD-00201805", "§6.2")),
        (5, "SERVICEIFDLEARFCNGRP", "DlEarfcn", "EARFCN", "Target frequency. Index 0 is the only target if FDD multi-freq is OFF.", "Yes", "6.2 Inter-freq",
         "Downlink EARFCN that may carry the QCI after service-based inter-frequency HO.", R("LBFD-00201805", "§6.2 / LOFD-171207")),
        (6, "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRSRP", "—", "A4 base. Add QCI/operator offset. Must be better than coverage A2.", "Tune", "6.2 Inter-freq",
         "Base A4 RSRP threshold for service-based IFHO. Effective thd = this + QCI/operator offset.", R("LBFD-00201805", "Table 6-2")),
        (7, "INTERFREQHOGROUP", "InterFreqHoA4TimeToTrig", "not 5120 ms", "5120 ms disables this function.", "Yes", "6.2 Inter-freq",
         "TimeToTrig for the A4 used by service-based IFHO. 5120 ms disables the function.", R("LBFD-00201805", "Table 4-9 / §6.2")),
        (8, "CELLALGOSWITCH", "ServiceBasedMultiFreqHoSwitch", "as designed", "ON: CA-incapable UEs pick BW/load; CA UEs skip service HO.", "Tune", "6.2 Inter-freq",
         "When ON, CA-incapable UEs prefer high-priority, high-bandwidth, light-load targets; CA-capable UEs are not service-HO’d.", R("LOFD-171207", "§6.2")),
        (9, "CELLALGOSWITCH", "ServBasedHoBackSwitch", "as designed", "Allow next service HO back to the source frequency.", "Tune", "6.2 Inter-freq",
         "Allows a later service-based HO to select the previous source frequency as target.", R("LBFD-00201805", "§6.2")),
        (10, "ENODEBALGOSWITCH", "UtranServiceHoSwitch", "ON", "§6.3.", "Yes", "6.3 UTRAN",
         "Enables service-based inter-RAT handover to UTRAN for bound QCIs.", R("LOFD-001043", "§6.3")),
        (11, "ENODEBALGOSWITCH", "GeranServiceHoSwitch", "ON", "§6.4.", "Yes", "6.4 GERAN",
         "Enables service-based inter-RAT handover to GERAN for bound QCIs.", R("LOFD-001046", "§6.4")),
        (12, "SERVICEIRHOCFGGROUP", "InterRatHoState", "MUST_HO or PERMIT_HO", "MUST_HO forces IRAT for that QCI.", "Yes", "6.3 / 6.4",
         "Inter-RAT policy for the QCI: MUST_HO forces IRAT; PERMIT_HO allows it.", R("LOFD-001043", "§6.3 / §6.4")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_distance(wb):
    d = start(
        wb, "d", "7", "Distance-based Handover",
        "Document page 225. Overshoot / TA. Measurement-based only. Sub-groups 7.2–7.4.",
    )
    d.h1("Introduction — sub-groups first")
    d.info_box("Types by target  ·  every type shares this Common for sub-group", [
        "Inter-frequency   ·   LBFD-00201804   ·   §7.2   ·  DistBasedMeasObjType = EUTRAN",
        "IRAT to UTRAN   ·   LOFD-001072   ·   §7.3   ·  UTRAN",
        "IRAT to GERAN   ·   LOFD-001073   ·   §7.4   ·  GERAN",
        "Measurement-based only. TA starts the HO; A4/B1 still qualifies the target.",
    ], "When a UE stays on an overshooting cell (often a high-power low-band cell), coverage A2 may never fire before the neighbour list ends.")
    d.h1("7.1  Common for sub-group")
    d.info_box("Common for sub-group", [
        "eNodeB monitors distance to all UEs from uplink timing advance (precision about 100 to 150 m).",
        "Start: measured distance > DistBasedHoThd for 10 seconds.",
        "Stop: measured distance ≤ stop threshold for 10 seconds.",
        "Then Chapter 4 A4 (LTE) or B1 (IRAT).",
        "If several distance functions are ON, IF / UTRAN / GERAN meas can all be delivered; the target decision picks one.",
        "A4 must still be a higher RSRP requirement than coverage A2.",
        "Do not wait for coverage A2 on an overshoot lobe.",
    ])
    d.h1("Principle")
    d.info_box("Principle", [
        "CELLALGOSWITCH DistBasedHoSwitch = ON is the master for Chapter 7.",
        "Pick DistBasedMeasObjType: EUTRAN and/or UTRAN and/or GERAN.",
        "Use on overshoot cells (high site, low band into suburban). 10 s both ways.",
        "Still needs a real neighbour on the target layer.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Distance from TA. Precision about 100 to 150 m.",
        "Start: measured distance > DistBasedHoThd for 10 seconds.",
        "Stop: measured distance ≤ stop threshold for 10 seconds.",
        "LTE target A4: InterFreqHoA4ThdRSRP / RSRQ (same A4 as coverage IFHO). Must still be better than coverage A2.",
        "Example: DistBasedHoThd = 3000 m. TA says 3200 m for 10 s → start. Then A4: Mn = −90 dBm, Hys = 2, A4 = −105 dBm.  −92 > −105 → HO to the planned layer.",
        "If TA later falls to 2500 m for 10 s → stop distance measurement. Do not wait for coverage A2 on an overshoot lobe.",
    ])

    d.h2("7.2  Distance-based inter-frequency")
    d.info_box("Principle", [
        "CELLALGOSWITCH DistBasedHoSwitch = ON. DISTBASEDHO DistBasedMeasObjType EUTRAN = selected.",
        "LTE target A4: InterFreqHoA4ThdRSRP / RSRQ (same A4 as coverage IFHO).",
        "Stops drop at the end of an overshoot lobe. Does not wait for A2.",
        "Wrong thd = HO inside the planned cell or too late.",
    ])
    d.two_col(
        "Advantage",
        ["Stops drop at the end of an overshoot lobe.", "Does not wait for A2."],
        "Limitation",
        ["TA granularity 100–150 m.", "Wrong thd = HO inside the planned cell or too late.", "Still needs a real neighbour on the target layer."],
    )

    d.h2("7.3 / 7.4  Distance-based IRAT")
    d.info_box("Principle", [
        "DistBasedHoSwitch = ON. DistBasedMeasObjType UTRAN and/or GERAN.",
        "Target B1.",
        "Chapter 4 IRAT neighbours required.",
        "Coverage IRAT is still the edge rescue; this is overshoot only.",
    ])

    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "TA starts the HO; A4/B1 still qualifies the target.",
        "Use on overshoot cells (high site, low band into suburban).",
        "10 s both ways. A4 must stay above coverage A2.",
    ])

    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "CELLALGOSWITCH", "DistBasedHoSwitch", "ON", "Master for Chapter 7.", "Yes", "7.x All",
         "Master switch that enables distance-based handover (TA-based overshoot control).", R("LBFD-00201804", "§7.1")),
        (2, "DISTBASEDHO", "DistBasedMeasObjType", "EUTRAN", "§7.2 LTE target.", "Yes", "7.2 Inter-freq",
         "Selects E-UTRAN as the measurement / target RAT for distance-based HO.", R("LBFD-00201804", "§7.2")),
        (3, "DISTBASEDHO", "DistBasedMeasObjType", "UTRAN", "§7.3.", "Yes", "7.3 UTRAN",
         "Selects UTRAN as a distance-based HO target RAT.", R("LOFD-001072", "§7.3")),
        (4, "DISTBASEDHO", "DistBasedMeasObjType", "GERAN", "§7.4.", "Yes", "7.4 GERAN",
         "Selects GERAN as a distance-based HO target RAT.", R("LOFD-001073", "§7.4")),
        (5, "DISTBASEDHO", "DistBasedHoThd", "default then RF", "Start when TA-distance exceeds this for 10 s.", "Tune", "7.2",
         "Distance threshold. HO measurement starts when TA-estimated distance exceeds this for 10 seconds.", R("LBFD-00201804", "Table 7-2")),
        (6, "INTERFREQHOGROUP", "InterFreqHoA4ThdRSRP", "—", "Same A4 as coverage IFHO. Must be better than coverage A2.", "Tune", "7.2",
         "A4 RSRP threshold used to qualify the LTE target after distance trigger.", R("LBFD-00201804", "Table 7-3")),
        (7, "INTRARATHOCOMM", "InterFreqHoA4TrigQuan", "RSRP", "Default.", "Tune", "7.2",
         "Quantity used to trigger the A4 that qualifies a distance-based LTE target.", R("LBFD-00201804", "Table 7-5")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_ulq(wb):
    d = start(
        wb, "u", "8", "UL-Quality-based Handover",
        "Document page 245. Uplink MCS + IBLER. Measurement-based and blind. Sub-groups 8.2–8.3.",
    )
    d.h1("Introduction — sub-groups first")
    d.info_box("Types by target  ·  every type shares this Common for sub-group", [
        "Inter-frequency (measurement + blind)   ·   UlQualityInterFreqHoSwitch   ·   §8.2",
        "IRAT to UTRAN or GERAN   ·   UlQualityInterRATHoSwitch   ·   §8.3",
        "UL problem starts the HO; A4/B1 still qualifies DL of the target.",
    ], "UL-quality-based handover decreases drops caused by poor uplink while downlink RSRP may still look acceptable.")
    d.h1("8.1  Common for sub-group")
    d.info_box("Common for sub-group", [
        "Start measurement when uplink MCS is below a threshold AND (actual IBLER − target IBLER) is above a threshold.",
        "Stop when either condition clears.",
        "Target is A4 (LTE) or B1 (IRAT).",
        "If UL gets worse still and no A4 has arrived, blind redirection (same idea as coverage emergency blind).",
        "Add UlBadQualHoA4Offset on top of coverage A4.",
        "Blind only after A4 never comes and UL is worse by another 10% IBLER.",
    ])
    d.h1("Principle")
    d.info_box("Principle", [
        "This catches UL-limited UEs that coverage A2 never sees.",
        "MCS/IBLER is sensitive to scheduler and UL interference.",
        "A4 offset must not dump the UE onto a worse UL cell.",
        "If QCI-1 is blind-redirected, release cause to MME is User Inactivity.",
    ])

    d.h2("8.2  UL-quality inter-frequency")
    d.info_box("Principle", [
        "ENODEBALGOSWITCH HoAlgoSwitch UlQualityInterFreqHoSwitch = ON.",
        "Start meas: UL MCS index < UlBadQualMcsThd AND (actual IBLER − target IBLER) > UlBadQualIblerThd.",
        "Stop meas: MCS ≥ MCS thd OR IBLER gap ≤ IBLER thd.",
        "A4 threshold = InterFreqHoA4ThdRSRP + UlBadQualHoA4Offset.",
        "Blind: MCS < blind MCS thd AND (IBLER gap − 10%) > IBLER thd AND no A4 received.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Start meas:  UL MCS index < UlBadQualMcsThd   AND   (actual IBLER − target IBLER) > UlBadQualIblerThd",
        "Stop meas:  MCS ≥ MCS thd   OR   IBLER gap ≤ IBLER thd",
        "A4 threshold = InterFreqHoA4ThdRSRP + UlBadQualHoA4Offset   (same for RSRQ + offset)",
        "Example start: MCS thd = 6, IBLER thd = 5%. UE MCS = 4, actual IBLER = 12%, target IBLER = 10%.  4 < 6 AND (12−10)=2% > 5%?  2% > 5% is NO, so measurement does not start yet (IBLER not bad enough).",
        "If actual IBLER = 18%: gap = 8% > 5% AND MCS 4 < 6 → START. Then A4: base −105 dBm + UlBadQualHoA4Offset −3 dB → used A4 = −108 dBm. Mn −90, Hys 2.  −92 > −108 → HO.",
        "Blind: MCS < blind MCS thd  AND  (IBLER gap − 10%) > IBLER thd  AND  no A4 received. Example: gap 18%, thd 5%.  (18−10)=8% > 5% AND still no A4 → blind redirect.",
        "If the QCI-1 blind switch is ON, QCI-1 may be blind-redirected; release cause to MME is always User Inactivity.",
    ])
    d.two_col(
        "Advantage",
        ["Catches UL-limited UEs that coverage A2 never sees.", "Blind as last step."],
        "Limitation",
        ["MCS/IBLER sensitive to scheduler and UL interference.", "A4 offset must not dump the UE onto a worse UL cell.", "QCI-1 blind cause is User Inactivity — be aware in traces."],
    )

    d.h2("8.3  UL-quality IRAT")
    d.info_box("Principle", [
        "HoAlgoSwitch UlQualityInterRATHoSwitch = ON.",
        "Target B1 to UTRAN or GERAN.",
        "Chapter 4 IRAT neighbours required.",
        "Same MCS/IBLER start as 8.2.",
    ])

    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "UL problem starts the HO; A4/B1 still qualifies DL of the target.",
        "Add UlBadQualHoA4Offset on top of coverage A4.",
        "Blind only after A4 never comes and UL is worse by another 10% IBLER.",
    ])

    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "ENODEBALGOSWITCH", "UlQualityInterFreqHoSwitch", "ON", "§8.2.", "Yes", "8.2 Inter-freq",
         "Enables UL-quality-based inter-frequency handover (measurement-based and blind).", R("LBFD-002018", "§8.2")),
        (2, "ENODEBALGOSWITCH", "UlQualityInterRATHoSwitch", "ON", "§8.3.", "Yes", "8.3 IRAT",
         "Enables UL-quality-based inter-RAT handover to UTRAN or GERAN.", R("LBFD-002018", "§8.3")),
        (3, "CELLHOPARACFG / UL HO", "UlBadQualMcsThd", "—", "Start when UL MCS is below this. Confirm exact MO in MAE.", "Tune", "8.2",
         "Uplink MCS index threshold. Measurement starts when UL MCS is below this value (together with IBLER gap).", R("LBFD-002018", "Table 8-2")),
        (4, "CELLHOPARACFG / UL HO", "UlBadQualIblerThd", "—", "Start when actual IBLER − target IBLER exceeds this.", "Tune", "8.2",
         "IBLER-gap threshold. Measurement starts when (actual IBLER − target IBLER) exceeds this value.", R("LBFD-002018", "Table 8-2")),
        (5, "INTERFREQHOGROUP", "InterFreqHoA4ThdRSRP", "—", "A4 base.", "Tune", "8.2",
         "Base A4 RSRP threshold for UL-quality IFHO before the UL-quality offset is added.", R("LBFD-002018", "Table 8-3")),
        (6, "INTERFREQHOGROUP", "UlBadQualHoA4Offset", "—", "Added to A4 RSRP and RSRQ for this function only.", "Tune", "8.2",
         "Offset added to the A4 RSRP/RSRQ threshold for UL-quality-based HO only. Effective A4 = base + this offset.", R("LBFD-002018", "Table 8-3")),
        (7, "INTRARATHOCOMM", "InterFreqHoA4TrigQuan", "RSRP", "Default.", "Tune", "8.2",
         "Quantity used to trigger the A4 that qualifies a UL-quality LTE target.", R("LBFD-002018", "Table 8-5")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_cqi(wb):
    d = start(
        wb, "q", "9", "CQI-based Inter-Frequency Handover (FDD)",
        "Document page 273. FDD only. Unnecessary HO. Uses Chapter 4 A4. Not in the FDD §2.3 snap you sent — it is a later eRAN21.1 chapter.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "Measurement-based inter-frequency only. No separate A2 family.",
        "Starter is serving CQI (and usually a persistence timer), not coverage A2.",
        "Target is still event A4 from Chapter 4.",
        "FDD only. Unnecessary HO: all QCIs must be admitted.",
        "Fully off if InterFreqHoA4TimeToTrig = 5120 ms (eRAN21.1 Table 4-9).",
    ], "CQI-based inter-frequency handover starts when downlink CQI of the serving cell is poor even if RSRP has not yet crossed coverage A2.")
    d.h1("Principle")
    d.info_box("Principle", [
        "Enable the CQI-based inter-frequency HO bit on CELLHOPARACFG / CELLALGOSWITCH (confirm the exact bit name in MAE on this eRAN21.1).",
        "Chapter 4 flags, NRT and object cap already correct.",
        "A4 threshold must be a higher RSRP requirement than coverage A2.",
        "Does not replace Ch.5 coverage. Confirm license / bit in MAE.",
        "Do not copy book example A4 dBm into Value.",
    ])
    d.callout("CORE", "Core setting", [
        "Enable the CQI-based inter-frequency HO bit on CELLHOPARACFG / CELLALGOSWITCH (confirm the exact bit name in MAE on this eRAN21.1).",
        "INTERFREQHOGROUP InterFreqHoA4TimeToTrig must not be 5120 ms. In eRAN21.1 Table 4-9 that value disables frequency-priority, CQI and service-based inter-frequency HO.",
        "Chapter 4 flags, NRT and object cap already correct.",
        "A4 threshold better than coverage A2.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Start: serving CQI stays below the CQI threshold for the configured period.",
        "Target A4: Mn + Ofn + Ocn − Hys > A4_thd (+ any CQI-specific A4 offset if the version has one).",
        "Example: CQI thd = 6, UE CQI = 4 for the CQI timer → start. Serving RSRP may still be −85 dBm (coverage A2 not fired).",
        "Then A4 = −105 dBm, Mn = −90, Hys = 2.  −92 > −105 → HO. Keep A4 better than coverage A2.",
        "Stop / ping-pong guard: A4_thd must be better than coverage A2 so the UE is not immediately measured out of the new cell.",
    ])
    d.two_col(
        "Advantage",
        ["Moves UEs whose DL throughput is already bad while RSRP still looks ‘in coverage’.", "Re-uses the A4 engine — no second event family."],
        "Limitation",
        ["FDD only.", "CQI follows MCS/scheduler; a too-low CQI thd causes extra IFHO.", "Fully off if A4 TTT = 5120 ms.", "Unnecessary HO: all-QCI admit."],
    )
    d.callout("CONDITION", "Conditions", "Requires Ch.4. Does not replace Ch.5 coverage. Confirm license / bit in MAE. Do not copy book example A4 dBm into Value.")
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "CQI starts; A4 qualifies.",
        "Coverage A2 must stay worse than that A4 (A4_thd > A2_thd).",
        "5120 ms A4 TTT means this feature is off.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "CELLHOPARACFG / CELLALGOSWITCH", "CQI-based IFHO switch", "ON", "Confirm exact bit name in MAE on eRAN21.1.", "Yes", "9 CQI",
         "Enables CQI-based inter-frequency handover (FDD). Starts when serving CQI is poor while coverage A2 may not have fired.", R("LBFD-002018", "§9")),
        (2, "INTERFREQHOGROUP", "InterFreqHoA4TimeToTrig", "not 5120 ms", "5120 ms disables CQI IFHO (eRAN21.1 Table 4-9).", "Yes", "9 CQI",
         "TimeToTrig for A4. 5120 ms disables CQI-based IFHO.", R("LBFD-002018", "Table 4-9 / §9")),
        (3, "INTERFREQHOGROUP", "InterFreqHoA4ThdRSRP", "—", "Must be better than coverage A2. Not the book MML example.", "Tune", "9 CQI",
         "Absolute neighbour RSRP threshold of event A4 used to qualify the CQI-based IFHO target.", R("LBFD-002018", "§9")),
        (4, "INTERFREQHOGROUP", "InterFreqHoA4Hyst", "—", "Hys in A4 formula.", "Tune", "9 CQI",
         "Hysteresis Hys in the A4 formula for CQI-based IFHO.", R("LBFD-002018", "Table 4-8 / §9")),
        (5, "INTRARATHOCOMM", "InterFreqHoA4TrigQuan", "RSRP", "Recommended.", "Tune", "9 CQI",
         "Quantity used to trigger event A4 for CQI-based IFHO.", R("LBFD-002018", "§9")),
        (6, "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG", "selected", "From Chapter 4. Still required.", "Yes", "→ Ch.4",
         "Whether this frequency is included in connected-mode measurement. Required for any IFHO including CQI-based.", R("LBFD-002018", "§4.1.4")),
    ]
    for row in rows:
        d.param_row(*row, link_sheet=S["b"] if row[0] == 6 else None)
    return d


def sheet_sreq(wb):
    d = start(
        wb, "r", "10", "Service-Request-based Inter-Frequency Handover",
        "Document page 281. Starts on bearer setup / modify, not on A2. Measurement-based only.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "Measurement-based only. Starts on bearer setup / modify, not on A2.",
        "Different from Ch.6: Ch.6 steers an already-running service; Ch.10 steers at the request.",
        "Target A4. Own threshold SrvReqHoA4ThdRsrp — not mixed with coverage A2.",
        "FDD: cell-level ServiceReqInterFreqHoSwitch (not the eNodeB-level bit).",
        "TDD: mutually exclusive with service-based HO. A4 TTT ≠ 5120 ms.",
    ], "After a setup or modification request, if that QCI is the highest configured priority, the eNodeB may move the UE to the frequency that should carry the service, then the EPC sets up the bearer.")
    d.h1("Principle")
    d.info_box("Principle", [
        "CELLALGOSWITCH HoAllowedSwitch ServiceReqInterFreqHoSwitch = ON.",
        "QCI bind: same ServiceIfHoCfgGroup / ServiceIfDlEarfcnGrp idea as Ch.6, with PERMIT_HO.",
        "A4RptWaitingTimer caps how long gap-assisted meas may run so a missing report does not stall user-plane.",
        "Highest-priority QCI among configured QCIs. Serving frequency not already the target. Chapter 4 flags.",
        "Keep SrvReqHoA4ThdRsrp a higher RSRP requirement than coverage A2.",
    ])
    d.callout("CORE", "Core setting", [
        "CELLALGOSWITCH HoAllowedSwitch ServiceReqInterFreqHoSwitch = ON.",
        "Disable or not collide with Ch.6 on TDD (mutually exclusive in the document).",
        "QCI bind: same ServiceIfHoCfgGroup / ServiceIfDlEarfcnGrp idea as Ch.6, with PERMIT_HO.",
        "A4 TTT ≠ 5120 ms.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "A4 thd = SrvReqHoA4ThdRsrp (+ SrvReqHoA4ThdRsrq if used)  — its own A4, not mixed with coverage A2.",
        "Example: SrvReqHoA4ThdRsrp = −102 dBm, coverage A2 = −110 dBm. Used service-request A4 (−102) is higher than A2 (−110).",
        "Mn = −90, Hys = 2.  −92 > −102 → ENTER. After HO, serving −90: −90 + 2 = −88 < −110? NO, so coverage measurement does not start immediately.",
        "Wait for A4: ServiceIfHoCfgGroup.A4RptWaitingTimer. Example 3 s. If no A4 in 3 s, stop gap measurement so user-plane is not stalled.",
        "VoipExProtSwitch: if ON and VoLTE exception, set up QCI-1 if PERMIT_HO in initial context, then start IF meas after access.",
    ])
    d.two_col(
        "Advantage",
        ["UE lands on the right layer as the service starts.", "FDD can use several target EARFCNs if the multi-freq parameter is ON."],
        "Limitation",
        ["TDD exclusive with service-based HO.", "Gap meas must be time-capped.", "Unnecessary HO admit rule."],
    )
    d.callout("CONDITION", "Conditions", "Highest-priority QCI among configured QCIs. Serving frequency not already the target. Chapter 4 flags.")
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Request → A4 to the service layer → bearer.",
        "Own A4 thd (SrvReqHoA4ThdRsrp). Keep it above coverage A2.",
        "Cell-level switch on FDD.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "CELLALGOSWITCH", "ServiceReqInterFreqHoSwitch", "ON", "FDD: cell-level. Do not use the legacy eNodeB bit on FDD.", "Yes", "10",
         "Cell-level switch that enables service-request-based inter-frequency handover at bearer setup / modify.", R("LBFD-002018", "§10")),
        (2, "SERVICEIFHOCFGGROUP", "InterFreqHoState", "PERMIT_HO", "Allows this QCI to move at request.", "Yes", "10",
         "Whether the QCI is permitted to leave the serving frequency when the service is requested.", R("LBFD-002018", "§10.1.1")),
        (3, "SERVICEIFHOCFGGROUP", "A4RptWaitingTimer", "—", "Max wait for A4 during gap meas.", "Tune", "10",
         "Maximum time the eNodeB waits for an A4 report during gap-assisted service-request measurement.", R("LBFD-002018", "§10.1.2")),
        (4, "INTERFREQHOGROUP", "SrvReqHoA4ThdRsrp", "—", "Own A4. Must be better than coverage A2.", "Tune", "10",
         "Dedicated A4 RSRP threshold for service-request-based IFHO (not mixed with coverage A2).", R("LBFD-002018", "§10.1.3")),
        (5, "INTERFREQHOGROUP", "SrvReqHoA4ThdRsrq", "—", "If RSRQ trigger is used.", "Tune", "10",
         "Dedicated A4 RSRQ threshold for service-request-based IFHO when RSRQ is used.", R("LBFD-002018", "§10.1.3")),
        (6, "INTERFREQHOGROUP", "InterFreqHoA4TimeToTrig", "not 5120 ms", "Shared A4 TTT. 5120 ms disables this class of IFHO.", "Yes", "10",
         "TimeToTrig for A4. 5120 ms disables service-request IFHO together with FreqPri and CQI.", R("LBFD-002018", "Table 4-9 / §10")),
        (7, "CELLALGOSWITCH", "VoipExProtSwitch", "as designed", "VoLTE exception protection path.", "Tune", "10",
         "When ON, after a VoLTE exception the eNodeB may set up QCI-1 if PERMIT_HO and then start IF measurement.", R("LBFD-002018", "§10.1.1")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_freqpri(wb):
    d = start(
        wb, "f", "11", "Frequency-Priority-based Inter-Frequency Handover",
        "Document page 299. Fig 11-1 / 11-2 in eRAN21.1. Place service on high band; keep low band for coverage. Not MLB.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "Measurement-based FreqPri. Optional blind in same-coverage (FreqPriorIFBlindHOSwitch).",
        "FDD↔TDD counts as inter-frequency.",
        "Not MLB. Place service on high band; keep low band for coverage.",
        "Same-coverage multi-band: start when serving is good enough (A1 family), target A4.",
        "Different-coverage: A1 is not used as the stop so the UE can still go high-band.",
        "A4 TTT ≠ 5120 ms.",
    ], "When serving quality is already good, the eNodeB may still move the UE to a higher-priority frequency (typically capacity / high band).")
    d.h1("Principle")
    d.info_box("Principle", [
        "CELLALGOSWITCH FreqPriorityHoSwitch FreqPriorIFHOSwitch = ON.",
        "MlbBasedFreqPriHoSwitch = ON when MLB is on, so heavy load is owned by MLB not FreqPri.",
        "A2BasedFreqPriHoSwitch = OFF in same-coverage multi-band (document recommendation).",
        "EUTRANINTERNFREQ FreqPriBasedHoMeasFlag = ENABLE and MeasPriorityForFreqPriHo set.",
        "Never reverse-pair with MLB on the same frequency (A→B FreqPri vs B→A MLB).",
        "VoipMeasFreqPriSwitch can stop FreqPri meas when voice starts.",
    ])
    d.callout("CORE", "Core setting", [
        "CELLALGOSWITCH FreqPriorityHoSwitch FreqPriorIFHOSwitch = ON.",
        "MlbBasedFreqPriHoSwitch = ON when MLB is on, so heavy load is owned by MLB not FreqPri.",
        "A2BasedFreqPriHoSwitch = OFF in same-coverage multi-band (document recommendation).",
        "FreqPriorIFBlindHOSwitch = ON only if blind is required in same-coverage.",
        "EUTRANINTERNFREQ FreqPriBasedHoMeasFlag = ENABLE and MeasPriorityForFreqPriHo set.",
        "A4 TTT ≠ 5120 ms.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "FreqPri A1: FreqPriInterFreqHoA1ThdRsrp. Example −90 dBm, Hys 2, Ms −85.  −85 − 2 = −87 > −90 → serving is good, FreqPri may start (same-coverage).",
        "Target A4: InterFreqLoadBasedHoA4ThdRsrp + FreqPriHoA4ThldRsrpOffset. Example base −105 dBm, per-EARFCN offset +2 dB → used A4 = −103 dBm.",
        "Mn = −90, Hys = 2.  −92 > −103 → ENTER to high band. Need not beat serving; serving can still be −85 dBm.",
        "Waiting: FreqPriIFHoWaitingTimer. Example 1 s. After incoming unnecessary HO: FreqPriInHoProtectionTimer example 5 s so the UE is not bounced back at once.",
    ])
    d.callout("CONDITION", "Conditions that block start (document)", [
        "HO_USE_VOIP_FREQ_ALLOWED deselected for all QCIs on the UE when VoipMeasFreqPriSwitch is ON — voice can stop FreqPri meas.",
        "If MlbBasedFreqPriHoSwitch is ON, FreqPri does not run after MLB has already triggered from this cell.",
        "LoadTriggerFreqPriHoSwitch extra checks: overlap, load info, neighbour not in UE-number MLB, no PCI conflict.",
        "No reverse MLB target on a FreqPri pair (A→B FreqPri vs B→A MLB).",
        "Book MML examples A1/A2 −100/−105 dBm and A4 −110/−120 dBm are examples, not design values.",
    ])
    d.two_col(
        "Advantage",
        ["Keeps coverage layer for A2 rescue; parks traffic on high band.", "Blind option for same-coverage."],
        "Limitation",
        ["Fights MLB if reverse pair or MlbBasedFreqPriHoSwitch is off under load.", "5120 ms A4 TTT turns it off.", "VoipMeasFreqPriSwitch stops FreqPri meas when voice starts."],
    )
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "A1 (good serving) + A4 (high-priority freq good enough).",
        "Not a load algorithm. Turn MlbBasedFreqPriHoSwitch on with MLB.",
        "Never reverse-pair with MLB. A4 TTT not 5120 ms. RSRP.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "CELLALGOSWITCH", "FreqPriorIFHOSwitch", "ON", "Master measurement-based FreqPri.", "Yes", "11",
         "Enables frequency-priority-based inter-frequency handover (measurement-based).", R("LBFD-002018", "§11")),
        (2, "CELLALGOSWITCH", "MlbBasedFreqPriHoSwitch", "ON if MLB used", "Lets MLB own heavy load.", "Yes", "11",
         "When MLB is triggered, FreqPri yields so load movement is owned by Intra-RAT MLB.", R("LBFD-002018", "§11 / Table 11-7")),
        (3, "CELLALGOSWITCH", "A2BasedFreqPriHoSwitch", "OFF if same-coverage", "Document: deselect in multi-band same-coverage.", "Tune", "11",
         "Uses FreqPri A2 (serving poor) as a FreqPri start/stop condition. Deselect in same-coverage multi-band.", R("LBFD-002018", "§11")),
        (4, "CELLALGOSWITCH", "FreqPriorIFBlindHOSwitch", "ON if blind needed", "Same-coverage blind only.", "Tune", "11",
         "Enables frequency-priority blind handover in multi-band same-coverage.", R("LBFD-002018", "Table 10-6 / §11")),
        (5, "CELLALGOSWITCH", "LoadTriggerFreqPriHoSwitch", "as designed", "Extra load/overlap/PCI checks.", "Tune", "11",
         "Adds load, overlap and PCI checks before a frequency-priority HO is started.", R("LBFD-002018", "§11")),
        (6, "EUTRANINTERNFREQ", "FreqPriBasedHoMeasFlag", "ENABLE", "This EARFCN is a FreqPri meas object.", "Yes", "11",
         "Whether this neighbouring frequency is a measurement object for frequency-priority HO.", R("LBFD-002018", "§11")),
        (7, "EUTRANINTERNFREQ", "MeasPriorityForFreqPriHo", "priority", "Higher = preferred FreqPri target.", "Tune", "11",
         "Priority of this frequency among frequency-priority measurement objects. Higher value = preferred.", R("LBFD-002018", "§11")),
        (8, "INTERFREQHOGROUP", "FreqPriInterFreqHoA1ThdRsrp", "—", "Serving-good A1. Not the book example dBm.", "Tune", "11",
         "Event A1 RSRP threshold that indicates serving is good enough to start or keep FreqPri.", R("LBFD-002018", "§11")),
        (9, "INTERFREQHOGROUP", "FreqPriInterFreqHoA2ThdRsrp", "—", "Only if A2-based mode is used.", "Tune", "11",
         "Event A2 RSRP threshold used when A2-based frequency-priority mode is enabled.", R("LBFD-002018", "§11")),
        (10, "INTERFREQHOGROUP", "InterFreqLoadBasedHoA4ThdRsrp", "—", "Target A4. Better than coverage A2.", "Tune", "11",
         "A4 RSRP threshold of the high-priority (usually high-band) target. May add FreqPriHoA4ThldRsrpOffset per EARFCN.", R("LBFD-002018", "§11")),
        (11, "INTERFREQHOGROUP", "InterFreqHoA4TimeToTrig", "not 5120 ms", "Disables FreqPri if 5120 ms.", "Yes", "11",
         "TimeToTrig for A4. 5120 ms disables frequency-priority IFHO.", R("LBFD-002018", "Table 4-9 / §11")),
        (12, "INTRARATHOCOMM", "FreqPriInterFreqHoA1TrigQuan", "RSRP", "Recommended.", "Tune", "11",
         "Quantity used to trigger FreqPri event A1. RSRP is recommended.", R("LBFD-002018", "§11")),
        (13, "INTRARATHOCOMM", "FreqPriIFHoWaitingTimer", "—", "Wait for FreqPri A4 report.", "Tune", "11",
         "Time the eNodeB waits for a frequency-priority A4 measurement report.", R("LBFD-002018", "§11")),
        (14, "INTRARATHOCOMM", "FreqPriInHoProtectionTimer", "> 0", "After incoming unnecessary HO.", "Tune", "11",
         "Protects a UE after an incoming unnecessary HO so FreqPri does not bounce it back immediately.", R("LBFD-002018", "§11")),
        (15, "EUTRANINTERFREQNCELL", "BlindHoPriority", "≥16 if used", "Optional preferred blind neighbour. Example in book uses 17.", "Tune", "11",
         "Priority used to pick a preferred neighbouring cell for frequency-priority blind HO.", R("LBFD-002018", "§11")),
    ]
    for row in rows:
        d.param_row(*row)
    return d


def sheet_speed(wb):
    d = start(
        wb, "p", "12", "Speed-based Inter-Frequency Handover (FDD)",
        "Document page 320. FDD. Requires coverage-based inter-frequency HO (Ch.5.3) first.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "FDD only. Does not replace coverage A2.",
        "High-speed UEs should use a large-coverage layer (usually low band). Low-speed UEs can stay on a capacity layer.",
        "Speed only changes the inter-frequency target policy once the UE is classified as high speed.",
        "Requires coverage-based inter-frequency HO (Ch.5.3) first.",
        "Do not confuse with HighSpeedUserRedirectSwitch (railway dedicated network).",
    ], "UE speed / mobility state is estimated from handover count (and related high-speed options).")
    d.h1("Principle")
    d.info_box("Principle", [
        "Ch.5.3 InterFreqCoverHoSwitch = ON first.",
        "Enable the speed-based inter-frequency HO switch (confirm bit name in MAE on eRAN21.1).",
        "Chapter 4 object cap must include the coverage-layer EARFCN.",
        "A4/A5 or A3 family already correct for coverage IFHO.",
        "Target still must pass coverage A3 or A4/A5. Speed HO is not a substitute for A2.",
    ])
    d.callout("CORE", "Core setting", [
        "Ch.5.3 InterFreqCoverHoSwitch = ON first.",
        "Enable the speed-based inter-frequency HO switch (confirm bit name in MAE on eRAN21.1).",
        "Chapter 4 object cap must include the coverage-layer EARFCN.",
        "A4/A5 or A3 family already correct for coverage IFHO.",
    ])
    d.callout("CALC", "Calculation + example (not a design)", [
        "Mobility state from HO count in a window (normal / medium / high). Confirm the exact parameter names and window in MAE / eRAN21.1 §12.1.2.",
        "Example: HoNumThd = 4, HighSpeedHoNumThd = 8, SpeedStateJudgePeriod = 60 s, SpeedStateTimer = 120 s, SpeedStateValidTime = 30 s.",
        "If the eNodeB counts 6 intra-frequency HOs inside the 60 s judge period: 6 ≥ 4 → enter high-speed, start the 120 s timer, and after 30 s apply the high-speed parameter set.",
        "If the count later stays ≥ 8, keep high-speed. If it falls below 4, return to normal speed.",
        "High-speed: prefer lower-frequency / larger-coverage target. Target still must pass coverage A3 or A4/A5 (same numbers as Chapter 5).",
        "Do not set coverage A2 so low that a high-speed UE never leaves a dying small cell — speed HO is not a substitute for A2.",
    ])
    d.two_col(
        "Advantage",
        ["Fewer ping-pongs of highway UEs on small high-band cells.", "Re-uses Ch.5.3; few extra parameters."],
        "Limitation",
        ["FDD chapter.", "Dead if Ch.5.3 is off.", "HO-count speed estimate is coarse in dense urban.", "Not MLB."],
    )
    d.callout("CONDITION", "Conditions", "FDD. Coverage IFHO already on. Do not confuse with HighSpeedUserRedirectSwitch (railway dedicated network), which is a different option.")
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Speed only chooses the layer; coverage events still qualify the cell.",
        "Turn Ch.5.3 on, then the speed bit.",
        "Keep A2/A4 consistent with the coverage layer you want highway UEs to use.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "CELLHOPARACFG", "InterFreqCoverHoSwitch", "ON", "Prerequisite. Open Chapter 5.", "Yes", "→ Ch.5.3",
         "Coverage-based inter-frequency HO must be on before speed-based IFHO can be enabled.", R("LBFD-00201802", "§5.3 / §12")),
        (2, "CELLALGOSWITCH / HO", "Speed-based IFHO switch", "ON", "Confirm exact bit in MAE on eRAN21.1.", "Yes", "12",
         "Enables speed-based inter-frequency handover (FDD). High-speed UEs are steered to the coverage layer.", R("LBFD-002018", "§12")),
        (3, "INTERFREQHOGROUP", "Coverage A2 / A4 as Ch.5.3", "—", "Speed HO re-uses coverage events. No separate A2 family.", "Tune", "12",
         "Speed-based HO does not have its own A2 family; it re-uses coverage inter-frequency A2/A3/A4/A5.", R("LBFD-002018", "§12")),
        (4, "EUTRANINTERNFREQ", "FREQ_MEAS_FLAG", "selected", "Coverage-layer EARFCN must be measurable.", "Yes", "→ Ch.4",
         "The coverage-layer frequency must be a connected-mode measurement object.", R("LBFD-002018", "§4.1.4 / §12")),
    ]
    for i, row in enumerate(rows):
        link = S["c"] if i == 0 else (S["b"] if i == 3 else None)
        d.param_row(*row, link_sheet=link)
    return d


def sheet_utran_mplmn(wb):
    d = start(
        wb, "m", "13", "Separate Mobility Policies to UTRAN for Multi PLMN (FDD)",
        "Document page 330. Feature ID LOFD-070216. FDD only. Enhances Chapter 4 target-policy, does not start a new event.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "FDD only. Feature ID LOFD-070216.",
        "Enhances Chapter 4 target-policy. Does not start a new event. No new A2/B1.",
        "In RAN sharing, different operators’ UTRAN RNCs do not have the same PS HO, VoIP, SRVCC, RIM or ultra-flash CSFB capability.",
        "The eNodeB picks the policy from PLMN + RNC, instead of one global UTRAN assumption.",
        "Still need coverage IRAT (Ch.5.4) if you want the HO itself.",
    ])
    d.h1("Principle")
    d.info_box("Principle", [
        "ENODEBALGOSWITCH MultiOpCtrlSwitch UtranSepOpMobilitySwitch = ON.",
        "ADD UTRANNETWORKCAPCFG with MCC, MNC, RncId and the capability bits.",
        "PsHoCapCfg — whether that RNC supports PS handover (else redirect).",
        "VoipCapCfg — whether VoIP can be handed over.",
        "SrvccCapCfg — SRVCC to that UTRAN.",
        "SiByRimCapCfg — CN-based RIM for SI. eCoordinator RIM is not affected.",
        "UltraFlashCsfbCapCfg — ultra-flash CSFB capability.",
        "Prerequisite: Ch.5.4 UtranPsHoSwitch or UtranRedirectSwitch as required by the policy you want.",
    ])
    d.callout("CORE", "Core setting", "ENODEBALGOSWITCH MultiOpCtrlSwitch UtranSepOpMobilitySwitch = ON. ADD UTRANNETWORKCAPCFG with MCC, MNC, RncId and the capability bits. Prerequisite: Ch.5.4 UtranPsHoSwitch or UtranRedirectSwitch as required by the policy you want.")
    d.two_col(
        "Advantage",
        ["Fewer HO attempts toward an RNC that cannot do PS HO / SRVCC.", "Per-operator UTRAN behaviour on one LTE RAN."],
        "Limitation",
        ["FDD only (no TDD equivalent of LOFD-070216).", "Wrong capability bits cause redirect when HO was possible, or HO fail when redirect was needed.", "Does not replace neighbour planning."],
    )
    d.callout("CONDITION", "Conditions", "FDD. UTRAN neighbours exist. Flash/ultra-flash/SRVCC licenses if those bits are set. Related: Ch.5.4 and CS Fallback / SRVCC documents.")
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Policy table per PLMN+RNC, used at Ch.4 target decision.",
        "Switch UtranSepOpMobilitySwitch.",
        "Still need coverage IRAT (Ch.5.4) if you want the HO itself.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "ENODEBALGOSWITCH", "UtranSepOpMobilitySwitch", "ON", "Master.", "Yes", "13",
         "Enables separate UTRAN mobility policies per PLMN / RNC at target-decision time.", R("LOFD-070216", "§13")),
        (2, "UTRANNETWORKCAPCFG", "Mcc / Mnc / RncId", "operator RNC", "Key of the capability row.", "Yes", "13",
         "Identifies the UTRAN RNC (PLMN + RNC ID) whose capabilities are configured.", R("LOFD-070216", "Table 11-1 / §13")),
        (3, "UTRANNETWORKCAPCFG", "PsHoCapCfg", "ON if RNC can PS HO", "Else eNodeB will not choose PS HO to that RNC.", "Yes", "13",
         "Whether that RNC supports PS handover. If off, the eNodeB will not choose PS HO toward it.", R("LOFD-070216", "§13.1")),
        (4, "UTRANNETWORKCAPCFG", "VoipCapCfg", "as RNC", "VoIP HO capability.", "Tune", "13",
         "Whether VoIP services can be handed over to that UTRAN RNC.", R("LOFD-070216", "§13.1")),
        (5, "UTRANNETWORKCAPCFG", "SrvccCapCfg", "as RNC", "SRVCC capability.", "Tune", "13",
         "Whether SRVCC to that UTRAN RNC is supported.", R("LOFD-070216", "§13.1")),
        (6, "UTRANNETWORKCAPCFG", "SiByRimCapCfg", "as RNC", "CN-based RIM only. Not eCoordinator RIM.", "Tune", "13",
         "Whether that RNC supports CN-based RIM for SI acquisition (CSFB / redirect method).", R("LOFD-070216", "§13.1")),
        (7, "UTRANNETWORKCAPCFG", "UltraFlashCsfbCapCfg", "as RNC", "Ultra-flash CSFB.", "Tune", "13",
         "Whether that RNC supports ultra-flash CSFB from E-UTRAN.", R("LOFD-070216", "§13.1")),
        (8, "ENODEBALGOSWITCH", "UtranPsHoSwitch / UtranRedirectSwitch", "ON as needed", "Prerequisite from Chapter 5.4.", "Yes", "→ Ch.5.4",
         "Coverage IRAT to UTRAN must already be enabled; this chapter only specialises the policy per PLMN.", R("LOFD-001019", "§5.4 / §13")),
    ]
    for i, row in enumerate(rows):
        d.param_row(*row, link_sheet=S["c"] if i == 7 else None)
    return d


def sheet_geran_mplmn(wb):
    d = start(
        wb, "g", "14", "Separate Mobility Policies to GERAN for Multi PLMN",
        "Document page 337. Feature ID LOFD-111204 (TDD: TDLOFD-131210). Same idea as Chapter 13, for GERAN.",
    )
    d.h1("Introduction")
    d.info_box("Overview of types in this chapter", [
        "Feature ID LOFD-111204 (TDD: TDLOFD-131210).",
        "Same idea as Chapter 13, for GERAN. No new event.",
        "Different operators’ GERAN BSCs do not share one capability set.",
        "Selects GERAN HO / CCO / SI-by-RIM policy from PLMN + BSC at Chapter 4 target decision.",
        "Still need Ch.5.5 for the actual coverage IRAT.",
    ])
    d.h1("Principle")
    d.info_box("Principle", [
        "MultiOpCtrlSwitch GeranSepOpMobilitySwitch = ON (confirm exact bit in MAE).",
        "Configure per-BSC GERAN capabilities (GERANNETWORKCAPCFG or the name shown in MAE) keyed by MCC/MNC/BSC.",
        "HO / CCO / RIM bits as supported by that BSC.",
        "Prerequisite: Ch.5.5 GeranRedirectSwitch (and related GERAN HO switches) already match the policy you want.",
        "Does not create GERAN neighbours.",
    ])
    d.callout("CORE", "Core setting", "MultiOpCtrlSwitch GeranSepOpMobilitySwitch = ON (confirm exact bit in MAE). Configure per-BSC GERAN capabilities. Prerequisite: Ch.5.5 GeranRedirectSwitch (and related GERAN HO switches) already match the policy you want.")
    d.two_col(
        "Advantage",
        ["Stops HO/CCO toward a BSC that cannot do it.", "Per-operator GSM behaviour on shared LTE."],
        "Limitation",
        ["Wrong bits → extra fail or extra redirect.", "Does not create GERAN neighbours.", "Still need Ch.5.5 for the actual coverage IRAT."],
    )
    d.callout("CONDITION", "Conditions", "GERAN NRT present. Related CS Fallback document for SI-by-RIM / CCO. TDD ID is TDLOFD-131210.")
    d.h1("Combined summary")
    d.info_box("Combined summary", [
        "Ch.13 = UTRAN per PLMN. Ch.14 = GERAN per PLMN.",
        "Both sit on Ch.4 target decision.",
        "Coverage IRAT chapters remain the ones that start the HO.",
    ])
    d.h1("Parameter list")
    d.param_heads()
    rows = [
        (1, "ENODEBALGOSWITCH", "GeranSepOpMobilitySwitch", "ON", "Confirm exact bit name in MAE.", "Yes", "14",
         "Enables separate GERAN mobility policies per PLMN / BSC at target-decision time.", R("LOFD-111204", "§14")),
        (2, "GERANNETWORKCAPCFG", "Mcc / Mnc / BscId", "operator BSC", "Key of the capability row. Confirm MO name in MAE.", "Yes", "14",
         "Identifies the GERAN BSC (PLMN + BSC ID) whose capabilities are configured.", R("LOFD-111204", "§14")),
        (3, "GERANNETWORKCAPCFG", "capability bits", "as BSC", "HO / CCO / RIM as supported by that BSC.", "Tune", "14",
         "GERAN HO, CCO and SI-by-RIM capabilities of that BSC. Confirm option names in MAE.", R("LOFD-111204", "§14")),
        (4, "ENODEBALGOSWITCH", "GeranRedirectSwitch", "ON as needed", "Prerequisite from Chapter 5.5.", "Yes", "→ Ch.5.5",
         "Coverage IRAT to GERAN must already be enabled; this chapter only specialises the policy per PLMN.", R("LOFD-001020", "§5.5 / §14")),
    ]
    for i, row in enumerate(rows):
        d.param_row(*row, link_sheet=S["c"] if i == 3 else None)
    return d


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb = Workbook()
    wb.active.title = "tmp"
    sheet_toc_fix(wb)
    sheet_overview(wb)
    sheet_basic(wb)
    sheet_coverage(wb)
    sheet_service(wb)
    sheet_distance(wb)
    sheet_ulq(wb)
    sheet_cqi(wb)
    sheet_sreq(wb)
    sheet_freqpri(wb)
    sheet_speed(wb)
    sheet_utran_mplmn(wb)
    sheet_geran_mplmn(wb)
    del wb["tmp"]
    # drop unused stub if created
    if "00 Contents" in wb.sheetnames and wb.sheetnames[0] != "00 Contents":
        pass
    wb._sheets = [wb[S[k]] for k in ORDER if S[k] in wb.sheetnames]
    wb.properties.title = f"Connected Mode eRAN21.1 Feature Sheets {VER}"
    wb.properties.subject = DOC
    wb.properties.creator = "Feature-parameter summary"
    wb.properties.keywords = VER
    wb.save(OUT)
    print("Wrote", OUT)
    print("sheets", wb.sheetnames)


if __name__ == "__main__":
    main()
