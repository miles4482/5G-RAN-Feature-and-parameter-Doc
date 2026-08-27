"""Feature Activation table: SN | MML with Value | Target | Feature | Additional Comments | Ref."""

LC = "<LocalCellId>"
EARFCN = "<DlEarfcn>"
GID = "<HoGroupId>"

F_BASIC = "Basic"
F_C_INTRA = "Coverage-based Intra-Frequency Handover"
F_C_INTER = "Coverage-based Inter-Frequency Handover"
F_C_UTRAN = "Coverage-based Inter-RAT Handover to UTRAN"
F_C_GERAN = "Coverage-based Inter-RAT Handover to GERAN"
F_C_CSPS = "Coverage-based E-UTRAN to UTRAN CS/PS Steering"
F_S_IF = "Service-based Inter-Frequency Handover"
F_S_UTRAN = "Service-based Inter-RAT Handover to UTRAN"
F_S_GERAN = "Service-based Inter-RAT Handover to GERAN"
F_D_IF = "Distance-based Inter-Frequency Handover"
F_D_UTRAN = "Distance-based Inter-RAT Handover to UTRAN"
F_D_GERAN = "Distance-based Inter-RAT Handover to GERAN"
F_UL_IF = "UL-Quality-based Inter-Frequency Handover"
F_UL_IRAT = "UL-Quality-based Inter-RAT Handover"
F_CQI = "CQI-based Inter-Frequency Handover (FDD)"
F_SREQ = "Service-Request-based Inter-Frequency Handover"
F_FREQPRI = "Frequency-Priority-based Inter-Frequency Handover"
F_SPEED = "Speed-based Inter-Frequency Handover (FDD)"
F_U_MPLMN = "Separate Mobility Policies to UTRAN for Multi PLMN"
F_G_MPLMN = "Separate Mobility Policies to GERAN for Multi PLMN"

# Row: (mml, target, feature, additional_comments, ref)
BASIC_MML = [
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, HoModeSwitch=UtranPsHoSwitch-1&UtranRedirectSwitch-1&GeranRedirectSwitch-1;",
     "//Setting policy switches of E-UTRAN to UTRAN or GERAN handovers",
     F_BASIC, "Replace <LocalCellId>. Bit -1 = ON. Leave a bit -0 if that RAT is not used.",
     "LBFD-002018 · §4 / §5.4 / §5.5"),
    (f"MOD ENODEBALGOSWITCH: HoModeSwitch=UtranVoipCapSwitch-1&GeranPsHoSwitch-1&GeranNaccSwitch-1&GeranCcoSwitch-1, RedirectSwitch=GeranFlashRedirectSwitch-1&UtranFlashRedirectSwitch-1;",
     "//Setting VoIP / GERAN PS HO / NACC / CCO and flash-redirect capability",
     F_BASIC, "Capability bits. Set only what the target RNC/BSC actually supports.",
     "LBFD-002018 · §4"),
    (f"MOD EUTRANINTERNFREQ: DlEarfcn={EARFCN}, FreqMeasFlag=MEASURE;",
     "//Setting FREQ_MEAS_FLAG so the frequency is measured",
     F_BASIC, "Silent no-HO if this flag is off.",
     "LBFD-002018 · §4.1.4"),
    (f"MOD EUTRANINTERNFREQ: DlEarfcn={EARFCN}, HoTrgFreqForbidMeasFlag=BOOLEAN_FALSE;",
     "//Deselecting HO_TRG_FREQ_FORBID_MEAS_FLAG for required HO targets",
     F_BASIC, "Must be off for frequencies that are allowed as HO targets.",
     "LBFD-002018 · §4.1.4"),
    (f"MOD EUTRANINTERFREQNCELL: LocalCellId={LC}, Mcc=<MCC>, Mnc=<MNC>, eNodeBId=<eNBId>, CellId=<CID>, NoHoFlag=PERMIT_HO;",
     "//Allowing the neighbouring cell as an HO target",
     F_BASIC, "Replace PLMN / eNodeB / cell IDs with the live neighbour.",
     "LBFD-002018 · §4.1.4"),
    (f"MOD CELLUEMEASCONTROLCFG: LocalCellId={LC}, MaxNonIntraMeasObjNum=<N>;",
     "//Setting non-intra measurement object capacity",
     F_BASIC, "N ≥ objects this cell must measure. Over cap = random drop, not MLB.",
     "LBFD-002018 · Table 4-3"),
    (f"MOD INTRARATHOCOMM: InterFreqHoA1A2TrigQuan=RSRP, InterFreqHoA4TrigQuan=RSRP;",
     "//Setting A1/A2 and A4 trigger quantity to RSRP",
     F_BASIC, "RSRQ moves with scheduler load. RSRP recommended.",
     "LBFD-002018 · §4.1.4"),
    (f"MOD INTERFREQHOGROUP: LocalCellId={LC}, InterFreqHoGroupId={GID}, InterFreqHoA4TimeToTrig=320;",
     "//Setting A4 TimeToTrig (must not be 5120 ms)",
     F_BASIC, "5120 ms disables FreqPri / CQI / service-based IFHO (Table 4-9). 320 is an example, not a design.",
     "LBFD-002018 · Table 4-9"),
    (f"MOD HOMEASCOMM: SMeasure=<from MR>;",
     "//Setting SMeasure from MR",
     F_BASIC, "Calibrate from MR. A too-high SMeasure silently kills A4. Do not copy book command-example dBm.",
     "LBFD-002018 · §4.1.5"),
    (f"MOD CELLHO: LocalCellId={LC}, IntraRatHoRprtAmount=1;",
     "//Setting event-triggered intra-RAT reporting",
     F_BASIC, "Amount = 1 means event-triggered. Do not periodical-ise HO events.",
     "LBFD-002018 · §4.1.5"),
    (f"MOD CELLHO: LocalCellId={LC}, InterRatHoRprtAmount=1;",
     "//Setting event-triggered inter-RAT reporting",
     F_BASIC, "Amount = 1 means event-triggered for B1/B2.",
     "LBFD-002018 · §4.1.5"),
]

MML_COVERAGE = [
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=IntraFreqCoverHoSwitch-1;",
     "Core switch to activate feature",
     F_C_INTRA, "Always-on A3. No A2. No blind.",
     "LBFD-00201801 · §5.2"),
    (f"MOD INTRARATHOCOMM: IntraFreqHoA3TrigQuan=RSRP, IntraFreqHoA3RprtQuan=SAME_AS_TRIG_QUAN;",
     "Setting the triggering quantity of event A3",
     F_C_INTRA, "RSRP recommended. Reporting same as trigger.",
     "LBFD-00201801 · §5.2.4"),
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, CellHoAlgoSwitch=InterFreqCoverHoSwitch-1;",
     "Core switch to activate feature",
     F_C_INTER, "Measurement-based inter-frequency coverage HO.",
     "LBFD-00201802 · §5.3"),
    (f"MOD EUTRANINTERNFREQ: DlEarfcn={EARFCN}, InterFreqHoEventType=EventA3;",
     "Setting the target event and the A2 family",
     F_C_INTER, "Use EventA3 or EventA4 or EventA5. Do not mix A2 families.",
     "LBFD-00201802 · §5.1.4 / §5.3"),
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, CellHoAlgoSwitch=IfCoverPreBlindHoSwitch-1;",
     "Setting preferential blind handover",
     F_C_INTER, "ON only if the neighbouring cell/frequency fully contains the source. Otherwise -0.",
     "LBFD-00201802 · §5.3.1"),
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, CellHoAlgoSwitch=EmcInterFreqBlindHoSwitch-1;",
     "Setting emergency blind redirection",
     F_C_INTER, "Separate worse A2 (BlindHoA1A2Thd). Leave -0 if not required.",
     "LBFD-00201802 · §5.3.1"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, CellAlgoSwitch=ReduceInvalidA1A2RptSigSwitch-1;",
     "Setting A2-first / A1-after-A2 signalling",
     F_C_INTER, "Cuts extra A1/A2 signalling at RRC setup.",
     "LBFD-00201802 · §5.3.1"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UtranPsHoSwitch-1;",
     "Core switch to activate feature",
     F_C_UTRAN, "Measurement-based PS HO to UTRAN. IRAT A2 then B1/B2. Offload TTT < 3 s.",
     "LOFD-001019 · §5.4"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UtranRedirectSwitch-1;",
     "Setting blind / redirect to UTRAN",
     F_C_UTRAN, "Blind path. Also listed under Basic HoModeSwitch; keep consistent.",
     "LOFD-001019 · §5.4"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=SrvccRatSteeringSwitch-1;",
     "Setting voice RAT pick after IRAT A2",
     F_C_UTRAN, "QCI-1 / SRVCC. RatLayerSwitch is legacy.",
     "LOFD-001019 · Table 5-4"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=GeranRedirectSwitch-1;",
     "Core switch to activate feature",
     F_C_GERAN, "Coverage IRAT / redirect to GERAN. Blind IRAT is not the normal path for QCI-1.",
     "LOFD-001020 · §5.5"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=PsRatSteeringSwitch-1;",
     "Setting data RAT pick after IRAT A2",
     F_C_GERAN, "Non-QCI-1. Measure only the highest-priority RAT for PS.",
     "LOFD-001019 · Table 5-4"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, FreqLayerSwitch=UtranFreqLayerMeasSwitch-1;",
     "Core switch to activate feature",
     F_C_CSPS, "Needs §5.4 first. Then set UTRANNFREQ CsPriority / PsPriority. Priority_0 = do not use.",
     "LOFD-001078 · §5.6"),
]

MML_SERVICE = [
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=ServiceBasedInterFreqHoSwitch-1;",
     "Core switch to activate feature",
     F_S_IF, "eNodeB master. Cell bit below is also required.",
     "LBFD-00201805 · §6.2"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, HoAlgoSwitch=SrvBasedInterFreqHoSw-1;",
     "Setting the cell allow switch",
     F_S_IF, "Both eNodeB and cell bits must be ON.",
     "LBFD-00201805 · §6.2"),
    (f"MOD SERVICEIFHOCFGGROUP: LocalCellId={LC}, ServiceIfHoCfgGroupId=<GrpId>, InterFreqHoState=PERMIT_HO;",
     "Setting QCI permission to leave the serving frequency",
     F_S_IF, "Bind QCI on CNOPERATORQCIPARA. Put target EARFCN on SERVICEIFDLEARFCNGRP.",
     "LBFD-00201805 · §6.2"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UtranServiceHoSwitch-1;",
     "Core switch to activate feature",
     F_S_UTRAN, "Event B1. Bind QCI. Offload TTT < 3 s.",
     "LOFD-001043 · §6.3"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=GeranServiceHoSwitch-1;",
     "Core switch to activate feature",
     F_S_GERAN, "Event B1. Same SERVICEIRHOCFGGROUP bind as UTRAN.",
     "LOFD-001046 · §6.4"),
]

MML_DISTANCE = [
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, HoAlgoSwitch=DistBasedHoSwitch-1;",
     "Core switch to activate feature",
     F_D_IF, "Master for Chapter 7. TA-based overshoot. 10 s start/stop.",
     "LBFD-00201804 · §7.1"),
    (f"MOD DISTBASEDHO: LocalCellId={LC}, DistBasedMeasObjType=EUTRAN;",
     "Setting E-UTRAN as the distance target RAT",
     F_D_IF, "Then A4 still qualifies the neighbour. A4 must stay better than coverage A2.",
     "LBFD-00201804 · §7.2"),
    (f"MOD DISTBASEDHO: LocalCellId={LC}, DistBasedMeasObjType=UTRAN;",
     "Setting UTRAN as a distance target RAT",
     F_D_UTRAN, "Leave off if IRAT overshoot is not used. Event B1.",
     "LOFD-001072 · §7.3"),
    (f"MOD DISTBASEDHO: LocalCellId={LC}, DistBasedMeasObjType=GERAN;",
     "Setting GERAN as a distance target RAT",
     F_D_GERAN, "Leave off if IRAT overshoot is not used. Event B1.",
     "LOFD-001073 · §7.4"),
]

MML_ULQ = [
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UlQualityInterFreqHoSwitch-1;",
     "Core switch to activate feature",
     F_UL_IF, "UL MCS + IBLER starts IFHO. A4 = coverage A4 + UlBadQualHoA4Offset.",
     "LBFD-002018 · §8.2"),
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UlQualityInterRATHoSwitch-1;",
     "Core switch to activate feature",
     F_UL_IRAT, "Same UL start; target B1. Leave -0 if IRAT is not required.",
     "LBFD-002018 · §8.3"),
]

MML_CQI = [
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, CellHoAlgoSwitch=<confirm CQI-based IFHO bit in MAE>-1;",
     "Core switch to activate feature",
     F_CQI, "Confirm the exact bit name in MAE on this eRAN21.1. Do not invent a bit. A4 TTT must not be 5120 ms.",
     "LBFD-002018 · §9"),
]

MML_SREQ = [
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, HoAllowedSwitch=ServiceReqInterFreqHoSwitch-1;",
     "Core switch to activate feature",
     F_SREQ, "FDD: cell-level. Starts on bearer setup / modify. Own A4 SrvReqHoA4ThdRsrp.",
     "LBFD-002018 · §10"),
    (f"MOD SERVICEIFHOCFGGROUP: LocalCellId={LC}, ServiceIfHoCfgGroupId=<GrpId>, InterFreqHoState=PERMIT_HO;",
     "Setting QCI permission at service request",
     F_SREQ, "Cap wait with A4RptWaitingTimer. Keep A4 better than coverage A2.",
     "LBFD-002018 · §10.1.1"),
]

MML_FREQPRI = [
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, FreqPriorityHoSwitch=FreqPriorIFHOSwitch-1;",
     "Core switch to activate feature",
     F_FREQPRI, "Measurement-based FreqPri. A1 (serving good) then A4 (high-priority freq).",
     "LBFD-002018 · §11"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, FreqPriorityHoSwitch=MlbBasedFreqPriHoSwitch-1;",
     "Setting FreqPri to yield to MLB under heavy load",
     F_FREQPRI, "ON when MLB is used so heavy load is owned by MLB, not FreqPri.",
     "LBFD-002018 · §11 / Table 11-7"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, FreqPriorityHoSwitch=A2BasedFreqPriHoSwitch-0;",
     "Deselecting A2-based FreqPri in same-coverage multi-band",
     F_FREQPRI, "Document recommendation for same-coverage: -0.",
     "LBFD-002018 · §11"),
    (f"MOD EUTRANINTERNFREQ: DlEarfcn={EARFCN}, FreqPriBasedHoMeasFlag=ENABLE;",
     "Setting this EARFCN as a FreqPri measurement object",
     F_FREQPRI, "Also set MeasPriorityForFreqPriHo. A4 TTT must not be 5120 ms.",
     "LBFD-002018 · §11"),
]

MML_SPEED = [
    (f"MOD CELLHOPARACFG: LocalCellId={LC}, CellHoAlgoSwitch=InterFreqCoverHoSwitch-1;",
     "Setting coverage inter-frequency HO first (prerequisite)",
     F_SPEED, "Ch.5.3 must be ON before speed-based IFHO.",
     "LBFD-00201802 · §5.3 / §12"),
    (f"MOD CELLALGOSWITCH: LocalCellId={LC}, HoAlgoSwitch=<confirm speed-based IFHO bit in MAE>-1;",
     "Core switch to activate feature",
     F_SPEED, "Confirm the exact bit name in MAE. Do not confuse with HighSpeedUserRedirectSwitch.",
     "LBFD-002018 · §12"),
]

MML_UTRAN_MPLMN = [
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=UtranPsHoSwitch-1;",
     "Setting coverage IRAT to UTRAN (prerequisite)",
     F_U_MPLMN, "This chapter does not start a new event.",
     "LOFD-001019 · §5.4 / §13"),
    (f"MOD ENODEBALGOSWITCH: MultiOpCtrlSwitch=UtranSepOpMobilitySwitch-1;",
     "Core switch to activate feature",
     F_U_MPLMN, "FDD LOFD-070216. Policy per PLMN / RNC.",
     "LOFD-070216 · §13"),
    (f"ADD UTRANNETWORKCAPCFG: Mcc=<MCC>, Mnc=<MNC>, RncId=<RncId>, PsHoCapCfg=<as RNC>;",
     "Setting per-RNC UTRAN capability",
     F_U_MPLMN, "One row per operator RNC. Set Voip / SRVCC / SI-by-RIM / ultra-flash as that RNC supports.",
     "LOFD-070216 · §13.1"),
]

MML_GERAN_MPLMN = [
    (f"MOD ENODEBALGOSWITCH: HoAlgoSwitch=GeranRedirectSwitch-1;",
     "Setting coverage IRAT to GERAN (prerequisite)",
     F_G_MPLMN, "This chapter does not start a new event.",
     "LOFD-001020 · §5.5 / §14"),
    (f"MOD ENODEBALGOSWITCH: MultiOpCtrlSwitch=GeranSepOpMobilitySwitch-1;",
     "Core switch to activate feature",
     F_G_MPLMN, "Confirm exact bit in MAE. LOFD-111204.",
     "LOFD-111204 · §14"),
    (f"ADD GERANNETWORKCAPCFG: Mcc=<MCC>, Mnc=<MNC>, BscId=<BscId>;",
     "Setting per-BSC GERAN capability",
     F_G_MPLMN, "Confirm MO name in MAE. One row per operator BSC.",
     "LOFD-111204 · §14"),
]


def write_activation(d, extra=None, include_basic=True, note=None):
    extra = extra or []
    d.act_title()
    d.act_subtitle("Step by Step MML Commands")
    d.para(note or (
        "Basic rows first (Chapter 4), then this feature in document order. "
        "Replace <LocalCellId> / <DlEarfcn> / <HoGroupId> / PLMN / cell IDs before paste to MAE. "
        "Bit -1 = ON, -0 = OFF. Do not copy book command-example dBm into the command."
    ))
    d.activation_heads()
    sn = 1
    rows = (BASIC_MML if include_basic else []) + list(extra)
    if not rows:
        d.activation_row(1, "—", "No activation command in this chapter.", "—", "—", "—")
        return d
    for mml, target, feature, comments, ref in rows:
        d.activation_row(sn, mml, target, feature, comments, ref)
        sn += 1
    return d
