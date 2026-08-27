#!/usr/bin/env python3
"""Teaching-value charts for Events A1–A5, B1, B2.

Y-axis is RSRP in dBm. Trigger and leave points show the calculated value.
Numbers match the Measurement Event CALC boxes (understanding only, not a design).
"""

import os

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch

OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "docs/4G_LTE_Mobility_Management/figures",
)

BLACK = "#111111"
NAVY = "#1F4E79"
TEAL = "#2E75B6"
DPI = 130
HYS = 2.0
TTT = 320.0  # ms
TMAX = 2400.0


def _curve(tps, yps, n=1200):
    t = np.linspace(0, TMAX, n)
    return t, np.interp(t, tps, yps)


def _cross(t, y, level, direction):
    s = np.sign(y - level)
    d = np.diff(s)
    hits = np.where(d > 0)[0] if direction == "up" else np.where(d < 0)[0]
    if len(hits) == 0:
        return None
    i = int(hits[0])
    y0, y1 = y[i] - level, y[i + 1] - level
    frac = 0.0 if y1 == y0 else -y0 / (y1 - y0)
    return float(t[i] + frac * (t[i + 1] - t[i]))


def _at(t, y, x):
    return float(np.interp(x, t, y))


def _save(fig, name):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name)
    fig.savefig(path, dpi=DPI, facecolor="white", bbox_inches="tight", pad_inches=0.16)
    plt.close(fig)
    from PIL import Image
    im = Image.open(path).convert("RGB")
    target_w = 1100
    if im.size[0] != target_w:
        nh = int(round(im.size[1] * target_w / float(im.size[0])))
        im = im.resize((target_w, nh), Image.Resampling.LANCZOS)
        im.save(path, "PNG")
    print("wrote", path, im.size)
    return path


def _style(ax, ymin, ymax):
    ax.set_xlim(0, TMAX)
    ax.set_ylim(ymin, ymax)
    ax.set_ylabel("RSRP (dBm)", color=BLACK, fontsize=10)
    ax.set_xlabel("Time (ms)", color=BLACK, fontsize=10)
    ax.set_xticks([0, 400, 800, 1200, 1600, 2000, 2400])
    ax.tick_params(colors=BLACK, labelsize=8)
    for sp in ax.spines.values():
        sp.set_color(BLACK)
        sp.set_linewidth(1.0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, ls=":", lw=0.5, color="#D0D5DA")


def _box(ax, lines, loc="upper left"):
    text = "Teaching values  ·  not a live design\n" + "\n".join(lines)
    ax.text(
        0.012 if "left" in loc else 0.988,
        0.985,
        text,
        transform=ax.transAxes,
        ha="left" if "left" in loc else "right",
        va="top",
        fontsize=7.6,
        color=NAVY,
        family="DejaVu Sans",
        bbox=dict(boxstyle="round,pad=0.35", facecolor="#EEF5FA", edgecolor=NAVY, linewidth=0.8),
        linespacing=1.25,
        zorder=6,
    )


def _thresh(ax, y, label):
    ax.axhline(y, ls=(0, (4, 3)), lw=1.1, color=BLACK, zorder=2)
    ax.text(20, y + 0.55, label, fontsize=8, color=BLACK, va="bottom", zorder=5)


def _ttt(ax, t0, t1, y, label="TimeToTrig = 320 ms"):
    ax.annotate(
        "",
        xy=(t1, y),
        xytext=(t0, y),
        arrowprops=dict(arrowstyle="<->", color=NAVY, lw=1.05),
        zorder=5,
    )
    ax.text((t0 + t1) / 2.0, y + 0.7, label, ha="center", va="bottom", fontsize=7.4, color=NAVY, zorder=5)


def _point(ax, x, y, title, detail, ha="left"):
    ax.plot(x, y, "o", ms=7, color=NAVY, zorder=7)
    ax.axvline(x, ls=(0, (1.4, 2.2)), lw=0.85, color=BLACK, zorder=1)
    dx = 40 if ha == "left" else -40
    ax.annotate(
        f"{title}\n{detail}",
        xy=(x, y),
        xytext=(x + dx, y + (3.8 if y > -100 else -3.8)),
        fontsize=7.3,
        color=NAVY,
        ha=ha,
        va="center",
        bbox=dict(boxstyle="round,pad=0.28", facecolor="white", edgecolor=NAVY, linewidth=0.7),
        arrowprops=dict(arrowstyle="->", color=NAVY, lw=0.8),
        zorder=8,
    )


def _caption(fig, text):
    fig.text(0.50, 0.012, text, ha="center", va="bottom", fontsize=9.5, color=BLACK)


def fig_a1():
    # Trigger uses Ms = −95. Leave uses Ms = −105. Same teaching set as the CALC box.
    tps = [0, 480, 800, 1150, 1580, 1900, 2400]
    yps = [-108, -98, -95, -90, -102, -105, -109]
    t, ms = _curve(tps, yps)
    enter = _cross(t, ms - HYS, -100, "up")
    leave = _cross(t, ms + HYS, -100, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.35))
    _style(ax, -114, -82)
    _thresh(ax, -100, "Thresh = −100 dBm")
    ax.plot(t, ms, color=BLACK, lw=1.7, label="Ms")
    ax.plot(t, ms - HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)), label="Ms − Hys")
    ax.plot(t, ms + HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)), label="Ms + Hys")
    ax.text(1180, _at(t, ms, 1180) + 0.8, "Ms", fontsize=8)
    ax.text(520, _at(t, ms - HYS, 520) - 1.6, "Ms − Hys", fontsize=7.4)
    ax.text(1680, _at(t, ms + HYS, 1680) + 0.7, "Ms + Hys", fontsize=7.4)
    _ttt(ax, enter, trig, -112.6)
    _ttt(ax, leave, term, -112.6)
    _point(ax, trig, _at(t, ms - HYS, trig), "Trigger A1", f"Ms−Hys = {_at(t, ms - HYS, trig):.0f} dBm  >  −100")
    _point(ax, term, _at(t, ms + HYS, term), "Leave A1", f"Ms+Hys = {_at(t, ms + HYS, term):.0f} dBm  <  −100", ha="right")
    _box(ax, [
        "Ms = −95 dBm (enter) / −105 dBm (leave)",
        "Hys = 2 dB    Thresh = −100 dBm    TTT = 320 ms",
        "Enter: Ms − Hys > Thresh     (−95 − 2 = −97 > −100)",
        "Leave: Ms + Hys < Thresh     (−105 + 2 = −103 < −100)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-3  Entering and leaving of event A1   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_3_event_a1_teaching.png")


def fig_a2():
    tps = [0, 480, 800, 1150, 1580, 1900, 2400]
    yps = [-88, -102, -105, -110, -98, -95, -88]
    t, ms = _curve(tps, yps)
    enter = _cross(t, ms + HYS, -100, "down")
    leave = _cross(t, ms - HYS, -100, "up")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.35))
    _style(ax, -116, -80)
    _thresh(ax, -100, "Thresh = −100 dBm")
    ax.plot(t, ms, color=BLACK, lw=1.7)
    ax.plot(t, ms - HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.plot(t, ms + HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.text(1180, _at(t, ms, 1180) - 1.5, "Ms", fontsize=8)
    ax.text(520, _at(t, ms + HYS, 520) + 0.8, "Ms + Hys", fontsize=7.4)
    ax.text(1680, _at(t, ms - HYS, 1680) - 1.5, "Ms − Hys", fontsize=7.4)
    _ttt(ax, enter, trig, -114.5)
    _ttt(ax, leave, term, -114.5)
    _point(ax, trig, _at(t, ms + HYS, trig), "Trigger A2", f"Ms+Hys = {_at(t, ms + HYS, trig):.0f} dBm  <  −100")
    _point(ax, term, _at(t, ms - HYS, term), "Leave A2", f"Ms−Hys = {_at(t, ms - HYS, term):.0f} dBm  >  −100", ha="right")
    _box(ax, [
        "Ms = −105 dBm (enter) / −95 dBm (leave)",
        "Hys = 2 dB    Thresh = −100 dBm    TTT = 320 ms",
        "Enter: Ms + Hys < Thresh     (−105 + 2 = −103 < −100)",
        "Leave: Ms − Hys > Thresh     (−95 − 2 = −97 > −100)",
    ], loc="upper right")
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-4  Entering and leaving of event A2   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_4_event_a2.png")


def fig_a3():
    t = np.linspace(0, TMAX, 1200)
    ms = np.full_like(t, -95.0)
    off = 2.0
    right = ms + off  # Ofs=Ocs=0
    tps = [0, 500, 820, 1150, 1600, 1920, 2400]
    yps = [-108, -91, -90, -85, -95, -100, -108]
    mn = np.interp(t, tps, yps)
    enter = _cross(t, (mn - HYS) - right, 0.0, "up")
    leave = _cross(t, (mn + HYS) - right, 0.0, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.45))
    _style(ax, -114, -78)
    ax.plot(t, right, color=BLACK, lw=1.05, ls=(0, (4, 3)))
    ax.plot(t, ms, color=BLACK, lw=1.35)
    ax.plot(t, mn, color=BLACK, lw=1.7)
    ax.plot(t, mn - HYS, color=BLACK, lw=1.0, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn + HYS, color=BLACK, lw=1.0, ls=(0, (2.2, 1.8)))
    ax.text(40, -95 + 0.9, "Ms = −95 dBm", fontsize=7.4)
    ax.text(40, -93 + 0.9, "Ms + Off = −93 dBm", fontsize=7.4)
    ax.text(1180, _at(t, mn, 1180) + 0.8, "Mn", fontsize=8)
    _ttt(ax, enter, trig, -112.6)
    _ttt(ax, leave, term, -112.6)
    _point(ax, trig, _at(t, mn - HYS, trig), "Trigger A3", f"Left = {_at(t, mn - HYS, trig):.0f}  >  Right = −93")
    _point(ax, term, _at(t, mn + HYS, term), "Leave A3", f"Left = {_at(t, mn + HYS, term):.0f}  <  Right = −93", ha="right")
    _box(ax, [
        "Ms = −95 dBm    Mn = −90 dBm (enter) / −100 dBm (leave)",
        "Hys = 2 dB    Off = 2 dB    Ofn=Ocn=Ofs=Ocs=0    TTT = 320 ms",
        "Enter: Mn − Hys > Ms + Off     (−90 − 2 = −92 > −93)",
        "Leave: Mn + Hys < Ms + Off     (−100 + 2 = −98 < −93)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-5  Entering and leaving of event A3   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_5_event_a3.png")


def fig_a4():
    tps = [0, 480, 800, 1150, 1580, 1900, 2400]
    yps = [-118, -103, -90, -88, -107, -110, -118]
    t, mn = _curve(tps, yps)
    th = -105.0
    enter = _cross(t, mn - HYS, th, "up")
    leave = _cross(t, mn + HYS, th, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.35))
    _style(ax, -124, -78)
    _thresh(ax, th, "Thresh = −105 dBm")
    ax.plot(t, mn, color=BLACK, lw=1.7)
    ax.plot(t, mn - HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn + HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.text(1180, _at(t, mn, 1180) + 0.8, "Mn", fontsize=8)
    ax.text(500, _at(t, mn - HYS, 500) - 1.6, "Mn − Hys", fontsize=7.4)
    ax.text(1680, _at(t, mn + HYS, 1680) + 0.7, "Mn + Hys", fontsize=7.4)
    _ttt(ax, enter, trig, -122.5)
    _ttt(ax, leave, term, -122.5)
    _point(ax, trig, _at(t, mn - HYS, trig), "Trigger A4", f"Mn−Hys = {_at(t, mn - HYS, trig):.0f} dBm  >  −105")
    _point(ax, term, _at(t, mn + HYS, term), "Leave A4", f"Mn+Hys = {_at(t, mn + HYS, term):.0f} dBm  <  −105", ha="right")
    _box(ax, [
        "Mn = −90 dBm (enter) / −110 dBm (leave)",
        "Hys = 2 dB    Thresh = −105 dBm    Ofn=Ocn=0    TTT = 320 ms",
        "Enter: Mn − Hys > Thresh     (−90 − 2 = −92 > −105)",
        "Leave: Mn + Hys < Thresh     (−110 + 2 = −108 < −105)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-6  Entering and leaving of event A4   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_6_event_a4.png")


def fig_a5():
    t = np.linspace(0, TMAX, 1200)
    # Serving drops through Thresh1 = −110; neighbour rises through Thresh2 = −105.
    ms = np.interp(t, [0, 400, 720, 1100, 2400], [-95, -112, -115, -115, -115])
    mn = np.interp(t, [0, 450, 820, 1200, 1650, 1970, 2400], [-118, -103, -90, -88, -107, -110, -118])
    th1, th2 = -110.0, -105.0
    enter_s = _cross(t, ms + HYS, th1, "down")
    enter_n = _cross(t, mn - HYS, th2, "up")
    enter = max(enter_s, enter_n)
    leave = _cross(t, mn + HYS, th2, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.55))
    _style(ax, -126, -78)
    _thresh(ax, th1, "Thresh1 = −110 dBm")
    _thresh(ax, th2, "Thresh2 = −105 dBm")
    ax.plot(t, ms, color=BLACK, lw=1.55)
    ax.plot(t, mn, color=BLACK, lw=1.7)
    ax.plot(t, ms + HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn - HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn + HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.text(40, -95 + 1.0, "Ms", fontsize=8)
    ax.text(1220, _at(t, mn, 1220) + 0.8, "Mn", fontsize=8)
    _ttt(ax, enter, trig, -124.2)
    _ttt(ax, leave, term, -124.2)
    _point(ax, trig, _at(t, mn - HYS, trig), "Trigger A5", f"Ms+Hys={_at(t, ms + HYS, trig):.0f}<−110  and  Mn−Hys={_at(t, mn - HYS, trig):.0f}>−105")
    _point(ax, term, _at(t, mn + HYS, term), "Leave A5", f"Mn+Hys = {_at(t, mn + HYS, term):.0f} dBm  <  −105", ha="right")
    _box(ax, [
        "Enter: Ms = −115 dBm, Mn = −90 dBm.  Leave: Mn = −110 dBm.",
        "Hys = 2 dB    Thresh1 = −110    Thresh2 = −105    TTT = 320 ms",
        "Enter: (−115+2=−113 < −110) AND (−90−2=−92 > −105)",
        "Leave: Mn + Hys < Thresh2     (−110 + 2 = −108 < −105)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-7  Entering and leaving of event A5   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_7_event_a5.png")


def fig_b1():
    tps = [0, 480, 800, 1150, 1580, 1900, 2400]
    yps = [-112, -98, -92, -88, -108, -110, -114]
    t, mn = _curve(tps, yps)
    th = -100.0
    enter = _cross(t, mn - HYS, th, "up")
    leave = _cross(t, mn + HYS, th, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.35))
    _style(ax, -120, -78)
    _thresh(ax, th, "Thresh = −100 dBm")
    ax.plot(t, mn, color=BLACK, lw=1.7)
    ax.plot(t, mn - HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn + HYS, color=BLACK, lw=1.05, ls=(0, (2.2, 1.8)))
    ax.text(1180, _at(t, mn, 1180) + 0.8, "Mn (IRAT)", fontsize=8)
    _ttt(ax, enter, trig, -118.4)
    _ttt(ax, leave, term, -118.4)
    _point(ax, trig, _at(t, mn - HYS, trig), "Trigger B1", f"Mn−Hys = {_at(t, mn - HYS, trig):.0f} dBm  >  −100")
    _point(ax, term, _at(t, mn + HYS, term), "Leave B1", f"Mn+Hys = {_at(t, mn + HYS, term):.0f} dBm  <  −100", ha="right")
    _box(ax, [
        "IRAT Mn = −92 (enter) / −110 (leave).  Inter-RAT neighbour.",
        "Hys = 2 dB    Thresh = −100 dBm    Ofn = 0    TTT = 320 ms",
        "Enter: Mn − Hys > Thresh     (−92 − 2 = −94 > −100)",
        "Leave: Mn + Hys < Thresh     (−110 + 2 = −108 < −100)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-8  Entering and leaving of event B1   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_8_event_b1.png")


def fig_b2():
    t = np.linspace(0, TMAX, 1200)
    ms = np.interp(t, [0, 400, 720, 1100, 2400], [-95, -112, -115, -115, -115])
    mn = np.interp(t, [0, 450, 820, 1200, 1650, 1970, 2400], [-114, -98, -92, -88, -108, -110, -116])
    th1, th2 = -110.0, -100.0
    enter_s = _cross(t, ms + HYS, th1, "down")
    enter_n = _cross(t, mn - HYS, th2, "up")
    enter = max(enter_s, enter_n)
    leave = _cross(t, mn + HYS, th2, "down")
    trig, term = enter + TTT, leave + TTT
    fig, ax = plt.subplots(figsize=(11.2, 4.55))
    _style(ax, -124, -78)
    _thresh(ax, th1, "Thresh1 = −110 dBm")
    _thresh(ax, th2, "Thresh2 = −100 dBm")
    ax.plot(t, ms, color=BLACK, lw=1.55)
    ax.plot(t, mn, color=BLACK, lw=1.7)
    ax.plot(t, ms + HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn - HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.plot(t, mn + HYS, color=BLACK, lw=0.95, ls=(0, (2.2, 1.8)))
    ax.text(40, -95 + 1.0, "Ms", fontsize=8)
    ax.text(1220, _at(t, mn, 1220) + 0.8, "Mn (IRAT)", fontsize=8)
    _ttt(ax, enter, trig, -122.4)
    _ttt(ax, leave, term, -122.4)
    _point(ax, trig, _at(t, mn - HYS, trig), "Trigger B2", f"Ms+Hys={_at(t, ms + HYS, trig):.0f}<−110  and  Mn−Hys={_at(t, mn - HYS, trig):.0f}>−100")
    _point(ax, term, _at(t, mn + HYS, term), "Leave B2", f"Mn+Hys = {_at(t, mn + HYS, term):.0f} dBm  <  −100", ha="right")
    _box(ax, [
        "Enter: Ms = −115 dBm, IRAT Mn = −92.  Leave: Mn = −110.",
        "Hys = 2 dB    Thresh1 = −110    Thresh2 = −100    TTT = 320 ms",
        "Enter: (−115+2=−113 < −110) AND (−92−2=−94 > −100)",
        "Leave: Mn + Hys < Thresh2     (−110 + 2 = −108 < −100)",
    ])
    fig.subplots_adjust(left=0.08, right=0.98, top=0.97, bottom=0.16)
    _caption(fig, "Figure 4-9  Entering and leaving of event B2   ·   teaching values marked on the trigger / leave points")
    return _save(fig, "fig_4_9_event_b2.png")


def main():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "axes.unicode_minus": False,
        "figure.facecolor": "white",
        "savefig.facecolor": "white",
    })
    fig_a1()
    fig_a2()
    fig_a3()
    fig_a4()
    fig_a5()
    fig_b1()
    fig_b2()


if __name__ == "__main__":
    main()
