#!/usr/bin/env python3
"""Huawei-FPD-style entering/leaving schematics for Events A2–A5, B1, B2.

Event A1 keeps the cropped document Figure 4-3. These figures follow the same
layout: RSRP/RSRQ vs Time, Thresh, measurement ± Hys, TimeToTrig, trigger and
termination marks.
"""

import os

import matplotlib.pyplot as plt
import numpy as np
OUT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "docs/4G_LTE_Mobility_Management/figures",
)

BLACK = "#111111"
DPI = 120


def _cross(t, y, level, direction):
    s = np.sign(y - level)
    d = np.diff(s)
    hits = np.where(d > 0)[0] if direction == "up" else np.where(d < 0)[0]
    if len(hits) == 0:
        return None
    i = int(hits[0])
    y0, y1 = y[i] - level, y[i + 1] - level
    if y1 == y0:
        return float(t[i])
    frac = -y0 / (y1 - y0)
    return float(t[i] + frac * (t[i + 1] - t[i]))


def _interp(t, y, x):
    return float(np.interp(x, t, y))


def _axes(ax):
    ax.set_xlim(0, 10)
    ax.set_ylim(-1.85, 2.15)
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_color(BLACK)
        sp.set_linewidth(1.0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_ylabel("RSRP/RSRQ", color=BLACK, fontsize=10, labelpad=8)
    ax.text(10.12, -1.85, "Time", ha="left", va="top", fontsize=10, color=BLACK)
    ax.tick_params(length=0)


def _thresh(ax, y, label, x=0.18):
    ax.plot([0.15, 9.7], [y, y], ls=(0, (4, 3)), lw=1.05, color=BLACK)
    ax.text(x, y + 0.10, label, fontsize=9, color=BLACK, va="bottom")


def _ttt(ax, x0, x1, y=-1.58, label="TimeToTrig"):
    ax.annotate(
        "",
        xy=(x1, y),
        xytext=(x0, y),
        arrowprops=dict(arrowstyle="<->", color=BLACK, lw=0.9),
    )
    ax.text((x0 + x1) / 2, y + 0.10, label, ha="center", va="bottom", fontsize=8, color=BLACK)


def _vmark(ax, x, ymin=-1.72, ymax=1.95):
    ax.plot([x, x], [ymin, ymax], ls=(0, (1.5, 2)), lw=0.9, color=BLACK)


def _caption(fig, text):
    fig.text(0.50, 0.035, text, ha="center", va="bottom", fontsize=10, color=BLACK)


def _save(fig, name):
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name)
    fig.savefig(path, dpi=DPI, facecolor="white", bbox_inches="tight", pad_inches=0.18)
    plt.close(fig)
    from PIL import Image
    im = Image.open(path).convert("RGB")
    target_w = 1100
    if im.size[0] != target_w:
        nh = int(round(im.size[1] * target_w / im.size[0]))
        im = im.resize((target_w, nh), Image.Resampling.LANCZOS)
        im.save(path, "PNG")
    print("wrote", path, im.size)
    return path


def _hill(t, peak=4.15, amp=2.15, base=-0.95, width=1.55):
    return base + amp * np.exp(-0.5 * ((t - peak) / width) ** 2)


def fig_threshold_event(name, fig_no, meas, enter_dir, extra_label=None):
    """A1-style: one measurement vs one Thresh. enter_dir 'up' (A1/A4/B1) or 'down' (A2)."""
    t = np.linspace(0, 10, 900)
    if enter_dir == "up":
        y = _hill(t)
    else:
        y = -_hill(t, peak=4.15, amp=2.15, base=-0.95, width=1.55)
    hys = 0.28
    thresh = 0.0
    ttt = 0.85

    if enter_dir == "up":
        enter = _cross(t, y - hys, thresh, "up")
        leave = _cross(t, y + hys, thresh, "down")
        enter_curve, leave_curve = y - hys, y + hys
        enter_lbl, leave_lbl = f"{meas}-Hys", f"{meas}+Hys"
    else:
        enter = _cross(t, y + hys, thresh, "down")
        leave = _cross(t, y - hys, thresh, "up")
        enter_curve, leave_curve = y + hys, y - hys
        enter_lbl, leave_lbl = f"{meas}+Hys", f"{meas}-Hys"

    trig = enter + ttt
    term = leave + ttt

    fig, ax = plt.subplots(figsize=(11.0, 3.05))
    _axes(ax)
    _thresh(ax, thresh, "Thresh")
    ax.plot(t, y, color=BLACK, lw=1.55, solid_capstyle="round")
    ax.plot(t, y - hys, color=BLACK, lw=1.05, ls=(0, (2.2, 2.0)))
    ax.plot(t, y + hys, color=BLACK, lw=1.05, ls=(0, (2.2, 2.0)))

    ax.text(5.55, _interp(t, y, 5.55) + 0.12, meas, fontsize=9, color=BLACK)
    ax.text(2.05, _interp(t, enter_curve, 2.05) + (-0.28 if enter_dir == "up" else 0.14), enter_lbl, fontsize=8, color=BLACK)
    ax.text(6.55, _interp(t, leave_curve, 6.55) + (0.12 if enter_dir == "up" else -0.28), leave_lbl, fontsize=8, color=BLACK)
    if extra_label:
        ax.text(0.25, 1.95, extra_label, fontsize=8, color=BLACK, va="top")

    _ttt(ax, enter, trig)
    _ttt(ax, leave, term)
    _vmark(ax, trig)
    _vmark(ax, term)
    ax.text(trig, -1.82, f"Triggering of event {name}", ha="center", va="top", fontsize=8.5)
    ax.text(term, -1.82, f"Termination of event {name} reporting", ha="center", va="top", fontsize=8.5)
    fig.subplots_adjust(left=0.07, right=0.97, top=0.92, bottom=0.22)
    _caption(fig, f"Figure {fig_no}  Entering and leaving of event {name}")
    return _save(fig, f"fig_{fig_no.replace('-', '_').replace(' ', '_').lower()}_event_{name.lower()}.png")


def fig_a3():
    t = np.linspace(0, 10, 900)
    ms = 0.05 - 0.08 * (t / 10.0)
    off = 0.35
    serving = ms + off
    mn = _hill(t, peak=4.2, amp=2.05, base=-1.05, width=1.50)
    hys = 0.26
    ttt = 0.85
    left = mn - hys
    right = serving
    enter = _cross(t, left - right, 0.0, "up")
    leave = _cross(t, (mn + hys) - serving, 0.0, "down")
    trig, term = enter + ttt, leave + ttt

    fig, ax = plt.subplots(figsize=(11.0, 3.25))
    _axes(ax)
    ax.plot(t, serving, color=BLACK, lw=1.05, ls=(0, (4, 3)))
    ax.plot(t, ms, color=BLACK, lw=1.35)
    ax.plot(t, mn, color=BLACK, lw=1.55)
    ax.plot(t, mn - hys, color=BLACK, lw=1.0, ls=(0, (2.2, 2.0)))
    ax.plot(t, mn + hys, color=BLACK, lw=1.0, ls=(0, (2.2, 2.0)))
    ax.text(0.35, _interp(t, ms, 0.35) + 0.12, "Ms", fontsize=9)
    ax.text(0.35, _interp(t, serving, 0.35) + 0.12, "Ms + Ofs + Ocs + Off", fontsize=8)
    ax.text(5.45, _interp(t, mn, 5.45) + 0.12, "Mn", fontsize=9)
    ax.text(2.15, _interp(t, mn - hys, 2.15) - 0.28, "Mn − Hys", fontsize=8)
    ax.text(6.55, _interp(t, mn + hys, 6.55) + 0.10, "Mn + Hys", fontsize=8)
    _ttt(ax, enter, trig)
    _ttt(ax, leave, term)
    _vmark(ax, trig)
    _vmark(ax, term)
    ax.text(trig, -1.82, "Triggering of event A3", ha="center", va="top", fontsize=8.5)
    ax.text(term, -1.82, "Termination of event A3 reporting", ha="center", va="top", fontsize=8.5)
    fig.subplots_adjust(left=0.07, right=0.97, top=0.90, bottom=0.22)
    _caption(fig, "Figure 4-5  Entering and leaving of event A3")
    return _save(fig, "fig_4_5_event_a3.png")


def fig_two_thresh(name, fig_no, neigh_label):
    """A5 / B2: serving vs Thresh1 and neighbour vs Thresh2."""
    t = np.linspace(0, 10, 900)
    ms = 1.15 - 1.55 / (1.0 + np.exp(-(t - 2.15) * 2.6))
    mn = _hill(t, peak=4.45, amp=2.20, base=-1.20, width=1.55)
    hys = 0.24
    th1, th2 = -0.10, 0.55
    ttt = 0.80
    enter_s = _cross(t, ms + hys, th1, "down")
    enter_n = _cross(t, mn - hys, th2, "up")
    enter = max(x for x in (enter_s, enter_n) if x is not None)
    leave = _cross(t, mn + hys, th2, "down")
    if leave is None or leave < enter + 0.2:
        raise RuntimeError(f"{name}: could not find leaving crossing")
    trig, term = enter + ttt, leave + ttt

    fig, ax = plt.subplots(figsize=(11.0, 3.35))
    _axes(ax)
    _thresh(ax, th1, "Thresh1")
    _thresh(ax, th2, "Thresh2", x=0.18)
    ax.plot(t, ms, color=BLACK, lw=1.45)
    ax.plot(t, mn, color=BLACK, lw=1.55)
    ax.plot(t, ms + hys, color=BLACK, lw=0.95, ls=(0, (2.2, 2.0)))
    ax.plot(t, mn - hys, color=BLACK, lw=0.95, ls=(0, (2.2, 2.0)))
    ax.plot(t, mn + hys, color=BLACK, lw=0.95, ls=(0, (2.2, 2.0)))
    ax.text(0.30, _interp(t, ms, 0.30) + 0.12, "Ms", fontsize=9)
    ax.text(5.55, _interp(t, mn, 5.55) + 0.12, neigh_label, fontsize=9)
    ax.text(1.55, _interp(t, ms + hys, 1.55) + 0.10, "Ms + Hys", fontsize=8)
    ax.text(2.35, _interp(t, mn - hys, 2.35) - 0.26, f"{neigh_label} − Hys", fontsize=8)
    ax.text(6.45, _interp(t, mn + hys, 6.45) + 0.10, f"{neigh_label} + Hys", fontsize=8)
    _ttt(ax, enter, trig)
    _ttt(ax, leave, term)
    _vmark(ax, trig)
    _vmark(ax, term)
    ax.text(trig, -1.82, f"Triggering of event {name}", ha="center", va="top", fontsize=8.5)
    ax.text(term, -1.82, f"Termination of event {name} reporting", ha="center", va="top", fontsize=8.5)
    fig.subplots_adjust(left=0.07, right=0.97, top=0.90, bottom=0.22)
    _caption(fig, f"Figure {fig_no}  Entering and leaving of event {name}")
    fname = f"fig_{fig_no.replace('-', '_')}_event_{name.lower()}.png"
    return _save(fig, fname)


def main():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "axes.unicode_minus": False,
        "figure.facecolor": "white",
        "savefig.facecolor": "white",
    })
    fig_threshold_event("A2", "4-4", "Ms", "down")
    fig_a3()
    fig_threshold_event("A4", "4-6", "Mn", "up")
    fig_two_thresh("A5", "4-7", "Mn")
    fig_threshold_event("B1", "4-8", "Mn", "up", extra_label="Inter-RAT neighbouring cell")
    fig_two_thresh("B2", "4-9", "Mn")


if __name__ == "__main__":
    main()
