"""Rebuild Huawei eRAN21.1 mobility figures from the feature-book procedures.

The original PDF bitmaps are not in this workspace. These charts follow the
documented flows: Idle Fig 4-1 / 5-1, Connected Fig 4-1 / 11-1, MLB Fig 3-1 /
4-1 / 4-4 / 4-5.
"""

import os
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Polygon

FIGDIR = "/workspace/docs/4G_LTE_Mobility_Management/figures"
BLUE = "#005596"
GREEN = "#008000"
GOLD = "#C65911"
GREY = "#F2F2F2"
DARK = "#1F4E79"


def _setup(w=13.5, h=3.2):
    fig, ax = plt.subplots(figsize=(w, h), dpi=140)
    ax.set_xlim(0, w)
    ax.set_ylim(0, h)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    return fig, ax


def _box(ax, x, y, w, h, text, fc=BLUE, tc="white", size=8):
    p = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.12",
                       linewidth=0.8, edgecolor=DARK, facecolor=fc)
    ax.add_patch(p)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", color=tc,
            fontsize=size, fontweight="bold", wrap=True)


def _diamond(ax, x, y, w, h, text, fc="#FFF2CC"):
    cx, cy = x + w / 2, y + h / 2
    pts = [(cx, y + h), (x + w, cy), (cx, y), (x, cy)]
    ax.add_patch(Polygon(pts, closed=True, facecolor=fc, edgecolor=DARK, linewidth=0.8))
    ax.text(cx, cy, text, ha="center", va="center", fontsize=7.5, fontweight="bold", color=GREEN)


def _arrow(ax, x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=12,
                                 linewidth=1.4, color=GREEN, shrinkA=0, shrinkB=0))


def _title(ax, text):
    ymin, ymax = ax.get_ylim()
    ax.text(0.15, ymax - 0.12, text, fontsize=10, fontweight="bold", color=BLUE, ha="left", va="top")


def _save(fig, name):
    os.makedirs(FIGDIR, exist_ok=True)
    path = os.path.join(FIGDIR, name)
    fig.tight_layout(pad=0.25)
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def hflow(title, boxes, fname, colors=None):
    fig, ax = _setup()
    _title(ax, title)
    n = len(boxes)
    bw = min(1.85, 11.8 / n - 0.25)
    gap = 0.28
    total = n * bw + (n - 1) * gap
    x0 = (13.5 - total) / 2
    y = 0.85
    h = 1.35
    for i, t in enumerate(boxes):
        fc = (colors[i] if colors else BLUE)
        x = x0 + i * (bw + gap)
        _box(ax, x, y, bw, h, t, fc=fc, size=7.4)
        if i < n - 1:
            _arrow(ax, x + bw, y + h / 2, x + bw + gap, y + h / 2)
    return _save(fig, fname)


def build_all():
    paths = {}
    paths["chain"] = hflow(
        "Document chain  |  Idle Fig 5-1  →  Connected Fig 4-1  →  MLB Fig 3-1",
        ["Idle\ncamp / reselect", "RRC\nconnect", "Coverage\nA2 / A5", "FreqPri\nA1 / A4", "MLB\nload HO / idle", "Next idle\nT320"],
        "fig_chain.png",
        [BLUE, DARK, GREEN, DARK, GOLD, BLUE],
    )
    paths["idle_41"] = hflow(
        "Idle Mode Management  Fig 4-1  — idle functions",
        ["PLMN\nselect", "Cell\nselection", "Cell\nreselection", "Dedicated\npriority", "SI\nbroadcast", "Idle MLB\ninterface"],
        "fig_idle_4_1.png",
    )
    paths["idle_51"] = hflow(
        "Idle Mode Management  Fig 5-1  — selection / reselection sequence",
        ["Power-on /\nstored cell", "Criterion S\nSrxlev / Squal", "Camp +\nread SIB3/5", "Measure\nintra / inter", "Reselect\nhigh/eq/low", "RRC on\ncamped cell"],
        "fig_idle_5_1.png",
    )
    # Reselection decision — Tables 5-1 higher, 5-2 equal, 5-3/5-4 lower
    fig, ax = _setup(13.5, 4.4)
    _title(ax, "Idle Mode Management  Tables 5-1 to 5-4  — reselection decision")
    _box(ax, 0.25, 1.55, 2.3, 1.3, "Camped > 1 s\nthen compare\nCellReselPriority", size=8)
    _arrow(ax, 2.55, 2.2, 3.15, 2.2)
    _diamond(ax, 3.15, 1.4, 2.6, 1.6, "Target\npriority?")
    _arrow(ax, 5.75, 2.85, 6.35, 3.15)
    _box(ax, 6.4, 2.7, 3.3, 1.25, "HIGHER  |  Table 5-1\nThreshXhigh for Treselection\n→ reselect higher layer", fc=GREEN, size=7.4)
    _arrow(ax, 5.75, 2.2, 6.35, 2.05)
    _box(ax, 6.4, 1.45, 3.3, 1.15, "EQUAL  |  Table 5-2\nRn > Rs  (Qhyst / QoffsetFreq)", fc=DARK, size=7.4)
    _arrow(ax, 4.45, 1.4, 6.35, 0.85)
    _box(ax, 6.4, 0.2, 6.7, 1.1, "LOWER  |  Tables 5-3 / 5-4   serving < ThrshServLow  AND  target > ThreshXlow\n(only if no higher-priority candidate already qualifies)", fc=GOLD, size=7.4)
    paths["idle_reselect"] = _save(fig, "fig_idle_reselect.png")

    paths["conn_41"] = hflow(
        "Mobility Management in Connected Mode  Fig 4-1  — HO procedure §§4.1.1–4.1.8",
        ["Start HO\nfunction", "Meas or\nblind", "Deliver\nmeas config", "UE report\nA1–A5", "Pick target\n+ admit", "Execute HO\n/ punish"],
        "fig_connected_4_1.png",
    )
    fig, ax = _setup()
    _title(ax, "Connected Mode  A1–A5  (Tables 4-8, 5-16, 5-22, 5-18)  — RSRP recommended")
    items = [
        ("A1", "Serving good\nstop coverage meas\nstart FreqPri", GREEN),
        ("A2", "Serving poor\nstart inter-freq\nmeas (by family)", GOLD),
        ("A3", "Neighbour\nrelatively better\nOfn/Ocn", DARK),
        ("A4", "Neighbour\nabsolutely good\nMLB / FreqPri gate", BLUE),
        ("A5", "Serving poor\nAND target good\ncoverage protect", GREEN),
    ]
    for i, (k, t, fc) in enumerate(items):
        x = 0.4 + i * 2.6
        _box(ax, x, 1.55, 0.7, 0.7, k, fc=fc, size=10)
        _box(ax, x + 0.75, 0.7, 1.7, 1.7, t, fc=GREY, tc="black", size=7)
        if i < 4:
            _arrow(ax, x + 2.45, 1.55, x + 2.55, 1.55)
    paths["conn_events"] = _save(fig, "fig_connected_a1_a5.png")

    fig, ax = _setup()
    _title(ax, "Connected Mode  Fig 11-1 / 11-2  — frequency-priority HO (not MLB)")
    _box(ax, 0.4, 0.7, 3.4, 1.8, "Serving good (A1)\nKeep low band\nfor coverage", fc=GREEN, size=8)
    _arrow(ax, 3.9, 1.6, 4.6, 1.6)
    _box(ax, 4.6, 0.7, 3.6, 1.8, "High-priority freq\nA4 target good enough\nplace service on high band", size=8)
    _arrow(ax, 8.3, 1.6, 9.0, 1.6)
    _box(ax, 9.0, 0.7, 4.0, 1.8, "No reverse MLB pair\nA4 TTT ≠ 5120 ms\nMlbBasedFreqPriHoSwitch", fc=GOLD, size=8)
    paths["conn_freqpri"] = _save(fig, "fig_connected_11.png")

    paths["mlb_31"] = hflow(
        "Intra-RAT MLB  Fig 3-1  — load balancing procedure",
        ["Eval load\nN / C", "Trigger\nthd + offset", "Admit\ntarget", "Select\nUEs", "A4/A5 HO\nor idle release", "Penalty\n/ stop"],
        "fig_mlb_3_1.png",
        [BLUE, GOLD, DARK, BLUE, GREEN, GOLD],
    )
    fig, ax = _setup()
    _title(ax, "Intra-RAT MLB  Fig 4-1  — equalisation vs offload")
    _diamond(ax, 0.35, 0.7, 3.3, 1.8, "Peer load\navailable?")
    ax.text(3.85, 2.35, "YES", fontsize=8, fontweight="bold", color=GREEN, ha="left")
    _arrow(ax, 3.65, 2.05, 4.45, 2.25)
    _box(ax, 4.5, 1.7, 3.8, 1.1, "Equalisation\nuse source and target load", fc=GREEN, size=8)
    ax.text(3.85, 0.85, "NO", fontsize=8, fontweight="bold", color=GOLD, ha="left")
    _arrow(ax, 3.65, 1.15, 4.45, 0.9)
    _box(ax, 4.5, 0.35, 3.8, 1.1, "Offload\nno full target-load check", fc=GOLD, size=8)
    _arrow(ax, 8.4, 2.25, 9.2, 2.25)
    _arrow(ax, 8.4, 0.9, 9.2, 0.9)
    _box(ax, 9.2, 0.5, 3.8, 2.2, "Then admit target\nselect UEs\nA4/A5 or idle T320", size=8)
    paths["mlb_41"] = _save(fig, "fig_mlb_4_1.png")

    fig, ax = _setup()
    _title(ax, "Intra-RAT MLB  Figs 4-4 / 4-5  — idle transfer vs connected transfer")
    _box(ax, 0.3, 1.7, 6.2, 1.15, "Idle: RRC release + dedicated priority + T320\nnext session only  |  less gap/HO cost", fc=BLUE, size=8)
    _box(ax, 0.3, 0.35, 6.2, 1.15, "Connected: measurement-based HO A4/A5\nmoves the UE now  |  HO cause Reduce Load", fc=GOLD, size=8)
    _box(ax, 7.0, 0.35, 6.0, 2.5, "Class order (idle):\nNG-RAN\n> E-UTRAN low-load\n> E-UTRAN high-load\n> UTRAN > GERAN", fc=GREEN, size=8)
    paths["mlb_idle_conn"] = _save(fig, "fig_mlb_4_4_4_5.png")
    return paths


if __name__ == "__main__":
    print(build_all())
