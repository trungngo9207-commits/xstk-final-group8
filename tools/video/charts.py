"""Static chart images for the B1/B2 guide video (matplotlib -> PNG)."""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "charts")
os.makedirs(OUT, exist_ok=True)

for f in ["/usr/share/fonts/truetype/ubuntu/Ubuntu-R.ttf", "/usr/share/fonts/truetype/ubuntu/Ubuntu-B.ttf"]:
    font_manager.fontManager.addfont(f)

SURFACE = "#fcfcfb"
INK = "#0b0b0b"
INK2 = "#52514e"
MUTED = "#898781"
GRID = "#e1e0d9"
AXIS = "#c3c2b7"
NAPA = "#2a78d6"   # categorical slot 1
OTHER = "#eb6834"  # categorical slot 2

plt.rcParams.update({
    "font.family": ["Ubuntu", "DejaVu Sans"],
    "font.size": 15,
    "axes.edgecolor": AXIS,
    "axes.labelcolor": INK2,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "axes.facecolor": SURFACE,
    "figure.facecolor": SURFACE,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.color": GRID,
    "grid.linewidth": 1,
    "axes.axisbelow": True,
})

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
d = pd.read_excel(os.path.join(REPO, "data", "Wine_Dataset.xlsx"))
price = d["Price"].astype(float).values
napa = d["Napa valley"].values
n = len(price)
mean = price.mean()
se = price.std(ddof=1) / np.sqrt(n)
tcrit = stats.t.ppf(0.975, n - 1)
lo, hi = mean - tcrit * se, mean + tcrit * se


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=100, facecolor=SURFACE)
    plt.close(fig)


# 1. CI number line with the H0 value -------------------------------------------------
fig, ax = plt.subplots(figsize=(4.6, 2.6))
ax.grid(False)
ax.spines["left"].set_visible(False)
ax.set_yticks([])
ax.set_xlim(22.5, 34)
ax.set_ylim(-1, 1.6)
ax.plot([lo, hi], [0, 0], color=NAPA, lw=6, solid_capstyle="round")
ax.plot([mean], [0], "o", ms=11, color=NAPA, mec=SURFACE, mew=2)
ax.axvline(25, color=INK2, lw=2, ls=(0, (4, 3)))
ax.text(25.25, 1.15, "H₀ value \\$25", ha="left", va="center", color=INK, fontsize=14)
ax.text(mean, 0.55, f"95% CI\n\\${lo:.2f} – \\${hi:.2f}", ha="center", va="center", color=INK, fontsize=14)
ax.text(mean, -0.6, f"x̄ = \\${mean:.2f}", ha="center", va="center", color=INK2, fontsize=13)
ax.set_xlabel("Mean price ($)", fontsize=13)
fig.tight_layout()
save(fig, "ci.png")

# 2. t distribution, observed t and critical values -----------------------------------
tobs = (mean - 25) / se
x = np.linspace(-5, 5, 800)
y = stats.t.pdf(x, n - 1)
fig, ax = plt.subplots(figsize=(4.6, 3.0))
ax.plot(x, y, color=INK2, lw=2)
for side in (-1, 1):
    m = (x * side) >= tcrit
    ax.fill_between(x[m], 0, y[m], color=NAPA, alpha=0.18, lw=0)
    ax.axvline(side * tcrit, color=MUTED, lw=1.5, ls=(0, (4, 3)))
    ax.axvline(side * tobs, color=NAPA, lw=2.5)
ax.text(tobs, 0.47, f"t = {tobs:.2f}", ha="center", color=INK, fontsize=14,
        bbox=dict(boxstyle="round,pad=0.25", fc=SURFACE, ec="none"))
ax.text(-tobs, 0.47, f"−{tobs:.2f}", ha="center", color=INK, fontsize=14,
        bbox=dict(boxstyle="round,pad=0.25", fc=SURFACE, ec="none"))
ax.text(0, 0.445, "t values if H₀ is true", ha="center", color=INK2, fontsize=12.5, bbox=dict(boxstyle="round,pad=0.2", fc=SURFACE, ec="none"))
ax.text(tcrit + 0.1, 0.2, "1.963", ha="left", color=MUTED, fontsize=12)
ax.set_ylim(0, 0.53)
ax.set_xlim(-5, 5)
ax.set_yticks([])
ax.spines["left"].set_visible(False)
ax.grid(False)
ax.set_xlabel("t statistic (df = 805)", fontsize=13)
fig.tight_layout()
save(fig, "tcurve.png")

# 3. SSE as a function of b: minimum at the mean --------------------------------------
b = np.linspace(10, 50, 400)
sse = ((price[:, None] - b[None, :]) ** 2).sum(axis=0) / 1e6
fig, ax = plt.subplots(figsize=(8.6, 5.2))
ax.plot(b, sse, color=NAPA, lw=2.5)
smin = ((price - mean) ** 2).sum() / 1e6
ax.plot([mean], [smin], "o", ms=11, color=NAPA, mec=SURFACE, mew=2.5, zorder=5)
ax.annotate(f"lowest point: b = x̄ = {mean:.2f}", xy=(mean, smin), xytext=(mean, smin + 0.11),
            fontsize=17, color=INK, ha="center", arrowprops=dict(arrowstyle="-", color=INK2, lw=1.2))
ax.set_xlabel("Candidate value for β₀ (dollars)")
ax.set_ylabel("Sum of squared errors  Σ(Price − b)²  (millions)")
ax.set_xlim(10, 50)
fig.tight_layout()
save(fig, "sse.png")

# 4/5. Strip plots of the two groups ---------------------------------------------------
rng = np.random.default_rng(7)
groups = [(0, "Other wines", OTHER, price[napa == 0]), (1, "Napa Valley", NAPA, price[napa == 1])]


def strip(name, ymax, with_line):
    fig, ax = plt.subplots(figsize=(8.6, 6.0))
    ax.grid(axis="x", visible=False)
    for xpos, label, col, vals in groups:
        shown = vals[vals <= ymax]
        jit = rng.uniform(-0.17, 0.17, len(shown))
        ax.scatter(xpos + jit, shown, s=22, color=col, alpha=0.45, lw=0)
        m = vals.mean()
        ax.plot([xpos - 0.25, xpos + 0.25], [m, m], color=INK, lw=3, solid_capstyle="round", zorder=4)
        if xpos == 0:
            ax.text(xpos - 0.29, m, f"mean \\${m:.2f}", va="center", ha="right", fontsize=15, color=INK, zorder=5)
        else:
            ax.text(xpos + 0.36, m, f"mean \\${m:.2f}", va="center", fontsize=15, color=INK, zorder=5)
    if with_line:
        m0, m1 = groups[0][3].mean(), groups[1][3].mean()
        ax.plot([0, 1], [m0, m1], color=INK2, lw=2, ls=(0, (5, 3)), zorder=3)
        ax.plot([0, 1.3], [m0, m0], color=MUTED, lw=1, ls=":", zorder=2)
        ax.annotate("", xy=(1.3, m1), xytext=(1.3, m0),
                    arrowprops=dict(arrowstyle="<->", color=INK, lw=1.6))
        ax.text(0.5, 72, f"slope = {m1:.2f} − {m0:.2f}\n= {m1 - m0:.2f} dollars",
                ha="center", va="center", fontsize=16, color=INK,
                bbox=dict(boxstyle="round,pad=0.35", fc=SURFACE, ec=AXIS))
    hidden = (price > ymax).sum()
    if hidden:
        ax.text(1.0, ymax * 0.97, f"{hidden} Napa wines above \\${ymax:.0f} not shown", ha="center",
                va="top", fontsize=12.5, color=MUTED)
    ax.set_xticks([0, 1])
    ax.set_xticklabels([f"Other wines\n(Napa = 0, n = {len(groups[0][3])})",
                        f"Napa Valley\n(Napa = 1, n = {len(groups[1][3])})"], fontsize=15, color=INK2)
    ax.set_xlim(-0.95, 1.75)
    ax.set_ylim(0, ymax)
    ax.set_ylabel("Price per bottle ($)")
    fig.tight_layout()
    save(fig, name)


strip("strip_full.png", 460, False)
strip("strip_reg.png", 150, True)

# 6. Histogram of Price -----------------------------------------------------------------
fig, ax = plt.subplots(figsize=(8.6, 5.0))
ax.hist(price, bins=np.arange(0, 460, 10), color=NAPA, edgecolor=SURFACE, linewidth=1.5)
ax.grid(axis="x", visible=False)
ax.axvline(np.median(price), color=INK, lw=1.5, ls=(0, (4, 3)))
ax.axvline(mean, color=INK2, lw=1.5)
ax.text(60, 240, f"dashed line: median \\${np.median(price):.0f}", ha="left", fontsize=15, color=INK)
ax.text(60, 212, f"solid line: mean \\${mean:.2f}", ha="left", fontsize=15, color=INK)
ax.annotate("long right tail:\na few bottles cost \\$300–440", xy=(390, 4), xytext=(320, 80),
            fontsize=15, color=INK, ha="center", arrowprops=dict(arrowstyle="->", color=INK2, lw=1.3))
ax.set_xlabel("Price per bottle ($)")
ax.set_ylabel("Number of wines")
ax.set_xlim(0, 450)
fig.tight_layout()
save(fig, "hist.png")

print("charts written to", OUT, "| mean", round(mean, 4), "CI", round(lo, 4), round(hi, 4), "t", round(tobs, 4))
