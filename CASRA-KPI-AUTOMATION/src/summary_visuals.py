import matplotlib.pyplot as plt

dates = ["Nov-Dec 2025", "Jan-Feb 2026", "Mar-Apr 2026", "May-June 2026"]
outflow = [1227, 1318, 1166, 1058]
backlog = [1008, 734, 674, 729]

pct_change = [None] + [
    (backlog[i] - backlog[i - 1]) / backlog[i - 1] * 100
    for i in range(1, len(backlog))
]

bar_color = "#000000"
line_color = "#F3CB02"
bubble_face = "#B7E4C7"
bubble_edge = "#5FA777"
bubble_text = "#1B4332"

fig, ax = plt.subplots(figsize=(10, 6))
x = list(range(len(dates)))

# Bars (bottom layer)
ax.bar(x, outflow, width=0.5, color=bar_color, label="PNQ Outflow", zorder=2)
for xi, val in zip(x, outflow):
    ax.text(xi, val + 25, f"{val:,}", ha="center", va="bottom",
            fontsize=10, fontweight="bold")

ax.set_ylim(0, 3400)  # bars occupy lower half of the plot

# Line on a secondary axis, scaled to float above the bars
ax2 = ax.twinx()
ax2.plot(x, backlog, color=line_color, linewidth=3, marker="o",
         markersize=9, markerfacecolor=line_color, markeredgecolor="black",
         markeredgewidth=1.2, zorder=4, label="PNQ Backlog")
for xi, val in zip(x, backlog):
    ax2.text(xi, val - 55, f"{val:,}", ha="center", va="top",
             fontsize=10, fontweight="bold", color="#8A7600")

# Backlog range ~674-1008; map it into the upper region of the plot
ax2.set_ylim(-500, 1300)  # pushes the line band into the top half
ax2.axis("off")

# Bubbles + arrows between backlog points (drawn on ax2 coords)
bubble_y = 1180
for i in range(1, len(backlog)):
    xm = (i - 1 + i) / 2
    change = pct_change[i]
    sign = "+" if change >= 0 else ""
    ax2.annotate(
        f"{sign}{change:.0f}%",
        xy=(xm, bubble_y), ha="center", va="center",
        fontsize=11, fontweight="bold", color=bubble_text, zorder=6,
        bbox=dict(boxstyle="round,pad=0.45", facecolor=bubble_face,
                  edgecolor=bubble_edge, linewidth=1.3),
    )
    ax2.annotate("", xy=(i, backlog[i] + 35), xytext=(xm + 0.12, bubble_y - 40),
                 arrowprops=dict(arrowstyle="-|>", color=bubble_edge,
                                 linewidth=1.3, connectionstyle="arc3,rad=-0.2"),
                 zorder=5)
    ax2.annotate("", xy=(i - 1, backlog[i - 1] + 35), xytext=(xm - 0.12, bubble_y - 40),
                 arrowprops=dict(arrowstyle="-", color=bubble_edge,
                                 linewidth=1.3, connectionstyle="arc3,rad=0.2"),
                 zorder=5)

ax.set_xticks(x)
ax.set_xticklabels(dates, fontsize=11)
ax.set_ylabel("PNQ Outflow", fontsize=11)
ax.set_title("PNQ Outflow vs Backlog", fontsize=15, fontweight="bold", pad=15)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

# Combined legend
h1, l1 = ax.get_legend_handles_labels()
h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h2 + h1, l2 + l1, loc="upper center", bbox_to_anchor=(0.5, -0.08),
          ncol=2, frameon=False, fontsize=10)

plt.tight_layout()
plt.savefig("pnq_outflow_backlog_chart.png", dpi=200, bbox_inches="tight")