# -*- coding: utf-8 -*-
"""Benchmark charts: dual-track grouped bar ranking + quadrant map.

Usage: python make_charts.py --scores <scores.csv> [--outdir eval/assets]
scores.csv header: skill,fakeo,text   (fakeo=图片复刻均分, text=纯文字均分, 5 分制)
Rows are plotted in file order (= final ranking order).
"""
import argparse
import csv

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["STKaiti", "KaiTi", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False
RED = "#B01F24"; GOLD = "#F59E0B"; LINE = "#D4D4D2"; SILVER = "#C98A10"

ap = argparse.ArgumentParser()
ap.add_argument("--scores", required=True)
ap.add_argument("--outdir", default=".")
a = ap.parse_args()

rows = list(csv.DictReader(open(a.scores, encoding="utf-8-sig")))
names = [r["skill"] for r in rows]
fu = [float(r["fakeo"]) for r in rows]
wen = [float(r["text"]) for r in rows]

# ---- 图1：分组柱状图 ----
fig, ax = plt.subplots(figsize=(11.6, 5.2), dpi=150)
x = range(len(names)); w = 0.36
b1 = ax.bar([i - w / 2 for i in x], fu, w, label="图片复刻均分", color=RED)
b2 = ax.bar([i + w / 2 for i in x], wen, w, label="纯文字均分", color="#4A4A4A")
for bars in (b1, b2):
    for b in bars:
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.06, "%.2f" % b.get_height(),
                ha="center", fontsize=10, color="#1C1917")
ax.set_xticks(list(x))
ax.set_xticklabels(names, fontsize=9.5)
ax.set_ylim(0, 5); ax.set_ylabel("盲评均分（5 分制）", fontsize=11)
ax.axhline(2.5, color=GOLD, linestyle="--", linewidth=1)
ax.text(len(names) - 0.4, 2.56, "推荐线 2.5", fontsize=9, color=GOLD, ha="right")
ax.legend(loc="upper right", fontsize=11, frameon=False)
ax.spines[["top", "right"]].set_visible(False)
for i in range(3):
    ax.get_xticklabels()[i].set_color(RED); ax.get_xticklabels()[i].set_fontweight("bold")
ax.text(0.01, 0.97, "前 3 名推荐 ／ 其余跌破推荐线或不推荐", transform=ax.transAxes,
        ha="left", va="top", fontsize=10, color="#8A8A8A")
plt.tight_layout()
plt.savefig(a.outdir + "/bar_ranking.png", facecolor="white")
plt.close()

# ---- 图2：四象限定位图 ----
fig, ax = plt.subplots(figsize=(8.6, 6.4), dpi=150)
ax.axvline(2.5, color=LINE, linewidth=1); ax.axhline(2.5, color=LINE, linewidth=1)
ax.add_patch(plt.Rectangle((2.5, 2.5), 2.6, 2.6, facecolor="#FDEAEA", zorder=0))
ax.text(4.9, 4.92, "双优区", fontsize=12, color=RED, ha="right", fontweight="bold")
ax.text(1.15, 4.92, "文字强", fontsize=10, color="#8A8A8A", ha="left")
ax.text(1.15, 1.1, "待改进区", fontsize=10, color="#8A8A8A", ha="left")
for i, (name, xx, yy) in enumerate(zip(names, fu, wen)):
    big = i == 0
    silver = i == 1
    ax.scatter(xx, yy, s=420 if big else 250, color=RED if big else (SILVER if silver else "#8A8A8A"),
               edgecolors="white", linewidths=1.6, zorder=3)
    ax.annotate(name, (xx, yy), textcoords="offset points",
                xytext=(0, -30 if big else 16), ha="center",
                fontsize=10.5 if big else 9.5,
                color=RED if big else ("#9A6A08" if silver else "#5A5A5A"),
                fontweight="bold" if (big or silver) else "normal")
ax.set_xlim(0.8, 5.1); ax.set_ylim(0.8, 5.1)
ax.set_xlabel("图片复刻均分 →", fontsize=11)
ax.set_ylabel("纯文字均分 →", fontsize=11)
ax.spines[["top", "right"]].set_visible(False)
plt.tight_layout()
plt.savefig(a.outdir + "/quadrant.png", facecolor="white")
plt.close()
print("charts saved to", a.outdir)
