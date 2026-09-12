# -*- coding: utf-8 -*-
"""Smoke-test the benchmark skill scripts in a temp sandbox."""
import io
import os
import shutil
import subprocess
import sys

SKILL = r"C:\Users\Administrator\.zcode\skills\ppt-skill-benchmark\scripts"
EVAL = r"C:\Users\Administrator\ZCodeProject\eval"
TMP = os.path.join(EVAL, "_skilltest")
shutil.rmtree(TMP, ignore_errors=True)
os.makedirs(TMP)

log = []

# 1) verify_pptx.py
r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(SKILL, "verify_pptx.py"),
                    os.path.join(EVAL, "A7_pptmaster", "mode1_text", "deck18.pptx")],
                   capture_output=True, text=True, encoding="utf-8")
log.append("verify_pptx: rc=%d %s" % (r.returncode, r.stdout.strip().splitlines()[-1] if r.stdout else r.stderr[:200]))

# 2) render_pptx_pages.py (small width for speed)
r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(SKILL, "render_pptx_pages.py"),
                    os.path.join(EVAL, "A7_pptmaster", "mode1_text", "deck18.pptx"),
                    "--outdir", os.path.join(TMP, "render"), "--width", "640"],
                   capture_output=True, text=True, encoding="utf-8")
log.append("render: rc=%d %s" % (r.returncode, (r.stdout or r.stderr).strip()[:120]))

# 3) make_collages.py
r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(SKILL, "make_collages.py"),
                    os.path.join(TMP, "render"), "--out", os.path.join(TMP, "collage.png")],
                   capture_output=True, text=True, encoding="utf-8")
log.append("collage: rc=%d %s" % (r.returncode, (r.stdout or r.stderr).strip()[:120]))

# 4) make_charts.py
csvp = os.path.join(TMP, "scores.csv")
io.open(csvp, "w", encoding="utf-8-sig").write(
    "skill,fakeo,text\nGordenPPTSkill,4.22,4.39\nppt-master,3.00,2.94\nclaude-pptx,2.72,2.83\n")
r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(SKILL, "make_charts.py"),
                    "--scores", csvp, "--outdir", TMP],
                   capture_output=True, text=True, encoding="utf-8")
log.append("charts: rc=%d %s | files=%s" % (r.returncode, (r.stdout or r.stderr).strip()[:80],
         os.path.exists(os.path.join(TMP, "bar_ranking.png")) and os.path.exists(os.path.join(TMP, "quadrant.png"))))

# 5) build_workbench.py on a minimal sandbox workbench
wb = os.path.join(TMP, "workbench")
os.makedirs(wb, exist_ok=True)
io.open(os.path.join(wb, "manifest.js"), "w", encoding="utf-8").write(
    'window.WB_DATA = {"items": []}')
io.open(os.path.join(wb, "answers.js"), "w", encoding="utf-8").write(
    'window.WB_ANSWERS = {};')
r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(SKILL, "build_workbench.py"),
                    "--workbench", wb, "--render-dir", os.path.join(TMP, "render"),
                    "--track", "文字赛道", "--skill", "demo-skill"],
                   capture_output=True, text=True, encoding="utf-8")
log.append("workbench: rc=%d %s" % (r.returncode, (r.stdout or r.stderr).strip()[:160]))

io.open(os.path.join(EVAL, "pm_skilltest.txt"), "w", encoding="utf-8").write("\n".join(log))
print("ok")
