# -*- coding: utf-8 -*-
"""Add one deck's rendered pages into a blind-scoring workbench (incremental).

Usage (one run per deck):
  python build_workbench.py --workbench <工作台目录> --render-dir <render_dir> \
      --track 文字赛道 --skill "skill 显示名" [--code ab12x] [--pages 18]

- Generates/uses a random 5-char blind code; copies pages to <workbench>/img/<code>-pNN.png
- Merges manifest.js / answers.js; existing items, order and scores untouched.
- Requires workbench/index.html with incremental order-merge logic (old order kept,
  new ids appended shuffled); if index.html lacks it, add it before use.
"""
import argparse
import json
import os
import random
import shutil

ap = argparse.ArgumentParser()
ap.add_argument("--workbench", required=True)
ap.add_argument("--render-dir", required=True)
ap.add_argument("--track", required=True, choices=["文字赛道", "复刻赛道"])
ap.add_argument("--skill", required=True)
ap.add_argument("--code", default=None)
ap.add_argument("--pages", type=int, default=18)
a = ap.parse_args()

wb = a.workbench
img_dir = os.path.join(wb, "img")
os.makedirs(img_dir, exist_ok=True)


def load_js(path):
    raw = open(path, encoding="utf-8").read()
    return json.loads(raw.split("=", 1)[1].strip().rstrip(";").strip())


man_path = os.path.join(wb, "manifest.js")
ans_path = os.path.join(wb, "answers.js")
man = load_js(man_path)
ans = load_js(ans_path)

used = {it["deck"] for it in man["items"]}
if a.code:
    code = a.code
    assert code not in used, "blind code already used: " + code
else:
    rnd = random.Random()
    while True:
        code = "".join(rnd.choice("abcdefghjkmnpqrstuvwxyz23456789") for _ in range(5))
        if code not in used:
            break

for i in range(1, a.pages + 1):
    src = os.path.join(a.render_dir, "slide%02d.png" % i)
    name = "%s-p%02d.png" % (code, i)
    shutil.copy(src, os.path.join(img_dir, name))
    man["items"].append({"id": "%s-p%02d" % (code, i), "track": a.track, "deck": code,
                         "page": i, "src": "img/" + name})

with open(man_path, "w", encoding="utf-8") as f:
    f.write("window.WB_DATA = " + json.dumps(man, ensure_ascii=False))

ans[code] = {"skill": a.skill, "track": a.track,
             "orig": os.path.basename(a.render_dir.rstrip("/\\"))}
with open(ans_path, "w", encoding="utf-8") as f:
    f.write("window.WB_ANSWERS = " + json.dumps(ans, ensure_ascii=False) + ";")

print("blind code:", code, "| track:", a.track, "| pages:", a.pages)
print("total items:", len(man["items"]), "| decks:", len(ans))
