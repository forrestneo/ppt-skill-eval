# -*- coding: utf-8 -*-
"""Stitch one deck's rendered pages into a 3x6 (or custom) collage.

Usage: python make_collages.py <render_dir> --out <collage.png> [--prefix slide] [--pages 18] [--thumb 1100]
Expects files named <prefix>01.png .. <prefix><pages>.png in render_dir.
"""
import argparse
import os

from PIL import Image

ap = argparse.ArgumentParser()
ap.add_argument("render_dir")
ap.add_argument("--out", required=True)
ap.add_argument("--prefix", default="slide")
ap.add_argument("--pages", type=int, default=18)
ap.add_argument("--thumb", type=int, default=1100)
ap.add_argument("--cols", type=int, default=3)
a = ap.parse_args()

pages = []
for pg in range(1, a.pages + 1):
    fp = os.path.join(a.render_dir, "%s%02d.png" % (a.prefix, pg))
    if not os.path.exists(fp):
        fp = os.path.join(a.render_dir, "%s%d.png" % (a.prefix, pg))
    if not os.path.exists(fp):
        raise SystemExit("missing page: " + fp)
    pages.append(Image.open(fp))

thumbs = [im.resize((a.thumb, int(im.height * a.thumb / im.width))) for im in pages]
th = thumbs[0].height
rows = (a.pages + a.cols - 1) // a.cols
gap = 10
W = a.cols * a.thumb + (a.cols + 1) * gap
H = rows * th + (rows + 1) * gap
sheet = Image.new("RGB", (W, H), (245, 243, 240))
for i, t in enumerate(thumbs):
    c, r = i % a.cols, i // a.cols
    sheet.paste(t, (gap + c * (a.thumb + gap), gap + r * (th + gap)))
sheet.save(a.out, optimize=True)
print("saved", a.out, sheet.size)
