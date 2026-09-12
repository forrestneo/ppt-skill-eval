# -*- coding: utf-8 -*-
"""Render a PPTX to per-page PNGs (LibreOffice -> PDF -> pymupdf).

Usage: python render_pptx_pages.py <deck.pptx> --outdir <render_dir> [--width 1920]
"""
import argparse
import os
import subprocess
import tempfile

import pymupdf

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"

ap = argparse.ArgumentParser()
ap.add_argument("pptx")
ap.add_argument("--outdir", required=True)
ap.add_argument("--width", type=int, default=1920)
a = ap.parse_args()

os.makedirs(a.outdir, exist_ok=True)
tmpdir = tempfile.mkdtemp(prefix="pptbench_")
subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, a.pptx],
               capture_output=True, timeout=300)
pdf = os.path.join(tmpdir, os.path.splitext(os.path.basename(a.pptx))[0] + ".pdf")
if not os.path.exists(pdf):
    raise SystemExit("LibreOffice conversion failed: " + pdf)

doc = pymupdf.open(pdf)
n = 0
for i, page in enumerate(doc):
    zoom = a.width / page.rect.width
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
    pix.save(os.path.join(a.outdir, "slide%02d.png" % (i + 1)))
    n += 1
doc.close()
os.remove(pdf)
print("rendered", n, "pages ->", a.outdir)
