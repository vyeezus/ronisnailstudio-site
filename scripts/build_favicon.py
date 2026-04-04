#!/usr/bin/env python3
"""Make favicon: transparent outer white, tight crop, scale circle to fill canvas."""
from __future__ import annotations

from collections import deque
from pathlib import Path

import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
# Master asset (white canvas); run: python3 scripts/build_favicon.py
SRC = ROOT / "public" / "favicon-source.png"
OUT_PUBLIC = ROOT / "public" / "favicon.png"
OUT_ROOT = ROOT / "favicon.png"

# Pixels this light and connected to image edge become transparent (background only).
WHITE_THRESH = 248
# Output square size (browsers downscale; larger source = sharper + bigger-looking circle).
OUT_SIZE = 512
# Fraction of canvas the content bbox should fill (padding inside square).
FILL = 0.96


def main() -> None:
    img = Image.open(SRC).convert("RGBA")
    data = np.array(img, dtype=np.uint8)
    h, w = data.shape[:2]

    def is_bg_white(r: int, g: int, b: int) -> bool:
        return int(r) >= WHITE_THRESH and int(g) >= WHITE_THRESH and int(b) >= WHITE_THRESH

    visited = np.zeros((h, w), dtype=bool)
    q: deque[tuple[int, int]] = deque()

    for x in range(w):
        q.append((x, 0))
        q.append((x, h - 1))
    for y in range(h):
        q.append((0, y))
        q.append((w - 1, y))

    while q:
        x, y = q.popleft()
        if x < 0 or x >= w or y < 0 or y >= h or visited[y, x]:
            continue
        r, g, b, _ = data[y, x].tolist()
        if not is_bg_white(r, g, b):
            continue
        visited[y, x] = True
        data[y, x, 3] = 0
        for dx, dy in ((0, 1), (0, -1), (1, 0), (-1, 0)):
            q.append((x + dx, y + dy))

    # Crop to visible pixels
    alpha = data[:, :, 3]
    ys, xs = np.where(alpha > 8)
    if len(xs) == 0:
        raise SystemExit("No visible pixels after background removal")
    left, top, right, bottom = int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1
    data = data[top:bottom, left:right]
    ch, cw = data.shape[:2]

    # Square canvas: center content, then scale so bbox uses FILL of min side
    side = max(cw, ch)
    pad_w = (side - cw) // 2
    pad_h = (side - ch) // 2
    square = np.zeros((side, side, 4), dtype=np.uint8)
    square[pad_h : pad_h + ch, pad_w : pad_w + cw] = data

    pil = Image.fromarray(square).convert("RGBA")
    # Scale up so logo fills OUT_SIZE with small margin (larger in tab)
    inner = int(OUT_SIZE * FILL)
    pil = pil.resize((inner, inner), Image.Resampling.LANCZOS)
    out = Image.new("RGBA", (OUT_SIZE, OUT_SIZE), (0, 0, 0, 0))
    margin = (OUT_SIZE - inner) // 2
    out.paste(pil, (margin, margin), pil)

    out.save(OUT_PUBLIC, format="PNG", optimize=True)
    out.save(OUT_ROOT, format="PNG", optimize=True)
    print(f"Wrote {OUT_PUBLIC} and {OUT_ROOT} ({OUT_SIZE}x{OUT_SIZE})")


if __name__ == "__main__":
    main()
