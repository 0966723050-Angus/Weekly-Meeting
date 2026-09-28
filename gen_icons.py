# 產生 PWA PNG 圖示(與 icons/icon.svg 同一設計)
from PIL import Image, ImageDraw
import os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
TEAL, MINT, WHITE = (15, 118, 110, 255), (153, 246, 228, 255), (255, 255, 255, 255)
LIGHT = (159, 200, 196, 255)

def make(size, path, maskable=False):
    s = size / 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    if maskable:
        d.rectangle([0, 0, size, size], fill=TEAL)
        k, off = 0.78, size * 0.11  # 內縮到安全區
    else:
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=112 * s, fill=TEAL)
        k, off = 1.0, 0
    f = lambda v: off + v * s * k
    d.rounded_rectangle([f(112), f(120), f(400), f(392)], radius=28 * s * k, fill=WHITE)
    d.rounded_rectangle([f(112), f(120), f(400), f(184)], radius=28 * s * k, fill=MINT)
    d.rectangle([f(112), f(156), f(400), f(184)], fill=MINT)
    for y, w in ((214, 144), (266, 112), (318, 128)):
        d.rounded_rectangle([f(148), f(y), f(204), f(y + 26)], radius=6 * s * k, fill=TEAL)
        d.rounded_rectangle([f(220), f(y), f(220 + w), f(y + 26)], radius=6 * s * k, fill=LIGHT)
    img.save(path)

make(192, os.path.join(OUT, "icon-192.png"))
make(512, os.path.join(OUT, "icon-512.png"))
make(512, os.path.join(OUT, "icon-maskable-512.png"), maskable=True)
print("ok")
