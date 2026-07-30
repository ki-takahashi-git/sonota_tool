"""
Generate app icon for the file renamer tool.
Requires: pip install pillow
"""
from PIL import Image, ImageDraw, ImageFont
import struct, zlib, os, sys

def make_icon(out_path):
    sizes = [256, 64, 48, 32, 16]
    images = []

    for size in sizes:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        p = size / 256  # scale factor

        # --- background rounded rect ---
        bg_color = (44, 95, 138)   # #2c5f8a (same as GUI header)
        r = int(32 * p)
        draw.rounded_rectangle([0, 0, size-1, size-1], radius=r, fill=bg_color)

        # --- folder shape ---
        fc = (255, 200, 60)        # folder yellow
        tab_w = int(80 * p)
        tab_h = int(22 * p)
        tab_x = int(28 * p)
        tab_y = int(48 * p)
        body_x1 = int(28 * p)
        body_y1 = int(66 * p)
        body_x2 = int(228 * p)
        body_y2 = int(168 * p)

        # folder tab
        draw.rounded_rectangle([tab_x, tab_y, tab_x + tab_w, tab_y + tab_h + 8],
                                radius=int(6*p), fill=fc)
        # folder body
        draw.rounded_rectangle([body_x1, body_y1, body_x2, body_y2],
                                radius=int(8*p), fill=fc)

        # --- arrow  A → B ---
        ac = (255, 255, 255)
        sw = max(2, int(8 * p))

        # left label "A"
        lx = int(50 * p)
        ly = int(90 * p)
        lsize = max(8, int(48 * p))
        try:
            font = ImageFont.truetype("arial.ttf", lsize)
        except Exception:
            font = ImageFont.load_default()
        draw.text((lx, ly), "A", fill=ac, font=font)

        # right label "B"
        rx = int(168 * p)
        draw.text((rx, ly), "B", fill=ac, font=font)

        # arrow line
        ay = int(118 * p)
        ax1 = int(96 * p)
        ax2 = int(160 * p)
        draw.line([(ax1, ay), (ax2, ay)], fill=ac, width=sw)

        # arrowhead
        aw = max(4, int(18 * p))
        ah = max(3, int(14 * p))
        draw.polygon([
            (ax2, ay),
            (ax2 - aw, ay - ah),
            (ax2 - aw, ay + ah)
        ], fill=ac)

        # --- pencil at bottom-right ---
        pen_color = (230, 126, 34)  # #e67e22 orange
        px = int(148 * p)
        py = int(172 * p)
        pw = int(28 * p)
        ph = int(72 * p)
        angle_offset = int(8 * p)
        # pencil body
        draw.polygon([
            (px, py),
            (px + pw, py),
            (px + pw + angle_offset, py + ph - int(16*p)),
            (px - angle_offset, py + ph - int(16*p))
        ], fill=pen_color)
        # pencil tip
        draw.polygon([
            (px - angle_offset, py + ph - int(16*p)),
            (px + pw + angle_offset, py + ph - int(16*p)),
            (px + int(pw/2), py + ph)
        ], fill=(240, 220, 160))

        images.append(img)

    images[0].save(out_path, format="ICO", sizes=[(s, s) for s in sizes],
                   append_images=images[1:])
    print(f"Icon saved: {out_path}")

if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
    make_icon(out)
