"""Tegner ikonene til PPTimagePoint (src/assets/icon-*.png)."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter

OUT = Path(__file__).resolve().parent.parent / "src" / "assets"
SS = 8  # supersampling


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(len(a)))


def gradient(size):
    stops = [(0.0, (88, 44, 255)), (0.55, (232, 62, 140)), (1.0, (255, 150, 60))]
    g = Image.new("RGBA", (size, size))
    px = g.load()
    for y in range(size):
        for x in range(size):
            t = (x * 0.55 + (size - 1 - y) * 0.45) / (size - 1)
            t = 1 - t
            for i in range(len(stops) - 1):
                t0, c0 = stops[i]
                t1, c1 = stops[i + 1]
                if t <= t1:
                    c = lerp(c0, c1, (t - t0) / (t1 - t0))
                    break
            px[x, y] = c + (255,)
    return g


def sparkle(d, cx, cy, r, fill):
    w = r * 0.28
    d.polygon([(cx, cy - r), (cx + w, cy - w), (cx + r, cy), (cx + w, cy + w),
               (cx, cy + r), (cx - w, cy + w), (cx - r, cy), (cx - w, cy - w)], fill=fill)


def icon(size, detailed=True):
    S = size * SS
    base = gradient(S)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], int(S * 0.22), fill=255)
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    im.paste(base, (0, 0), mask)

    # svak glans oeverst
    gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(gloss).ellipse([-S * 0.3, -S * 0.9, S * 1.3, S * 0.45], fill=(255, 255, 255, 38))
    gloss = gloss.filter(ImageFilter.GaussianBlur(S * 0.06))
    im = Image.alpha_composite(im, Image.composite(gloss, Image.new("RGBA", (S, S)), mask))

    d = ImageDraw.Draw(im)
    m = S * (0.2 if detailed else 0.16)
    x0, y0, x1, y1 = m, S * 0.27, S - m, S - m * 0.95
    fw = x1 - x0
    if detailed:
        # bildeflate: venstre halvdel hel, hoeyre halvdel rutenett (gjennomsiktig)
        frame = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        fd = ImageDraw.Draw(frame)
        fd.rounded_rectangle([x0, y0, x1, y1], int(S * 0.06), fill=(255, 255, 255, 255))
        cs = fw / 6
        split = x0 + fw * 0.5
        for i in range(6):
            for j in range(8):
                cx = x0 + i * cs
                cy = y0 + j * cs
                if cx >= split and (i + j) % 2 == 0:
                    fd.rectangle([cx, cy, cx + cs, cy + cs], fill=(222, 214, 240, 255))
        fmask = Image.new("L", (S, S), 0)
        ImageDraw.Draw(fmask).rounded_rectangle([x0, y0, x1, y1], int(S * 0.06), fill=255)
        im.paste(frame, (0, 0), fmask)
        # fjell og sol, bare paa venstre side (hoeyre er "fjernet")
        clip = Image.new("L", (S, S), 0)
        ImageDraw.Draw(clip).rounded_rectangle([x0, y0, split, y1], int(S * 0.06), fill=255)
        art = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        ad = ImageDraw.Draw(art)
        ad.ellipse([x0 + fw * 0.1, y0 + fw * 0.08, x0 + fw * 0.3, y0 + fw * 0.28], fill=(255, 176, 46, 255))
        ad.polygon([(x0, y1), (x0 + fw * 0.33, y0 + fw * 0.33), (x0 + fw * 0.55, y0 + fw * 0.62), (x0 + fw * 0.7, y0 + fw * 0.5), (x1, y1)], fill=(54, 34, 110, 255))
        im.paste(art, (0, 0), Image.composite(clip, Image.new("L", (S, S), 0), art.split()[3]))
        # stiplet skillelinje der bakgrunnen er fjernet
        dd = ImageDraw.Draw(im)
        seg = fw / 10
        y = y0
        while y < y1:
            dd.line([(split, y), (split, min(y1, y + seg * 0.55))], fill=(88, 44, 255, 255), width=max(2, int(S * 0.018)))
            y += seg
    else:
        d.polygon([(x0, y1), (x0 + fw * 0.36, y0 + fw * 0.18), (x0 + fw * 0.6, y0 + fw * 0.5), (x0 + fw * 0.76, y0 + fw * 0.36), (x1, y1)], fill=(255, 255, 255, 255))
    # gnist oeverst til hoeyre
    d = ImageDraw.Draw(im)
    sparkle(d, S * 0.79, S * 0.2, S * (0.15 if detailed else 0.17), (255, 255, 255, 255))
    if detailed:
        sparkle(d, S * 0.6, S * 0.12, S * 0.055, (255, 255, 255, 230))
    return im.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    for s in (16, 32, 64, 80, 128):
        icon(s, detailed=s >= 32).save(OUT / f"icon-{s}.png")
    print("ok")
