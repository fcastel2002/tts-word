"""Genera el logo de TTS Word: hoja con una palabra resaltada y ondas de sonido.

Escribe los PNG de addin/assets/ y installer/tts-word.ico. Requiere Pillow:
    python installer/make_icons.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
S = 1024  # se dibuja grande y se achica, para tener bordes suaves

BLUE_TOP, BLUE_BOTTOM = (43, 124, 211), (24, 82, 168)
PAGE = (255, 255, 255)
LINE = (178, 196, 222)
MARK = (255, 196, 46)
WORD = (24, 82, 168)


def draw(simple: bool) -> Image.Image:
    """simple: menos líneas y trazos más gruesos, para 16 y 32 px."""
    gradient = Image.new("RGB", (1, S))
    for y in range(S):
        t = y / (S - 1)
        gradient.putpixel((0, y), tuple(round(a + (b - a) * t) for a, b in zip(BLUE_TOP, BLUE_BOTTOM)))
    gradient = gradient.resize((S, S))
    mask = Image.new("L", (S, S))
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=220, fill=255)
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    img.paste(gradient, (0, 0), mask)
    d = ImageDraw.Draw(img)

    left, right, top, bottom = 150, 590, 170, 854
    d.rounded_rectangle((left, top, right, bottom), radius=56, fill=PAGE)

    x0 = left + 70
    if simple:
        rows = [(330, 300, False), (512, 300, True), (694, 220, False)]
        thick = 90
    else:
        rows = [(300, 300, False), (420, 260, False), (540, 300, True), (660, 230, False), (780, 160, False)]
        thick = 56
    for y, width, marked in rows:
        if marked:
            # la palabra que se está leyendo: resaltada en amarillo
            d.rounded_rectangle((x0 - 28, y - thick, x0 + width + 28, y + thick), radius=28, fill=MARK)
            d.rounded_rectangle((x0, y - thick // 2, x0 + width, y + thick // 2), radius=thick // 2, fill=WORD)
        else:
            d.rounded_rectangle((x0, y - thick // 2, x0 + width, y + thick // 2), radius=thick // 2, fill=LINE)

    cx, cy = right - 10, 512
    for radius in (150, 260) if not simple else (170,):
        width = 64 if not simple else 96
        box = (cx - radius, cy - radius, cx + radius, cy + radius)
        d.arc(box, start=-42, end=42, fill=PAGE, width=width)
    return img


def main() -> None:
    full, simple = draw(False), draw(True)
    assets = ROOT / "addin" / "assets"
    for size in (16, 32, 64, 80, 128):
        source = simple if size <= 32 else full
        source.resize((size, size), Image.LANCZOS).save(assets / f"icon-{size}.png")
    full.resize((300, 300), Image.LANCZOS).save(assets / "logo-filled.png")
    # El .ico guarda una imagen por tamaño; las chicas salen de la versión simple.
    ico = [(simple if s <= 32 else full).resize((s, s), Image.LANCZOS) for s in (16, 24, 32, 48, 64, 128, 256)]
    ico[-1].save(ROOT / "installer" / "tts-word.ico", sizes=[i.size for i in ico], append_images=ico[:-1])


if __name__ == "__main__":
    main()
