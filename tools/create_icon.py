"""Generate app icon for XLS to XLSX Converter."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"


def rounded_line(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[int, int]],
    width: int,
    fill: tuple[int, int, int, int] = (12, 12, 12, 255),
) -> None:
    draw.line(points, fill=fill, width=width, joint="curve")
    radius = width // 2
    for x, y in points:
        draw.ellipse(
            (x - radius, y - radius, x + radius, y + radius), fill=fill
        )


def build_icon(size: int = 1024) -> Image.Image:
    scale = size / 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def s(value: int) -> int:
        return int(round(value * scale))

    stroke = s(13)

    # Container/background square
    container = (s(24), s(24), s(232), s(232))
    draw.rounded_rectangle(
        container, radius=s(44), fill=(255, 255, 255, 255),
        outline=(10, 10, 10, 255), width=s(8)
    )

    # --- Left document (.xls) ---
    draw.rounded_rectangle(
        (s(50), s(60), s(125), s(200)), radius=s(8),
        outline=(10, 10, 10, 255), width=stroke
    )
    # Folded corner
    draw.line(
        [(s(125), s(60)), (s(125), s(90)), (s(95), s(90)), (s(95), s(60))],
        fill=(10, 10, 10, 255), width=stroke
    )
    # Text lines on left doc
    for y in [s(115), s(140), s(165)]:
        draw.rounded_rectangle(
            (s(65), y, s(110), y + s(6)), radius=s(3),
            fill=(10, 10, 10, 255)
        )

    # Label "XLS" on left doc
    draw.text(
        (s(72), s(88)),
        "XLS",
        fill=(10, 10, 10, 255),
        font_size=s(12),
        stroke_width=0,
    )

    # --- Arrow (rightward) ---
    arrow_y = s(128)
    # Arrow shaft
    draw.line(
        [(s(105), arrow_y), (s(145), arrow_y)],
        fill=(10, 10, 10, 255), width=stroke
    )
    # Arrow head
    draw.polygon(
        [
            (s(145), arrow_y - s(16)),
            (s(170), arrow_y),
            (s(145), arrow_y + s(16)),
        ],
        fill=(10, 10, 10, 255),
    )

    # --- Right document (.xlsx) ---
    draw.rounded_rectangle(
        (s(135), s(60), s(210), s(200)), radius=s(8),
        outline=(10, 10, 10, 255), width=stroke
    )
    # Folded corner
    draw.line(
        [(s(210), s(60)), (s(210), s(90)), (s(180), s(90)), (s(180), s(60))],
        fill=(10, 10, 10, 255), width=stroke
    )
    # Text lines on right doc
    for y in [s(115), s(140), s(165)]:
        draw.rounded_rectangle(
            (s(150), y, s(195), y + s(6)), radius=s(3),
            fill=(10, 10, 10, 255)
        )

    # Checkmark on right doc
    green = (16, 128, 64, 255)
    draw.line(
        [(s(155), s(185)), (s(168), s(198)), (s(195), s(168))],
        fill=green, width=s(11)
    )

    return image


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    image = build_icon()
    png_path = ASSETS / "app.png"
    ico_path = ASSETS / "app.ico"
    image.resize((256, 256), Image.Resampling.LANCZOS).save(png_path)
    image.save(
        ico_path,
        sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
    )
    print(f"Created {png_path}")
    print(f"Created {ico_path}")


if __name__ == "__main__":
    main()
