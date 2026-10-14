from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "assets" / "pdf_toolbox.ico"


def main() -> None:
    size = 256
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    draw.rounded_rectangle((34, 18, 222, 238), radius=24, fill="#ffffff", outline="#cbd5e1", width=8)
    draw.polygon([(166, 18), (222, 74), (166, 74)], fill="#e8f0fe", outline="#cbd5e1")
    draw.rounded_rectangle((49, 142, 207, 205), radius=13, fill="#2563eb")

    try:
        font = ImageFont.truetype("arialbd.ttf", 52)
    except OSError:
        font = ImageFont.load_default()
    draw.text((128, 174), "PDF", font=font, fill="#ffffff", anchor="mm")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    image.save(OUTPUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    main()
