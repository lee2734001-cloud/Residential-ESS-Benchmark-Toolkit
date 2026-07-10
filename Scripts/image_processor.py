from pathlib import Path
from PIL import Image


def combine_vertical(top_path: Path, bottom_path: Path, output_path: Path, top_width_ratio: float = 0.72) -> Path:
    """Combine an inverter image above a battery image on a transparent canvas."""
    top = Image.open(top_path).convert("RGBA")
    bottom = Image.open(bottom_path).convert("RGBA")

    target_top_width = max(1, int(bottom.width * top_width_ratio))
    top_height = max(1, int(top.height * target_top_width / top.width))
    top = top.resize((target_top_width, top_height), Image.LANCZOS)

    gap = max(10, int(bottom.height * 0.025))
    canvas_width = max(top.width, bottom.width)
    canvas_height = top.height + gap + bottom.height
    canvas = Image.new("RGBA", (canvas_width, canvas_height), (255, 255, 255, 0))
    canvas.alpha_composite(top, ((canvas_width - top.width) // 2, 0))
    canvas.alpha_composite(bottom, ((canvas_width - bottom.width) // 2, top.height + gap))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path)
    return output_path


def image_size(path: Path) -> tuple[int, int]:
    with Image.open(path) as image:
        return image.size
