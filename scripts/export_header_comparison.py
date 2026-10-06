"""Create diagnostic header crops only; never modify source slide images."""
import argparse
import json
from pathlib import Path
from PIL import Image, ImageDraw


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--contract', required=True)
    parser.add_argument('--images', nargs='+', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    contract = json.loads(Path(args.contract).read_text(encoding='utf-8-sig'))
    width = contract['canvas_px']['width']
    height = contract['canvas_px']['height']
    crop = contract['header_crop_px']
    if width <= 0 or height <= 0 or not (0 <= crop[0] < crop[2] <= width and 0 <= crop[1] < crop[3] <= height):
        raise ValueError('Invalid canvas or crop coordinates')
    output = Path(args.output)
    if output.exists():
        raise FileExistsError(output)
    rows = []
    for name in args.images:
        with Image.open(name) as original:
            if abs(original.width / original.height / (width / height) - 1) > .01:
                raise ValueError(f'Aspect ratio mismatch: {name}')
            normalized = original.convert('RGB').resize((width, height), Image.Resampling.LANCZOS)
        rows.append(normalized.crop(crop))
    row_height = crop[3] - crop[1] + 30
    sheet = Image.new('RGB', (crop[2] - crop[0], row_height * len(rows)), 'white')
    draw = ImageDraw.Draw(sheet)
    for index, row in enumerate(rows):
        top = index * row_height
        draw.text((8, top + 5), f'Body {index + 1:02d}', fill='black')
        sheet.paste(row, (0, top + 30))
        # Guides indicate intended box origins, not detected glyph boundaries.
        for field in ('title', 'subtitle'):
            x = contract[field]['x'] - crop[0]
            y = top + 30 + contract[field]['y'] - crop[1]
            draw.line((x, top + 30, x, top + row_height - 1), fill='#e05858', width=1)
            draw.line((0, y, sheet.width - 1, y), fill='#e05858', width=1)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)
    print(json.dumps({'output': str(output.resolve()), 'pages': len(rows), 'automatic_acceptance': False}))


if __name__ == '__main__':
    main()
