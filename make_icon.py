from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


SIZE = 256
image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)

draw.rounded_rectangle((8, 8, 248, 248), radius=52, fill=(36, 105, 184, 255))
draw.rounded_rectangle((47, 38, 177, 218), radius=14, fill=(255, 255, 255, 255))
draw.polygon([(145, 38), (177, 70), (145, 70)], fill=(203, 221, 242, 255))
draw.rounded_rectangle((95, 118, 226, 196), radius=16, fill=(230, 89, 42, 255))

font_paths = [
    Path(r"C:\Windows\Fonts\seguisb.ttf"),
    Path(r"C:\Windows\Fonts\arialbd.ttf"),
]
font_path = next((path for path in font_paths if path.exists()), None)
pdf_font = ImageFont.truetype(str(font_path), 38) if font_path else ImageFont.load_default()
ppt_font = ImageFont.truetype(str(font_path), 38) if font_path else ImageFont.load_default()

draw.text((57, 50), "PDF", font=pdf_font, fill=(36, 105, 184, 255), stroke_width=1)
draw.text((136, 127), "PPT", font=ppt_font, fill=(255, 255, 255, 255), stroke_width=1)
draw.polygon([(65, 134), (105, 134), (105, 120), (132, 144), (105, 168), (105, 154), (65, 154)], fill=(36, 105, 184, 255))

output = Path(__file__).with_name("pdf2ppt.ico")
preview = Path(__file__).with_name("pdf2ppt-icon.png")
image.save(preview, "PNG", optimize=True)
image.save(output, sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (128, 128), (256, 256)])
print(output)
print(preview)
