from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


SIZE = 256
image = Image.new("RGBA", (SIZE, SIZE), (28, 82, 150, 255))
draw = ImageDraw.Draw(image)

draw.rounded_rectangle((18, 18, 238, 238), radius=42, fill=(36, 105, 184, 255))
draw.rounded_rectangle((47, 38, 177, 218), radius=14, fill=(255, 255, 255, 255))
draw.polygon([(145, 38), (177, 70), (145, 70)], fill=(203, 221, 242, 255))
draw.rounded_rectangle((95, 118, 226, 196), radius=16, fill=(230, 89, 42, 255))

font_paths = [
    Path(r"C:\Windows\Fonts\seguisb.ttf"),
    Path(r"C:\Windows\Fonts\arialbd.ttf"),
]
font_path = next((path for path in font_paths if path.exists()), None)
font = ImageFont.truetype(str(font_path), 48) if font_path else ImageFont.load_default()
small_font = (
    ImageFont.truetype(str(font_path), 29) if font_path else ImageFont.load_default()
)

draw.text((62, 54), "PDF", font=small_font, fill=(36, 105, 184, 255))
draw.text((114, 128), "PPT", font=small_font, fill=(255, 255, 255, 255))
draw.polygon([(65, 134), (105, 134), (105, 120), (132, 144), (105, 168), (105, 154), (65, 154)], fill=(36, 105, 184, 255))

output = Path(__file__).with_name("pdf2ppt.ico")
image.save(output, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(output)
