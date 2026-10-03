# 生成一键采集图标（多尺寸 .ico）——需要改图标样式时运行
# 用法: D:\Python\python.exe make_icon.py
from PIL import Image, ImageDraw, ImageFont

OUT = '一键采集.ico'
FONT_CANDIDATES = [
    r'C:\Windows\Fonts\msyh.ttc',
    r'C:\Windows\Fonts\msyhbd.ttc',
    r'C:\Windows\Fonts\simhei.ttf',
    r'C:\Windows\Fonts\simkai.ttf',
]
S = 256


def pick_font(size):
    for f in FONT_CANDIDATES:
        try:
            return ImageFont.truetype(f, size)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_icon(size):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 圆角方块底（品牌蓝，与思迅门户蓝接近）
    pad = max(1, size // 16)
    r = size // 5
    d.rounded_rectangle([pad, pad, size - pad, size - pad], radius=r, fill=(14, 144, 216, 255))
    # 白色"汇"字
    fsize = int(size * 0.52)
    font = pick_font(fsize)
    ch = '汇'
    bbox = d.textbbox((0, 0), ch, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (size - w) / 2 - bbox[0]
    y = (size - h) / 2 - bbox[1] - size * 0.03
    d.text((x, y), ch, font=font, fill=(255, 255, 255, 255))
    # 底部三条迷你柱状图（销售汇总意象）
    bw = max(2, size // 16)
    gap = bw
    base_y = size - pad - int(size * 0.12)
    heights = [int(size * 0.08), int(size * 0.14), int(size * 0.11)]
    total = bw * 3 + gap * 2
    bx = (size - total) // 2
    for i, hh in enumerate(heights):
        x0 = bx + i * (bw + gap)
        d.rectangle([x0, base_y - hh, x0 + bw, base_y], fill=(255, 255, 255, 230))
    return img


master = draw_icon(S)
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
master.save(OUT, format='ICO', sizes=sizes)
# 另存一张 png 便于预览
master.save('_icon_preview.png')
print('saved:', OUT, 'sizes:', sizes)
