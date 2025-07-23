from PIL import Image, ImageDraw
import os

# 创建一个简单的32x32图标
icon_size = (32, 32)
icon = Image.new('RGBA', icon_size, color=(255, 255, 255, 0))
draw = ImageDraw.Draw(icon)

# 绘制一个简单的图形（蓝色圆形）
center = (icon_size[0] // 2, icon_size[1] // 2)
radius = min(icon_size) // 2 - 2
draw.ellipse(
    [(center[0] - radius, center[1] - radius), 
     (center[0] + radius, center[1] + radius)], 
    fill=(0, 102, 204, 255)
)

# 保存为ICO文件
new_icon_path = "new_favicon.ico"
icon.save(new_icon_path, format="ICO")

print(f"创建了新图标: {new_icon_path}")
print(f"文件大小: {os.path.getsize(new_icon_path)} 字节")

# 验证新创建的图标
try:
    img = Image.open(new_icon_path)
    print(f"图标格式: {img.format}, 尺寸: {img.size}, 模式: {img.mode}")
    print("图标文件有效")
except Exception as e:
    print(f"验证图标时出错: {e}")