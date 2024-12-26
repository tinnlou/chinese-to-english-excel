from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    # 创建一个 256x256 的图像，带透明背景
    size = 256
    image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    
    # 绘制圆形背景
    circle_color = '#007AFF'  # Apple蓝色
    circle_bbox = [20, 20, size-20, size-20]
    draw.ellipse(circle_bbox, fill=circle_color)
    
    # 绘制翻译符号 (简化的"中A"图标)
    text_color = 'white'
    try:
        # 尝试使用微软雅黑字体
        font = ImageFont.truetype("msyh.ttc", 120)
    except:
        # 如果没有找到，使用默认字体
        font = ImageFont.load_default()
    
    # 绘制文字
    text = "译"
    # 获取文字大小
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    
    # 计算文字位置使其居中
    x = (size - text_width) // 2
    y = (size - text_height) // 2
    
    # 绘制文字
    draw.text((x, y), text, fill=text_color, font=font)
    
    # 保存为ICO文件
    image.save('icon.ico', format='ICO', sizes=[(256, 256)])
    
    print("图标文件已创建: icon.ico")

if __name__ == "__main__":
    create_icon() 