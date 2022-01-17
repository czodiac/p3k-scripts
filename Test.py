#https://code-maven.com/python-write-text-on-images-pil-pillow
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import sys
#import os

#print(os.listdir()) # 디렉토리 맞는지 확인
img = Image.open('./Bg/a.png')
draw = ImageDraw.Draw(img)
engFont = ImageFont.truetype('./Fonts/timesbd.ttf', 50)
chiFont = ImageFont.truetype('./Fonts/MaShanZheng-Regular.ttf', 50)
korFont = ImageFont.truetype('./Fonts/JejuHallasan.ttf', 50)

# 0,0 coordinates
draw.text((50, 50), "Dang Jun", (12,13,14), font=engFont)
draw.text((50, 120), "党均", (12,13,14), font=chiFont)
draw.text((50, 190), "당균31", (12,13,14), font=korFont)
img.save('test.png')