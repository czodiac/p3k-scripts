import random
from openpyxl import load_workbook as load  # To read/modify Excel file
from urllib.request import urlopen as uReq

# Create image: https://code-maven.com/python-write-text-on-images-pil-pillow
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import sys
#import os 
#print(os.listdir()) # 디렉토리 맞는지 확인

# Open Source Excel file
srcFile = "HeroData.xlsx"
srcWB = load(filename=srcFile)
srcWBSheet = srcWB['Sheet1']

# Tire 금카드 최소 능력시
goldMinStat = 77
silverMinStat = 59

# Tier 컬러
goldCardColor = (118, 95, 39)
bronzeCardColor = (66, 34, 4)
silverCardColor = (27, 24, 23)
platinumCardColor = (34, 65, 39)

goldCardBarColor = (249, 239, 214)
bronzeCardBarColor = (230, 157, 87)
silverCardBarColor = (220, 220, 220)
platinumCardBarColor = (193, 226, 198)

# 플래티넘 카드들
platinums = ['조조','유비','손권','제갈량','주유','사마의','동탁','여포','왕윤','원소','조비','손책','육손','여몽','가후','노숙','진궁','순욱','곽가','법정','방통','서서','관우','장비','황충','조운','마초','전위','태사자','감녕','하후돈','장료','초선','손상향','대교','소교']

# Font
engFont = ImageFont.truetype('./Fonts/Random11Bold-OVABe.otf', 80)
statFont = ImageFont.truetype('./Fonts/AreaKilometer50-ow3xB.ttf', 90)
chiFont = ImageFont.truetype('./Fonts/wangfonts/wt014.ttf', 155)
korFont = ImageFont.truetype('./Fonts/JejuHallasan.ttf', 70)
# PastiRegular-mLXnm.otf PlatNomor-WyVnn.ttf
barFont = ImageFont.truetype('./Fonts/PublicPixel-0W6DP.ttf', 45)

def addStatBar(yPos, val):
    yPos = int(yPos)

    # First add stat number
    draw.text((340, yPos), val, nameColor, font=statFont)

    # Add dark stat bar
    spacePixel = 8
    barYPosFix = 10
    startXPos = 480
    for i in range(int(val)):
        updatedXPos = startXPos+(i*spacePixel)
        draw.text((updatedXPos, yPos+barYPosFix), '|', nameColor, font=barFont)

    # Add brighit stat bar
    whiteBarUpdatedXPos = updatedXPos + spacePixel
    for j in range(100-int(val)):
        newUpdatedXPos = whiteBarUpdatedXPos+(j*spacePixel)
        draw.text((newUpdatedXPos, yPos+barYPosFix), '|', brightBarColor, font=barFont)

# Iterate every row from source file
totalCnt = 0
goldCnt = 0
silverCnt = 0
bronzeCnt = 0
platinumCnt = 0
for i, row in enumerate(srcWBSheet.iter_rows()):
    if i > 1:
        totalCnt=totalCnt+1
        engName = row[0].internal_value.upper()  
        korName = row[1].internal_value
        chiName = row[2].internal_value
        appearYear = str(row[3].internal_value)
        birthYear = str(row[4].internal_value)
        command = str(row[5].internal_value)
        war = str(row[6].internal_value)
        brain = str(row[7].internal_value)
        politics = str(row[8].internal_value)
        charm = str(row[9].internal_value)
        warBrainChar = str(row[10].internal_value)
        totalStat = str(row[11].internal_value)
        #print(str(i)+': '+engName+' '+korName+' '+chiName+' '+appearYear+' '+birthYear+' '+command+' '+war+' '+brain+' '+politics+' '+charm+' '+warBrainChar+' '+totalStat)
        print(str(i))
        
        # 카드 종류별 다른 상수들
        if korName in platinums:
            bgImg = './frame_3stats_platinum.png'
            nameColor = platinumCardColor
            brightBarColor = platinumCardBarColor
            platinumCnt=platinumCnt+1
        elif int(war) > goldMinStat or int(brain) > goldMinStat or int(charm) > goldMinStat:
            bgImg = './frame_3stats_gold.png'
            nameColor = goldCardColor
            brightBarColor = goldCardBarColor
            goldCnt = goldCnt+1
        elif int(war) > silverMinStat or int(brain) > silverMinStat or int(charm) > silverMinStat:
            bgImg = './frame_3stats_silver.png'
            nameColor = silverCardColor
            brightBarColor = silverCardBarColor
            silverCnt=silverCnt+1
        else:
            bgImg = './frame_3stats_bronze.png'
            nameColor = bronzeCardColor
            brightBarColor = bronzeCardBarColor
            bronzeCnt = bronzeCnt+1
            
        bgImg = Image.open(bgImg)
        draw = ImageDraw.Draw(bgImg)

        # 영어 X 포지션
        engICnt = engName.count('I') * 30
        engBlankCnt = engName.count(' ') * 20
        restEngPixelCnt = (len(engName) - engName.count('I') - engName.count(' ')) * 50
        totalEngPixelCnt = restEngPixelCnt + engICnt + engBlankCnt

        engXPos = 135
        engXTotalLen = 1220
        engXPos = engXPos + ((engXTotalLen - totalEngPixelCnt)/2)
        draw.text((engXPos, 335), engName, nameColor, font=engFont)

        # 중국어 X 포지션
        chiXPos = 590
        if len(chiName) == 3:
            chiXPos = chiXPos - 80
        elif len(chiName) == 4:
            chiXPos = chiXPos - 145
        
        draw.text((chiXPos+3, 170), chiName, (77, 58, 51), font=chiFont)  # 12,13,14 검정
        draw.text((chiXPos, 167), chiName, (215, 15, 15), font=chiFont)  # 12,13,14 검정

        # 한글 X 포지션
        korXPos = 1110
        if len(korName) == 3:
            korXPos = korXPos - 50
        elif len(korName) == 4:
            korXPos = korXPos - 115
        
        draw.text((korXPos, 1690), korName, nameColor, font=korFont)
        
        # 능력치
        addStatBar(540, war)
        addStatBar(720, brain)
        addStatBar(900, charm)
        #draw.text((statXPos, 1130+adjustAmt), politics, getColorForStat(politics), font=statFont)
        #draw.text((statXPos, 1330+adjustAmt), command, getColorForStat(command), font=statFont)

        bgImg.save('./Img/'+str(i-2)+'.png')

f = open("CardCount.txt", "w+")
f.write("Platinum: "+str(platinumCnt)+"\r\n")
f.write("Gold: "+str(goldCnt)+"\r\n")
f.write("Silver: "+str(silverCnt)+"\r\n")
f.write("Bronze: "+str(bronzeCnt)+"\r\n")
f.write("============================\r\n")
f.write("Total: "+str(totalCnt)+"\r\n")
f.close()
print('All done.')
