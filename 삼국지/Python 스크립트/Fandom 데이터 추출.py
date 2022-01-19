# https://the3kingdoms.fandom.com/wiki/Special:AllPages?from=Shishun+Rui

from openpyxl import load_workbook as load # To read/modify Excel file
from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup # BeautifulSoup ver 4

# Open Excel file
file = "KOR_CHI_656_GAME10DATA.xlsx"
wb = load(filename = file)
wbSheet = wb['Sheet1']

# Read first HTML
baseUrl = "https://the3kingdoms.fandom.com"
myUrl = 'https://the3kingdoms.fandom.com/wiki/Special:AllPages?from='
masterPages = [myUrl+'Agui',myUrl+'Guan+Suo',myUrl+'Shishun+Rui']

for masterPage in masterPages:
    uClient = uReq(masterPage)
    pageHtml = uClient.read()
    pageSoup = soup(pageHtml, "html.parser")
    container = pageSoup.findAll("div",{"class":"mw-allpages-body"})
    pages = container[0].findAll("li")

    # Iterate through pages
    checkFor = ['Alliance','Assault','Avengers','Battle','Battles','Breakup','Campaign','Forces','Invasion','Capture','Black','Yellow','Wiki','Attack','Style name','Rebellion','Mystic','Dynasty','Revolt','Latest','Mutiny','Main','Power','Reconquest','Blog'] 
    for myPage in pages:
        if any(myPage.a["href"].find(check) > -1 for check in checkFor):
            print('Skipping ' + myPage.a["href"])
        else:
            myUrl = baseUrl + myPage.a["href"]
            uClient = uReq(myUrl)
            pageHtml = uClient.read()
            pageSoup = soup(pageHtml, "html.parser")
            toFind = pageSoup.select('th:contains("Real name:")')
            if len(toFind) > 0:
                engName = toFind[0].parent.td.text.replace("\n","")
                chiName = pageSoup.select('th:contains("Chinese name:")')[0].parent.td.text.replace("\n","")
                
                for row in wbSheet.iter_rows():
                    if row[1].internal_value == chiName:
                        row[12].value = engName
                        print(engName + ' updated')
                        wb.save('Updated_List.xlsx')
                        break