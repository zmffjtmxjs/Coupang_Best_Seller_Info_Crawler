from bs4 import BeautifulSoup
from selenium import webdriver

import time
import os
import urllib
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

query_url = "http://corners.gmarket.co.kr/Bestsellers"

f_dir = "E:/coding/3years/python/GMarket_Best_Seller_Info_Crawler/"

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

resultName = s + '-' + 'G마켓'

f_dir += resultName

os.makedirs(f_dir)
os.chdir(f_dir)
os.makedirs(f_dir + '/images')

fileName = f_dir + '/' + resultName 
imageName = f_dir + '/images/'

path = "E:/coding/3years/chrome driver/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)

#모든 이미지 표시를 위한 스크롤링
for i in range(20):
    driver.execute_script('window.scrollBy(0, 1000);')
    time.sleep(0.5)
    
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('#gBestWrap > div > div:nth-child(5) > div:nth-child(3) > ul')
slist = reple_result[0].find_all('li')

cnt = 1
ranking = []
title = []
fullPrice = []
currentPrice = []
discountPer = []
imgs = []


for li in slist:
    
    try:
        getTitle = li.find('a', class_='itemname').get_text().strip()
    except:
        getTitle = ''
    
    try:
        getSrc = li.find('img', class_='lazy')['src']
    except:
        getSrc = ''
        
    try:
        getFullPrice = li.find('div', class_='o-price').find('span').find('span').get_text().strip()
    except:
        getFullPrice = ''
        
    try:
        getCurrentPrice = li.find('div', class_='s-price').find('span').get_text().strip()
    except:
        getCurrentPrice = ''
    
    try:
        getDiscountPer = li.find('div', class_='s-price').find('em').get_text().strip()
    except:
        getDiscountPer = ''
    
    print("-" * 40)
    print("1.판매순위 : " + str(cnt))
    print("2.제품소개 : " + getTitle)
    print("3.원래가격 : " + getFullPrice)
    print("4.판매가격 : " + getCurrentPrice)
    print("5.할인율 : " + getDiscountPer)
    
    f = open(fileName + '.txt', 'a', encoding = 'UTF-8')
    f.write("-" * 40 + "\n")
    f.write("1.판매순위 : " + str(cnt) + '\n')
    f.write("2.제품소개 : " + getTitle + '\n')
    f.write("3.원래가격 : " + getFullPrice + '\n')
    f.write("4.판매가격 : " + getCurrentPrice + '\n')
    f.write("5.할인율 : " + getDiscountPer + '\n')
    
    #가져온 값 배열에 저장
    ranking.append(cnt)
    title.append(getTitle)
    fullPrice.append(getFullPrice)
    currentPrice.append(getCurrentPrice)
    discountPer.append(getDiscountPer)
    
    #이미지 다운로드
    if(getSrc != ''):
        try:
            urllib.request.urlretrieve(getSrc, imageName + str(cnt) + '.jpg')
            imgs.append(imageName + str(cnt) + '.jpg')
        except:
            imgs.append('')
    else:
        imgs.append('')
    print('\n')
    cnt += 1
    
driver.quit()

#검색 결과를 다양한 형태로 저장하기

amazon_best_seller = pd.DataFrame()
amazon_best_seller['판매순위'] = ranking
amazon_best_seller['제품소개'] = pd.Series(title)
amazon_best_seller['원래가격'] = pd.Series(fullPrice)
amazon_best_seller['판매가격'] = pd.Series(currentPrice)
amazon_best_seller['할인율'] = pd.Series(discountPer)

#엑셀 형태로 저장하기
amazon_best_seller.to_excel(fileName + '.xlsx', index = True)

#그림추가
wb = load_workbook(filename = fileName + '.xlsx', read_only = False, data_only = False)
ws = wb.active

for i in range(0, len(imgs)):
    if(imgs[i] != ''):                                                  #이미지 파일 누락 시 건너뜀
        img = Image(imgs[i])                                            #추가 할 이미지 파일 위치
        
        cellNum = i + 2                                                 #셀 크기 조절 대상을 이미지 저장 위치에 맞춤
        
        ws.row_dimensions[cellNum].height = img.height * 0.75 + 16      #이미지 크기에 맞게 높이 조절
        ws.column_dimensions['C'].width = 102                           #제목 최대 길이에 맞게 넓이 조절
        
        ws.add_image(img, 'C' + str(cellNum))                           #이미지를 엑셀에 추가

wb.save(fileName + '.xlsx')

print('정상적으로 처리되었습니다.')