from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import Workbook
import sys
import urllib.request
import cgi
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

#user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36'
user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'

headers={'User-Agent':user_agent,} 
data = dict({'sku': [], 'image': [], 'title': [], 'discount': [], 'price': [], 'current_price': [], 'description': []})
url = "https://eg.oriflame.com/men/hair"
driver = webdriver.Chrome()
driver.get(url)
try:
    x=driver.find_element(By.XPATH, "/html/body/div[10]/div[1]/div/a/span")
    #x=driver.find_element(By.CSS_SELECTOR,"span.k-icon .k-i-close")
    x.click()
    print("YEsssssssssss")
except:
    print("No")


while(True):
    try:
        driver.find_element(By.CSS_SELECTOR, ".k-button.secondary").click()
        print("ok")
    except:
      break

product = driver.find_elements(By.CLASS_NAME, 'product-box-1uxw1kc')
print(len(product))

#get urls of each product to open
for prodindex in range(0,len(product)):
    #pindex = "orf" + str(prodindex + 1)  
    pindex = str(prodindex + 1)  
    product_url= product[prodindex].get_attribute("href")
    print(product_url)
    driver_product = webdriver.Chrome()
    try:
        driver_product.get(product_url)
    

        try:
            x=driver_product.find_element(By.XPATH, "/html/body/div[12]/div[1]/div/a/span")
            x.click()
            print("YEsssssssssss")
        except:
            print("No")
        driver_product.implicitly_wait(20) # gives an implicit wait for 20 seconds
        title =driver_product.find_element(By.CLASS_NAME,'summary__name_SYLp7')
        # print(title.text)
        driver_product.implicitly_wait(20) # gives an implicit wait for 20 seconds
        price = driver_product.find_element(By.CLASS_NAME,'summary__price_KIqGP')
        #print("PRiceeeee",price.text)
        try:
            icon_dis = driver_product.find_elements(By.CSS_SELECTOR,'span.product-detail-MuiIconButton-label .expand-icon__kO6x')
            icon_dis[1].click()
            print("iconnnnnnn clickeeeeeeeeeeeeeeeeeeeedddd")
            dis = driver_product.find_elements(By.CLASS_NAME,'wysiwyg_rcqSs')
            #print(dis[1].text)
            data['description'].append(dis[1].text)

        except:
            dis = ""
            data['description'].append(dis)

        imges=driver_product.find_elements(By.CLASS_NAME,'image__source__XcDT')
        for i in range(len(imges)):
            image_url = imges[i].get_attribute('src')
            print(image_url)
            request=urllib.request.Request(image_url,None,headers) #The assembled request
            response = urllib.request.urlopen(request)
            galleryimg = response.read()
            # print(image)
            imagename = str(pindex) + "orf"+ str(i+1) + ".jpg"
            with open(imagename , "wb") as file:
                file.write(galleryimg)
        
    except:
        continue

    data['sku'].append(pindex)
    data['image'].append(imagename)
    data['title'].append(title.text)
    data['current_price'].append(price.text)
    driver_product.close()    

    # data['sku'].append(pindex)
    # data['image'].append(imagename)
    # data['title'].append(title.text)
    # data['current_price'].append(price.text)
    # driver_product.close()


# print(data['image'])
df = pd.DataFrame(data, columns= ['sku', 'image', 'title', 'description', 'current_price'])
df.set_index('sku')
df.rename(columns= {'sku': 'رمز المنتج (SKU)','image': 'الصور', 'title': 'الأسم', 'description': 'الوصف', 'current_price': 'سعر الشراء'}, inplace=True)

# print(df)
df.to_excel("wearable" + '.xlsx', index = False, header=True)