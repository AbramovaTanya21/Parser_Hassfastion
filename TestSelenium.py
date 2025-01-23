
  
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

import openpyxl
from openpyxl import Workbook,load_workbook
import xlsxwriter
import os

# Авторизация на ХассФэсшен 
def AuthorizationForHasfesshen(driver):
    driver.get("https://hassfashion.ru/auth/?login=yes")

    Auht_Log = driver.find_element(By.NAME, "USER_LOGIN")
    Auht_Log.send_keys('Olga_Guschina@mail.ru')

    Auht_Password = driver.find_element(By.NAME, "USER_PASSWORD")
    Auht_Password.send_keys('optopt')

    Auth_Imp = driver.find_element(By.NAME, "Login")
    Auth_Imp.click()
    
    time.sleep(10) 
    
#Получение страниц каталога
def GettingСategories(driver):
   
    # Обращение к эксель файлу
    file_name = "HassCatalog.xlsx"
    file_path = "G:\\NRU\\SP\\Parsing\\TestSelenium\\TestSelenium\\HassCatalog.xlsx" 
   
    wb = load_workbook(file_path)
    ws = wb.active 
    LinkPages = []
    

    for index, row in enumerate(range(2, ws.max_row + 1)):
         current_collection = ws[row][0].value 
         if index > 0:
            last_collection = ws[row - 1][0].value  
            if last_collection !=  current_collection:
                print(f'Start parsing page for next colltction: {last_collection}'  )
                Goods = ParsingPage(driver, LinkPages) 
                RecordingInExcel(driver,Goods,last_collection)
                LinkPages = []              
         LinkPages.append(ws[row][2].value) 
         
    Goods = ParsingPage(driver, LinkPages) 
    RecordingInExcel(driver,Goods,last_collection)
    
#Парсинг страницы
def ParsingPage(driver, LinkPages):    
  for LinkPage in LinkPages:   
      driver.get(LinkPage) 
      time.sleep(10)
      class TabInd:
            NAME = 0
            ARTICLE = 1
            BRAND = 2
            PRICE = 3
            SIZE = 4
            DESCRIPTION = 5
            PHOTO = 6
            LINK = 7
      # Отбор товаров в наличие
      Links = [] 
      OnSales = driver.find_elements(By.XPATH, "//div[@class='res_cards']/div[not(@style)]") 
      for OnSale in OnSales:   
         a = OnSale.find_element(By.TAG_NAME, "a")
         Links.append(a.get_attribute("href")) 
      # Сбор данных со страницы товара в структуры и запись в список
      Goods = []
      for Link in Links:
        driver.get(Link)         
        #Имя
        Name1 = driver.find_element(By.XPATH,"//div[@class ='title h3 mobile']").get_attribute("innerText")
        Color = driver.find_element(By.XPATH,"//div[@class ='df jcsb rel']/span[2]").get_attribute("innerText")
        Name = f"{Name1} {Color}"
        #Артикл
        Article = driver.find_element(By.XPATH,"//div[@class ='vendor df jcsb']/span[2]").get_attribute("innerText")
        #Цена
        Price = driver.find_element(By.XPATH,"//span[@class='price']").get_attribute("innerText").strip("₽").replace("\xa0","")
        #Размер
        Size = ""
        Sizes = driver.find_elements(By.XPATH,"//div[contains(@class,'offer-size size df aic jcc txt_bolder')][not(contains(@class,'disable'))]")
        for Siz in Sizes:      
            Size += Siz.get_attribute("innerText") + " "         
        # Описание 
        Description = "" 
        symbols = "АВ:"
        for sbl in symbols:
             Composition = driver.find_element(By.XPATH,"//span[@class='compound txt_bolder']").get_attribute("innerText").replace(sbl, "") 
        try:
             Discript = driver.find_element(By.XPATH,"//span[@class='description_text txt typing']").get_attribute("innerText")   
        except:
            Discript = " - "
        Description = Discript + " " + Composition 
        #Картинки
        Picture = []
        Pictures = driver.find_elements(By.XPATH,"//div[@class ='miniatures_item df fdc']//img")
        for index, Pict in enumerate(Pictures): 
            if index % 2 != 0:
                Picture.append(Pict.get_attribute("src"))
        # Запись данных в экземляр структуры StructureOfProducts    
        StructureOfProduct = {
            TabInd.NAME : Name,
            TabInd.ARTICLE : Article,
            TabInd.BRAND : 'Hassfastion',
            TabInd.PRICE : Price, 
            TabInd.SIZE :Size,
            TabInd.DESCRIPTION : Description,
            TabInd.PHOTO : Picture,
            TabInd.LINK : Link,
            }
        # Запись структуры в список   
        Goods.append(StructureOfProduct)
      print(Goods)
      return Goods
      RecordingInExcel(driver,Goods,CategoryName)
    
#Запись в эксель      
def RecordingInExcel(driver,Goods,CategoryName): 
    # Создание\загрузка эксель файла
    file_name = "hasf-parser1.xlsx"
    file_path = "G:\\NRU\\SP\\Parsing\\TestSelenium\\TestSelenium\\hasf-parser1.xlsx" 
    if os.path.exists(file_name):
        wb = load_workbook(file_path)
        print(f"Файл '{file_name}' успешно загружен.")
    else:
         wb = Workbook()  
         print(f"Файл '{file_name}' успешно создан.")     
    # Обращение к книге и листам
    if CategoryName in wb.sheetnames:
        ws = wb[CategoryName]
        print(f"Лист '{CategoryName}' уже существует. Данные будут обновлены.")
    else:
        ws = wb.create_sheet(title=CategoryName)
        print(f"Создан новый лист '{CategoryName}'.")
       
    # Добавление данных
        headers = ['Название', 'Артикл', 'Бренд', 'Цена', 'Размер', 'Описание', 'Изображение', 'Изображение1', 'Изображение2', 'Изображение3', 'Ссылка на товар']
        ws.append(headers)
    
    # Определение активную строки заполнения    
    start_row = ws.max_row + 1  
    # for row_idx in range(1, ws.max_row + 1):
    #      row_values = [cell.value for cell in ws[row_idx]]
    #      if not any(row_values):  
    #           start_row = row_idx
    #           break
            
    for index, item in enumerate(Goods, start=start_row):  
        ws.cell(row=index, column=1, value=item[0])  
        ws.cell(row=index, column=2, value=item[1])
        ws.cell(row=index, column=3, value=item[2])
        ws.cell(row=index, column=4, value=item[3])
        ws.cell(row=index, column=5, value=item[4])
        ws.cell(row=index, column=6, value=item[5])
        for img_index, img in enumerate(item[6]):
            if img_index < 4:
                ws.cell(row=index, column=7 + img_index, value=img) 
        ws.cell(row=index, column=11, value=item[7])
       
    wb.save(file_path)
    print(f"Данные по категории {CategoryName} успешно записаны в файл '{file_name}'.") 
    
    try:  
        FL = wb["Sheet"]
        FL.title = "Лист1"
        wb.remove(wb["Лист1"])
        # wb.move_sheet(wb["Лист1"], offset = len(wb.sheetnames) - 1)
        wb.save(file_path)   
    except:
        pass
     
#Главная процедура   
s=Service('G:\\NRU\\SP\\Parsing\\selenium\\chromedriver\\win64\\130.0.6723.69\\chromedriver-win64\\chromedriver.exe')
driver = webdriver.Chrome(service=s)
AuthorizationForHasfesshen(driver)
GettingСategories(driver)

input()

    ## file_puth = "G:\NRU\SP\Parsing\TestSelenium\TestSelenium\hasf-parser.xlsx"
    # wb = openpyxl.load_workbook(file_puth) 
    # Auht_page = driver.current_window_handle
    # print(Auht_page)
    # driver.window_handles
    # driver.switch_to.window()

    #
    # password = driver.find_element(By.NAME, "password")
    # password.send_keys('k39cisdjt0_nj')
    # time.sleep(2)

    # sabmit = driver.find_elements(By.TAG_NAME, "input")
    # sabmit[0].click()

    # print(login)




# Auth = driver.find_element(By.CLASS_NAME, 'btn_DaWUW')
# Auth.click()
# Auth2 = driver.find_element(By.CLASS_NAME, 'content_A9YMTcontent_A9YMT')
# Auth2.click()




# password = driver.find_element(By.CLASS_NAME, 'text-style-ui-body input_Eg63j')
# password.send_keys_for_element('k3956577')

# sabmit = driver.find_element(By.CLASS_NAME, 'class="text-style-ui-body-bold uiButton_A9YMT primary_A9YMT large_A9YMT"')
# sabmit.click()

# time.sleep(3)

# # driver.quit()
# driver.get("https://www.nn.ru/")


# time.sleep(2)
# login = driver.find_element(By.CLASS_NAME, 'text-style-ui-body input_Eg63j')
# login.send_keys_for_element('Olg')

# password = driver.find_element(By.CLASS_NAME, 'text-style-ui-body input_Eg63j')
# password.send_keys_for_element('k3956577')

# sabmit = driver.find_element(By.CLASS_NAME, 'class="text-style-ui-body-bold uiButton_A9YMT primary_A9YMT large_A9YMT"')
# sabmit.click()

# time.sleep(3)
