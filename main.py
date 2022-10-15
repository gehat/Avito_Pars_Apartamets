import selenium.common.exceptions
from selenium import webdriver
from selenium_stealth import stealth
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import xlsxwriter


class browsr():
    def __init__(self, url,count_page):
        self.url = url
        self.page=1
        self.count_page=count_page   # number of pages on the site Optional parameter

        self.__book = xlsxwriter.Workbook(r"C:\Users\jamem\Desktop\parser.xlsx", options={'strings_to_urls': False}) # Create xl file  on path
        self.__page = self.__book.add_worksheet('Квартиры')

        self.__row = 0
        self.__column = 0

        self.__page.set_column("A:A", 100)
        self.__page.set_column("B:B", 100)
        self.__page.set_column("C:C", 50)
        self.__page.set_column("D:D", 50)
        self.__page.set_column("E:E", 50)

    def drive(self): # inicialization parametrs
        options = webdriver.ChromeOptions()
        options.add_argument("start-min")

        # options.add_argument("--headless")

        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        self.__driver = webdriver.Chrome(options=options,
                                         executable_path=r"C:\Users\jamem\AppData\Local\Programs\Python\Python310\Scripts\chromedriver.exe") #locale file

        stealth(self.__driver,
                languages=["en-US", "en"],
                vendor="Google Inc.",
                platform="Win32",
                webgl_vendor="Intel Inc.",
                renderer="Intel Iris OpenGL Engine",
                fix_hairline=True,
                )
        # avito = f"https://www.avito.ru/tver/kvartiry/prodam-ASgBAgICAUSSA8YQ?cd={1}"
        # self.driver.get(f"{self.url}{self.page}")
        # block = self.driver.find_element(By.CLASS_NAME, 'items-items-kAJAg')
        # self.pos = block.find_elements(By.CLASS_NAME, 'iva-item-body-KLUuy')

    def wrt(self, name,description, price,path): # writer xl
        print(name)
        self.__page.write(self.__row, self.__column, name)
        self.__page.write(self.__row, self.__column+1, description)
        self.__page.write(self.__row, self.__column + 2, price[0])
        if len(price) == 1:
            self.__page.write(self.__row, self.__column + 3, 'None')
        else:
            self.__page.write(self.__row, self.__column + 3, price[1])
        self.__page.write(self.__row, self.__column+4, path)

        self.__row += 1

    @property
    def Get_Elements(self):
        self.drive()
        for i in range(1, self.count_page):# number of pages on the site
            try:
                self.__driver.get(f"{self.url}{self.page}")
            finally:
                block = self.__driver.find_element(By.CLASS_NAME, 'items-items-kAJAg')
                pos = block.find_elements(By.CLASS_NAME, 'iva-item-body-KLUuy')
                #path = block.find_element(By.CLASS_NAME, 'iva-item-sliderLink-uLz1v').get_attribute('href')
                for p in pos:
                    path=p.find_element(By.CLASS_NAME,'iva-item-titleStep-pdebR').find_element(By.TAG_NAME,'a').get_attribute('href')
                    try:
                        description=p.find_element(By.CLASS_NAME,'iva-item-descriptionStep-C0ty1').find_element(By.CLASS_NAME,'iva-item-text-Ge6dR').text
                    except:
                        description=None
                    name = p.find_element(By.CLASS_NAME, 'iva-item-titleStep-pdebR').text
                    price = p.find_element(By.CLASS_NAME, 'iva-item-priceStep-uq2CQ').text.split('\n')
                    print(description)
                    self.wrt(name, description,price,path)
                self.page += 1
        self.__book.close()

    # @property
    # def Get_all(self):
    #     for i in range(1, 3):
    #         self.Get_Elements
    #         self.page += 1
    #         print(self.page)
    #     self.__book.close()


avito = browsr(
    'https://www.avito.ru/tver/kvartiry/prodam-ASgBAgICAUSSA8YQ?cd=1&p=',2)  # insert address without page pointer

avito.Get_Elements
