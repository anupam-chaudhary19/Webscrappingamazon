import smtplib
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import openpyxl
from email.message import EmailMessage
# opt = Options
# opt.add_argument("--headless")

#driver = webdriver.Chrome(executable_path="C:/Users/Anupam/PycharmProjects/eBanking/Drivers/chromedriver.exe", chrome_options=opt)
driver = webdriver.Chrome(executable_path="C:/Users/Anupam/PycharmProjects/eBanking/Drivers/chromedriver.exe")
driver.implicitly_wait(10)
driver.maximize_window()
driver.get('https://www.amazon.in/')
driver.find_element_by_id("twotabsearchtextbox").send_keys("iphones")
driver.find_element_by_xpath("//input[@value='Go']").click()
driver.find_element_by_xpath("//*[@id='p_89/Apple']/span/a/span").click()
phonenames = driver.find_elements_by_xpath("//span[contains(@class,'a-size-medium a-color-base a-text-normal')]")
prices = driver.find_elements_by_xpath("//span[contains(@class,'a-price-whole')]")

myphone = []
myprice = []

for phone in phonenames:
    myphone.append(phone.text)

for price in prices:
    myprice.append(price.text)

finallist = zip(myphone, myprice)

wb = Workbook()
wb['Sheet'].title = 'iphonesData'
sh1 = wb.active
sh1.append(['Name', 'Price'])
for x in list(finallist):
    sh1.append(x)
    wb.save("C:\\Users\\Anupam\\PycharmProjects\\WebscrappingAmazon\\iphone_data.xlsx")

msg = EmailMessage()
msg['Subject'] = 'Iphone_Data'
msg['From'] = 'Automation Team'
msg['To'] = 'havisha2020@gmail.com'

with open('Email_Template.txt') as myfile:
    data = myfile.read()
    msg.set_content(data)

with open("C:\\Users\\Anupam\\PycharmProjects\\WebscrappingAmazon\\iphone_data.xlsx", "rb") as f:
    file_data = f.read()
    file_name = f.name
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    server.login("havisha2020@gmail.com", "Radar@8080")
    server.send_message(msg)

driver.quit()
