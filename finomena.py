from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait # wait a few second before click
from selenium.webdriver.support import expected_conditions as EC # wait a few second before click
import time
import xlsxwriter

fund_name_list = ['ASP-DIGIBLOC-SSF', 'ASP-DIGIBLOCRMF', 'SCBSEMI(SSFE)', 'ASP-DIGIBLOC', 'SCBSEMI(SSF)',
              #    'SCBSEMI(A)', 'SCBSEMI(P)', 'KFJPINDX-I', 'SCBBLOC(E)']

# get normal fund data
fund_response_list = []
fund_category_list = []
dividend_list = []
fund_initial_date = []
fund_fee_sell_Frontendfee = []
fund_fee_Backendfee = []
fund_fee_management = []
#purchase_detail_list = []
#purchase_detail_dict = {}
#purchase_detail_0_list = []

# get stat fund data
yelid_list = []
m3_pct_change_list = []
m6_pct_change_list = []
y1_pct_change_list = []
y3_pct_change_list = []
y5_pct_change_list = []

fund_sd_list = []
m3_sd_list = []
m6_sd_list = []
y1_sd_list = []
y3_sd_list = []
y5_sd_list = []

fund_sr_list = []
m3_sr_list = []
m6_sr_list = []
y1_sr_list = []
y3_sr_list = []
y5_sr_list = []

fund_dd_list = []
m3_dd_list = []
m6_dd_list = []
y1_dd_list = []
y3_dd_list = []
y5_dd_list = []

path = r"C:\Users\kaewt\OneDrive\Desktop\Quant_Spyder\chromedriver-win64\chromedriver-win64\chromedriver.exe"

service = webdriver.chrome.service.Service(path)
service.start()
driver = webdriver.Chrome(service = service)

# Get yelid
for fund in fund_name_list:
    url = f"https://www.finnomena.com/fund/{fund}"
    driver.get(url)
    driver.implicitly_wait(2)
    fund_des = driver.find_element(By.XPATH, "//div[@class='fund-detail']").text
    # สร้าง บลจ. และเก็บไว้อยู่ใน list
    fund_response_list.append(fund_des.split('\n')[2])
    # ประเภทกองทุน
    fund_category_list.append(fund_des.split("\n")[3])
    # policy ปันผล(dividend)
    dividend_list.append(fund_des.split("\n")[7])
    # วันจดทะเบียนกองทุน
    fund_initial_date.append(fund_des.split("\n")[13])
    # ค่าธรรมเนียมขาย Front – End Fee 
    fund_fee_sell_Frontendfee.append(fund_des.split("\n")[8])
    # ค่าธรรมเนียมรับซื้อคืน Back – End Fee
    fund_fee_Backendfee.append(fund_des.split("\n")[9])
    # ค่าใช้จ่ายกองทุนรวม Management Fee
    fund_fee_management.append(fund_des.split("\n")[10])
    driver.implicitly_wait(2)
    print(f'get yelid for {fund} succesfully')

df = pd.DataFrame(list(zip(fund_name_list, fund_response_list, fund_category_list, dividend_list, fund_initial_date,\
                           fund_fee_sell_Frontendfee, fund_fee_Backendfee, fund_fee_management)),
               columns =['fund_name' , 'บลจ', 'fund_category', 'divivend', 'inital_fund_date',\
                         'front_end_fee', 'back_end_fee', 'management_fee'])

# Get Performance
for fund in fund_name_list:
    url = f"https://www.finnomena.com/fund/{fund}"
    driver.get(url)
    driver.implicitly_wait(2)
    # get data ผลการดำเนินงานและปันผล
    feerate_button = driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/nav[2]/ul/li[3]/a")
    feerate_button.send_keys(Keys.ENTER)
    driver.implicitly_wait(2)
    yelid = driver.find_elements(By.XPATH, "//*[contains(@class, 'table-row item-row border-row')]")

    for detail in yelid:
        text = detail.text
        yelid_list.append(text)
    
    m3_pct_change_list.append(yelid_list[0].split("\n")[1])
    m6_pct_change_list.append(yelid_list[1].split("\n")[1])
    y1_pct_change_list.append(yelid_list[2].split("\n")[1])
    y3_pct_change_list.append(yelid_list[3].split("\n")[2])
    y5_pct_change_list.append(yelid_list[4].split("\n")[2])
    yelid_list = []
    print(f'get performance for {fund} succesfully')

per_df = pd.DataFrame(list(zip(fund_name_list, m3_pct_change_list, m6_pct_change_list, y1_pct_change_list, y3_pct_change_list, y5_pct_change_list)),
             columns = ['fund_name', '3m_pct_change', '6m_pct_change', '1y_pct_change', '3y_pct_change', '5y_pct_change'])

# Get sd
for fund in fund_name_list:
    url = f"https://www.finnomena.com/fund/{fund}"
    driver.get(url)
    driver.implicitly_wait(2)
    # get SD
    feerate_button = driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/nav[2]/ul/li[3]/a")
    feerate_button.send_keys(Keys.ENTER)
    driver.implicitly_wait(2)
    fund_sd = driver.find_elements(By.XPATH, "//*[contains(@class, 'item-box')]")
    
    for detail in fund_sd:
        text = detail.text
        fund_sd_list.append(text)
    
    m3_sd_list.append(fund_sd_list[1].split('\n')[3])
    m6_sd_list.append(fund_sd_list[1].split('\n')[6])
    y1_sd_list.append(fund_sd_list[1].split('\n')[9])
    y3_sd_list.append(fund_sd_list[1].split('\n')[13])
    y5_sd_list.append(fund_sd_list[1].split('\n')[17])
    fund_sd_list = []
    print(f'get sd for {fund} succesfully')
    
sd_df = pd.DataFrame(list(zip(fund_name_list, m3_sd_list, m6_sd_list, y1_sd_list, y3_sd_list, y5_sd_list)),
             columns = ['fund_name', '3m_sd', '6m_sd', '1y_sd', '3y_sd', '5y_sd'])

# get Sharpe ratio
for fund in fund_name_list:
    url = f"https://www.finnomena.com/fund/{fund}/performance/sharpeRatio"
    driver.get(url)
    driver.implicitly_wait(2)
    # get Sharpe ratio
    fund_sr = driver.find_elements(By.XPATH, "//*[contains(@class, 'fund-deviation')]")

    for detail in fund_sr:
        text = detail.text
        fund_sr_list.append(text)

    m3_sr_list.append(fund_sr_list[0].split('\n')[3])
    m6_sr_list.append(fund_sr_list[0].split('\n')[6])
    y1_sr_list.append(fund_sr_list[0].split('\n')[9])
    y3_sr_list.append(fund_sr_list[0].split('\n')[13])
    y5_sr_list.append(fund_sr_list[0].split('\n')[17])
    fund_sr_list = []
    print(f'get sharpe ratio for {fund} succesfully')

sr_df = pd.DataFrame(list(zip(fund_name_list, m3_sr_list, m6_sr_list, y1_sr_list, y3_sr_list, y5_sr_list)),
             columns = ['fund_name', '3m_sr', '6m_sr', '1y_sr', '3y_sr', '5y_sr'])


# get drawdown
for fund in fund_name_list:
    url = f"https://www.finnomena.com/fund/{fund}/performance/maxDrawdown"
    driver.get(url)
    driver.implicitly_wait(2)
    # get drawdown
    
    fund_dd = driver.find_elements(By.XPATH, "//*[contains(@class, 'fund-deviation')]")

    for detail in fund_dd:
        text = detail.text
        fund_dd_list.append(text)

    m3_dd_list.append(fund_dd_list[0].split('\n')[3])
    m6_dd_list.append(fund_dd_list[0].split('\n')[6])
    y1_dd_list.append(fund_dd_list[0].split('\n')[9])
    y3_dd_list.append(fund_dd_list[0].split('\n')[12])
    y5_dd_list.append(fund_dd_list[0].split('\n')[15])
    fund_dd_list = []
    print(f'get drawdown for {fund} succesfully')

dd_df = pd.DataFrame(list(zip(fund_name_list, m3_dd_list, m6_dd_list, y1_dd_list, y3_dd_list, y5_dd_list)),
             columns = ['fund_name', '3m_dd', '6m_dd', '1y_dd', '3y_dd', '5y_dd'])

df  
per_df
sd_df
sr_df
dd_df

with pd.ExcelWriter('multiple.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='fund_description')
    per_df.to_excel(writer, sheet_name='per_pct_change')
    sd_df.to_excel(writer, sheet_name='SD')
    sr_df.to_excel(writer, sheet_name='Sharpe_Ratio')
    dd_df.to_excel(writer, sheet_name='Maximun_Drwadown')

driver.close()






















