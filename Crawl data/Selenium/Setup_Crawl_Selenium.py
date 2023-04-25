from time import sleep, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_experimental_option("detach", True)

browser = webdriver.Chrome(executable_path="chromedriver.exe",chrome_options=chrome_options)
browser.set_window_position(-1000, 0)
browser.maximize_window()

def get_proxy(browser, chrome_options):
    browser.get('https://free-proxy-list.net/')
    table = browser.find_elements(By.CSS_SELECTOR,"table.table-striped.table-bordered tbody tr")
    proxy_list = []
    asean_list = ['VN']
    for row in table:
        proxy_list.append(row.text)
    df = pd.DataFrame({"col": proxy_list})['col'].str.split(" ", expand=True)
    df_proxy = df.iloc[:, 0:3].copy()
    df_proxy = df_proxy.rename(columns={0: 'ip', 1: 'port', 2:'code_nation'})
    
    index = df_proxy[df_proxy['code_nation'].isin(asean_list)].index
    df_proxy_asean = df_proxy.loc[index]
    df_proxy.to_csv("E:\MyDesktop\ThaiTran\Personal_Project\Project_Shopee_ETL_Visualization\Data\proxy.csv")
    df_proxy_asean.to_csv("E:\MyDesktop\ThaiTran\Personal_Project\Project_Shopee_ETL_Visualization\Data\proxy_aseans.csv")
    random_proxy = df_proxy_asean[['ip','port']].sample(1)
    proxy_format = random_proxy[['ip', 'port']].apply(":".join, axis=1)
    proxy = proxy_format.iloc[0]
    print(proxy,"---DONE---", sep="\n")
    browser.close()
    return proxy

if __name__ == '__main__':
    proxy = get_proxy(browser, chrome_options)
    sleep(2)
    chrome_options.add_argument('--proxy-server=%s' % proxy)
    browser_shopee = webdriver.Chrome(executable_path="chromedriver.exe",chrome_options=chrome_options)
    browser_shopee.set_window_position(-1000, 0)
    browser_shopee.maximize_window()
    browser_shopee.get('https://whoer.net/')
    
# asean_list = ['BN','KH','TL','ID','LA','MY','MM','PH','SG','TH','VN']

#42.118.46.60