from setup import *

#setup selenium
text_search = "mx master 3s"
page_start = 0
link = "https://shopee.vn/search?keyword={}&page={}&sortBy=sales"
link_firstpage = link
link_firstpage = link_firstpage.format(text_search, page_start)
start_time = time()
browser.get(link_firstpage)
browser.implicitly_wait(10)

def scrollPage(browser):
    for i in range(5):
        browser.execute_script("window.scrollBy(0, 1000);")
        sleep(2)

def getProducts(browser, number_of_page):
    number_of_page += 1
    browser.implicitly_wait(10)
    products = WebDriverWait(browser, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".col-xs-2-4 a")))
    link_list = []
    name_list = []
    send_from_list = []
    quantity_sold_in_month_list = []
    
    for product in products:
        name_list.append(product.find_element(By.CSS_SELECTOR, "a div div:nth-child(2) div div div").text)
        link_list.append(product.get_attribute('href'))
        quantity_sold_in_month_list.append(product.find_element(By.CSS_SELECTOR, "a div div:nth-child(2) div:nth-child(3)").text)
        send_from_list.append(product.find_element(By.CLASS_NAME, "zGGwiV").text)

    dataframe = pd.DataFrame(
        {
            'name': name_list,
            'quantity_sound_in_month': quantity_sold_in_month_list,
            'link_shopee': link_list,
            'send_from': send_from_list
        })
    
    # dataframe.to_csv("E:\MyDesktop\ThaiTran\Personal_Project\Project_Shopee_ETL_Visualization\Data\Test{}.csv".format(number_of_page))
    # print(dataframe.to_markdown)
    print("Page {}: {} rows, {} columns".format(number_of_page, dataframe.shape[0], dataframe.shape[1]))
    browser.implicitly_wait(10)
    print("Page: Done {} !!".format(number_of_page))
    print("___________________________")

def getPages(browser, link, text_search, total_pages):
    for i in range(1, total_pages):
        link_copy = link
        link_copy = link_copy.format(text_search, i)
        browser.get(link_copy)
        browser.implicitly_wait(10)
        scrollPage(browser)
        getProducts(browser, i)
    browser.close()

total_pages = 0
try:
    total_pages = int(WebDriverWait(browser, 1).until(
        EC.presence_of_element_located((By.CLASS_NAME, "shopee-mini-page-controller__total"))
    ).text)
    print("\nTotal pages: ", total_pages)
    print("___________________________", )
    if total_pages >= 10:
        total_pages = 10
        print("Just crawl first 10 pages!")
    sleep(3)
    scrollPage(browser)
    getProducts(browser, page_start)
    getPages(browser, link, text_search, total_pages)
    
except:
    end_of_result = browser.find_element(By.CLASS_NAME, "shopee-search-empty-result-section__title").text
    check_end_result = "Không tìm thấy kết quả nào"
    if end_of_result == check_end_result:
        print("!!! Search results have expired !!!")
        browser.close()
        
print("--- Runtime: {} seconds ---".format(round(time() - start_time), 0))
