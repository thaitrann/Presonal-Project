from setup import *

#get link
link = 'https://shopee.vn/Chu%E1%BB%99t-Logitech-MX-Master-3S-Master-3-I-Master-3-For-Mac-Kh%C3%B4ng-D%C3%A2y-B%E1%BA%A3o-h%C3%A0nh-12-Th%C3%A1ng-ch%C3%ADnh-h%C3%A3ng-i.4046711.3862731850?sp_atk=becc1820-8d38-4a60-9a45-f6132ca47907&xptdk=becc1820-8d38-4a60-9a45-f6132ca47907'
start_time = time()
browser.get(link)
browser.implicitly_wait(10)

#get overall product
product_name = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
    div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > \
        div.flex.flex-auto.RBf1cu > div > div._44qnta > span").text

product_avg_star = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
    div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > div.flex.flex-auto.RBf1cu > \
        div > div.flex.X5u-5c > div:nth-child(1) > div._1k47d8._046PXf").text

product_total_reviews = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
    div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > div.flex.flex-auto.RBf1cu > \
        div > div.flex.X5u-5c > div:nth-child(2) > div._1k47d8").text

product_total_sold = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
    div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > div.flex.flex-auto.RBf1cu > \
        div > div.flex.X5u-5c > div.flex.jgUbWJ > div.P3CdcB").text

product_price_range = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
    div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > div.flex.flex-auto.RBf1cu > \
        div > div:nth-child(3) > div > div > div:nth-child(1) > div > div > div").text

browser.implicitly_wait(10)
product_supplement_list = browser.find_element(By.CLASS_NAME, "h-y3ij").text.split("\n")

product_free_ship = "Miễn phí vận chuyển"
product_insurance = "Bảo hiểm Thiết bị điện tử"
# print(product_name, product_avg_star, product_total_reviews, product_total_sold, product_price_range, "---", sep="\n")
print(product_supplement_list)
dataframe = pd.DataFrame(
    {
        'name': product_name,
        'avg_star': product_avg_star,
        'total_reviews': product_total_reviews,
        'total_sold': product_total_sold,
        'price_range': product_price_range,
        'supplement_list': [product_supplement_list]
    })

dataframe.to_csv("E:\MyDesktop\ThaiTran\Personal_Project\Project_Shopee_ETL_Visualization\Data\Test.csv")
#get price variation
browser.implicitly_wait(10)
product_price_variation = browser.find_elements(By.CLASS_NAME, "product-variation")
for p in product_price_variation:
    p.click()
    sleep(0.5)
    name_variation = p.text
    price = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > div:nth-child(1) > \
        div.ndOSOO > div > div.container > div.product-briefing.flex.card.s9-a-0 > div.flex.flex-auto.RBf1cu > \
            div > div:nth-child(3) > div > div > div:nth-child(1) > div > div > div").text
    if p.get_attribute("aria-disabled") == "true":
        print(name_variation, "Out Of Stock","---" , sep="\n")
    else:
        print(name_variation, price,"---" , sep="\n")

sleep(1)
#get product shop
shop_name = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.jwlMoy > div > div.VlDReK").text

shop_favourite_tag = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.jwlMoy > a > div.NP8R3A").text

shop_link = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div.ndOSOO > div > div.container > div.UwHWuz > \
        div.NLeTwo.page-product__shop > div.jwlMoy > div > div.Uwka-w > a").get_attribute("href")

shop_total_review = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > div:nth-child(1) > span").text

shop_total_product = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > a > span").text

shop_response_rate = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > div:nth-child(2) > span").text

shop_response_time = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > div:nth-child(5) > span").text

shop_participation_time = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) \
    > div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > div:nth-child(3) > span").text

shop_follower = browser.find_element(By.CSS_SELECTOR, "#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div.container > div.UwHWuz > div.NLeTwo.page-product__shop \
        > div.Po6c6I > div:nth-child(6) > span").text

# print(shop_name, shop_favourite_tag, shop_link, shop_total_review, shop_total_product, \
#     shop_response_rate, shop_response_time, shop_participation_time, shop_follower, "---", sep="\n")

#detail review product
detail_review_product = browser.find_element(By.CSS_SELECTOR,"#main > div > div:nth-child(3) > \
    div:nth-child(1) > div > div > div > div.UwHWuz > div.page-product__content > \
        div.page-product__content--left > div:nth-child(2) > div > div > div.product-rating-overview \
            > div.product-rating-overview__filters").text

print(detail_review_product)

sleep(1)
browser.close()
