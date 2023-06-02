from setup import *

start_time = time.time()

def get_products_detail():
    #connect mongodb
    collection_products_serp = products_tiki["collection_products_serp"]
    collection_products_detail = products_tiki["collection_products_detail"]
    count_documents = collection_products_serp.count_documents({})
    print('Number of documents: {}'.format(count_documents))
    count_product = 0
    products_detail = []
    for x in collection_products_serp.find({},{ "_id": 0, "product_id": 1, "seller_id": 1, "seller_product_id": 1 }):
        url_detail_products = 'https://tiki.vn/api/v2/products/{}?platform=web&spid={}'.format(x['product_id'], x['seller_product_id'])
        url_seller = 'https://tiki.vn/api/shopping/v2/widgets/seller?seller_id={}&mpid={}&spid={}'.format(x['seller_id'], x['product_id'], x['seller_product_id'])
        url_review = 'https://tiki.vn/api/v2/reviews?limit=5&include=comments,contribute_info,attribute_vote_summary&sort=score%7Cdesc,id%7Cdesc,stars%7Call&page=1&spid={}&product_id={}&seller_id={}'.format(x['seller_product_id'], x['product_id'], x['seller_id'])

        headers = ''
        params = {}
        
        response_detail_products = requests.get(url=url_detail_products, headers=headers, params=params)
        products_detail_json = response_detail_products.json()
        
        response_seller = requests.get(url=url_seller, headers=headers, params=params)
        seller_json = response_seller.json().get('data').get('seller')
        
        response_review = requests.get(url=url_review, headers=headers, params=params)
        review_json = response_review.json()

        products_detail.append({
            'product_id': products_detail_json.get('id'),
            'seller_id': products_detail_json.get('current_seller').get('id'),
            'seller_name': products_detail_json.get('current_seller').get('name'),
            'seller_level': seller_json.get('store_level'),
            'all_time_quantity_sold': products_detail_json.get('all_time_quantity_sold'),
            'category_id': products_detail_json.get('categories').get('id'),
            'category_name': products_detail_json.get('categories').get('name'),
            'day_ago_created': products_detail_json.get('day_ago_created'),
            'discount': products_detail_json.get('discount'),
            'discount_rate': products_detail_json.get('discount_rate'),
            'favourite_count': products_detail_json.get('favourite_count'),
            'gift_item_title': products_detail_json.get('gift_item_title'),
            'inventory_status': products_detail_json.get('inventory_status'),
            'inventory_type': products_detail_json.get('inventory_type'),
            'list_price': products_detail_json.get('list_price'),
            'quantity_sold': (products_detail_json.get('quantity_sold').get('value') \
                if products_detail_json.get('quantity_sold') else 0),
            'rating_average': products_detail_json.get('rating_average'),
            'review_count': products_detail_json.get('review_count'),
            'url': products_detail_json.get('short_url'),
            'configurable_products': ([[item['child_id'], item['name'], item['option1'], \
                item['price'], item['inventory_status']] for item in products_detail_json.get('configurable_products')]\
                    if products_detail_json.get('configurable_products') else []),
            '1_star_count': review_json.get('stars').get('1').get('count'),
            '1_star_percent': review_json.get('stars').get('1').get('percent'),
            '2_star_count': review_json.get('stars').get('2').get('count'),
            '2_star_percent': review_json.get('stars').get('2').get('percent'),
            '3_star_count': review_json.get('stars').get('3').get('count'),
            '3_star_percent': review_json.get('stars').get('3').get('percent'),
            '4_star_count': review_json.get('stars').get('4').get('count'),
            '4_star_percent': review_json.get('stars').get('4').get('percent'),
            '5_star_count': response_review.json().get('stars').get('5').get('count'),
            '5_star_percent': review_json.get('stars').get('5').get('percent'),
            'completion_time': datetime.now()
            })
        count_product += 1
        print("Finished {} product!".format(count_product))
        time.sleep(random.randrange(1, 2))
        
    #backup csv
    pd.DataFrame(products_detail).to_csv("data_backup/get_products_detail.csv")
    #insert data to mongodb
    collection_products_detail.insert_many(products_detail)
    print("--- Done! ---")
    
get_products_detail()
myclient.close()
print("--- Runtime: {} seconds ---".format(round(time.time() - start_time), 0))


            