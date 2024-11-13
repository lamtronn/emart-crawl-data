# Import the requests library
# python -X utf8 main.py

import requests
from bs4 import BeautifulSoup
import re
from openpyxl import load_workbook

pages = [1,2,3,4,5]
row_index = 7

try:
    # Import excel file
    workbook = load_workbook(filename="test.xlsx")
    sheet = workbook['Bản đăng tải']
except Exception as error: 
    print('Error', error)

for page in pages: 
    # Define the URL of the website to scrape
    URL = f"https://phi.emartmall.com.vn/index.php?route=product/category&path=138&page={page}"
    headers = requests.utils.default_headers()
    headers.update(
        {
            'User-Agent': 'My User Agent 1.0',
        }
    )
    # Send a GET request to the specified URL and store the response in 'resp'
    htmlDoc = requests.get(URL, headers)

    productURLList = []
    productIdList = []
   
    if htmlDoc.status_code == 200:
        # get the page sourse
        soup = BeautifulSoup(htmlDoc.content, "html.parser")

        #extract product list per page
        buttons = soup.find_all('button', class_='btn-action')
        for button in buttons:
            match = re.search(r"wishlist\.add\('(\d+)'\)", button['onclick'])
            productIdList.append(match.group(1))
            
        print(productIdList)
        
        for id in productIdList:
            print(id + "=========================")
            try:
                # Ngành hàng 100800
                # Tên sản phẩm
                # Mô tả sản phẩm
                # Giá
                # Kho hàng
                # Cân nặng
                # Ảnh bìa
                
                product_url = f"https://phi.emartmall.com.vn/index.php?route=product/product&product_id={id}"
                htmlDoc_product = requests.get(product_url, headers)
                if htmlDoc_product.status_code == 200:
                    product_soup = BeautifulSoup(htmlDoc_product.content, "html.parser")           
                    product_name = [child for child in product_soup.find('h1', class_="heading").strings][0]
                    product_description = " ".join([child for child in product_soup.find(id="tab-description").strings])
                    product_price = [child for child in product_soup.find('ul', class_="list-unstyled price").strings][1][:-1].replace('.','')
                    product_quantity = 10
                    product_weight = 1
                    
                    product_img = product_soup.find(id="pdt-image")['src']
                    data = {}
                    # data['product_id'] = id
                    # data['product_name'] = product_name
                    # data['product_description'] = product_description
                    # data['product_price'] = product_price
                    # data['product_quantity'] = product_quantity
                    # data['product_weight'] = product_weight
                    # data['product_img'] = product_img
                        
                    sheet[f'A{row_index}'] = '100780'
                    sheet[f'B{row_index}'] = product_name
                    sheet[f'C{row_index}'] = product_description
                    sheet[f'K{row_index}'] = float(product_price) * 1.4
                    sheet[f'L{row_index}'] = product_quantity
                    sheet[f'P{row_index}'] = product_img
                    sheet[f'Y{row_index}'] = product_weight
                    sheet[f'AC{row_index}'] = "Bật"
                    sheet[f'AD{row_index}'] = "Bật"
                    sheet[f'AE{row_index}'] = "Bật"
                    sheet[f'AF{row_index}'] = "Bật"
                    row_index = row_index + 1
                    # json_data = json.dumps(data)
                    print(data)
                else:
                    print(f"Failed to retrieve the webpage. Status code: {htmlDoc.status_code}")
            except Exception as error: 
                print('Error', error)
                
        workbook.save(filename="output.xlsx")
            
    else:
        print(f"Failed to retrieve the webpage. Status code: {htmlDoc.status_code}")