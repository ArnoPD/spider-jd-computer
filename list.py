# -*- coding: utf-8 -*-

import requests
import xlwings as xw
from bs4 import BeautifulSoup
import re
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
import time
import sys
import os

# 关注的参数名
COMPUTER_CARE_PARAMETERS = ['处理器', '硬盘容量', '内存容量', '显卡型号', '显存容量', '商品毛重', '色域', '刷新率']
# 数据保存文件名
FILE_NAME = 'computer_jd.xlsx'


def requests_page(url):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0"}
    r = requests.get(url, headers=headers)  # 增加headers, 模拟浏览器

    # print(r.encoding)  # 查看网页返回的字符集类型
    # print(r.apparent_encoding)  # 自动判断字符集类型
    # 编码不一致

    r.encoding = r.apparent_encoding

    return r.text


# url
# target_name 浏览器拖动到的元素名
def selenium_page(url, target_name=None):
    try:
        # 无界面运行
        options = Options()
        options.add_argument('--headless')

        browser = webdriver.Firefox(options=options)
        browser.get(url)

        if target_name is not None:
            # 列表页
            target = browser.find_element_by_id(target_name)
            browser.execute_script("arguments[0].scrollIntoView();", target)  # 拖动到可见的元素，动态加载商品
            time.sleep(2)
            wait = WebDriverWait(browser, 10)
            wait.until(
                EC.presence_of_all_elements_located(
                    (By.CLASS_NAME, "gl-item")
                )
            )

            # page_attr = wait.until(
            #     EC.presence_of_all_elements_located(
            #         (By.CSS_SELECTOR, '#J_bottomPage > span.p-skip > em:nth-child(1) > b')
            #     )
            # )
            # page_total = page_attr[0].text

        else:
            time.sleep(2)
            # wait = WebDriverWait(browser, 10, 0.5)
            # wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "summary-quan")))

        soup = browser.page_source
        browser.close()

        return soup
    except Exception as error:
        print(error)


def get_list(url):
    # html = requests_page(url)
    html = selenium_page(url, 'J_bottomPage')
    soup = BeautifulSoup(html, 'html.parser')
    con = soup.find(id='J_goodsList')

    products = con.find('ul').find_all('li')
    products_count = len(products)
    link_header = "https:"

    data = []
    index = 1
    for i in products:
        product_id = i['data-sku']
        html_a = i.find('a')
        href = html_a['href']
        name = i.find(class_='p-name').em.get_text()
        # shop = i.find(class_='p-shop').a.string
        price = i.find(class_='p-price').find(class_='J_' + str(product_id)).i.string
        if re.match('(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?', href) == None:
            href = link_header + href

        detail = get_detail(href)

        info = [
            product_id,
            name,
            href,
            float(price),
        ]
        info.extend(detail)
        data.append(info)

        process(products_count, index)
        index += 1

    return data


def get_detail(url):
    html = selenium_page(url)
    soup = BeautifulSoup(html, 'html.parser')
    tab_con = soup.find(id='detail').find(class_='tab-con')
    product_detail = tab_con.find(class_='p-parameter').parent
    # 商品介绍tab里的产品参数
    product_detail_parameter_list = get_detail_parameter(product_detail)

    product_intro = soup.find(class_='product-intro')
    price = product_intro.find(class_='summary-price').find(class_='price').string  # 价格
    coupons_attr = product_intro.find(id='summary-quan')

    soup.find()
    discount_price = price
    discount = '无'
    if coupons_attr.contents:
        # 获取优惠券列表
        coupons_texts = get_coupon(coupons_attr)
        # 折扣价
        discount_price = use_coupons(coupons_texts, price)
        # 所有优惠券文本
        discount = ' '.join(coupons_texts)

    more_param_row = [
        discount,
        discount_price
    ]

    more_param_row.extend(product_detail_parameter_list)

    return more_param_row


def get_coupon(coupons_attr):
    # 所有优惠券节点
    coupons = coupons_attr.find(class_='lh').find_all(class_='text')

    # 优惠券文本
    coupons_texts = []
    for coupon in coupons:
        coupon_text = coupon.string  # 优惠券文本
        coupons_texts.append(coupon_text)

    return coupons_texts


# 使用优惠券计算优惠价
def use_coupons(coupons, price):
    price = float(price)
    for coupon in coupons:
        reg = re.compile(r"(?<=满).*")
        match = reg.search(coupon)
        if match is not None:
            (demand, discount) = match.group().split('减')  # demand：要求金额，discount：优惠金额
            demand = int(demand)
            discount = int(discount)
            price = price - discount if price >= demand else price

    return price


# 获取商品介绍中的产品参数节点信息
# detail BeautifulSoup Tag Object
def get_detail_parameter(detail):
    detail_parameters_attr = detail.find(class_='parameter2').findAll('li')

    detail_parameters = []
    for i in COMPUTER_CARE_PARAMETERS:
        detail_parameters.append(extract_key_parameters(detail_parameters_attr, i))

    return detail_parameters


# 提取产品关键参数
# parameters_attr BeautifulSoup Tag Object
# key_parameter_name String
def extract_key_parameters(parameters_attr, key_parameter_name):
    for j in parameters_attr:
        key_parameter = j.find(text=re.compile("(?<=^" + key_parameter_name + ").*"))
        if key_parameter is not None:
            return j['title']

    return '无'


# 显示进度条
# count 需要处理次数
# index 当前
def process(count, index):
    process_len = 50  # 进度条长度
    percent = index * 100.0 / count
    arrow_num = round(index * process_len / count)
    process_content = '\r' + '[' + '>' * arrow_num + ']' \
                      + '%.2f' % percent + '%'
    sys.stdout.write(process_content)
    sys.stdout.flush()


def save_excel(data):
    # 设置excel属性
    app = xw.App(visible=False, add_book=False)
    path = FILE_NAME

    if os.path.exists(path):
        # 打开一个存在的excel文件
        table = app.books.open(path)
    else:
        # 添加新的excel，可能需要授权
        table = app.books.add()
        table.save(path)
        # 获取不到sheet

    sheet = table.sheets['Sheet1']
    # 获取有效的数据行数
    row = sheet.used_range.last_cell.row
    if row == 1:
        sheet.range('A1').value = [
            'id', '产品名', '链接', '价格', '优惠券', '折扣价', 'CPU',
            '硬盘', '内存', '显卡', '显存', '重量', '色域', '刷新率'
        ]

    # 新行
    row += 1
    sheet.range('A' + str(row)).value = data

    table.save()
    # 关闭excel
    table.close()
    # 退出设置excel属性
    app.quit()


def main():
    start_time = time.time()

    link = "https://search.jd.com/search?keyword=%E7%AC%94%E8%AE%B0%E6%9C%AC&enc=utf-8&qrst=1&rt=1&stop=1&vt=2&wq=%E7%AC%94%E8%AE%B0%E6%9C%AC&ev=149_76033%7C%7C122693%7C%7C90279%7C%7C122695%7C%7C17977%5E244_73389%7C%7C73390%7C%7C73391%7C%7C95704%7C%7C115688%7C%7C115689%5E5280_88794%7C%7C1070%5E25_90269%7C%7C88804%7C%7C88803%7C%7C79830%7C%7C116151%5Eexprice_0-4000%5E1107_90357%7C%7C88782%5Eexbrand_%E8%81%94%E6%83%B3%EF%BC%88Lenovo%EF%BC%89%7C%7C%E6%88%B4%E5%B0%94%EF%BC%88DELL%EF%BC%89%7C%7CThinkPad%7C%7C%E5%8D%8E%E4%B8%BA%EF%BC%88HUAWEI%EF%BC%89%7C%7C%E6%83%A0%E6%99%AE%EF%BC%88HP%EF%BC%89%7C%7C%E5%8D%8E%E7%A1%95%EF%BC%88ASUS%EF%BC%89%7C%7C%E5%B0%8F%E7%B1%B3%EF%BC%88MI%EF%BC%89%7C%7C%E5%AE%8F%E7%A2%81%EF%BC%88acer%EF%BC%89%7C%7C%E8%8D%A3%E8%80%80%EF%BC%88honor%EF%BC%89%7C%7C%E6%9C%BA%E6%A2%B0%E9%9D%A9%E5%91%BD%EF%BC%88MECHREVO%EF%BC%89%7C%7C%E7%A5%9E%E8%88%9F%EF%BC%88HASEE%EF%BC%89%5E&page={}&s=61&click=0"
    page_count = 2
    urls = [link . format(str(i)) for i in range(1, page_count * 2, 2)]

    index = 1
    for url in urls:
        print('page:' + str(index))
        data = get_list(url)
        save_excel(data)
        index += 1

    end_time = time.time()
    print('总共耗时：%s' % (end_time - start_time))
    print('ok')


if __name__ == '__main__':
    main()
