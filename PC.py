from time import sleep
import urllib.request,urllib.error
import xlwt
from selenium import webdriver
from selenium.webdriver.support.select import Select
import numpy as np
import pandas as pd
import requests
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
# 创建一个sheet对象，一个sheet对象对应Excel文件中的一张表格。
sheet = book.add_sheet('Output', cell_overwrite_ok=True)
# 其中的Output是这张表的名字,cell_overwrite_ok，表示是否可以覆盖单元格，其实是Worksheet实例化的一个参数，默认值是False
# 向表中添加数据标题
sheet.write(0, 0, '企业名称')  # 其中的'0-行, 0-列'指定表中的单元，'X'是向该单元写入的内容
sheet.write(0, 1, '所在区域')
sheet.write(0, 2, '产品名称')
sheet.write(0, 3, '计量单位')
sheet.write(0, 4, '价格')
sheet.write(0, 5, '产品类别')
sheet.write(0, 6, '图片')

a=0
driver = webdriver.Chrome()
page=0
for page in range(0,40):
    driver.get("http://www.casicloud.com/bde/search?keyword=&business_type=4&province=%E8%B4%B5%E5%B7%9E&page="+str(page))
    elements = driver.find_elements_by_class_name("col-r-lg-5")
    page = page + 1
    for element in elements:
        a=a+1
        element.find_element_by_class_name("goods-img").click()
        sleep(3)

        # 当前打开的所有窗口
        windows = driver.window_handles
        # 转换到最新打开的窗口
        driver.switch_to.window(windows[-1])
        es = []
        try:
            x=driver.find_element_by_class_name("content-top").text
            es.append(x)
        except:
            es.append("爬取失败")
#价格
        try:
            y = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[2]/div/table/tbody[1]/tr[2]/td[2]/p/span[2]").text
            es.append(y)
        except:
            es.append("爬取失败")
#单位
        try:
            y1 = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[2]/div/table/tbody[1]/tr[3]/td[2]/p/span[2]").text
            es.append(y1)
        except:
            es.append("爬取失败")
#类别
        try:
            y2 = driver.find_element_by_xpath("/html/body/div[7]/div/div[2]/div[2]/div/table/tbody[1]/tr[1]/td[2]/p/span").text
            es.append(y2)
        except:
            es.append("爬取失败")
#公司名称
        try:
            y3 = driver.find_element_by_xpath("/html/body/div[8]/div[1]/div[1]/p[1]/span").text
            es.append(y3)
        except:
            es.append("爬取失败")
#所在区域
        try:
            y4 = driver.find_element_by_xpath("/html/body/div[8]/div[1]/div[1]/p[3]/span").text
            es.append(y4)
        except:
            es.append("爬取失败")


        try:
            url = driver.find_element_by_class_name("product").find_element_by_css_selector('img').get_attribute('src')
            name = url.split("/")[-1]
            r = requests.get(url)

            # 将获取到的图片二进制流写入本地文件
            with open("./image/" + name + ".png", 'wb') as f:
                # 对于图片类型的通过r.content方式访问响应内容，将响应内容写入baidu.png中
                f.write(r.content)
            es.append(name+".png")

        except:
            es.append("没有图片")
        driver.close()

        # 当前打开的所有窗口
        windows = driver.window_handles
        # 转换到最新打开的窗口
        driver.switch_to.window(windows[-1])

        sheet.write(a, 0, es[4])
        sheet.write(a, 1, es[5])
        sheet.write(a, 2, es[0])
        sheet.write(a, 3, es[2])
        sheet.write(a, 4, es[1])
        sheet.write(a, 5, es[3])
        sheet.write(a, 6, es[6])
        sleep(1)
    book.save('研发.xls')
driver.quit()