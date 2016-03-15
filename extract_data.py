# !/usr/bin/env python
# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ET
import codecs
from Tkinter import Tk
from tkFileDialog import askopenfilename


NS = {'default': "urn:schemas-microsoft-com:office:spreadsheet", 'o': "urn:schemas-microsoft-com:office:office", 'x': "urn:schemas-microsoft-com:office:excel", 'ss': "urn:schemas-microsoft-com:office:spreadsheet", 'html': "http://www.w3.org/TR/REC-html40"}
ROW = ['询价号', '电厂码', '电厂名', '采购员', '邮箱', '结束时间']
PRODUCT = ['物资编码', '品名规格', '单位', '数量', '行', '通知供应商']


def xmlparser(file):
    tree = ET.parse(file) # 生成ElementTree对象
    root = tree.getroot() # 提取树根
    sheet = root.findall('default:Worksheet', NS)[1] # 提取第2个sheet
    table = sheet.find('default:Table', NS) # 提取上面的sheet的table


    abbr_comp = table[1][1].find('default:Data', NS) # 提取公司简称
    abbr_comp = '' if abbr_comp is None else abbr_comp.text # 如果公司简称不存在用空字符代替
    abbr, company = extract_company(abbr_comp) # 分离电厂码电厂名

    # 询价号
    order_id = table[3][4].find('default:Data', NS)
    order_id = '' if order_id is None else order_id.text

    # 结束时间
    end_time = table[4][4].find('default:Data', NS)
    end_time = '' if end_time is None else end_time.text[0:10] + ' ' + end_time.text[11:19]

    # 采购员
    agent = table[4][8].find('default:Data', NS)
    agent = '' if agent is None else agent.text

    # 邮箱
    email = table[6][9].find('default:Data', NS)
    email = '' if email is None else email.text


    order = [order_id, abbr, company, agent, email, end_time]

    # 把订单信息写入文件
    with codecs.open('1.csv',  'w', encoding='gb2312') as f1:
        f1.write((','.join(ROW) + '\r\n').decode('utf-8'))
        f1.write(','.join(order))

    # 计算商品种类，比如 <Worksheet ss:Protected="1" ss:Name="行 (1 - 3)"> 表示有3种商品
    num_of_products = extract_range(sheet.get('{' + NS['ss'] + '}Name'))

    # 商品都是从sheet的第14行开始，用于定位
    start = 14
    # 提取每一行商品的信息，保存在一个array，然后在将这个array放到 p 中
    p = []
    for i in xrange(start, start + num_of_products):
        pid = table[i][2][0].text
        pname = remove_id(table[i][1][0].text)
        punit = table[i][4][0].text
        pamount = table[i][5][0].text
        prow = table[i][1][0].text
        pnotice = table[i][14]
        if len(pnotice):
            pnotice = pnotice[0].text
        else:
            pnotice = ''
        p.append([pid, pname, punit, pamount, prow, pnotice])

    # 把商品信息写入文件
    with codecs.open('2.csv',  'w', encoding='gb2312') as f2:
        f2.write((','.join(PRODUCT) + '\r\n').decode('utf-8'))
        for r in p:
            f2.write(','.join(r) + '\r\n')

# 提取公司名
def extract_company(title):
    abbr = ''
    company = ''
    for i in title:
        j = ord(i)
        if j >= ord('A') and j <= ord('Z'):
            abbr += i
        elif j >= ord('0') and j <= ord('9'):
            break
        else:
            company += i
    return (abbr, company)

# 提取商品种类
def extract_range(row_range):
    num = ''
    for i in row_range[::-1]:
        j = ord(i)
        if j >= ord('0') and j <= ord('9'):
            num = i + num
        elif i == ' ':
            break
    return int(num)

# 分离商品编码和品名
def remove_id(name):
    for i, j in enumerate(name):
        if j == ' ':
            return name[(i + 1):]

def read_file_path():
    with open('filepath.txt', 'r') as f:
        return f.read().splitlines()

# 用户界面打开文件
def get_file_path_gui():
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()
    return filename

def main():
    xmlparser(get_file_path_gui())

if __name__ == '__main__':
    main()
