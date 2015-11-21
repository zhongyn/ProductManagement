# !/usr/bin/env python
# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ET
import codecs

NS = {'default': "urn:schemas-microsoft-com:office:spreadsheet", 'o': "urn:schemas-microsoft-com:office:office", 'x': "urn:schemas-microsoft-com:office:excel", 'ss': "urn:schemas-microsoft-com:office:spreadsheet", 'html': "http://www.w3.org/TR/REC-html40"}
ROW = ['询价号', '电厂码', '电厂名', '采购员', '邮箱', '结束时间']
PRODUCT = ['物资编码', '品名规格', '单位', '数量', '行', '通知供应商']


def xmlparser(file):
    tree = ET.parse(file)
    root = tree.getroot()
    sheet = root.findall('default:Worksheet', NS)[1]
    table = sheet.find('default:Table', NS)


    abbr_comp = table[1][1].find('default:Data', NS)
    abbr_comp = '' if abbr_comp is None else abbr_comp.text
    abbr, company = extract_company(abbr_comp)

    order_id = table[3][4].find('default:Data', NS)
    order_id = '' if order_id is None else order_id.text

    end_time = table[4][4].find('default:Data', NS)
    end_time = '' if end_time is None else end_time.text[0:10] + ' ' + end_time.text[11:19]

    agent = table[4][8].find('default:Data', NS)
    agent = '' if agent is None else agent.text

    email = table[6][9].find('default:Data', NS)
    email = '' if email is None else email.text


    order = [order_id, abbr, company, agent, email, end_time]

    with codecs.open('1.csv',  'w', encoding='gb2312') as f1:
        f1.write((','.join(ROW) + '\r\n').decode('utf-8'))
        f1.write(','.join(order))

    num_of_products = extract_range(sheet.get('{' + NS['ss'] + '}Name'))

    start = 14
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


    with codecs.open('2.csv',  'w', encoding='gb2312') as f2:
        f2.write((','.join(PRODUCT) + '\r\n').decode('utf-8'))
        for r in p:
            f2.write(','.join(r) + '\r\n')


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
    

def extract_range(row_range):
    num = ''
    for i in row_range[::-1]:
        j = ord(i)
        if j >= ord('0') and j <= ord('9'):
            num = i + num
        elif i == ' ':
            break
    return int(num)

def remove_id(name):
    for i, j in enumerate(name):
        if j == ' ':
            return name[(i + 1):]

def read_file_path():
    with open('filepath.txt', 'r') as f:
        return f.read().splitlines()

def main():
    for f in read_file_path():
        xmlparser(f)

if __name__ == '__main__':
    main()





