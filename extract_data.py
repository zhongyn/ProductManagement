# !/usr/bin/env python
# -*- coding: utf-8 -*-
import xml.etree.ElementTree as ET

NS = {'default': "urn:schemas-microsoft-com:office:spreadsheet", 'o': "urn:schemas-microsoft-com:office:office", 'x': "urn:schemas-microsoft-com:office:excel", 'ss': "urn:schemas-microsoft-com:office:spreadsheet", 'html': "http://www.w3.org/TR/REC-html40"}
ROW = ['询价号', '电厂码', '电厂名', '采购员', '邮箱', '结束时间']
PRODUCT = ['物资编码', '品名规格', '单位', '数量', '行', '通知供应商']


def xmlparser(file):
    tree = ET.parse(file)
    root = tree.getroot()
    sheet = root.findall('default:Worksheet', NS)[1]
    table = sheet.find('default:Table', NS)


    abbr, company = extract_company(table[1][1][0].text)
    print abbr, company.encode('utf-8')
    order_id = table[3][4][0].text
    end_time = table[4][4][0].text

    agent = table[4][8][0].text
    print agent.encode('utf-8')
    email = table[6][9][0].text

    order = [order_id, abbr, company, agent, email, end_time]
    print order

    f1 = open('result1.csv',  'w')
    f1.write(','.join(ROW) + '\n')
    f1.write(','.join([i.encode('utf-8') for i in order]) + '\n\n\n\n')

    num_of_products = extract_range(sheet.get('{' + NS['ss'] + '}Name'))
    print num_of_products

    start = 14
    p = []
    for i in xrange(start, start + num_of_products):
        pid = table[i][2][0].text
        pname = remove_id(table[i][1][0].text)
        print pname.encode('utf-8')
        punit = table[i][4][0].text
        pamount = table[i][5][0].text
        prow = table[i][1][0].text
        pnotice = table[i][14]
        if len(pnotice):
            pnotice = pnotice[0].text
        else:
            pnotice = ''
        p.append([pid, pname, punit, pamount, prow, pnotice])

    print p

    f2 = open('result2.csv',  'w')

    f2.write(','.join(PRODUCT) + '\n')
    for r in p:
        f2.write(','.join([i.encode('utf-8') for i in r]) + '\n')


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
    print row_range.encode('utf-8')
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


if __name__ == '__main__':
    xmlparser('RFQ412457-Response.xml')






