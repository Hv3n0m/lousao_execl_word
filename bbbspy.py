# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import xlwings as xw
import os

'''
Date:2019-9-6
Auth:DASHU
Version:2
抓取漏扫结果导入到excel和word。
'''

app = xw.App(visible=True, add_book=False)
wb = app.books.add()
ws = wb.sheets.active
sht = wb.sheets[0]




## 兼容问题，修改路径兼容mac
#url = os.getcwd() + '/'
url = os.getcwd() + '\\'
print (url)

soup = BeautifulSoup(open('index.html', encoding='utf8'),features='html.parser')
hosts = soup.select('#content>div:nth-child(6)>div:nth-child(2)>table>tbody>tr>td>a')


ip_shu = len(hosts)
print ('存活IP个数', ip_shu)


#单元格位置和数量从A2开始
a_list = []
b_list = []
c_list = []
d_list = []
e_list = []
f_list = []
for ip in range(2,ip_shu+2):
    a = 'A' + str(ip)
    b = 'B' + str(ip)
    c = 'C' + str(ip)
    d = 'D' + str(ip)
    e = 'E' + str(ip)
    f = 'F' + str(ip)
    a_list.append(a)
    b_list.append(b)
    c_list.append(c)
    d_list.append(d)
    e_list.append(e)
    f_list.append(f)

ip_list = []
gaoshu_list = []
zhongshu_list = []
dishu_list = []
zongshu_list = []
ld_list = []



for host in hosts:
    print('\n---------------------------------------------------------------------------\n')
    print (host.get_text(), host.get('href'))
    soup = BeautifulSoup(open(url + host.get('href') ,encoding='utf8'),features="html.parser")
    vuls = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    #level_dangers = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    ip_list.append(host.get_text())
    #存入字典并转发成列表
    temp = []
    high_list = []
    middle_list = []
    low_list = []
    zong_list = []
    for vul in vuls:
        ## 威胁等级
        level = vul.get('class')[0]
        ## 漏洞名称
        v = vul.get_text()
        ## 去重
        if v in temp:
            continue
        temp.append(v)
        ## 漏洞等级分类并替换漏洞等级为中文
        if level == 'level_danger_low':
            low_list.append( {
            'level_danger' : '低',
            'vul' : v,
        })
        elif level == 'level_danger_middle':
            middle_list.append( {
            'level_danger' : '中',
            'vul' : v,
        })
        else:
            high_list.append( {
            'level_danger' : '高',
            'vul' : v,
        })
    for t in temp:
        t = (t + '\r')
        ld_list.append(t)
        print (t)
 

    zong_list = low_list + middle_list + high_list
    gaoshu = len(high_list)
    gaoshu_list.append(gaoshu)
    zhongshu = len(middle_list)
    zhongshu_list.append(zhongshu)
    dishu = len(low_list)
    dishu_list.append(dishu)
    zongshu = len(zong_list)
    zongshu_list.append(zongshu)


sht.range('A1:B1:C1:D1:E1:F1').value = ['IP地址','高','中','低','总','漏洞详情']
    #sht.range('A2:B2:C2:D2:E2').value = [host.get_text(),gao,zhong,di,zong_len]   
for A,IP in zip (a_list,ip_list):        
    sht.range(A).value = IP
for B,G in zip (b_list,gaoshu_list):        
    sht.range(B).value = G
for C,ZH in zip (c_list,zhongshu_list):        
    sht.range(C).value = ZH
for D,Di in zip (d_list,dishu_list):        
    sht.range(D).value = Di
for E,Z in zip (e_list,zongshu_list):        
    sht.range(E).value = Z
for F,LD in zip (f_list,ld_list):
    sht.range(F).value = LD

wb.save('bbb.xlsx')
wb.close()
app.quit()
exit()

print('\n---------------------------------------------------------------------------\n')

