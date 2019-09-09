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
f_list = []
for ip in range(2,ip_shu+2):
    f = 'F' + str(ip)
    f_list.append(f)

ip_list = []
ld_list = []


for host in hosts:
    print('\n---------------------------------------------------------------------------\n')
    print (host.get_text(), host.get('href'))
    soup = BeautifulSoup(open(url + host.get('href') ,encoding='utf8'),features="html.parser")
    vuls = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    #level_dangers = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    #存入字典并转发成列表
    temp = []
    for vul in vuls:
        ## 威胁等级
        level = vul.get('class')[0]
        ## 漏洞名称
        v = vul.get_text()
        ## 去重
        if v in temp:
            continue
        temp.append(v)
    ld_list.append(temp)

a = []
b = len(ld_list)

for i in range(0,b):
    for ld in ld_list[i]:
        ld = (ld + '\n')
        a.append(ld)
  
print ("********************")
# print (a)
# print (len(a))

zzz = zip(f_list,a)
print (zzz(k,v))

'''
#sht.range('F2').value = test_a
print ("****************************")

for F,test_a in zip (f_list,a):
    print (F,test_a)

'''
'''
sht.range('F1').value = ['漏洞详情']
    #sht.range('A2:B2:C2:D2:E2').value = [host.get_text(),gao,zhong,di,zong_len]   

for F,LD in zip (f_list,ld_list):
    sht.range(F).value = LD
'''
wb.save('ccc.xlsx')
wb.close()
app.quit()
exit()

print('\n---------------------------------------------------------------------------\n')

