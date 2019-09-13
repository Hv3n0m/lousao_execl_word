# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import xlwings as xw
import os
import time

'''
Date:2019-9-13
Auth:DASHU
Version:4

抓取漏扫结果导入到excel和word。
已解决导入excel的功能，下次更新导入word
'''

app = xw.App(visible=True, add_book=False)
wb = app.books.add()
ws = wb.sheets.active
sht = wb.sheets[0]
sht.range('A1:B1:C1:D1:E1:F1').value = ['IP地址','高','中','低','总','漏洞详情']


## 兼容问题，修改路径兼容mac
#url = os.getcwd() + '/'
url = os.getcwd() + '\\'
print (url)
host = 'host/192.168.50.200.html'


gaoshu = []
zhongshu = []
dishu = []
zongshu = []



soup = BeautifulSoup(open('index.html', encoding='utf8'),features='html.parser')
hosts = soup.select('#content>div:nth-child(6)>div:nth-child(2)>table>tbody>tr>td>a')
ips = len(hosts)
print ('存活IP个数', ips)

ips_list = []
vals = []
for host in hosts:
    soup = BeautifulSoup(open(url + host.get('href') ,encoding='utf8'),features="html.parser")
    vuls = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    ips_list.append(host.get_text())
    
    high_list = []
    middle_list = []
    low_list = []
    val = []
    
    for vul in vuls:
        # 威胁等级
        level = vul.get('class')
        # 漏洞名称
        v = vul.get_text()
        # 去重
        if v in val:
            continue
        # 漏洞详情
        val.append(v)
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
    # 漏洞详情list转str    
    val = "\n".join(val)
    #print (val)
    vals.append(val)

    # 漏洞个数
    gao = len(high_list)
    gaoshu.append(gao)
    zhong = len(middle_list)
    zhongshu.append(zhong)
    di = len(low_list)
    dishu.append(di)
    zong = len(high_list) + len(middle_list) + len(low_list)
    zongshu.append(zong)



aa = []
bb = []
cc = []
dd = []
ee = []
ff = []
for ip in range(2,ips+2):
    a = 'A' + str(ip)
    b = 'B' + str(ip)
    c = 'C' + str(ip)
    d = 'D' + str(ip)
    e = 'E' + str(ip)
    f = 'F' + str(ip)
    aa.append(a)
    bb.append(b)
    cc.append(c)
    dd.append(d)
    ee.append(e)
    ff.append(f)




for A,IP in zip (aa,ips_list):        
    sht.range(A).value = IP
for B,G in zip (bb,gaoshu):        
    sht.range(B).value = G
for C,ZH in zip (cc,zhongshu):        
    sht.range(C).value = ZH
for D,Di in zip (dd,dishu):        
    sht.range(D).value = Di
for E,Z in zip (ee,zongshu):        
    sht.range(E).value = Z
for F,v in zip(ff,vals):
    sht.range(F).value = v





wb.save('ccc.xlsx')
wb.close()
app.quit()
exit()

print('\n---------------------------------------------------------------------------\n')

