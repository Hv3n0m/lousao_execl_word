# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import xlwings as xw
import os

'''
Date:2019-8-26
Auth:DASHU
Version:2
抓取漏扫结果导入到excel和word。
'''




## 兼容问题，修改路径兼容mac
#url = os.getcwd() + '/'
url = os.getcwd() + '\\'
print (url)

app = xw.App(visible=True, add_book=False)
wb = app.books.add()
ws = wb.sheets.active
sht = wb.sheets[0]

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
gaoshu = []
zhongshu_list = []
dishu_list = []
zongshu_list = []



for host in hosts:
    print('\n---------------------------------------------------------------------------\n')
    print (host.get_text(), host.get('href'))
    soup = BeautifulSoup(open(url + host.get('href') ,encoding='utf8'),features="html.parser")
    vuls = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    level_dangers = soup.select('#vuln_list > tbody > tr > td > ul > li > div > span')
    
    ip_list.append(host.get_text())
    #存入字典并转发成列表
    temp = []
    high_list = []
    middle_list = []
    low_list = []
    for level_danger,vul in zip (level_dangers,vuls):
        ## 威胁等级
        level = level_danger.get('class')[0]
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

    for gg in high_list:
        print (gg['level_danger'], gg['vul'])

    for zz in middle_list:
        print (zz['level_danger'], gg['vul'])

    for dd in low_list:
        print (dd['level_danger'], gg['vul'])
    

    #漏洞个数

    

    gao = len(high_list)
    print ('高危漏洞', gao)
    zhong = len(middle_list)
    print ('中危漏洞' ,zhong)
    di = len(low_list)
    print ('低危漏洞' , di)
    zong = len(high_list) + len(middle_list) + len(low_list)
    print ('总体漏洞' , zong)

    gaoshu.append(gao)
    zhongshu_list.append(zhong)
    dishu_list.append(di)
    zongshu_list.append(zong)
    

    zong_list = high_list + middle_list + low_list
    #漏洞详情添加换行
    ld_list = []

    for lou in zong_list:
        loudong = (lou['vul'])
        ld_list.append(loudong)
        ld_list.append('\n')
    ld_list = ' '.join(ld_list)
    print (ld_list)
    print (type(ld_list))




sht.range('A1:B1:C1:D1:E1:F1').value = ['IP地址','高','中','低','总','漏洞详情']
    #sht.range('A2:B2:C2:D2:E2').value = [host.get_text(),gao,zhong,di,zong_len]   
for A,IP in zip (a_list,ip_list):        
    sht.range(A).value = IP
for B,G in zip (b_list,gaoshu):        
    sht.range(B).value = G
for C,ZH in zip (c_list,zhongshu_list):        
    sht.range(C).value = ZH
for D,Di in zip (d_list,dishu_list):        
    sht.range(D).value = Di
for E,Z in zip (e_list,zongshu_list):        
    sht.range(E).value = Z
for F in f_list:        
    sht.range(F).value = ld_list


wb.save('aaa.xlsx')
wb.close()
app.quit()
exit()

print('\n---------------------------------------------------------------------------\n')
