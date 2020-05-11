import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup


headers = {
    'Authorization':'Basic ****',
    'Cookie':'user_key=***',
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.10240',
    'Host':'192.168.50.58:60001'
}
url = 'https://IP:端口/report.php?do=threport&at=att_info'
r = requests.post(url, headers=headers, verify=False)
r.encoding = 'utf-8'
html = r.text
soup = BeautifulSoup(html, 'html.parser')
soup = soup.select('#att_info > div > div:nth-child(1) > div > div:nth-child(2)')

print (soup)
