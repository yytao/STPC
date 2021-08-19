import urllib3
import xlrd
from xlutils.copy import copy
from bs4 import BeautifulSoup
import re

#初始化链接字典
url = {
    '要闻':"http://chisa.edu.cn/rmtnews1/ssyl/",
    '时评':"http://chisa.edu.cn/rmtnews1/guandian/"
};

titleKeyword    = "";
contentKeyword  = "";
beginDate       = "20210819";
endDate         = "20210819";
wbook = xlrd.open_workbook('./article.xls')
newbook = copy(wbook)
newsheet = newbook.get_sheet(0)

# Return Array
def getContent(fatherUrl, url):
    http = urllib3.PoolManager()
    date = re.findall(r'.*?t(\d+)_.*?', (url))
    date = ''.join(date)
    if date > endDate or date < beginDate :
        return False;
    url = (fatherUrl+url);
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        return None

    content = response.data.decode()
    html = BeautifulSoup(content, features='html.parser')
    divList = html.find_all("div", {"class", "leftpart"})

    try:
        item = divList[0]
        item = str(item)
        item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
        item = re.sub('\n', '', item)
        item = re.sub('\s', '', item)

        title = re.findall(r'<h1class="content_title">(.*?[\u4E00-\u9FA5].*?[\u4E00-\u9FA5].*?[\u4E00-\u9FA5]+)</h1>', item)

        print(title)
        
    except BaseException as err:
        print(err)
        return None

    


#处理顶级列表urlData
def processData(url, row = 0):
    
    http = urllib3.PoolManager()
    
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        return None
    #获取状态码，如果是200表示获取成功
    code = response.status
    
    #读取内容
    if 200 == code:
        content = response.data.decode()
        html = BeautifulSoup(content, features='html.parser')
        divList = html.find_all("div", {"class", "hnews block nopic"})
        for item in divList:
            item = str(item)
            item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
            item = re.sub('\n', '', item)
            item = re.sub('\s', '', item)

            articleContentHerf = re.findall(r'<divclass="txtconthline">.*?<ahref="(.*?.html)".*?>.*?</a>.*?</div>', (item))
            articleContentHerf = ''.join(articleContentHerf)
            
            articleContent = getContent(url, articleContentHerf)
            exit()



for (key,item) in url.items() :
    processData(item)