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

#初始化参数
keyword    = "一带一路";
cloumn = "";    #默认为title，传入both则同时匹配content
beginDate       = "20210819";
endDate         = "20210819";

#复制一份article.xls，其实每次就是以新文件覆盖了旧文件
wbook = xlrd.open_workbook('./article.xls')
newbook = copy(wbook)
newsheet = newbook.get_sheet(0)

def matchByKeyword():

    result = False

    #if cloumn==None:
    
    return result
    

# 此函数接收文章url，抓取内容并根据条件进行返回
def getContent(fatherUrl, url):
    http = urllib3.PoolManager()

    #匹配url内的日期，首先进行日期范围判断
    date = re.findall(r'.*?t(\d+)_.*?', (url))
    date = ''.join(date)
    if date > endDate or date < beginDate :
        return False;
    
    # 拼写完整url，抓取内容
    url = (fatherUrl+url);
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        return None

    #格式化抓取数据
    content = response.data.decode()
    html = BeautifulSoup(content, features='html.parser')

    #单独抓取到PC模板的leftpart内容
    divList = html.find_all("div", {"class", "leftpart"})

    try:
        #格式化数据
        item = divList[0]
        item = str(item)
        item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
        item = re.sub('\n', '', item)
        item = re.sub('\s', '', item)

        #定义result字典，用于返回给调用函数
        result = {}
        result['title'] = re.findall(r'<h1class="content_title">(.*?)</h1>', item)
        result['title'] = ''.join(result['title'])

        result['source'] = re.findall(r'<divclass="from">.*?</script>来源：(.*?)<script>.*?</div>', item)
        result['source'] = ''.join(result['source'])

        result['editor'] = re.findall(r'<pclass="more">责任编辑：(.*?)</p>', item)
        result['editor'] = ''.join(result['editor'])
        
        result['href'] = url

        print(result)        
        
    except BaseException as err:
        print(err)
        return None

    


#处理顶级列表urlData
def processData(url, row = 0):
    
    #创建http连接池
    http = urllib3.PoolManager()
    
    #抓取一级目录列表
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
        
        #将每个文章列表分离为单独的元素
        divList = html.find_all("div", {"class", "hnews block nopic"})
        for item in divList:

            #格式化数据
            item = str(item)
            item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
            item = re.sub('\n', '', item)
            item = re.sub('\s', '', item)
            
            #匹配获取到文章的url
            articleContentHerf = re.findall(r'<divclass="txtconthline">.*?<ahref="(.*?.html)".*?>.*?</a>.*?</div>', (item))
            articleContentHerf = ''.join(articleContentHerf)
            
            #url传递给getContent函数
            articleContent = getContent(url, articleContentHerf)


for (key,item) in url.items() :
    processData(item)