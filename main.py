import urllib3
import xlrd
from xlutils.copy import copy
from bs4 import BeautifulSoup
import re
import sys

# 初始化链接字典
url = {
    '要闻': "http://chisa.edu.cn/rmtnews1/ssyl/",
    '时评': "http://chisa.edu.cn/rmtnews1/guandian/",
    '海外': "http://chisa.edu.cn/rmtnews1/haiwai/",
    '人才': "http://chisa.edu.cn/rmtnews1/rencai/",
    '综合': "http://chisa.edu.cn/rmtnews1/zonghe/",
    '原创': "http://chisa.edu.cn/rmtycgj/",
    '创业': "http://chisa.edu.cn/rmtnews1/chuangye/",
    '留学': "http://chisa.edu.cn/rmtnews1/chuguo/"
}

#url = {
#    '原创': "http://www.chisa.edu.cn/rmtycgj/",
#}

# 初始化参数
KEYWORD = []  # 该元组的条件为&&
COLUMN = ['from']#默认只检索title，新增content或者from
CONTENT_KEYWORD = ['新华社'];   # 是否给comtent自定义keyword检索，留空则共用KEYWORK变量

BEGINDATE = "20210101"
ENDDATE = "20211012"
ROWLINE = 0

# 复制一份article.xls，其实每次就是以新文件覆盖了旧文件
wbook = xlrd.open_workbook('./article.xls')
newbook = copy(wbook)
newsheet = newbook.get_sheet(0)
newsheet.width = 0x0d00 + 50

def writeInFile(content):
    global ROWLINE
    row = ROWLINE
    for item in content:
        if item:
            for val in item:
                global newsheet
                newsheet.write(row, 0, str(val['title']))
                newsheet.write(row, 1, str(val['source']))
                newsheet.write(row, 2, str(val['editor']))
                newsheet.write(row, 3, str(val['href']))
                row = row + 1
    ROWLINE = row + 1

# 在参数content中，匹配元组KEYWORD中的每一个元素，如有某个元素匹配不到，则返回False
def matchByKeyword(content, keyword = KEYWORD):
    result = True
    for item in keyword:
        if item not in content:
            result = False
    return result


# 此函数接收文章url，抓取内容并根据条件进行返回
def getContent(fatherUrl, url):
    http = urllib3.PoolManager()

    # 匹配url内的日期，首先进行日期范围判断
    date = re.findall(r'.*?t(\d+)_.*?', (url))
    date = ''.join(date)
    if date > ENDDATE or date < BEGINDATE:
        return None

    # 拼写完整url，抓取内容
    url = (fatherUrl+url)
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        return None

    # 格式化抓取数据
    content = response.data.decode()
    html = BeautifulSoup(content, features='html.parser')

    # 单独抓取到PC模板的leftpart内容
    #divList = html.find_all("div", {"class", "leftpart"})
    divList = html.find_all("html")
    try:
        # 格式化数据
        item = divList[0]
        item = str(item)
        item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
        item = re.sub('\n', '', item)
        item = re.sub('\s', '', item)

        title = re.findall(r'<h1class="content_title">(.*?)</h1>', item)
        title = re.findall(r'<title>(.*?)</title>', item)
        title = ''.join(title)
        
        # title条件筛选
        if 'title' in COLUMN:
            titleIsLegal = matchByKeyword(title, KEYWORD)
            if titleIsLegal == False:
                return None

        # content条件筛选
        if 'content' in COLUMN:
            content = re.findall(
                r'<divclass="detail"id="js_content">(.*?)</div>', item)
            content = ''.join(content)
            keyword = CONTENT_KEYWORD if CONTENT_KEYWORD else KEYWORD
            contentIsLegal = matchByKeyword(content, keyword)
            if contentIsLegal == False:
                return None

        if 'from' in COLUMN:
            fromContent = re.findall(
                r'<divclass="from">(.*?)</div>', item)
            fromContent = ''.join(fromContent)
            keyword = CONTENT_KEYWORD if CONTENT_KEYWORD else KEYWORD
            fromContentIsLegal = matchByKeyword(fromContent, keyword)
            if fromContentIsLegal == False:
                return None
        

        # 定义result字典，用于返回给调用函数
        result = {}
        result['title'] = title

        result['source'] = re.findall(
            r'<divclass="from">.*?</script>来源：(.*?)<script>.*?</div>', item)
        result['source'] = ''.join(result['source'])

        result['editor'] = re.findall(r'<pclass="more">责任编辑：(.*?)</p>', item)
        result['editor'] = ''.join(result['editor'])

        result['href'] = url

        return result

    except BaseException as err:
        print(err)
        return None

# 处理顶级列表urlData


def processData(url, row=0, rootUrl=False):

    # 创建http连接池
    http = urllib3.PoolManager()

    # 抓取一级目录列表
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        return None
    # 获取状态码，如果是200表示获取成功
    code = response.status

    # 读取内容
    if 200 == code:
        content = response.data.decode()
        html = BeautifulSoup(content, features='html.parser')

        # 将每个文章列表分离为单独的元素
        result = []
        divList = html.find_all("div", {"class", "hnews block nopic"})
        for item in divList:

            # 格式化数据
            item = str(item)
            item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
            item = re.sub('\n', '', item)
            item = re.sub('\s', '', item)

            # 匹配获取到文章的url
            articleContentHerf = re.findall(
                r'<divclass="txtconthline">.*?<ahref="(.*?.html)".*?>.*?</a>.*?</div>', (item))
            articleContentHerf = ''.join(articleContentHerf)

            # url传递给getContent函数
            articleContent = getContent(
                rootUrl if rootUrl else url, articleContentHerf)
            if articleContent != None:
                result.append(articleContent)
        return result

for (key, item) in url.items():
    
    writeIn = []
    writeIn.append(processData(item))

    i = 1
    while i <= 50:
        writeIn.append(processData(item + "index_" + str(i) + ".html", 0, item))
        i = i + 1

    # 写入excel
    writeInFile(writeIn)

newbook.save('./article.xls')
