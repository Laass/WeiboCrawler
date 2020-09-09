import requests
import xlwt
import json

from bs4 import BeautifulSoup

# 载入配置文件
with open('./config.json', 'r', encoding='utf-8') as configFile:
    config = json.load(configFile)

# 检索起始页码和最大页码
startPage = config['startPage']
maxPage = config['maxPage']

# 检索关键字
options = config['options']

# HTTP 请求头
headers = config['headers']

# 设置代理
proxies = config['proxies']

blogList = []
fromList = []

pageIndex = startPage
while pageIndex <= maxPage:
    print('page = ' + str(pageIndex))
    url = 'https://s.weibo.com/weibo?q=' + "%20".join(options) + '&wvr=6&Refer=SWeibo_box&page=' + str(pageIndex)
    html = requests.get(url, headers=headers, proxies=proxies)
    soup = BeautifulSoup(html.text, 'html.parser')

    # newBlogs 包含当前页面所有博文内容及用户名信息
    newBlogs = soup.findAll('p', attrs={'class': 'txt', 'node-type': 'feed_list_content'})

    # newFroms 包含每一条博文的发布时间、Url、设备信息
    newFroms = soup.findAll('p', attrs={'class': 'from'})

    # 如果在当前检索页面已经没有内容，则直接终止
    if len(newBlogs) == 0:
        break

    blogList += newBlogs
    fromList += newFroms
    pageIndex += 1

print(len(blogList))
print(len(fromList))

# 新建一个工作薄
wb = xlwt.Workbook()
sheet1 = wb.add_sheet('sheet1', cell_overwrite_ok=True)  # cell_overwrite_ok=true 使同一个单元可以重设值

# 统计的内容如下：
sheet1.write(0, 0, 'Username')
sheet1.write(0, 1, 'Content')
sheet1.write(0, 2, 'BlogUrl')
sheet1.write(0, 3, 'Time')
sheet1.write(0, 4, 'Device')

for i in range(len(blogList)):
    username = ''
    content = ''
    blogUrl = ''
    time = ''
    device = ''

    # 获取某些博文的 nick-name 时出现 keyError，暂时未找到原因，先跳过这些博文的 nick-name
    if 'nick-name' in blogList[i].attrs:
        username = blogList[i].attrs['nick-name']

    content = blogList[i].text

    childList = []
    for aChild in fromList[i].children:
        if len(str(aChild)) > 0 and str(aChild)[0] == '<':
            childList.append(aChild)

    blogUrl = 'https:' + str(childList[0].attrs['href'])

    time = str(childList[0].text).replace(' ', '')

    # 某些博文没有设备信息
    if len(childList) > 1:
        device = childList[1].text

    sheet1.write(i+1, 0, username)
    sheet1.write(i+1, 1, content)
    sheet1.write(i+1, 2, blogUrl)
    sheet1.write(i+1, 3, time)
    sheet1.write(i+1, 4, device)

wb.save('./' + '_'.join(options) + '.xls')
