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
keywords = config['keywords']

# HTTP 请求头
headers = config['headers']

# 设置代理
if 'proxies' in config:
    proxies = config['proxies']
else:
    proxies = {}

# 博文列表
blogList = []

pageIndex = startPage
while pageIndex <= maxPage:
    print('page = ' + str(pageIndex))
    url = 'https://s.weibo.com/weibo?q=' + "%20".join(keywords) + '&wvr=6&Refer=SWeibo_box&page=' + str(pageIndex)
    html = requests.get(url, headers=headers, proxies=proxies)
    soup = BeautifulSoup(html.text, 'html.parser')
    newBlogs = soup.findAll('div', attrs={'class': 'content', 'node-type': 'like'})
    # 如果在当前检索页面已经没有新博文，则直接终止
    if len(newBlogs) == 0:
        print('no more content on page ' + str(pageIndex))
        break
    blogList += newBlogs
    pageIndex += 1

# 新建一个工作薄
wb = xlwt.Workbook()
sheet1 = wb.add_sheet('sheet1', cell_overwrite_ok=True)  # cell_overwrite_ok=true 使同一个单元可以重设值

# 表头
sheet1.write(0, 0, 'Username')
sheet1.write(0, 1, 'Content')
sheet1.write(0, 2, 'BlogUrl')
sheet1.write(0, 3, 'Time')
sheet1.write(0, 4, 'Device')

# 遍历所有博文
for i in range(len(blogList)):
    username = ''
    content = ''
    blogUrl = ''
    time = ''
    device = ''

    # 记录博客内容的元素
    contentElement = None
    # 记录博客来源信息的元素
    fromElement = None

    # 接下来要在该博文的子元素中找到 contentElement 与 fromElement
    for aChild in blogList[i].children:
        # 跳过空白子元素
        if 'attrs' not in dir(aChild):
            continue
        # 某些博文内容较长，存在全文元素 'feed_list_content_full' 以及非全文元素 'feed_list_content'，此时需要选用全文元素
        if 'node-type' in aChild.attrs and aChild.attrs['node-type'] == 'feed_list_content_full':
            contentElement = aChild
        # 如果是非全文的内容元素，并且之前没有记录过全文元素，则记录
        elif 'node-type' in aChild.attrs and aChild.attrs['node-type'] == 'feed_list_content' and contentElement == None:
            contentElement = aChild
        # 如果是 from 元素，则记录
        elif 'class' in aChild.attrs and 'from' in aChild.attrs['class']:
            fromElement = aChild

    username = contentElement.attrs['nick-name']
    content = contentElement.text

    childListOfFromElement = []

    # 接下来要在 fromElement 中找到发布时间子元素和设备信息子元素
    for aChild in fromElement:
        # 跳过空白子元素
        if 'attrs' not in dir(aChild):
            continue
        childListOfFromElement.append(aChild)

    blogUrl = 'https:' + str(childListOfFromElement[0].attrs['href'])
    time = str(childListOfFromElement[0].text).replace(' ', '')

    # 某些博文没有设备信息
    if len(childListOfFromElement) > 1:
        device = childListOfFromElement[1].text

    sheet1.write(i+1, 0, username)
    sheet1.write(i+1, 1, content)
    sheet1.write(i+1, 2, blogUrl)
    sheet1.write(i+1, 3, time)
    sheet1.write(i+1, 4, device)

wb.save('./' + '_'.join(keywords) + '.xls')