import requests
from bs4 import BeautifulSoup
import xlwt

# 检索起始页和最大页
startPage = 1
maxPage = 50


# 检索关键字
# options = ['腾讯会议', '卡', '热', '电']
options = ['腾讯']

# HTTP 请求头
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36',
    'Cookie': 'SINAGLOBAL=5530116431946.579.1598008684649; UOR=,,login.sina.com.cn; _s_tentry=login.sina.com.cn; Apache=4805368567449.919.1599619236716; ULV=1599619236780:9:7:6:4805368567449.919.1599619236716:1599563553506; login_sid_t=672b8a3d051b13446bd7aba308de3e49; cross_origin_proto=SSL; ALF=1631155442; SSOLoginState=1599619443; SCF=Ap4t2IOhQCYT5mSe6VDfof3YgYuFVWKy1KXLix8xQ0vks8WC-TYaxE85Dk87yMYXcAyaW66wOyWqHba5GXJ7BkI.; SUB=_2A25yXDEjDeRhGeNI7VQZ9y7KyzyIHXVRKCXrrDV8PUNbmtANLVPQkW9NSFeDTHgfum6fPblHzafmormPKZDHanUI; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WWAbREemV7B2HqNo2rJblYS5JpX5KzhUgL.Fo-cSoqRS05ceh52dJLoIpjLxKqLBoBLB-2LxKqLB.BL1hMLxK-LBKBLBK.t; SUHB=06ZFUY3b8SgLjH; wvr=6; webim_unReadCount=%7B%22time%22%3A1599619750004%2C%22dm_pub_total%22%3A0%2C%22chat_group_client%22%3A0%2C%22chat_group_notice%22%3A0%2C%22allcountNum%22%3A0%2C%22msgbox%22%3A0%7D'
}

blogList = []
fromList = []

pageIndex = startPage
while pageIndex <= maxPage:
    print('page = ' + str(pageIndex))
    url = 'https://s.weibo.com/weibo?q=' + "%20".join(options) + '&wvr=6&Refer=SWeibo_box&page=' + str(pageIndex)
    html = requests.get(url, headers=headers)
    soup = BeautifulSoup(html.text, 'html.parser')

    # 这些 HTML 元素包含
    newBlogs = soup.findAll('p', attrs={'class': 'txt', 'node-type': 'feed_list_content'})
    newFroms = soup.findAll('p', attrs={'class': 'from'})

    # 如果在当前检索页面已经没有内容，则直接终止
    if len(newBlogs) == 0:
        break

    blogList += newBlogs
    fromList += newFroms
    pageIndex += 1

print(len(blogList))
print(len(fromList))

wb = xlwt.Workbook()
sheet1 = wb.add_sheet('sheet1', cell_overwrite_ok=True)  # cell_overwrite_ok=true 使同一个单元可以重设值
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

    # 获取某些博文的 nick-name 时出现 keyError，暂时未找到原因，先跳过这些博文
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
