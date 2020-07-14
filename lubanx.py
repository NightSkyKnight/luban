# coding=utf-8

import requests
import random
import json
import time
import xlwt
import re


# 存放IP
ip_list = []
# 浏览器标识
user_agents = ['Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3704.400 QQBrowser/10.4.3587.400', 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36',
               'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SE 2.X MetaSr 1.0; SE 2.X MetaSr 1.0; .NET CLR 2.0.50727; SE 2.X MetaSr 1.0)', 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.4094.1 Safari/537.36']
# 商品URL
list_url = ['链接']
# 商品名称
list_page = ['产品']
# 商品价格
list_price = ['单价']
# 商品公司名称
list_companyName = ['公司名']
# 联系电话
list_mobile = ['电话']
# 销量
list_sell_num = ['销量']
# 每个商品的ID
list_goodsId = []

# 获取代理IP


def get_ip():
    global ip_list
    ip_shumu = 0
    ip_chengg = 0
    # 请求头
    header = {
        'User-Agent': user_agents[random.randint(0, len(user_agents)-1)],
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Host': 'www.89ip.cn'
    }
    # 获取页数
    print('获取IP中。。。')
    for i in range(1, 2):
        ip_url = 'http://www.89ip.cn/index_{}.html'.format(i)
        try:
            ip_html = requests.get(ip_url, headers=header)
        except ConnectionError:
            time.sleep(2)
            ip_html = requests.get(ip_url, headers=header)
        if ip_html.status_code == 200:
            print("连接IP代理网站成功。。。")
        # 替换掉HTML的空格和换行
        html = ip_html.text.replace(" ", "").replace(
            "\n", "").replace("\t", "")
        # 匹配IP和端口的正则表达式
        r = re.compile('<tr><td>(.*?)</td><td>(.*?)</td><td>')
        # 匹配到的IP与端口
        ip_data = re.findall(r, html)
        ip_shumu += len(ip_data)
        for k in range(len(ip_data)):
            # 拼接IP与端口
            ip = "https://" + ip_data[k][0] + ":" + ip_data[k][1]
            ip_a = {"https://": ip}
            # 测试可不可用
            ping = requests.get("https://www.baidu.com", proxies=ip_a)
            if ping.status_code == 200:
                ip_list.append(ip_a)
                ip_chengg += 1
    print('获取到的IP数：{0}\n有效的IP数：{1}'.format(ip_shumu, ip_chengg))


# 获取商品列表
def get_list():
    global list_url
    global list_price
    header = {
        'Host': 'www.lubanx.cn',
        'Connection': 'close',
        'Cache-Control': 'max-age=0',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': user_agents[random.randint(0, len(user_agents)-1)],
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7',
        'Cookie': 'session=ZGU4NzEyMTAtZjMxYi00MjY1LTgzZTktMTA1YTI1NDM3NmI3; Hm_lvt_addf109ac734cd9bb1733b8328cc9c5d=1594195801; vue_admin_template_token=undefined; Hm_lpvt_addf109ac734cd9bb1733b8328cc9c5d=1594198160',
    }
    url = 'https://www.lubanx.cn/cod/jinritemai/goodsRank?p=1&s=50&keyword=&sortRule=0&firstCategory=0&secondCategory=0&thirdCategory=0&storeType=0&onShelf=0'
    get_list_data = requests.get(url, headers=header)
    # print(get_list_data.text)

    time.sleep(1)

    for i in range(len(get_list_data.json()['data']['list'])):
        list_url.append(get_list_data.json()['data']['list'][i]['goodsUrl'])
        list_price.append(get_list_data.json()['data']['list'][i]['price'])
        list_companyName.append(get_list_data.json()[
                                'data']['list'][i]['companyName'])
        list_page.append(get_list_data.json()['data']['list'][i]['pageTitle'])
        list_goodsId.append(get_list_data.json()['data']['list'][i]['goodsId'])


# 获取电话和销量
def get_page():
    global list_mobile
    header = {
        'Host': 'ec.snssdk.com',
        'Connection': 'close',
        'Cache-Control': 'max-age=0',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': user_agents[random.randint(0, len(user_agents)-1)],
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7'
    }
    for i in range(len(list_url)):
        # if i >= 5:
        #     return True
        try:
            url = 'https://ec.snssdk.com/product/lubanajaxstaticitem?id={}&page_id=&scope_type=5&b_type_new=0'.format(
                list_goodsId[i])
        except IndexError:
            return True

        get_page_data = requests.get(url, headers=header, proxies=ip_list[
            random.randint(0, len(ip_list) - 1)])
        # print(get_page_data.text)
        list_mobile.append(get_page_data.json()['data']['mobile'])
        list_sell_num.append(get_page_data.json()['data']['sell_num'])

        time.sleep(2)


def excel():
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('数据表')
    alignment = xlwt.Alignment()
    # 设置垂直居中
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    # 创建一个样式对象，初始化样式
    style = xlwt.XFStyle()
    style.alignment = alignment

    for i in range(len(list_url)):
        # for i in range(5):
        if i == 0:
            worksheet.write(0, 0, list_companyName[i], style)
            worksheet.write(0, 1, list_mobile[i], style)
            worksheet.write(0, 2, list_url[i], style)
            worksheet.write(0, 3, list_page[i], style)
            worksheet.write(0, 4, list_price[i], style)
            worksheet.write(0, 5, list_sell_num[i], style)
            continue

        worksheet.write(i, 0, list_companyName[i])
        worksheet.write(i, 1, list_mobile[i], style)
        worksheet.write(i, 2, list_url[i])
        worksheet.write(i, 3, list_page[i])
        worksheet.write(i, 4, list_price[i], style)
        worksheet.write(i, 5, list_sell_num[i], style)

    worksheet.col(0).width = 8000
    worksheet.col(1).width = 5000
    worksheet.col(2).width = 8500
    worksheet.col(3).width = 8000
    worksheet.col(4).width = 4000
    worksheet.col(5).width = 4000
    workbook.save('数据表.xls')


if __name__ == "__main__":
    get_ip()
    get_list()
    get_page()
    excel()
