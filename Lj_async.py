#!/usr/bin/env python  
# encoding: utf-8 
"""
Author: ISeeMoon
Python: 3.6
Software: PyCharm
File: Lj_async.py
Time: 2018/5/6 15:26
"""
import requests
from lxml import etree
import asyncio
import aiohttp
import pandas
import re
import math
import time


loction_info = '''    1→杭州
    2→武汉
    3→北京
    按ENTER确认：'''
loction_select = input(loction_info)
loction_dic = {'1':'hz',
               '2':'wh',
               '3':'bj'}
city_url = 'https://{}.lianjia.com/ershoufang/'.format(loction_dic[loction_select])

inter_list = [(0,1000)]

def half_inter(inter):
    lower = inter[0]
    upper = inter[1]
    delta = int((upper-lower)/2)
    inter_list.remove(inter)
    print('已经缩小价格区间',inter)
    inter_list.append((lower, lower+delta))
    inter_list.append((lower+delta, upper))

pagenum = {}
def get_num(inter):
    url = city_url + 'bp{}ep{}/'.format(inter[0],inter[1])
    r = requests.get(url).text
    num = int(etree.HTML(r).xpath("//h2[@class='total fl']/span/text()")[0].strip())
    pagenum[(inter[0],inter[1])] = num
    return num

judge = True
while judge:
    a = [get_num(x)>3000 for x in inter_list]
    if True in a:
        judge = True
    else:
        judge = False
    for i in inter_list:
        if get_num(i) > 3000:
            half_inter(i)
print('价格区间缩小完毕！')

url_list = []
for i in inter_list:
    totalpage = math.ceil(pagenum[i]/30)
    for j in range(1,totalpage+1):
        url = city_url + 'pg{}bp{}ep{}/'.format(j,i[0], i[1])
        url_list.append(url)
print('url列表获取完毕！')

info_lst = []
async def get_info(url):
    async with aiohttp.ClientSession() as session:
        async with session.get(url) as resp:
            r = await resp.text()
            nodelist = etree.HTML(r).xpath("//ul[@class='sellListContent']/li")
            print('-------------------------------------------------------------')
            print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url),len(url_list)))
            print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url), len(url_list)))
            print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url), len(url_list)))
            print('-------------------------------------------------------------')
            info_dic = {}
            for node in nodelist:
                info_dic['title'] = node.xpath(".//div[@class='title']/a/text()")[0]
                info_dic['href'] = node.xpath(".//div[@class='title']/a/@href")[0]
                info_dic['xiaoqu'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ','').split('|')[0]
                info_dic['huxing'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[1]
                info_dic['area'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[2]
                info_dic['chaoxiang'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[3]
                if len(node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')) >= 5:
                    info_dic['zhuangxiu'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[4]
                else:
                    info_dic['zhuangxiu'] = '/'
                if len(node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')) == 6:
                    info_dic['dianti'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[5]
                else:
                info_dic['dianti'] = '/'
                info_dic['louceng'] = re.findall('\((.*)\)',node.xpath(".//div[@class='positionInfo']/text()")[0])
                info_dic['nianxian'] = re.findall('\)(.*?)年',node.xpath(".//div[@class='positionInfo']/text()")[0])
                info_dic['guanzhu'] = ''.join(re.findall('[0-9]',node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ','').split('/')[0]))
                info_dic['daikan'] = ''.join(re.findall('[0-9]',node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ', '').split('/')[1]))
                info_dic['fabu'] = node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ', '').split('/')[2]
                info_dic['totalprice'] = node.xpath(".//div[@class='totalPrice']/span/text()")[0]
                info_dic['unitprice'] = node.xpath(".//div[@class='unitPrice']/span/text()")[0].replace('单价','')
                info_lst.append(info_dic)
                print('{}→房屋信息抓取完毕！'.format(info_dic['title']))
                info_dic = {}

start = time.time()
tasks = [asyncio.ensure_future(get_info(url)) for url in url_list]
loop = asyncio.get_event_loop()
loop.run_until_complete(asyncio.wait(tasks))
end = time.time()
loop.close()
print('总共耗时{}秒'.format(end-start))

# df = pandas.DataFrame(info_lst)
# df.to_csv('C:\\test\\lj.csv',encoding='utf_8_sig')

#
# info_lst = []
# def get_info1(url):
#     r = requests.get(url).text
#     nodelist = etree.HTML(r).xpath("//ul[@class='sellListContent']/li")
#     print('-------------------------------------------------------------')
#     print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url),len(url_list)))
#     print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url), len(url_list)))
#     print('开始抓取第{}个页面的数据,共计{}个页面'.format(url_list.index(url), len(url_list)))
#     print('-------------------------------------------------------------')
#     info_dic = {}
#     for node in nodelist:
#         info_dic['title'] = node.xpath(".//div[@class='title']/a/text()")[0]
#         info_dic['href'] = node.xpath(".//div[@class='title']/a/@href")[0]
#         info_dic['xiaoqu'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ','').split('|')[0]
#         info_dic['huxing'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[1]
#         info_dic['area'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[2]
#         info_dic['chaoxiang'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[3]
#         if len(node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')) >= 5:
#             info_dic['zhuangxiu'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[4]
#         else:
#             info_dic['zhuangxiu'] = '/'
#         if len(node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')) == 6:
#             info_dic['dianti'] = node.xpath(".//div[@class='houseInfo']")[0].xpath('string(.)').replace(' ', '').split('|')[5]
#         else:
#             info_dic['dianti'] = '/'
#         info_dic['louceng'] = re.findall('\((.*)\)',node.xpath(".//div[@class='positionInfo']/text()")[0])
#         info_dic['nianxian'] = re.findall('\)(.*?)年',node.xpath(".//div[@class='positionInfo']/text()")[0])
#         info_dic['guanzhu'] = ''.join(re.findall('[0-9]',node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ','').split('/')[0]))
#         info_dic['daikan'] = ''.join(re.findall('[0-9]',node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ', '').split('/')[1]))
#         info_dic['fabu'] = node.xpath(".//div[@class='followInfo']/text()")[0].replace(' ', '').split('/')[2]
#         info_dic['totalprice'] = node.xpath(".//div[@class='totalPrice']/span/text()")[0]
#         info_dic['unitprice'] = node.xpath(".//div[@class='unitPrice']/span/text()")[0].replace('单价','')
#         info_lst.append(info_dic)
#         print('{}→房屋信息抓取完毕！'.format(info_dic['title']))
#         info_dic = {}
#
# start = time.time()
# for url in url_list:
#     get_info1(url)
# end = time.time()
# print('总共耗时{}秒'.format(end-start))