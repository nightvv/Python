#!/usr/bin/env python
# encoding: utf-8
"""
Author: ISeeMoon
Python: 3.6
Software: PyCharm
File: cos_ocr.py
Time: 2018/4/17 21:16
"""

from aip import AipOcr
import json
import time
import re
from pprint import pprint
from PIL import ImageEnhance
from PIL import  ImageFilter
from PIL import Image
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select

exce = load_workbook(r"C:\baiduOCR.xlsx")
sheet = exce.worksheets[0]
APP_ID = str(sheet['B1'].value)
API_KEY = sheet['B2'].value
SECRET_KEY = sheet['B3'].value
aipOcr = AipOcr(APP_ID, API_KEY, SECRET_KEY)

def get_file_content(filepath):
    with open(filepath,'rb') as fp:
        return fp.read()

imagepath = r"C:\jkhz.jpg"

image = Image.open(imagepath).convert('L')
size = image.size
cropimage = image.crop((size[0]*0.05,size[-1]*0.05,size[0]*0.95,size[-1]*0.7))
# cropimage.show()
sharpimage = cropimage.filter(ImageFilter.SHARPEN)
# sharpimage.show()
image_contrasted = ImageEnhance.Contrast(sharpimage).enhance(8)
# image_contrasted.show()
image_contrasted.save(imagepath)

im = get_file_content(imagepath)

recog_table = aipOcr.tableRecognitionAsync(im)
id = recog_table['result'][0]['request_id']
options_table = {
    'result_type':'json'
}

result_table = aipOcr.getTableRecognitionResult(id, options_table)

while 'error_code' in result_table or 'result' in result_table and result_table['result']['ret_code'] != 3:
    time.sleep(0.3)
    print('表格正在识别中...')
    aipOcr.getTableRecognitionResult(id, options_table)
    result_table = aipOcr.getTableRecognitionResult(id, options_table)
print('表格识别完毕!')

dic_table = json.loads(result_table['result']['result_data'])['forms'][0]['body']

info_dic = {}
for item in dic_table:
    x = item['row'][0]
    y = item['column'][0]
    word = item['word']
    info_dic[(x,y)] = word

def rowstr(x):
    i = 0
    judge = True
    str = ''
    while judge:
        if (x,i) in info_dic.keys():
            str += info_dic[(x,i)]
            i += 1
        else:
            judge = False
    return str
# pprint(info_dic)

JPMC = rowstr(0).split('中文')[-1]
JPYWMC = rowstr(1).split('外文')[-1]
XHGG = rowstr(2).split('送检')[0].split('格')[-1]
JPSL = rowstr(2).split('数量')[-1]
YPXZ = ''.join(rowstr(3).split('样品')[0:-1]).split(')')[-1].split('态')[-1]
SCRQHSCPH = rowstr(4).split('保质')[0].split('限')[0].split('生')[-1].split('或')[-1].split('期')[-1].split('号')[-1]
BZQ = rowstr(4).split('其')[-1].split('期')[-1].replace('习','')
SQQYMC = rowstr(8).split('称')[-1].split('生产')[0].replace('|','')
if len(rowstr(8).split(')')[-1]) != 0:
    SCGJDQ = ''.join(re.findall("[\u4e00-\u9fa5a]",rowstr(8).split(')')[-1]))
else:
    SCGJDQ = ''.join(re.findall("[\u4e00-\u9fa5a]", rowstr(8).split('国')[-1])).replace('()','').replace('地区','')
SQQYDZ = rowstr(9).split('址')[-1].split('联')[0].replace('名称','')
LXDH = ''.join(re.findall("\d|-|\.",rowstr(9).split('联系人')[0].split('联')[-1]))
LXR = ''.join(re.findall("[\u4e00-\u9fa5a-zA-Z0-9]",rowstr(9).split('人')[-1]))
ZHSBZRDWMC = rowstr(10).split('称')[-1].split('联')[0]
ZHSBDWDH = ''.join(re.findall("\d|-",rowstr(10)))
ZHSBDWLXR = rowstr(10).split('人')[-1].split('入')[-1].replace('|','')
ZHSBDWDZ = rowstr(11).split('址')[-1].split('传')[0].replace('|','')
ZHSBDWCZ = rowstr(11).split('真')[-1].split('邮')[0].replace('|','').replace('[','').replace('\\','').replace('「','')
ZHSBDWYB = rowstr(11).split('邮编')[-1].replace('|','')
CZ = '/'
JYYJ = '化妆品安全技术规范（2015年版）'
LYSL = ''
JPSLDW = rowstr(2).split('数量')[-1][-1]
SLBH = 'JF0182018'
JYZQ = '60'
YPLB = ''
ZJKS = '保化室'

print('''
      ---------------------------------------------------------------------------
      |中文名：{}
      |英文名：{}
      |规格：{}
      |颜色物态：{}
      |检品数量：{}
      |生产日期/批号：{}
      |保质期：{}
      |生产国家：{}
      ---------------------------------------------------------------------------
      |申请企业：{}
      |申请企业地址：{}
      |申请企业电话：{}
      |申请企业联系人：{}
      ---------------------------------------------------------------------------
      |在华企业：{}
      |在华企业地址：{}
      |在华企业电话：{}
      |在华企业联系人：{}
      |在华企业传真：{}
      |邮编：{}
      ---------------------------------------------------------------------------
      '''.format(JPMC,JPYWMC,XHGG,YPXZ,JPSL,SCRQHSCPH,BZQ,SCGJDQ,SQQYMC,SQQYDZ,LXDH,LXR,ZHSBZRDWMC,ZHSBDWDZ,ZHSBDWDH,
                 ZHSBDWLXR,ZHSBDWCZ,ZHSBDWYB))


def login():
    global wb
    wb = webdriver.Chrome("C:\chromedriver.exe")
    wb.implicitly_wait(30)
    wb.maximize_window()
    wb.get('http://10.0.144.50/web/Login.aspx')
    wb.find_element_by_xpath("//input[@id='txtUserCode']").send_keys(limis_name)
    wb.find_element_by_xpath("//input[@id='txtUserPassword']").send_keys(limis_pwd)
    wb.find_element_by_xpath("//a[@id='btnLogin']").click()
    time.sleep(2)

def selecttype():
    category = {'1': '食品', '2': '药品', '3': '保健品', '4': '化妆品', '5': '药包材'}
    category_table = PrettyTable(['类别编号', '类别名称'])
    category_table.add_row(['1', '食品'])
    category_table.add_row(['2', '药品'])
    category_table.add_row(['3', '保健品'])
    category_table.add_row(['4', '化妆品'])
    category_table.add_row(['5', '药包材'])
    print(category_table)
    index_category = input('请输入检品类别编号(按ENTER确认):')
    if index_category == str(1):
        pass
    elif index_category == str(2):
        pass
    elif index_category == str(3):
        pass
    elif index_category == str(4):
        print('\n\n已选定检品类别为→化妆品\n\n')
        HZ_type_table = PrettyTable(['小类编号', '小类名称'])
        HZ_type_table.add_row(['1', '进口非特'])
        HZ_type_table.add_row(['2', '国产特殊'])
        HZ_type_table.add_row(['3', '抽验检验'])
        print(HZ_type_table)
    else:
        pass
    type_index = input('请选择检品检验小类(按ENTER确认):')
    type_dic = {'食品': [],
                '药品': [],
                '保健品': [],
                '化妆品': {'1': '5', '2': '6', '3': '2'},
                '药包材': [], }
    index_1 = {'1': '1', '2': '2', '3': '5', '4': '3', '5': '4'}[index_category]
    index_2 = type_dic[category[index_category]][type_index]
    wb.find_element_by_xpath("//ul[@id='leftMenuTree']/li[1]/ul/li[1]/ul/li[{}]/ul/li[{}]".format(index_1, index_2)).click()
    time.sleep(2)


def save_hzjkft():
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe"))
    wb.find_element_by_xpath("//div[@class='datagrid-toolbar']//span[contains(text(),'添加')]").click()#添加按钮
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_first']"))
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_sample']"))
    wb.find_element_by_xpath("//div[@class='fitem']/ul[1]/li[8]/span/input[2]").send_keys(SCGJDQ) #生产国家地区
    wb.find_element_by_xpath("//div[@class='fitem']/ul[2]/li[2]/span/input[2]").send_keys(JPMC) #检品名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[2]/li[5]/span/input[2]").send_keys(JPYWMC) #检品英文名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[3]/li[2]/span/input[2]").send_keys(SQQYMC) #申请企业名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[4]/li[2]/input").send_keys(SQQYDZ) #申请企业地址
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[2]/input").send_keys(LXR) #联系人
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[5]/input").send_keys(LXDH) #联系电话
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[8]/input").send_keys(CZ) #传真
    wb.find_element_by_xpath("//div[@class='fitem']/ul[6]/li[2]/span/input[2]").send_keys(ZHSBZRDWMC) #在华申报责任单位名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[6]/li[5]/input").send_keys(ZHSBDWLXR) #在华申报单位联系人
    wb.find_element_by_xpath("//div[@class='fitem']/ul[7]/li[2]/input").send_keys(ZHSBDWDH) #在华申报单位电话
    wb.find_element_by_xpath("//div[@class='fitem']/ul[7]/li[5]/input").send_keys(ZHSBDWCZ) #在华申报单位传真
    wb.find_element_by_xpath("//div[@class='fitem']/ul[8]/li[2]/input").send_keys(ZHSBDWDZ) #在华申报单位地址
    wb.find_element_by_xpath("//div[@class='fitem']/ul[8]/li[5]/input").send_keys(ZHSBDWYB) #在华申报单位邮编
    wb.find_element_by_xpath("//div[@class='fitem']/ul[9]/li[2]/span/input[2]").send_keys(JYYJ) #检验依据
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[5]/input").send_keys(JPSL) #检品数量
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[8]/input").send_keys(LYSL) #留样数量
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[11]/input").send_keys(JPSLDW) #检品数量单位
    wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[2]/input").send_keys(SLBH) #受理编号
    Select(wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[5]/select")).select_by_value(JYZQ) #检验周期
    wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[8]/span/input[2]").send_keys(YPLB) #样品类别
    wb.find_element_by_xpath("//div[@class='fitem']/ul[13]/li[2]/span/input[2]").send_keys(XHGG) #型号规格
    wb.find_element_by_xpath("//div[@class='fitem']/ul[13]/li[5]/span/input[2]").send_keys(YPXZ) #样品性状
    wb.find_element_by_xpath("//div[@class='fitem']/ul[14]/li[2]/input").send_keys(SCRQHPH) #生产日期或生产批号
    wb.find_element_by_xpath("//div[@class='fitem']/ul[14]/li[5]/input").send_keys(BZQ) #保质期
    wb.find_element_by_xpath("//div[@class='fitem']/ul[15]/li[2]/select").send_keys(ZJKS) #主检科室
    #保存检品信息
    wb.find_element_by_xpath("//div[@id='active_area']/div").click()
    time.sleep(3)
    wb.switch_to.default_content()

login()
selecttype()
save_hzjkft()
