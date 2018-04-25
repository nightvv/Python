#!/usr/bin/env python
# encoding: utf-8 
"""
Author: ISeeMoon
Python: 3.6
Software: PyCharm
File: 1234.py
Time: 2018/4/22 22:08
"""

import time
from selenium.webdriver.support.select import Select
from Limis.hzgcLOAD import hzgc_API
from Limis.hzOCR import hzOCR
from selenium import webdriver
from prettytable import PrettyTable

print('本软件目前仅适用于【化妆品国抽】导入。\n【委托方所属地区】、【样品包装方式】和【抽样人员】需要自行核对！！！\n【委托方所属地区】、【样品包装方式】和【抽样人员】需要自行核对！！！\n【委托方所属地区】、【样品包装方式】和【抽样人员】需要自行核对！！！\n -------2018-04-23 Version 1.0-------')
limis_name = input('请输入Limis用户名(按ENTER确认):')
limis_pwd = input('请输入Limis密码(按ENTER确认):')

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
    global index_category,type_index
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

def save_hzgc(i):
    hzgc = hzgc_API()
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe"))
    wb.find_element_by_xpath("//div[@class='datagrid-toolbar']//span[contains(text(),'添加')]").click()#添加按钮
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_first']"))
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_sample']"))
    #检品名称
    wb.find_element_by_xpath("//div[@id='p-first']/ul[2]/li[2]/input").send_keys(hzgc.JPMC(i))
    #品牌PP
    wb.find_element_by_xpath("//div[@id='p-first']/ul[2]/li[5]/input").send_keys(hzgc.PP(i))
    #所属大类SSDL
    if hzgc.SSDL(i) in '染发类产品':
        Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[3]/li[2]/select")).select_by_value('染发类产品')
    elif hzgc.SSDL(i) in '防晒类产品':
        Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[3]/li[2]/select")).select_by_value('防晒类产品')
    elif hzgc.SSDL(i) in '祛痘产品':
        Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[3]/li[2]/select")).select_by_value('祛痘产品')    
    #所属小类SSXL
    wb.find_element_by_xpath("//div[@id='p-first']/ul[3]/li[5]/input").send_keys(hzgc.SSXL(i))
    #生产单位SCDW
    wb.find_element_by_xpath("//div[@id='p-first']/ul[4]/li[2]/span/input[2]").send_keys(hzgc.SCDW(i))
    #产地CD
    wb.find_element_by_xpath("//div[@id='p-first']/ul[4]/li[5]/input").send_keys(hzgc.CD(i))
    #是否进口SFJK
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[4]/li[8]/select")).select_by_value(hzgc.SFJK(i))
    #生产单位地址SCDWDZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[5]/li[2]/input").send_keys(hzgc.SCDWDZ(i))
    #经销商名称
    wb.find_element_by_xpath("//div[@id='p-first']/ul[6]/li[2]/input").send_keys(hzgc.JXSMC(i))
    #经销商省份JXSSF
    wb.find_element_by_xpath("//div[@id='p-first']/ul[6]/li[5]/input").send_keys(hzgc.JXSSF(i))
    #经销商地址JXSDZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[7]/li[2]/input").send_keys(hzgc.JXSDZ(i))
    #委托方WTF
    wb.find_element_by_xpath("//div[@id='p-first']/ul[8]/li[2]/span/input[2]").send_keys(hzgc.WTF(i))
    #所属地区SSDQ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[8]/li[5]/input").send_keys(hzgc.SSDQ(i))
    #所在地SZD
    wb.find_element_by_xpath("//div[@id='p-first']/ul[8]/li[8]/input").send_keys(hzgc.SZD(i))
    #委托方地址WTFDZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[9]/li[2]/input").send_keys(hzgc.WTFDZ(i))
    #供养单位GYDW
    wb.find_element_by_xpath("//div[@id='p-first']/ul[10]/li[2]/span/input[2]").send_keys(hzgc.GYDW(i))
    #供养单位性质GYDWXZ
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[10]/li[5]/select")).select_by_value(hzgc.GYDWXZ(i))
    #供养单位归属地GYDWGSD
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[10]/li[8]/select")).select_by_value(hzgc.GYDWGSD(i))
    #供养单位地址GYDWDZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[11]/li[2]/input").send_keys(hzgc.GYDWDZ(i))
    #供养单位联系人GYDWLXR
    wb.find_element_by_xpath("//div[@id='p-first']/ul[12]/li[2]/input").send_keys(hzgc.GYDWLXR(i))
    #供养单位联系电话GYDWLXDH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[12]/li[5]/input").send_keys(hzgc.GYDWLXDH(i))
    #供养单位邮箱GYDWYX
    wb.find_element_by_xpath("//div[@id='p-first']/ul[12]/li[8]/input").send_keys(hzgc.GYDWYX(i))
    #被抽样单位BCYDW
    wb.find_element_by_xpath("//div[@id='p-first']/ul[13]/li[2]/span/input[2]").send_keys(hzgc.BCYDW(i))
    #被抽样单位性质BCYDWXZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[13]/li[5]/input").send_keys(hzgc.BCYDWXZ(i))
    #被抽样单位所在地
    wb.find_element_by_xpath("//div[@id='p-first']/ul[14]/li[2]/input").send_keys(hzgc.BCYDWSZD(i))
    #地域类型
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[14]/li[5]/select")).select_by_value(hzgc.DYLX(i))
    #被抽样单位省份BCYDWSF
    wb.find_element_by_xpath("//div[@id='p-first']/ul[15]/li[2]/input").send_keys(hzgc.BCYDWSF(i))
    #被抽样单位归属地BCYDWGSD
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[15]/li[5]/select")).select_by_value(hzgc.BCYDWGSD(i)[:-1])
    #被抽样单位县名BCYDWXM
    wb.find_element_by_xpath("//div[@id='p-first']/ul[15]/li[8]/input").send_keys(hzgc.BCYDWXM(i))
    #被抽样单位地址
    wb.find_element_by_xpath("//div[@id='p-first']/ul[16]/li[2]/input").send_keys(hzgc.BCYDWDZ(i))
    #被抽样单位负责人
    wb.find_element_by_xpath("//div[@id='p-first']/ul[17]/li[2]/input").send_keys(hzgc.BCYDWFZR(i))
    #被抽样单位联系电话BCYDWLXDH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[17]/li[5]/input").send_keys(hzgc.BCYDWLXDH(i))
    #抽样基数
    wb.find_element_by_xpath("//div[@id='p-first']/ul[18]/li[2]/input").send_keys(hzgc.CYJS(i))
    #单价DJ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[18]/li[5]/input").send_keys(hzgc.DJ(i))
    #执行标准ZXBZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[18]/li[8]/input").send_keys(hzgc.ZXBZ(i))
    #规格GG
    wb.find_element_by_xpath("//div[@id='p-first']/ul[19]/li[2]/span/input[2]").send_keys(hzgc.GG(i))
    #包装BZ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[19]/li[5]/span/input[2]").send_keys(hzgc.BZ(i))
    #包装规格BZGG
    wb.find_element_by_xpath("//div[@id='p-first']/ul[20]/li[2]/span/input[2]").send_keys(hzgc.BZGG(i))
    #剂型JX
    wb.find_element_by_xpath("//div[@id='p-first']/ul[20]/li[5]/span/input[2]").send_keys(hzgc.JX(i))
    #检验依据JYYJ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[21]/li[2]/span/input[2]").send_keys(hzgc.JYYJ(i))
    #检验目的大项    
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[22]/li[2]/select")).select_by_value(hzgc.JYMDDX(i))
    #检验目的小项
    wb.find_element_by_xpath("//div[@id='p-first']/ul[22]/li[5]/input").send_keys(hzgc.JYMDXX(i))
    #批数
    PS = '1'
    wb.find_element_by_xpath("//div[@id='p-first']/ul[23]/li[2]/input").send_keys(PS)
    #检品数量JPSL
    wb.find_element_by_xpath("//div[@id='p-first']/ul[23]/li[5]/input").send_keys(hzgc.JPSL(i))
    #检品数量单位JPSLDW
    wb.find_element_by_xpath("//div[@id='p-first']/ul[23]/li[8]/input").send_keys(hzgc.JPSLDW(i))
    #留样数量LYSL
    wb.find_element_by_xpath("//div[@id='p-first']/ul[23]/li[11]/input").send_keys(hzgc.LYSL(i))
    #批准文号PZWH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[24]/li[2]/input").send_keys(hzgc.PZWH(i))
    #任务来源RWLY
    wb.find_element_by_xpath("//div[@id='p-first']/ul[24]/li[5]/input").send_keys(hzgc.RWLY(i))
    #报送分类BSFL
    wb.find_element_by_xpath("//div[@id='p-first']/ul[24]/li[8]/input").send_keys(hzgc.BSFL(i))
    #批号PH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[25]/li[2]/input").send_keys(hzgc.PH(i))
    #生产日期SCRQ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[25]/li[5]/input").send_keys(hzgc.SCRQ(i))
    #保质期BZQ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[25]/li[8]/input").send_keys(hzgc.BZQ(i))
    #颜色和物态YSHWT
    wb.find_element_by_xpath("//div[@id='p-first']/ul[26]/li[2]/input").send_keys(hzgc.YSHWT(i))
    #抽样单编号CYDBH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[26]/li[5]/input").send_keys(hzgc.CYDBH(i))
    #许可证号XKZH
    wb.find_element_by_xpath("//div[@id='p-first']/ul[27]/li[2]/input").send_keys(hzgc.XKZH(i))
    #主检科室ZJKS
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[28]/li[2]/select")).select_by_value(hzgc.ZJKS(i))
    #检验项目JYXM
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[28]/li[5]/select")).select_by_value(hzgc.JYXM(i))
    #检验周期JYZQ
    Select(wb.find_element_by_xpath("//div[@id='p-first']/ul[28]/li[8]/select")).select_by_value(hzgc.JYZQ(i))
    #抽样人员CYRY
    wb.find_element_by_xpath("//div[@id='p-first']/ul[31]/li[2]/input").send_keys(hzgc.CYRY(i))
    #抽样时间CYSJ
    wb.find_element_by_xpath("//div[@id='p-first']/ul[31]/li[5]/input").send_keys(hzgc.CYSJ(i))
    #保存检品信息
    wb.find_element_by_xpath("//div[@id='active_area']/div").click()
    time.sleep(3)
    wb.switch_to.default_content()


def save_hzjkft():
    jkft = hzOCR()
    jkft.baiduOCR()
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe"))
    wb.find_element_by_xpath("//div[@class='datagrid-toolbar']//span[contains(text(),'添加')]").click()#添加按钮
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_first']"))
    wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_sample']"))
    wb.find_element_by_xpath("//div[@class='fitem']/ul[1]/li[8]/span/input[2]").send_keys(jkft.SCGJDQ()) #生产国家地区
    wb.find_element_by_xpath("//div[@class='fitem']/ul[2]/li[2]/span/input[2]").send_keys(jkft.JPMC()) #检品名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[2]/li[5]/span/input[2]").send_keys(jkft.JPYWMC()) #检品英文名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[3]/li[2]/span/input[2]").send_keys(jkft.SQQYMC()) #申请企业名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[4]/li[2]/input").send_keys(jkft.SQQYDZ()) #申请企业地址
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[2]/input").send_keys(jkft.LXR()) #联系人
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[5]/input").send_keys(jkft.LXDH()) #联系电话
    wb.find_element_by_xpath("//div[@class='fitem']/ul[5]/li[8]/input").send_keys(jkft.CZ()) #传真
    wb.find_element_by_xpath("//div[@class='fitem']/ul[6]/li[2]/span/input[2]").send_keys(jkft.ZHSBZRDWMC()) #在华申报责任单位名称
    wb.find_element_by_xpath("//div[@class='fitem']/ul[6]/li[5]/input").send_keys(jkft.ZHSBDWLXR()) #在华申报单位联系人
    wb.find_element_by_xpath("//div[@class='fitem']/ul[7]/li[2]/input").send_keys(jkft.ZHSBDWDH()) #在华申报单位电话
    wb.find_element_by_xpath("//div[@class='fitem']/ul[7]/li[5]/input").send_keys(jkft.ZHSBDWCZ()) #在华申报单位传真
    wb.find_element_by_xpath("//div[@class='fitem']/ul[8]/li[2]/input").send_keys(jkft.ZHSBDWDZ()) #在华申报单位地址
    wb.find_element_by_xpath("//div[@class='fitem']/ul[8]/li[5]/input").send_keys(jkft.ZHSBDWYB()) #在华申报单位邮编
    wb.find_element_by_xpath("//div[@class='fitem']/ul[9]/li[2]/span/input[2]").send_keys(jkft.JYYJ()) #检验依据
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[5]/input").send_keys(jkft.JPSL()) #检品数量
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[8]/input").send_keys(jkft.LYSL()) #留样数量
    wb.find_element_by_xpath("//div[@class='fitem']/ul[11]/li[11]/input").send_keys(jkft.JPSLDW()) #检品数量单位
    wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[2]/input").send_keys(jkft.SLBH()) #受理编号
    Select(wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[5]/select")).select_by_value(jkft.JYZQ()) #检验周期
    wb.find_element_by_xpath("//div[@class='fitem']/ul[12]/li[8]/span/input[2]").send_keys(jkft.YPLB()) #样品类别
    wb.find_element_by_xpath("//div[@class='fitem']/ul[13]/li[2]/span/input[2]").send_keys(jkft.XHGG()) #型号规格
    wb.find_element_by_xpath("//div[@class='fitem']/ul[13]/li[5]/span/input[2]").send_keys(jkft.YPXZ()) #样品性状
    wb.find_element_by_xpath("//div[@class='fitem']/ul[14]/li[2]/input").send_keys(jkft.SCRQHSCPH()) #生产日期或生产批号
    wb.find_element_by_xpath("//div[@class='fitem']/ul[14]/li[5]/input").send_keys(jkft.BZQ()) #保质期
    wb.find_element_by_xpath("//div[@class='fitem']/ul[15]/li[2]/select").send_keys(jkft.ZJKS()) #主检科室
    #保存检品信息
    wb.find_element_by_xpath("//div[@id='active_area']/div").click()
    time.sleep(3)
    wb.switch_to.default_content()

def main():
    login()
    selecttype()
    if index_category == 4 and type_index == 3:
        hzgc_judge = input('''请在放入国抽导入表格并按【→1 导入】;【按→2 终止】''')
        for i in range(1,hzgc.sample_num):
            save_hzgc(i)
            print('已成功导入第{}个化妆品，共{}个化妆品'.format(i,hzgc.sample_num-1))
        print('已完成当前数据表格导入！')
        hzgc_judge = input('''请在放入国抽导入表格并按【→1 导入】;【按→2 终止】''')
    elif index_category == 4 and type_index == 1:
        jkhz_judge = input('请在放入申请表扫描件并按→1 继续导入\n按→2终止导入')
        while jkhz_judge:
            save_hzjkft()
            print('已完成当前化妆品导入！')
            jkhz_judge = input('''请在放入申请表扫描件并按【→1 导入】;【按→2 终止】''')

