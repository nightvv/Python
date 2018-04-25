from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
from openpyxl.styles import Font, colors, Alignment

limis_name = input('请输入Limis用户名(按ENTER确认):')
limis_pwd = input('请输入Limis密码(按ENTER确认):')


wb = webdriver.Chrome("C:\chromedriver.exe")
wb.maximize_window()
wb.get('http://10.0.144.50/web/Login.aspx')
wb.find_element_by_xpath("//input[@id='txtUserCode']").send_keys(limis_name)
wb.find_element_by_xpath("//input[@id='txtUserPassword']").send_keys(limis_pwd)
wb.find_element_by_xpath("//a[@id='btnLogin']").click()
time.sleep(2)

#index=1,2,3,4,5 进口，国内药品，化妆品，药包材，保健品
index = 3
sampleType1 = [1,2,3,4,5]
sampleType2 = ['进口','国内药品','化妆品','药包材','保健品']
wb.find_element_by_xpath("//ul[@id='leftMenuTree']/li[1]/ul/li[2]/ul/li[{}]".format(sampleType1[index-1])).click()
time.sleep(2)
wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@src='/ut/Ucc/DataList.aspx?view_id=522&template_id=381&jpfl={}&muid={}']".format(sampleType2[index-1],limis_name)))


def output():
    
    global sample_lst
    sample_lst = []
    for sampleNum in range(1,11):
        sample_info = {}
        wb.find_element_by_xpath("//tr[@id='datagrid-row-r2-2-{}']".format(sampleNum-1)).click()
        #修改
        wb.find_element_by_xpath("//span[@class='l-btn-text icon-edit']").click()
        time.sleep(1)
        print('第{}个样品'.format(sampleNum))

        wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_first']"))
        wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_sample']"))
        try:
            #检品名称
            sample_info['ChName'] = wb.find_element_by_xpath("//*[@id='p-first']/ul[2]/li[2]/span/span[1]").get_attribute("innerHTML").replace('amp;','')
            print('检品名称：',wb.find_element_by_xpath("//*[@id='p-first']/ul[2]/li[2]/span/span[1]").get_attribute("innerHTML").replace('amp;',''))
            sample_info['EnName'] = wb.find_element_by_xpath("//*[@id='p-first']/ul[2]/li[5]/span/span[1]").get_attribute("innerHTML").replace('amp;','')
            #print('英文名称：',wb.find_element_by_xpath("//*[@id='p-first']/ul[2]/li[5]/span/span[1]").get_attribute("innerHTML").replace('amp;',''))
            #生产国家
            sample_info['Country'] = wb.find_element_by_xpath('//*[@id="p-first"]/ul[1]/li[8]/span/span[1]').get_attribute("innerHTML")
            #print('生产国家：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[1]/li[8]/span/span[1]').get_attribute("innerHTML"))
            #申请企业 申请企业地址
            sample_info['Company_name'] = wb.find_element_by_xpath("//*[@id='p-first']/ul[3]/li[2]/span/span[1]").get_attribute("innerHTML")
            #print('申请企业：', wb.find_element_by_xpath("//*[@id='p-first']/ul[3]/li[2]/span/span[1]").get_attribute("innerHTML"))
            sample_info['Company_add'] = wb.find_element_by_xpath("//*[@id='field_1359']").get_attribute("value")
            #print('申请企业地址：', wb.find_element_by_xpath("//*[@id='field_1359']").get_attribute("value"))
            #联系人 联系电话 传真
            sample_info['Company_contact'] = wb.find_element_by_xpath('//*[@id="field_1347"]').get_attribute("value")
            #print('申请企业联系人：', wb.find_element_by_xpath('//*[@id="field_1347"]').get_attribute("value"))
            sample_info[' Company_tele'] = wb.find_element_by_xpath('//*[@id="field_1348"]').get_attribute("value")
            #print('申请企业联系电话：', wb.find_element_by_xpath('//*[@id="field_1348"]').get_attribute("value"))
            sample_info['Company_fax'] = wb.find_element_by_xpath('//*[@id="field_1362"]').get_attribute("value")
            #print('申请企业传真：', wb.find_element_by_xpath('//*[@id="field_1362"]').get_attribute("value"))
            #在华单位 在华联系人
            sample_info['Res_unit'] = wb.find_element_by_xpath('//*[@id="p-first"]/ul[6]/li[2]/span/span[1]').get_attribute("innerHTML")
            #print('在华申报单位：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[6]/li[2]/span/span[1]').get_attribute("innerHTML"))
            #print('在华联系人：', wb.find_element_by_xpath('//*[@id="field_1502"]').get_attribute("value"))
            #在华电话 在华传真
            #print('在华电话：', wb.find_element_by_xpath('//*[@id="field_1501"]').get_attribute("value"))
            #print('在华传真：', wb.find_element_by_xpath('//*[@id="field_1499"]').get_attribute("value"))
            #在华地址 在华邮编
            #print('在华地址：', wb.find_element_by_xpath('//*[@id="field_1500"]').get_attribute("value"))
            #print('在华邮编：', wb.find_element_by_xpath('//*[@id="field_1504"]').get_attribute("value"))
            #检验依据
            #print('检验依据：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[9]/li[2]/span/span[1]').get_attribute("innerHTML"))
            #检验目的大项 检验目的小项
            #print('检验目的大项：', wb.find_element_by_xpath('//*[@id="field_1330"]').get_attribute("value"))
            #print('检验目的小项：', wb.find_element_by_xpath('//*[@id="field_1331"]').get_attribute("value"))
            #批数 检品数量 留样数量 检品数量单位
            #print('批数：', wb.find_element_by_xpath('//*[@id="field_1333"]').get_attribute("value"))
            #print('检品数量：', wb.find_element_by_xpath('//*[@id="field_1354"]').get_attribute("value"))
            #print('留样数量：', wb.find_element_by_xpath('//*[@id="field_1350"]').get_attribute("value"))
            #print('检品单位：', wb.find_element_by_xpath('//*[@id="field_3100"]').get_attribute("value"))
            #受理编号 检验周期 样品类别
            sample_info['Accept_num'] = wb.find_element_by_xpath('//*[@id="field_1351"]').get_attribute("value")
            print('受理编号：', wb.find_element_by_xpath('//*[@id="field_1351"]').get_attribute("value"))
            #print(wb.find_element_by_xpath('').get_attribute("value"))
            #print('样品类别：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[12]/li[8]/span/span[1]').get_attribute("innerHTML"))
            #型号规格 样品性状
            #print('型号规格：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[13]/li[2]/span/span[1]').get_attribute("innerHTML"))
            #print('样品性状：', wb.find_element_by_xpath('//*[@id="p-first"]/ul[13]/li[5]/span/span[1]').get_attribute("innerHTML"))
            #生产日期或批号 限期使用日期
            sample_info['data_or_No'] = wb.find_element_by_xpath('//*[@id="field_1352"]').get_attribute("value")
            #print('生产日期或批号：', wb.find_element_by_xpath('//*[@id="field_1352"]').get_attribute("value"))
            #print('限期使用日期：', wb.find_element_by_xpath('//*[@id="field_1356"]').get_attribute("value"))
            #预收费
            print('预收费：', wb.find_element_by_xpath('//*[@id="field_1505"]').get_attribute("value"))
            #备注
            #print('备注：', wb.find_element_by_xpath('//*[@id="field_1334"]').get_attribute("value"))
            #报告说明
            #print('报告说明：', wb.find_element_by_xpath('//*[@id="field_3277"]').get_attribute("value"))
            sample_lst.append(sample_info)

            print('--------------------------------------------------------------------------')

            #下一步
            ##wb.switch_to.default_content()
            ##wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@src='/ut/Ucc/DataList.aspx?view_id=522&template_id=381&jpfl={}&muid={}']".format(sampleType2[index-1],limis_name)))
            ##wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@id='ifm_first']"))
            ##wb.find_element_by_xpath("//span[@class='l-btn-left']/span").click()
            ##time.sleep(1.5)
            ##wb.find_element_by_xpath("//div[@id='p-second']/a").click()
            ##time.sleep(1.5)
            ##wb.find_element_by_xpath("//div[@id='p-third']/a").click()
            ##time.sleep(1.5)
            ##wb.find_element_by_xpath("//div[@class='tool-bottom']/a").click()
            ##time.sleep(1.5)
            ###保存并确认
            ##wb.find_element_by_xpath("//div[@class='messager-button']").click()

            #关闭
            wb.switch_to.default_content()
            wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@src='/ut/Ucc/DataList.aspx?view_id=522&template_id=381&jpfl={}&muid={}']".format(sampleType2[index-1],limis_name)))
            wb.find_element_by_xpath('/html/body/div[5]/div[1]/div[2]/a').click()
        except:
            print('样品不是进口非特')
            print('--------------------------------------------------------------------------')
            wb.switch_to.default_content()
            wb.switch_to.frame(wb.find_element_by_xpath("//iframe[@src='/ut/Ucc/DataList.aspx?view_id=522&template_id=381&jpfl={}&muid={}']".format(sampleType2[index-1],limis_name)))
            wb.find_element_by_xpath('/html/body/div[5]/div[1]/div[2]/a').click()
            continue
            #关闭

def fengqian():
    wb = openpyxl.Workbook()
    wb.save("C:\化妆品limis封签.xlsx")
    wb = openpyxl.load_workbook('C:\化妆品limis封签.xlsx')
    sheet = wb.active

    bold_font = Font(bold=True,size=12)
    norm_font = Font(size=12)

    sheet.column_dimensions['A'].width = 3.38
    sheet.column_dimensions['B'].width = 69

    j = 0
    for i in range(0,len(sample_lst)):
        sheet.row_dimensions[j+1].height = 15.75
        sheet.row_dimensions[j+2].height = 15.75
        sheet.row_dimensions[j+3].height = 15.75
        sheet.row_dimensions[j+4].height = 15.75
        sheet.row_dimensions[j+5].height = 15.75
        sheet.row_dimensions[j+6].height = 15.75
        sheet['A{}'.format(j+1)] = '化'
        sheet['A{}'.format(j+1)].font = bold_font
        sheet['A{}'.format(j+2)] = '妆'
        sheet['A{}'.format(j+2)].font = bold_font
        sheet['A{}'.format(j+3)] = '品'
        sheet['A{}'.format(j+3)].font = bold_font
        sheet['A{}'.format(j+4)] = '封'
        sheet['A{}'.format(j+4)].font = bold_font
        sheet['A{}'.format(j+5)] = '签'
        sheet['A{}'.format(j+5)].font = bold_font
        sheet['B{}'.format(j+1)] = '检品名称：{}(批号{})'.format(sample_lst[i]['ChName'],sample_lst[i]['data_or_No'])
        sheet['B{}'.format(j+1)].font = norm_font
        sheet['B{}'.format(j+2)] = '外文名称：{}'.format(sample_lst[i]['EnName'])
        sheet['B{}'.format(j+2)].font = norm_font
        sheet['B{}'.format(j+3)] = '受理编号：{}'.format(sample_lst[i]['Accept_num'])
        sheet['B{}'.format(j+3)].font = norm_font
        sheet['B{}'.format(j+4)] = '申请企业：{}'.format(sample_lst[i]['Company_name'])
        sheet['B{}'.format(j+4)].font = norm_font
        sheet['B{}'.format(j+5)] = '在华申报责任单位：{}'.format(sample_lst[i]['Res_unit'])
        sheet['B{}'.format(j+5)].font = norm_font
        sheet['B{}'.format(j+6)] = '封样单位经手人/抽样封签日期：'
        sheet['B{}'.format(j+6)].font = norm_font
        j+=9
    wb.save("C:\化妆品limis封签.xlsx")

def main():
    output()
    fengqian()
	
judge = True
yunxing = input('请输入1或2并按ENTER确认（1开始执行，2退出）')
while judge:
    if yunxing == '1':
        main()
        yunxing = input('请输入1或2并按ENTER确认（1开始执行，2退出）')
    elif yunxing == '2':
        judge = False
		



