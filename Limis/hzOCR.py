#!/usr/bin/env python  
# encoding: utf-8 
"""
Author: ISeeMoon
Python: 3.6
Software: PyCharm
File: OCR.py
Time: 2018/4/25 20:10
"""
from aip import AipOcr
import json
import time
import re
from PIL import ImageEnhance
from PIL import  ImageFilter
from PIL import Image
from openpyxl import load_workbook

class hzOCR(object):

    info_dic = {}
    def baiduKey(self):
        exce = load_workbook(r"C:\baiduOCR.xlsx")
        sheet = exce.worksheets[0]
        APP_ID = str(sheet['B1'].value)
        API_KEY = sheet['B2'].value
        SECRET_KEY = sheet['B3'].value
        return {'APP_ID':APP_ID,
        'API_KEY':API_KEY,
        'SECRET_KEY':SECRET_KEY}

    def get_file_content(self,filepath):
        with open(filepath, 'rb') as fp:
            return fp.read()

    def baiduOCR(self):
        imagepath = r"C:\hzocr\jkhz.jpg"
        image = Image.open(imagepath).convert('L')
        size = image.size
        cropimage = image.crop((size[0] * 0.05, size[-1] * 0.05, size[0] * 0.95, size[-1] * 0.7))
        sharpimage = cropimage.filter(ImageFilter.SHARPEN)
        image_contrasted = ImageEnhance.Contrast(sharpimage).enhance(8)
        image_contrasted.save(imagepath)
        im = self.get_file_content(imagepath)
        aipOcr = AipOcr(self.baiduKey()['APP_ID'], self.baiduKey()['API_KEY'], self.baiduKey()['SECRET_KEY'])
        recog_table = aipOcr.tableRecognitionAsync(im)
        id = recog_table['result'][0]['request_id']
        options_table = {'result_type': 'json'}
        result_table = aipOcr.getTableRecognitionResult(id, options_table)
        while 'error_code' in result_table or 'result' in result_table and result_table['result']['ret_code'] != 3:
            time.sleep(0.3)
            print('表格正在识别中...')
            aipOcr.getTableRecognitionResult(id, options_table)
            result_table = aipOcr.getTableRecognitionResult(id, options_table)
        print('表格识别完毕!')
        dic_table = json.loads(result_table['result']['result_data'])['forms'][0]['body']

        for item in dic_table:
            x = item['row'][0]
            y = item['column'][0]
            word = item['word']
            self.info_dic[(x, y)] = word

    def rowstr(self, x):
        i = 0
        judge = True
        str = ''
        while judge:
            if (x, i) in self.info_dic.keys():
                str += self.info_dic[(x, i)]
                i += 1
            else:
                judge = False
        return str

    def JPMC(self):
        return self.rowstr(0).split('中文')[-1]
    def JPYWMC(self):
        return self.rowstr(1).split('外文')[-1]
    def XHGG(self):
        return self.rowstr(2).split('送检')[0].split('格')[-1]
    def JPSL(self):
        return self.rowstr(2).split('数量')[-1]
    def YPXZ(self):
        return ''.join(self.rowstr(3).split('样品')[0:-1]).split(')')[-1].split('态')[-1]
    def SCRQHSCPH(self):
        return self.rowstr(4).split('保质')[0].split('限')[0].split('生')[-1].split('或')[-1].split('期')[-1].split('号')[-1]
    def BZQ(self):
        return self.rowstr(4).split('其')[-1].split('期')[-1].replace('习', '')
    def SQQYMC(self):
        return self.rowstr(8).split('称')[-1].split('生产')[0].replace('|', '')
    def SCGJDQ(self):
        if len(self.rowstr(8).split(')')[-1]) != 0:
            return ''.join(re.findall("[\u4e00-\u9fa5a]", self.rowstr(8).split(')')[-1]))
        else:
            return ''.join(re.findall("[\u4e00-\u9fa5a]", self.rowstr(8).split('国')[-1])).replace('()', '').replace('地区', '')
    def SQQYDZ(self):
        return self.rowstr(9).split('址')[-1].split('联')[0].replace('名称', '')
    def LXDH(self):
        return ''.join(re.findall("\d|-|\.", self.rowstr(9).split('联系人')[0].split('联')[-1]))
    def LXR(self):
        return ''.join(re.findall("[\u4e00-\u9fa5a-zA-Z0-9]", self.rowstr(9).split('人')[-1]))
    def ZHSBZRDWMC(self):
        return self.rowstr(10).split('称')[-1].split('联')[0]
    def ZHSBDWDH(self):
        return ''.join(re.findall("\d|-", self.rowstr(10)))
    def ZHSBDWLXR(self):
        return self.rowstr(10).split('人')[-1].split('入')[-1].replace('|', '')
    def ZHSBDWDZ(self):
        return self.rowstr(11).split('址')[-1].split('传')[0].replace('|', '')
    def ZHSBDWCZ(self):
        return self.rowstr(11).split('真')[-1].split('邮')[0].replace('|', '').replace('[', '').replace('\\', '').replace('「', '')
    def ZHSBDWYB(self):
        return self.rowstr(11).split('邮编')[-1].replace('|', '')
    def CZ(self):
        return '/'
    def JYYJ(self):
        return '化妆品安全技术规范（2015年版）'
    def LYSL(self):
        return ''
    def JPSLDW(self):
        return self.rowstr(2).split('数量')[-1][-1]
    def SLBH(self):
        return 'JF0182018'
    def JYZQ(self):
        return '60'
    def YPLB(self):
        return ''
    def ZJKS(self):
        return '保化室'

