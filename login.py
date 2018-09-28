'''
作者：Leif
time：2018/9/27
version:1.0
说明：用于搜集无课表
有问题联系qq：919820268
'''
import random
import time
from urllib.parse import parse_qsl, urlparse
import re
import cv2
import requests
import xlwt
import xlrd 
import xlrd as xltest
from xlutils.copy import copy
from PIL import Image
from pytesseract import image_to_string
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
import csv
from urllib.request import urlopen



class loginin:

    def __init__(self,username,passwd):
        self.driver = webdriver.Chrome()
        self.username=username
        self.passwd=passwd
    
    def login(self):
        url = 'http://portal.tfswufe.edu.cn/web/guest/243/'
        self.driver.delete_all_cookies() #打开浏览器时清除所有cookie
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)

        self.driver.find_element_by_id('username').send_keys(self.username)
        self.driver.find_element_by_id('password').send_keys(self.passwd)

        # 保存截屏
        self.driver.save_screenshot('01.png')
        ran = Image.open("01.png")
        box = (986, 492, 1173, 558)
        ran.crop(box).save("captcha.png")

        # 图像识别
        image = Image.open("captcha.png")
        text = image_to_string(image)
        #print("验证码：" + text)
        authcode = text#input("请输入验证码：")#图像识别无法使用时候
        self.driver.find_element_by_id('authcode').send_keys(authcode)
        #time.sleep(2)

        self.driver.find_element_by_xpath("//input[@value='登录']").click()
        #print(self.driver.current_url)
        #print(url)
        ans = re.match(url,self.driver.current_url)
        #print(ans)
        if ans != None:
            print('登录成功')
        else:
            print('登录失败，请重试')
            # self.username=[]
            # self.password=[]
            self.login()
    def nameInTable(self,idnu):
        #获取无课数据放入excle表格
        stuName = self.driver.find_element_by_class_name('pClass').text#获取学生姓名
        wbk = xlrd.open_workbook('D:\\test.xls')
        newwb = copy(wbk)
        newws = newwb.add_sheet('stu'+str(idnu))
        # sheet.write(0,1,'test text')#第0行第一列写入内容
        # wbk.save('test.xls')
        zh_pattern = re.compile(u'[\u4e00-\u9fa5]+')#中文字符集合
        courseTime = ['_1_2','_3_5','_6_7','_8_10','_11_13','_14_15'] #课程的节数
        xnum = 1
        for y in courseTime: 
            for iday in range(1,8):#星期一到星期天
                txt = self.driver.find_element_by_id(str(iday)+y).text#获取课程内容
                #oldv = self.getOldVaule(xnum,iday)
                if zh_pattern.search(txt):#如果有中文则不操作，否则填入名字
                    newws.write(xnum,iday,'')
                else:
                    #sheet.write(x,i,'6')
                    #print(666)
                    newws.write(xnum,iday,stuName+',')
            xnum = xnum+1
        newwb.save('D:\\test.xls')
        # wbk1 = xlrd.open_workbook('D:\\test.xls')
        # newwb1 = copy(wbk)
        # newwb1.save('D:\\test1.xls')
        self.driver.close()#关闭浏览器
    # def getOldVaule(self,xc,yc):
    #     oldwb = xltest.open_workbook('D:\\test.xls')
    #     oldsh = oldwb.get_sheet(0)
    #     oldv = oldsh.cell(xc,yc).value
    #     return oldv
url = 'D:\\name.xlsx'#这个是你的记录学号密码的表


uname = []
pwd = []
namebk = xlrd.open_workbook(url)
sheet1 = namebk.sheets()[0]
nrows = sheet1.nrows   #行
ncols = sheet1.ncols   #列
#print(nrows,ncols)
# 判断是否为数字
def isNum(value):
    try:
        value + 1
    except TypeError:
        return False
    else:
        return True
#将学号密码导入数组
for row in range(1,nrows):
    cell_value0=sheet1.cell(row,1).value
    cell_value1=sheet1.cell(row,0).value
    cell_value1 = int(cell_value1)
    uname.append(cell_value1)
    if isNum(cell_value0):
        cell_value0= int(cell_value0)
        pwd.append(str(cell_value0))
    else:
        pwd.append(cell_value0)
#print(uname)
idu = 0
bk = xlwt.Workbook()#打开工作簿
sheet = bk.add_sheet('sheet totall',cell_overwrite_ok=True)#创建工作表
bk.save('D:\\test.xls')
#获取每个人的无课表情况
for un in uname:

    test = loginin(un,pwd[idu])
    test.login()
    test.nameInTable(idu)

    idu = idu+1