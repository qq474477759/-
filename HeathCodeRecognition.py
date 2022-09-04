from xml.dom.minidom import Element
import pytesseract
from PIL import Image
# from openpyxl.drawing.image import Image
import os
import shutil
import re
from enum import Enum
import cv2 as cv
import numpy as np
import xlrd
import xlwt

class CardType(Enum):
    HealthCard = 1,
    TravalCard = 2,
    UnknownCard = 9

class HeathCodeRecognition:
    def __init__(self, path):
        self.path = path
        self.imglist = []
        self.resultlist = []

    def findFile(self):
        items = os.listdir(self.path)
        subfixs = ['.png','.jpg','.jpeg']
        imglist = []
        for item in items:
            filePath = os.path.join(self.path, item)
            if os.path.isfile(filePath):
                if os.path.splitext(filePath)[1] in subfixs:  # 后缀名判断
                    imglist.append(filePath)
                else:
                    continue

        # print(imglist)
        retlist = []
        for imgpath in imglist:
            ret = self.recognize(imgpath)
            #分离imgpath，提取文件名
            imgname = os.path.split(imgpath)[1]

            retlist.append((ret[0], ret[1],ret[2], imgname))

        print(retlist)
        return retlist
    #write to excel
    def write2excel(self, resultlist):
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('sheet1')
        #第一列写类型，第二列写结果，第三列写文件名

        sheet.write(0, 0, '类型')
        sheet.write(0, 1, '结果')
        sheet.write(0, 2, '7天内到访地')
        sheet.write(0, 3, '来源文件名')
        for i in range(len(resultlist)):
            sheet.write(i+1, 0, resultlist[i][0])
            sheet.write(i+1, 2, resultlist[i][2])
            sheet.write(i+1, 3, resultlist[i][3])
            sheet.write(i+1, 1, resultlist[i][1])
        workbook.save('result.xls')

    def recognize(self, imgpath):
        img = cv.imread(imgpath)
 
        #转为白底黑字
        height, width, deep = img.shape
        gray = cv.cvtColor(img, cv.COLOR_BGR2GRAY)
        dst = np.zeros((height, width, 1), np.uint8)
        for i in range(0, height):                          # 反相 转白底黑字
            for j in range(0, width):
                grayPixel = gray[i, j]
                dst[i, j] = 255 - grayPixel
        ret, canny = cv.threshold(dst, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU)   # 二值化
        imgname = os.path.split(imgpath)[1]


        print("=======开始识别=======", imgname)
        #and English

        s = pytesseract.image_to_string(canny, lang="chi_sim")
        s = s.replace(" ", "").replace("\n","").replace("\r","")

        # print(pytesseract.get_languages(config=""))
        # s = pytesseract.image_to_data(img, lang="chi_sim")

        # 识别健康码, 查找时间戳 如 20:27:13
        if  "核酸" in s:
            idx = s.index("核酸")

            if idx:
                print("识别到<健康码>")
                #寻找"健康状态"
                RNA_idx = s.index("核酸")
                RNA_time_inx=s.index("省内")
                stats_inx=s.index("健康状态")


                #取长度
                RNA_result = 2
                RNA_name = s[RNA_idx+3: RNA_idx+3+RNA_result]
                RNA_time = s[RNA_time_inx+2: RNA_time_inx+2+4]
                RNA_name=RNA_time+RNA_name

                #取长度
                stats_result = 8
                stats_name = s[stats_inx: stats_inx+stats_result]
                print('核酸结果：',RNA_name)
                print('健康状态：',stats_name)
                result=RNA_name+"+"+stats_name
                return ("健康码",result,'')

            # flag2 = s.index("代办")
        if "动态行程卡" in s:
        # 识别行程码
         idx = -1
         print (s)
         idx = s.find("的动态行程卡")
         if idx:
            print("识别到<行程卡>")
            pp = re.findall("1[0-9]{2}[*|x|X]*[0-9]{4}", s)
            #截取两个部分之间的字符串
            location = s[s.index("途经:"):s.index("结果包含")]
            
            if pp:
                pn = pp[0]
                #获得pp中的数字
                pn = re.findall("[0-9]", pn)
                #将数字转换为字符串
                pn = ''.join(pn)
                pn = pn[0:3] + "****" + pn[-4:]
                print('匹配手机号：',pp)
                print(location[3:])
                return ("行程卡", pn,location[3:])

         return ("未知", "未知")
#read python file path
path = os.path.dirname(os.path.realpath(__file__)) 
ocr = HeathCodeRecognition(path+"/images")
retlist=ocr.findFile()
ocr.write2excel(retlist)