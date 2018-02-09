'''
python 3.4
1. 下载安装pytesseract
2.配置环境变量 C:\Program Files (x86)\Tesseract-OCR 至 Path中
添加系统环境变量 TESSDATA_PREFIX   C:\Program Files (x86)\Tesseract-OCR\tessdata

3.将 test.traineddata 文件拷贝到 C:\Program Files (x86)\Tesseract-OCR\tessdata 中

4.pip install -r requirements.pip

5.安装雷神模拟器

6. python main.py

7.使用微信进入挑战智力按照提示进行标点

8.不要拖动雷神模拟器等待刷分

'''

import win32api
import win32con
from ctypes import *
import time
import random
import pytesseract

from PIL import ImageGrab



class POINT(Structure):
    _fields_ = [("x", c_ulong), ("y", c_ulong)]

def get_mouse_point():
    """
    获取鼠标位置
    :return: dict，鼠标横纵坐标
    """
    po = POINT()
    windll.user32.GetCursorPos(byref(po))
    return int(po.x), int(po.y)

def mouse_move(x, y):
    """
    移动鼠标位置
    :param x: int, 目的横坐标
    :param y: int, 目的纵坐标
    :return: None
    """
    windll.user32.SetCursorPos(x, y)

def mouse_pclick(x=None, y=None, press_time=0.0):
    """
    模拟式长按鼠标
    :param x: int, 鼠标点击位置横坐标
    :param y: int, 鼠标点击位置纵坐标
    :param press_time: float, 点击时间，单位秒
    :return: None
    """
    if not x is None and not y is None:
        mouse_move(x, y)
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(press_time)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def click_xy(coordinate=None, press_time=0.0):
    """
    模拟式长按鼠标
    :param coordinate: tuple, 鼠标点击位置横坐标
    :param press_time: float, 点击时间，单位秒
    :return: None
    """
    if not coordinate is None:
        mouse_move(int(coordinate[0]), int(coordinate[1]))
        time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(press_time)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def get_static_mouse_point():
    """
    获取鼠标稳定的位置
    :return: dict，鼠标横纵坐标
    """
    last = get_mouse_point()
    while True:
        time.sleep(0.5)
        current = get_mouse_point()
        if last == current:
            return current
        last = current

def grap_img(area):
    '''
    根据区域进行截图
    :param area:
    :return:
    '''
    return ImageGrab.grab((int(area[0][0]),int(area[0][1]),int(area[1][0]),int(area[1][1])))

def get_str_by_img(img):
    '''
    根据图片进行文字识别
    :param img:
    :return:
    '''
    return pytesseract.image_to_string(img, lang="test",config="-psm 8 -c tessedit_char_whitelist=1234567890")

def grap_by_center(center):
    '''
    以一个坐标点为中心截一个100*100像素的图片并显示出来
    :param center:
    :return:
    '''
    grap_area =  ((center[0]-50,center[1]-50),(center[0]+50,center[1]+50))
    img = grap_img(grap_area)
    img.show()
    img.close()

def get_num_coordinate(numUpleft,numDownright):
    '''
    根据选取的数字区域，来解析该区域中每个格子（将区域划分为12个格子）的数字
    :param numUpleft:
    :param numDownright:
    :return dict:  返回一个字典 能够获取到每个数字对应的中心点坐标用于点击
    '''
    offsetX =  (numDownright[0] - numUpleft[0] ) / 3 #X轴分为三块每块的偏移量
    offsetY =  (numDownright[1] - numUpleft[1] ) / 4 #Y轴分为三块每块的偏移量
    num1Center = (numUpleft[0]+offsetX/2,numUpleft[1]+offsetY/2) #第一块中心点坐标
    num2Center = (numUpleft[0]+offsetX+offsetX/2,numUpleft[1]+offsetY/2) #第二块中点坐标
    num3Center = (numUpleft[0]+offsetX*2+offsetX/2,numUpleft[1]+offsetY/2) #第三块中点坐标
    num4Center = (numUpleft[0]+offsetX/2,numUpleft[1]+offsetY+offsetY/2) #第四块中点坐标
    num5Center = (numUpleft[0]+offsetX+offsetX/2,numUpleft[1]+offsetY+offsetY/2) #第五块中点坐标
    num6Center = (numUpleft[0]+offsetX*2+offsetX/2,numUpleft[1]+offsetY+offsetY/2) #第六块中点坐标
    num7Center = (numUpleft[0]+offsetX/2,numUpleft[1]+offsetY*2+offsetY/2) #第七块中点坐标
    num8Center = (numUpleft[0]+offsetX+offsetX/2,numUpleft[1]+offsetY*2+offsetY/2) #第八块中点坐标
    num9Center = (numUpleft[0]+offsetX*2+offsetX/2,numUpleft[1]+offsetY*2+offsetY/2) #第九块中点坐标
    #num10Center = (numUpleft[0]+offsetX/2,numUpleft[1]+offsetY*3+offsetY/2) #第七块中点坐标
    num11Center = (numUpleft[0]+offsetX+offsetX/2,numUpleft[1]+offsetY*3+offsetY/2) #第八块中点坐标
    #num12Center = (numUpleft[0]+offsetX*2+offsetX/2,numUpleft[1]+offsetY*3+offsetY/2) #第九块中点坐标


    num1Area =  ((numUpleft[0],numUpleft[1]),(numUpleft[0]+offsetX,numUpleft[1]+offsetY))#第一块数字范围坐标
    num2Area = ((numUpleft[0]+offsetX,numUpleft[1]),(numUpleft[0]+offsetX*2,numUpleft[1]+offsetY))#第二块数字范围坐标
    num3Area = ((numUpleft[0]+offsetX*2,numUpleft[1]),(numUpleft[0]+offsetX*3,numUpleft[1]+offsetY))#第三块数字范围坐标
    num4Area = ((numUpleft[0],numUpleft[1]+offsetY),(numUpleft[0]+offsetX,numUpleft[1]+offsetY*2))#第四块数字范围坐标
    num5Area = ((numUpleft[0]+offsetX,numUpleft[1]+offsetY),(numUpleft[0]+offsetX*2,numUpleft[1]+offsetY*2))#第五块数字范围坐标
    num6Area = ((numUpleft[0]+offsetX*2,numUpleft[1]+offsetY),(numUpleft[0]+offsetX*3,numUpleft[1]+offsetY*2))#第六块数字范围坐标
    num7Area = ((numUpleft[0],numUpleft[1]+offsetY*2),(numUpleft[0]+offsetX,numUpleft[1]+offsetY*3))#第七块数字范围坐标
    num8Area = ((numUpleft[0]+offsetX,numUpleft[1]+offsetY*2),(numUpleft[0]+offsetX*2,numUpleft[1]+offsetY*3))#第八块数字范围坐标
    num9Area = ((numUpleft[0]+offsetX*2,numUpleft[1]+offsetY*2),(numUpleft[0]+offsetX*3,numUpleft[1]+offsetY*3))#第九块数字范围坐标
    #num10Area = ((numUpleft[0],numUpleft[1]+offsetY*3),(numUpleft[0]+offsetX,numUpleft[1]+offsetY*4))#第七块数字范围坐标
    num11Area = ((numUpleft[0]+offsetX,numUpleft[1]+offsetY*3),(numUpleft[0]+offsetX*2,numUpleft[1]+offsetY*4))#第八块数字范围坐标
    #num12Area = ((numUpleft[0]+offsetX*2,numUpleft[1]+offsetY*3),(numUpleft[0]+offsetX*3,numUpleft[1]+offsetY*4))#第九块数字范围坐标

    num_list = [
        {"numCenter":num1Center,
         "numArea":num1Area},
        {"numCenter":num2Center,
         "numArea":num2Area},
        {"numCenter":num3Center,
         "numArea":num3Area},
        {"numCenter":num4Center,
         "numArea":num4Area},
        {"numCenter":num5Center,
         "numArea":num5Area},
        {"numCenter":num6Center,
         "numArea":num6Area},
        {"numCenter":num7Center,
         "numArea":num7Area},
        {"numCenter":num8Center,
         "numArea":num8Area},
        {"numCenter":num9Center,
         "numArea":num9Area},
        {"numCenter":num11Center,
         "numArea":num11Area}
    ]
    resultDict = {}
    pageId = 1;
    for num_dict in num_list :
        img1 = grap_img(num_dict["numArea"])
        img1.save("./graps/grap{}.jpg".format(pageId))
        str1 = get_str_by_img(img1)
        print("当前格子解析的数字是{}".format(str1))
        img1.close()
        resultDict[int(str1)]=num_dict["numCenter"]
        pageId = pageId+1
    return resultDict

def click_current_num(i,numDict):
    '''
    点击数字对应的中心点坐标
    :param i: 当前的数字
    :param numDict: 数字字段 key:数字 value: 数字对应的中心点坐标
    :return:
    '''
    if i < 10:
        # 1位数
        oneNum = i
        # print("{}的坐标点位{}".format(oneNum, numDict.get(oneNum)))
        # grap_by_center(numDict.get(oneNum)) #用于数字识别错误时DEBUG使用

        click_xy(numDict.get(oneNum), round(random.uniform(0.1, 0.4), 2))

    elif i < 100:
        # 2位数
        oneNum = int(i % 10)
        twoNum = int(i / 10)
        # grap_by_center(numDict.get(twoNum)) #用于数字识别错误时DEBUG使用
        # time.sleep(2)
        # grap_by_center(numDict.get(oneNum))#用于数字识别错误时DEBUG使用
        click_xy(numDict.get(twoNum), round(random.uniform(0.1, 0.4), 2))
        click_xy(numDict.get(oneNum), round(random.uniform(0.1, 0.4), 2))
    else:
        # 3位数
        oneNum = int(i % 10)
        twoNum = int(i / 10 % 10)
        threeNum = int(i / 100)
        # grap_by_center(numDict.get(threeNum))#用于数字识别错误时DEBUG使用
        # grap_by_center(numDict.get(twoNum))#用于数字识别错误时DEBUG使用
        # grap_by_center(numDict.get(oneNum))#用于数字识别错误时DEBUG使用
        click_xy(numDict.get(threeNum), round(random.uniform(0.1, 0.4), 2))
        click_xy(numDict.get(twoNum), round(random.uniform(0.1, 0.4), 2))
        click_xy(numDict.get(oneNum), round(random.uniform(0.1, 0.4), 2))
    return

if __name__ == '__main__':

    print("鼠标移动到剩余秒数左上角")
    time.sleep(1)
    remainUpleft = get_static_mouse_point() #剩余秒数左上角坐标
    print("获取到剩余秒数左上坐标为{}".format(remainUpleft))
    time.sleep(0.5)
    print("鼠标移动到剩余秒数右下角")
    time.sleep(1)
    remainDownright = get_static_mouse_point() #剩余秒数右下角坐标
    print("获取到剩余秒数右下坐标为{}".format(remainDownright))
    print("鼠标移动到数字区域左上角")
    time.sleep(1)
    numUpleft = get_static_mouse_point() #数字区域左上角坐标
    print("获取到数字区域左上坐标为{}".format(numUpleft))
    time.sleep(0.5)
    print("鼠标移动到数字区域右下角")
    time.sleep(1)
    numDownright = get_static_mouse_point() #数字区域右下角坐标
    print("获取到数字区域右下坐标为{}".format(numDownright))

    remainArea = (remainUpleft, remainDownright)  # 剩余秒数区域

    for i in range(501): #进行循环
        im1 = grap_img(remainArea) #获取目前剩余秒数的图片
        remainstr = get_str_by_img(im1) #获取目前剩余秒数的数值
        im1.close()
        # im2 = grap_img((numUpleft,numDownright)) #用于数字识别错误时DEBUG使用
        # im2.save("./graps/umarea.jpg")
        #确保当前剩余3秒以上。需要除以10是因为后缀s被识别为5需要去掉
        while int(int(remainstr)/10) < 3:
            print("剩余秒数小于3秒，等待{}秒....".format(int(remainstr)/10))
            time.sleep(int(remainstr)/10)
            im1 = grap_img(remainArea)
            remainstr = get_str_by_img(im1)
        #获取当前数字区域的数字字典
        numDict =  get_num_coordinate(numUpleft, numDownright)
        i=i+1
        print("目前需要输入的数字为{}".format(i))
        click_current_num(i,numDict) #输入当前数字
    print("点击完成")
    time.sleep(20)