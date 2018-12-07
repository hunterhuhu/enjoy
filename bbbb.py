from selenium import webdriver
from PIL import Image
from selenium.webdriver import ActionChains
import os,time,random
import xlrd
from xlwt import *
from xlrd import open_workbook
from xlutils.copy import copy
from selenium.webdriver.common.keys import Keys
#异常代码
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchFrameException

def get_track(distance):
    track = []
    current = 0
    mid = distance*3/4
    t = random.randint(2, 3)/10
    v=0
    while current < distance:
        if current < mid:
            a = 2
        else:
            a = -3
        v0 = v
        v = v0+a*t
        move = v0*t+1/2*a*t*t
        current += move
        track.append(round(move))
    return track
# 生成拖拽移动轨迹，加3是为了模拟滑过缺口位置后返回缺口的情况

def read_excel(workbook,num):
    # 获取所有sheet
    print(workbook.sheet_names())
    # 根据sheet索引或者名称获取sheet内容
    # sheet2 = workbook.sheet_by_index(1)  # sheet索引从0开始
    sheet1 = workbook.sheet_by_name('Sheet1')
    # sheet的名称，行数，列数
    # print(sheet1.name, sheet1.nrows, sheet1.ncols)
    # 获取整行和整列的值（数组）
    # rows = sheet1.row_values(3)  # 获取第四行内容
    # cols = sheet1.col_values(1)  # 获取第2列内容
    sheet1.cell(num, 0)
    # 获取单元格内容
    print(sheet1.cell(num, 0).value)
    return sheet1.cell(num, 0).value
    # print(sheet1.cell_value(1, 0).encode('utf-8'))
    # print(sheet1.row(1)[0].value.encode('utf-8'))
    # # 获取单元格内容的数据类型
    # print(sheet1.cell(1, 0).ctype)
# 打开文件
print('请将待查询的表格放入D盘，第一列为单号，第二列为状态，并将名称修改为‘dan1.xls’\n')
n = int(input("请输入每次查询的顺丰快递数量:"))
number = int(input("请输入从第几行开始查询:"))
workbook = xlrd.open_workbook(r'D:\dan1.xls')
wb = copy(workbook)
ws = wb.get_sheet(0)
danhao = ['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
status = ['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
status_xpath = ['0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0']
danhao[0] = read_excel(workbook,1)
driver= webdriver.Chrome()
driver.implicitly_wait(20)
driver.get('http://www.sf-express.com/cn/sc/dynamic_function/waybill/#search/bill-number/')
driver.maximize_window()
driver.implicitly_wait(1)
while 1:
    for i in range(0,n):
        danhao[i] = read_excel(workbook,number+i)
    number += n
    flag2=1
    while flag2:
        try:
            driver.find_element_by_xpath('//*[@id="function"]/div/div/div[1]/div/div[1]/div/label/span').click()#删除历史单号
            flag2 = 0
        except NoSuchElementException:
            time.sleep(0.2)
    time.sleep(0.5)
    for i in range(0, n):
        driver.find_element_by_class_name('token-input').send_keys(danhao[i]+' ')  # 输入运单号
        time.sleep(0.1)
    time.sleep(2)

    driver.find_element_by_id('queryBill').click()  # 查询
    time.sleep(0)
    flag = 1
    while flag :
        try:
            driver.switch_to.frame('tcaptcha_popup')
            flag = 0
        except NoSuchFrameException:
            flag = 1
    # getElementImage(driver.find_element_by_xpath('//*[@id="slideBlock"]'))
    flag = 9
    yundong = [0,260,240,220,260,240,220,260,240,220]
    while flag:
        flag3 = 1
        while flag3:# 直到找到缺块
            try:
                driver.find_element_by_xpath('//*[@id="tcaptcha_drag_thumb"]/div[2]')  # 直到找到缺块
                flag3 = 0
            except NoSuchElementException:
                flag3 = 1
                time.sleep(0.2)
        butten0 = driver.find_element_by_xpath('//*[@id="tcaptcha_drag_thumb"]/div[2]')  # 找到缺块
        action = ActionChains(driver)
        track_list = get_track(yundong[flag])  # 生成轨迹
        flag -= 1
        time.sleep(0.8)
        action.click_and_hold(butten0)  # 根据轨迹拖拽缺块
        for track in track_list:
            action.move_by_offset(track, 0)
        action.release(butten0).perform()  # 拖拽缺块
        time.sleep(2.5)#等待缺块归位
        try:
            driver.find_element_by_xpath('//*[@id="tcaptcha_drag_thumb"]/div[2]')#看是否还能找到缺块
            time.sleep(0.1)
            continue
        except NoSuchElementException:#找不到缺块证明已经查找完
            driver.switch_to.default_content()
            #查到当前
            flag1 = 1
            while flag1:
                time.sleep(0.2)
                try:
                    driver.find_element_by_xpath("//span[contains(text(),'查询中')]")
                except NoSuchElementException:#查询完成
                    for i in range(0, n):
                        status_xpath[i] = '//*[@id="waybill-'+danhao[i]+'"]/span'
                        status[i] = driver.find_element_by_xpath(status_xpath[i]).text
                        if status[i] == '运送中':
                            status[i] = '待收款'
                        elif status[i] == '已退回':
                            status[i] = '拒签退回'
                        elif status[i] == '已签收':
                            status[i] = '已签收'
                        else:
                            status[i] = '有问题'
                    for i in range(0,n):
                        ws.write(number-n+i, 2, status[i])
                    flag1 = 0
                    flag = 0
                    wb.save('D:\dan1.xls')
                    break
    time.sleep(0.2)
#务必记得加入quit()或close()结束进程，不断测试电脑只会卡卡西
driver.close()
