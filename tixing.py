#coding:utf-8
import sys

reload(sys)
sys.setdefaultencoding("utf-8")
import requests
import json
import csv,time,os
import xlwt
from xlutils.copy import copy
from xlrd import open_workbook
import winsound
import pymysql
from win32com.client import Dispatch
from playsound import playsound
from wxpy import *




def zhuaqu():
    # 爬虫地址
    url = "http://ces.sino-life.com:7001/SL_CES/marketingChannel/queryMarketCollectList.do"
    headers = {
    "Host": "ces.sino-life.com:7001",
    "Connection": "keep-alive",
    "Content-Length": "14",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "SF_AJAX": "true",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36",
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin": "http://ces.sino-life.com:7001",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Cookie":"""_sf_profile="63757272656e74556964:7765692e7a68656e673030364073696e6f2d6c6966652e636f6d@70726f66696c65496e6974466c6167:37@70726f66696c65436667466c6167:30"; JSESSIONID=F4B342ACFD9F88621B30384F40B508EE; SF_LOGIN_TIME="2019/09/04 14-24-20@2019/09/05 07-44-28"; BIGipServerCES_PRD_POOL_NEW=3129911488.9392.0000""",
    }
    data = {"channelType":"01","branchName":"葫芦岛中心支公司(862114)","startDate":"2019-09-5","endDate":"2019-09-5","productCode":"","page":"1","rows":"100"}
    response = requests.post(url=url, headers=headers,data=data, verify=False)
    jsonobj = json.loads(response.text)
    json_list =  jsonobj["rows"]

    return json_list

    # for jsonobj2 in json_list:
    #     name = jsonobj2['agentName']
    #     number = jsonobj2["agentNo"]
    #     money = jsonobj2["standardPrem"]
    #     place = jsonobj2["branchName"]
    #     xianzhong =jsonobj2["productName"]

def excel():
    values = zhuaqu()

    rexcel = open_workbook("Python.xls", formatting_info=False)  # 用wlrd提供的方法读取一个excel文件
    excel = copy(rexcel)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    work_sheet = excel.get_sheet(0)  # 用xlwt对象的方法获得要操作的sheet
    work_sheet.write(0, 0, u'姓名')
    work_sheet.write(0, 1, u'工号')
    work_sheet.write(0, 2, u'保费')
    work_sheet.write(0, 3, u'营服')
    work_sheet.write(0, 4, u'险种')
    work_sheet.write(0, 5, u'分区代码')
    row = 1
    rows_num = rexcel.sheets()[0].nrows
    print rows_num


    dict_obj = {
        "I00665000001":"本级",
        "I00665000002":"本级",
        "I00665001395":"本级",
        "I00665002293":"本级",
        "I000000000230848":"本级",
        "I000000000394104":"本级",
        "I00665000919":"建昌",
        "I000000000254136": "建昌",
        "I000000000342993": "建昌",
        "I000000000389946": "建昌",
        "I000000000712626": "建昌",
        "I00665003107": "二区",
        "I000000000051679": "二区",
        "I000000000321700": "二区",
        "I000000000372278": "二区",
        "I000000000374813": "二区",
        "I00665000920": "一区",
        "I00665002691": "一区",
        "I00665002761": "一区",
        "I00665002832": "一区",
        "I00665003197": "一区",
        "I000000000025399": "一区",
        "I000000000301265": "一区",
        "I00665001015": "兴城",
        "I000000000268655": "兴城"

    }



    print(values)
    for date in values:
        key = date["zoneCode"]
        if key in dict_obj:
            yingfu = dict_obj[key]



        work_sheet.write(row, 0, date["agentName"])
        work_sheet.write(row, 1, date["agentNo"])
        work_sheet.write(row, 2,date["standardPrem"])
        work_sheet.write(row, 3,u"%s"%yingfu)
        work_sheet.write(row, 4,date["productName"])
        work_sheet.write(row, 5,date["zoneCode"])

        row += 1

    print row


    if row > rows_num:
        # duration = 4000  # millisecond
        # freq = 700  # Hz
        # winsound.Beep(freq, duration)
        playsound('ts.mp3')

    excel.save("Python.xls")  # xlwt对象的保存方法，这时便覆盖掉了原来的excel
    #执行以下








if __name__ == '__main__':

    shuju = zhuaqu()
    while 1:
        ceshi = excel()




