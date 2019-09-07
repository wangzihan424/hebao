#coding:utf-8
import sys

reload(sys)
sys.setdefaultencoding("utf-8")

import json
str1 = "K:1|K1:2|K2:3"
str1 = str1.replace("|",'","')

str1 = str1.replace(":",'";"')
str2 = '{"'+str1+'"}'
print json.loads(str2)