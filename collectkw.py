#!/usr/bin/env python3.4
# -*- encoding:utf-8 -*-

#Author:lamkinying@qq.com
#Description
'''
a. excel多个sheet情况下,遍历表单
b. 表中只要"关键词名称"所在列下一行开始的所有字符串(关键词名称不一定是奇数列)
c. 去重
d. 输出到一个新的excel,一个关键字一行
e. 带上"关键词名称"所对应的"入库时间"
'''

import time
import xlrd
import xlsxwriter

op_file = "查看屏蔽关键词2018-04-04 - 副本.xlsx"
op_keyword = "关键词名称"
op_result = {}

#从sheet中的表格中收集屏蔽关键词
def cokwta(she):
    data = xlrd.open_workbook(op_file)
    table = data.sheet_by_name(she)
    #总列数
    ncols = table.ncols
    for col in range(ncols):
        #关键词名称列
        op_col_kwvalue = table.col_values(col)
        if op_keyword in op_col_kwvalue:
            #获取"关键词名称"在列表的索引
            op_keyword_index = op_col_kwvalue.index(op_keyword)
            #获取到的"关键词名称"列表
            op_col_kwvalue_get = op_col_kwvalue[op_keyword_index + 1:]
            while '' in op_col_kwvalue_get:
                op_col_kwvalue_get.remove('')
            #关键词名称列对应的入库时间列
            op_col_dtvalue = table.col_values(col + 1)
            op_col_dtvalue_get = op_col_dtvalue[op_keyword_index+1:op_keyword_index+len(op_col_kwvalue_get)+1]
            #"关键词名称"作为key，"入库时间"作为value。
            for k_w,d_t in zip(op_col_kwvalue_get,op_col_dtvalue_get):
                if k_w in op_result.keys():
                    #如果"关键词名称"出现多个."入库时间"比较新旧，新的作为value
                    if time.strftime(d_t) > time.strftime(op_result[k_w]):
                        op_result[k_w] = d_t
                else:
                    op_result[k_w] = d_t

data = xlrd.open_workbook(op_file)
#获取excel中所有sheet_name
shes = data.sheet_names()
for she in shes:
    cokwta(she)
#屏蔽关键词总数
print(len(op_result.keys()))
workbook = xlsxwriter.Workbook("new.xlsx")
worksheet = workbook.add_worksheet("关键词")
worksheet.write("A1", "关键词名称")
worksheet.write("B1", "入库时间")
n = 1
for k,v in op_result.items():
    n += 1
    worksheet.write("A"+str(n), k)
    worksheet.write("B"+str(n), v)
workbook.close()
