#!/usr/bin/env python3.4
# -*- encoding:utf-8 -*-

#Author:lamkinying@qq.com
#Description
'''
a. excel多个sheet情况下,遍历表单
b. 表中只要"关键词名称"所在列下一行开始的所有字符串(关键词名称不一定是奇数列)
c. 去重
d. 输出到一个新的excel,一个关键字一行
'''

import xlrd
import xlsxwriter

op_file = "查看屏蔽关键词2018-04-04 - 副本.xlsx"
op_keyword = "关键词名称"
op_result = []

#从sheet中的表格中收集屏蔽关键词
def cokwta(she):
    data = xlrd.open_workbook(op_file)
    table = data.sheet_by_name(she)
    #总列数
    ncols = table.ncols
    for col in range(ncols):
        #每列
        op_col_value = table.col_values(col)
        if op_keyword in op_col_value:
            while '' in op_col_value:
                op_col_value.remove('')
            #获取"关键词名称"在列表的索引
            op_keyword_index = op_col_value.index(op_keyword)
            op_result.extend(op_col_value[op_keyword_index + 1:])

data = xlrd.open_workbook(op_file)
#获取excel中所有sheet_name
shes = data.sheet_names()
for she in shes:
    cokwta(she)
#通过set去重
op_result = list(set(op_result))
#屏蔽关键词总数
print(len(op_result))
workbook = xlsxwriter.Workbook("new.xlsx")
worksheet = workbook.add_worksheet("关键词")
#通过enumerate函数返回索引，元素对应关系
for inx,val in enumerate(op_result):
    worksheet.write(("A"+str(inx+1)), val)
workbook.close()
