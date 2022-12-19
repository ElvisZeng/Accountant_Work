#! python3.7
#_*_ coding: utf-8, unicode
"""
@ Time: 2020-3-27
@ Author: Elvis_Zeng
@ File: BS_Combinator.py
@ Version: 0.0.1
"""
print('专为金蝶报表系统导出的合并报表设计\n可以将各个合并主体的合并报表复制到一张表上')
from os import getcwd, listdir, path, chdir
import xlwings as xw
from pandas import  concat, set_option, DataFrame

set_option('display.float_format', lambda x: '%.2f' % x)             # Set pandas show the number after point

workdir =  getcwd()              #get the current work direction
filelist = [u for u in listdir() if u[-5:]=='.xlsx']             #get the file list under the current work direction which was .xls files.


confirm = input('本文件夹内的可合并文件共有'+str(len(filelist))+'个。'+
                '\n分别是：'+ str(filelist) + 
                '\n请确认要处理的文件无误!!!!\n'+
                '输入ok确认要处理的表格文件\n'+
                '合并文件不正确按回车退出!\n'+
                'input: ')

def BS_Combine(x):         # deal with the excel file
    book = EXCEL.books.open(x)
    BS_sht = book.sheets['管理-资产负债表']
    sht_name = BS_sht['b1'].value[:-9]           #取合并主体名称
    Consolidate_Column = search_cell_address(BS_sht, '合并数\n(期末数)')
    #BS_sht[Consolidate_Column].value = sht_name
    BS_index_keys = BS_sht['a1:a110'].value
    BS_values = BS_sht[Consolidate_Column+':'+Consolidate_Column.split('$')[1]+'110'].value

    PS_sht = book.sheets['管理-利润表']
    PS_sht_field = [i.split('$') for i in PS_sht.used_range.address.split(':')]
    PS_index_keys = PS_sht['a2:a73'].value
    PS_values = PS_sht['b2:b73'].value
    
    Index_keys = BS_index_keys + PS_index_keys
    column_keys = [sht_name]
    Values = BS_values + PS_values
    DF = DataFrame(Values, index= Index_keys, columns=column_keys)
    DF = DF.fillna(0)
    book.close()
    return DF

def search_cell_address(x, i):
    for u in x.used_range:
        if u.value == i:
            return u.address
        
if confirm == "ok":
    EXCEL = xw.App(add_book=False)
    DFs = [BS_Combine(i) for i in filelist]
    Result = concat(DFs, axis = 1, sort = False)
    Result.to_excel('报表合并.xlsx')
    print(' 文件合并成功！\n 保存位置：', workdir, '\n 文件名称：报表合并.xlsx')
    input(' ***按Enter退出***')
    EXCEL.kill()

else:
    input("未确认合并文件！按Enter键退出！")
