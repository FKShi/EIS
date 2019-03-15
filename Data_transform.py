#_*_ coding: utf-8 _*_
# author: Fukun Shi
'''
本程序用于读取HP4294A通过VBA excel存储的数据，并将其转换成只包含测试数据的excel文档方便后续处理。
本程序基于python 3.6 + pandas利用jupyter notebook编写。也可以用于读取其他任意excel文档。
对每个sheet里特定行列的选择通过pandas里的read_excel函数实现，具体用法参考下面实例。
要注意的是sheet_name不是0和none时，读取的内容时ndarray形式，需要用DataFrame函数转化为pandas的dataframe。
在一个loop中实现每次循环命名一个文件是通过‘+filename+’.xlsx
'''
import pandas as pd
import numpy as np
import openpyxl as op

m = 201     # 这是测量的采集点数，默认201
filepath = 'path\\4294A_DataTransfer_0310.xlsx' # 路径名得用\\连接上下级
IS = pd.ExcelFile(filepath)   # 读取excel文件，这个类可以方便的找出sheet数目
nsheets = IS.sheet_names[2:]  # 提取每个数据对应的sheet名，用于后面的命名
nsheet = np.size(nsheets)     # 计算含有多少个sheets
print(nsheets)                # 打印出sheets的名字，方便验证后续的新文件名命名正确

for sheetname in range(nsheet):
    EIS = pd.read_excel('4294A_DataTransfer_0310.xlsx', sheet_name=[sheetname+2],header=[9], skiprows=0, skipfooter=801-m+1, usecols=[2,3,5]) # read_excel 各参数解释：首先是读取的文件路径和文件名，然后sheet_name指定要读取的sheets，skipfooter是从尾部开始忽略的行数，usecols指定要读取的列
    EIS = EIS[sheetname+2].values   #将数据转为ndarray,方便转为pandas的dataframe形式
    assert type(EIS)==np.ndarray
    Impedance = pd.DataFrame(EIS, columns=['Frequency','Data Real', 'Data Imag']) #将ndarray类型转化为pandas所用的dataframe，便于后续写入操作。
    filename = 'D:\\Nutstore\\PhD\Experiment\\Cooperation-Biosensor\\Hydrogel Plasma\\IDE\\13-03-2019 w medium thin layer\\'+str(nsheets[sheetname])+'.xlsx'
    Impedance.to_excel(filename, index=False) 
