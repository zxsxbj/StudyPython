#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np
import xlrd as xl
import xlwt as xlwt
from xlutils.copy import copy


# In[3]:


# 1、读取Excel中的数据
df = pd.read_excel('d://2019.xls', sheet_name='All ratios')
# 2、删除没有用的数据
new_df = df.drop(0,axis=0,inplace=False) # 删除第一行数据 inplace=True 修改原数据
# 3、输出
print('格式化输出：\n{0}'.format(new_df))


# In[4]:


# Plesovice@01 为样品原号 
# 获取标签
# print(df.columns.values) #获取值
# print(type(df.columns.values)) #属性
# print(df.columns.values.shape) #长度

# lab_arr=df.columns.values.astype(np.string_)
# print(lab_arr.dtype) #元素属性


print('数据类型', type(df))
# 4、修改数据
"""
#containVar：查找包含的字符
#stringVar：所要查找的字符串
"""
def containVarInString(containVar,stringVar):
    try:
        if isinstance(stringVar, str):
        
             if stringVar.find(containVar):
                 return True
             else:
                 return False
        else:
            return False
    except Exception as e:
        print(e)

"""
删除数据
"""
def data_deletion(data):
    lab = ['样品原号']
    for i in data:
        if( containVarInString('Unn',i)):
            lab.append(i)
    
    return lab 
    
"""
修改数据
"""

def data_modify(data):
    data = data_deletion(data)
    for i in data:
        indexs = data.index(i)+1
        if (indexs+1)%2 == 0:
            data.insert(indexs,'±σ(%)') 
    data[1] = 'IP/nA'
    return data 


# In[5]:


# 5、将修改数据添加到DataFrame中
# new_data = new_df.rename(columns=datas, inplace = True)
datas = data_modify(new_df.columns.values)
new_df.columns = datas
print(new_df)


# In[38]:


# 保存数据到Excel
new_df.to_excel('d://2019w.xls', sheet_name='123', index=False,header=True)


# In[14]:


data  = xlwt.Workbook() # 新建一个工作簿
sheet = data.add_sheet(u'sheet1') #新建一个sheet
style = xlwt.XFStyle() #设置样式
font = xlwt.Font() #设置字体
font.name = '楷体'
font.bold = True # 加粗
font.colour_index = 2 #红色字体 0黑色 1白色  2红色 3绿色  4蓝色 5黄色 6品红色 7青色

#下划线类型
# UNDERLINE_NONE,  UNDERLINE_SINGLE,UNDERLINE_SINGLE_ACC,UNDERLINE_DOUBLE, UNDERLINE_DOUBLE_ACC
#font.underline = xlwt.UNDERLINE_SINGLE # 双下划线
#font.escapement = xlwt.Font.ESCAPEMENT_SUPERSCRIPT # 设置上标
#font.family = xlwt.Font.FAMIL # Y_ROMAN
font.height = 0x190 # 0x190 是 16 进制， 换成 10 进制为 400，然后除以 20，就得到字体的大小为 20
font.italic = True #斜体
font.struck_out = True #删除线
style.font = font #将创建的字体应用到样式中

sheet.col(0).width = 6000 # 设置列宽
sheet.set_col_default_width(2)

# 设置单元格对齐方式
alignment = xlwt.Alignment()# 创建 alignment
# May be:HORZ_GENERAL,  HORZ_LEFT,  HORZ_CENTER,  HORZ_RIGHT,  HORZ_FILLED,HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.horz = xlwt.Alignment.HORZ_CENTER # 设 置 水 平 对 齐 为 居 中 
# May be:VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER  # 设 置 垂 直 对 齐 为 居 中
style.alignment = alignment  #应用 alignment 到 style 上

style.num_format_str = 'YYYY-MM-DD HH:MM:SS' # 设置时间格式
#sheet.write(1,1,datetime.datetime.now(),style)  #在第 2行第 2列插入当前时间， 格式为 style

#设置单元格背景颜色
pattern = xlwt.Pattern()  #创建 pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN  #设置填充模式为全部填充
pattern.pattern_fore_colour = 5  #设置填充颜色为 yellow 黄色
style.pattern = pattern  #把设置的 pattern 应用到 style 上

#设置单元格边框
borders = xlwt.Borders()  #创建 borders
# May be: NO_LINE, THIN,MEDIUM,  DASHED,  DOTTED,  THICK,  DOUBLE,  HAIR,  MEDIUM_DASHED,THIN_DASH_DOTTED,  MEDIUM_DASH_DOTTED,  THIN_DASH_DOT_DOTTED,
# MEDIUM_DASH_DOT_DOTTED,  SLANTED_MEDIUM_DASH_DOTTED,  or 0x00 through 0x0D.
borders.left = xlwt.Borders.DASHED  #设置左边框的类型为虚线 
borders.right = xlwt.Borders.THIN  #设置右边框的类型为细线
borders.top = xlwt.Borders.DOTTED  #设置上边框的类型为打点的
borders.bottom = xlwt.Borders.THICK  #设置底部边框类型为粗线
borders.left_colour = 0x10  #设置左边框线条颜色
borders.right_colour = 0x20
borders.top_colour = 0x30
borders.bottom_colour = 0x40
style.borders = borders  #将 borders 应用到 style上
sheet.write(3, 0, 'HuZhangdong', style)  #在第 4 行第 1 列写入 'HuZhangdong' ，格式引用style
data.save(u'e:\\3.xls')  #保存到 e:\\3.xls


# In[ ]:




