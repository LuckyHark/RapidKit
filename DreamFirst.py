#!/usr/bin/python
#--coding: utf-8 --

import pandas as pd #导入pandas模块并命名为pd
import xlrd #导入xlrd模板用于EXCEL读取
import xlwt #导入xlwt模板用于EXCEL写入
import numpy as np
import scipy.stats.stats as stats
import math
import scipy.stats
from scipy.stats import kstest, ttest_ind, levene
from scipy.stats import shapiro

#数据读取与数据初始化
FileRoad = 'D:/Python/My dream/SWIS.xls'
Outfile = 'D:/Python/My dream/NewData_dream.xls'
data = pd.read_excel(FileRoad)

dv_list = data.columns.values.tolist()
vt_list=[]; vn_list=[]; vd_list=[]; v_counting=[]; vf_list=[]; vfnum_list=[];list_alpha=[]

#根据变量字符名称判断变量类型、变量名
for dv_name in dv_list:
    dv_split = dv_name.split("_")
    if   dv_split[0] == "iv": vt = 1
    elif dv_split[0] == "dv": vt = 2
    elif dv_split[0] == "cv": vt = 3
    else:vt = 0
    vt_list.append(vt)
    vn_list.append(dv_split[1])
    #根据变量名长度判断是否为组间变量或量表变量
    if   len(dv_split)<=2: vf="NA";vf_list.append(vf);vfnum="NA";vfnum_list.append(vfnum)
    elif len(dv_split)> 2:
         if dv_split[2] == "sc":vf = 1;vf_list.append(vf);vfnum=dv_split[3];vfnum_list.append(vfnum)
         if dv_split[2] == "rp":vf = 2;vf_list.append(vf);vfnum=dv_split[3];vfnum_list.append(vfnum)

#统计全表变量类型数目
iv_counting = vt_list.count(1)
dv_counting = vt_list.count(2)
cv_counting = vt_list.count(3)

#根据变量数值判断变量分布类型(Shapiro正态、二值分布)
for ia in dv_list:
    xa = shapiro(data[ia])
    ya = list(xa)
    if ya[1]>0.05:
       ya[1]="1"
    else:
        ya[1] = "0"
    vd_list.append(ya[1])
#判断变量数值数目
for ib in dv_list:
    xb = np.unique(data[ib])
    yb = len(xb)
    v_counting.append(yb)

#生成变量信息表
dictionary = {'vName':vn_list,
              'vType':vt_list,
              'vDistribution':vd_list,
              'vCount':v_counting,
              'vFunction:':vf_list}

frame = pd.DataFrame(dictionary)

#提取含有sc字符的变量
vsc_list = []  # 存储包含‘sc’字段的列名
ie_list = []
for ie in dv_list:
    if 'sc' in ie:
        ie_list.append(ie) #含有sc变量的列名表
    else : xww=1
ie_data=data[ie_list] #含有sc变量的列表

vsc_name=[]
vsc_num = []
vsc_nt = []
for ig in ie_list:
    ig_split = ig.split("_")
    vsc_name.append(ig_split[1])
    vsc_nt = list(set(vsc_name))

#计数rp与sc变量的重复次数
dic = {}
for ik in vsc_name:
    dic[ik] = vsc_name.count(ik)
list_key = dic.keys()
dic_len = len(list_key)
list_test = []
tt=0
list_dd=[]

for il in list_key: #循环变量种类（如control和pro）
    xh1 = "vsc_sublist_%s = []"  % (il) #循环定义列表变量
    xh2 = "mean_%s=0" % (il) #循环定义均值变量
    exec xh1
    exec xh2
    xhh1 = eval("vsc_sublist_%s" % il)
    xhh2 = eval("mean_%s" % il)
    for z in dv_list: #循环创建list（分别包含control和pro）
        if il in z: xhh1.append(z)

    #计算Cronbach's α
    vsc_data = data[xhh1]
    total_row = vsc_data.sum(axis=1)
    sy = total_row.var()  # 观察值方差
    var_column = vsc_data.var()  # 项目值方差
    si = var_column.sum()  # 项目值方差和
    k_alpha = 3.00  # 项目数，即sc量表后的最大数字
    r = (k_alpha / (k_alpha - 1)) * ((sy - si) / sy)
    list_alpha.append(r)
    ##信度大于0.65计算均值
    if r>=0.65:
        vsc_mean = (vsc_data.mean(axis=1))
        data["mean_%s" % il]=vsc_mean
        print "【信度】","The reliability of", il, "is good.", "Cronbach's alpha = ", r
    if r<0.65:
        print "【信度】","The reliability of", il , "is unacceptable.", "Cronbach's alpha = ", r


#信度检验合格并求得均值的矩阵
dv_list = data.columns.values.tolist() #将合并后生产的变量列加入变量list
ra_list = []
for ra in dv_list:
    if 'sc' not in ra:
        ra_list.append(ra) #含有sc变量的列名表,即去除原始量表数据
data_mean = data[ra_list]
len_dv_list = len(ra_list)
#print data_mean

#分割变量全名取出变量名
ra_name_list = []
for name in ra_list:
    ra_name_list.append(name)

#判断相关算法并计算相关
result_r = []
cor_name_list = []
for ix in ra_list:
    index_r = ra_list.index(ix) #变量所在的索引值
    x = data_mean[ix]
    for ik in range(len_dv_list-index_r-1):
        y = data_mean[ra_list[index_r+ik]]
        result_r.append(scipy.stats.pearsonr(x,y))
        cor_name_list.append(ra_name_list[index_r]+" and "+ra_name_list[ik+1])

#判断p值显著性
for i in range(len(result_r)):
    p_r= result_r[i]
    p_picked = p_r[1]
    r_picked = float('%.2f' % p_r[0])
    if p_picked >0 and p_picked < 0.001:
        print "【相关】",cor_name_list[i+1], "is extremely significant,", "p < .001",", r =",r_picked
    if p_picked >0.001 and p_picked <=0.01:
        print "【相关】",cor_name_list[i + 1], "is significant,", "p =", float('%.2f' % p_picked), ", r =", r_picked
    if p_picked > 0.01 and p_picked <= 0.05:
        print "【相关】",cor_name_list[i + 1], "is significant,", "p =", float('%.2f' % p_picked), ", r =", r_picked
    if p_picked > 0.05 and p_picked <= 0.01:
        print "【相关】",cor_name_list[i + 1], "is marginally significant,", "p =", float('%.2f' % p_picked), ", r =", r_picked

#全表r值
r_result_data = data_mean.corr()
#print r_result_data

writer2 = pd.ExcelWriter(Outfile)
r_result_data.to_excel(writer2,'Correlation')
data_mean.to_excel(writer2,'Mean')
frame.to_excel(writer2,'VariableParameters')
writer2.save()