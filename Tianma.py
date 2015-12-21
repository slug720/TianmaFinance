# -*- coding: utf-8 -*-
"""
Created on Tue Dec  8 22:49:49 2015

@author: Administrator
"""
#==============================================================================
#==============================================================================
# 说明:
# 1.将Tianma.py和静态表拷贝到工作目录中；
# 2.在工作目录下建立Data文件夹，将银行数据拷贝至该文件夹;该文件夹下所有的带有银行关键字
# 的xls和xlsx文件都会被处理
# 3.双击执行Tianma.py；会在工作目录下生成Result文件夹及分类结果
#==============================================================================
#==============================================================================

from pandas import DataFrame, Series
import pandas as pd
import numpy as np
import math
import string
import os
from numpy import nan as NA
from os.path import join, getsize

WorkPath = 'E:/_Projects/Personal/SVN/_Projects/Python/TianmaFinance'
os.chdir(WorkPath)

WorkPath = os.getcwd()
DataPath = WorkPath + '/Data'
ResultPath = WorkPath + '/Result'
if not os.path.exists('Result'):
    os.mkdir('Result')  #如果无Result文件夹则创建Result文件夹

#读取分类规则
StatRule1 = pd.read_excel(WorkPath + '/静态表.xlsx',index_col = 0) #第一页读取字段关键字
StatRule2 = pd.read_excel(WorkPath + '/静态表.xlsx',1) #第二页读取分类规则
ClassifyRuleDF = StatRule2.ix[:,'最终结果']
ClassifyRuleDF.index = [StatRule2.银行,StatRule2.收支,StatRule2.大类,StatRule2.对方户名,StatRule2.关键字]

os.chdir(DataPath)

class StatFileClass:
    def __init__(self,temp=0):
        self.FileName = temp
        self.FileNameHead = self.FileName[0:self.FileName.find('.')]
        self.BankName = self.FileName[0:self.FileName.find('银行')+2]
        if '美元' in temp:
            self.Currency = 'USD'
        elif '日币' in temp:
            self.Currency = 'JPY'
        else:
            self.Currency = 'CNY'
        self.DateLable = StatRule1.ix[self.BankName,'交易日期字段']
        self.IncomeLable = StatRule1.ix[self.BankName,'收入字段']
        self.PayLable = StatRule1.ix[self.BankName,'支出字段']
        self.BalanceLable = StatRule1.ix[self.BankName,'当日余额字段']
        self.KeyLable1 = StatRule1.ix[self.BankName,'大类字段']
        self.KeyLable2 = StatRule1.ix[self.BankName,'子类字段'].split('+') #字符串list
        self.CountLable = StatRule1.ix[self.BankName,'户名字段']
        self.SkipRows = StatRule1.ix[self.BankName,'数据开始行']-1
        
        #判断收入支出类型
        if self.IncomeLable == self.PayLable: #中国银行,收入字段和支出字段相同
            self.RawData = pd.read_excel(self.FileName,skiprows = self.SkipRows,converters = {self.IncomeLable : str})
            self.IncomeData =  self.RawData.ix[:,self.IncomeLable].astype(float)
            self.IncomeType = Series(np.zeros(self.IncomeData.shape[0])) #初始化
            for i in range(self.IncomeData.shape[0]):
                if self.IncomeData[i] > 0:
                    self.IncomeType[i] = '收入'
                else:
                    self.IncomeType[i] = '支出'                  
        else:
            self.RawData = pd.read_excel(self.FileName,skiprows = self.SkipRows,converters = {self.IncomeLable : str, self.PayLable : str})
            self.IncomeData =  self.RawData.ix[:,self.IncomeLable].astype(float).fillna(0)
            self.PayData =  self.RawData.ix[:,self.PayLable].astype(float).fillna(0)
            self.IncomeType = Series(np.zeros(self.IncomeData.shape[0])) #初始化
            for i in range(self.IncomeData.shape[0]):
                if self.IncomeData[i] > self.PayData[i]:
                    self.IncomeType[i] = '收入'
                else:
                    self.IncomeType[i] = '支出'
                    
        #获取大类数据
        self.KeyWord1 = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        if self.KeyLable1 == '无':
            self.KeyWord1[:] = '无'
        else:
            self.KeyWord1 = self.RawData.ix[:,self.KeyLable1]
        
        #获取交易户名数据
        self.CountName = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        if self.CountLable == '无':
            self.CountName[:] = '无'
        else:
            self.CountData = self.RawData.ix[:,self.CountLable]
            self.CountData.fillna('无',inplace = True)
            for i in range(self.IncomeData.shape[0]):
                TempData = list(set(list(zip(*list(ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].index)))[0])) #获取户名的唯一值的list
                if len(TempData) == 1: #子类字段只有一个
                    self.CountName[i] = TempData[0]
                else:
                    bFindResult = False
                    for j in TempData:
                        if j in self.CountData[i]:
                            self.CountName[i] = j
                            bFindResult = True
                            break
                    if not bFindResult:
                        self.CountName[i] = '无'
                        
        #获取子类数据并分类
        self.KeyWord2 = self.RawData.ix[:,self.KeyLable2]
        self.KeyWord2.fillna('无',inplace = True)
        self.ClassifyResult = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        for i in range(self.KeyWord2.shape[0]):
            TempData = list(set(list(ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].ix[self.CountName[i]].index)))
            if len(TempData) == 1: #子类字段只有一个
                self.KeyWord2[i] = TempData[0]
            else:
                bFindResult = False
                for j in TempData:
                    IterList = j.split('+')
                    bFindResult2 = True
                    for k in IterList:
                        bFindResult3 = False
                        for m in list(self.KeyWord2.ix[i]):
                            if k in m:
                                bFindResult3 = True
                                break
                        if not bFindResult3:
                           bFindResult2 = False
                           break
                    if bFindResult2:
                        self.KeyWord2[i] = j
                        bFindResult = True
                        break
                if not bFindResult:
                    self.KeyWord2[i] = '无'
            if self.KeyWord2[i] in TempData:                
                self.ClassifyResult[i] = ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].ix[self.CountName[i]].ix[self.KeyWord2[i]]
            else:
                self.ClassifyResult[i] = '分类错误'
        if self.IncomeLable == self.PayLable: 
            self.ResultDF = pd.concat([self.IncomeData,self.IncomeType,self.KeyWord1,self.CountName,self.KeyWord2,self.ClassifyResult],axis = 1)
            self.ResultDF.columns = ['收入','收支类型','大类','对方户名','子类','分类结果']
        else:
            self.ResultDF = pd.concat([self.IncomeData,self.PayData,self.IncomeType,self.KeyWord1,self.CountName,self.KeyWord2,self.ClassifyResult],axis = 1)
            self.ResultDF.columns = ['收入','支出','收支类型','大类','对方户名','子类','分类结果']
        self.ResultFileName = ResultPath + '/' +self.FileNameHead + '分类结果.xlsx'       
        self.ResultDF.to_excel(self.ResultFileName,self.FileNameHead)


for FileName in os.listdir(DataPath):
    if ('银行'  in FileName) and ('xls' in FileName) and ('分类结果' not in FileName):
        print('正在处理:' + FileName)
        StatFileClass(FileName);
print('处理完毕!')
        
A = StatFileClass('中国银行美元一般户.xls')
A = StatFileClass('中国银行.xls')

A.CountName = Series(np.zeros(A.IncomeData.shape[0]))  #初始化
if A.CountLable == '无':
    A.CountName[:] = '无'
else:
    A.CountData = A.RawData.ix[:,A.CountLable]
    for i in range(A.IncomeData.shape[0]):
        TempData = list(set(list(zip(*list(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].index)))[0])) #获取户名的唯一值的list
        if len(TempData) == 1: #子类字段只有一个
            A.CountName[i] = TempData[0]
        else:
            bFindResult = False
            for j in TempData:
                if j in A.CountData[i]:
                    A.CountName[i] = j
                    bFindResult = True
                    break
            if not bFindResult:
                A.CountName[i] = '无'

A.KeyWord2 = A.RawData.ix[:,A.KeyLable2]
A.KeyWord2.fillna('无',inplace = True)
A.ClassifyResult = Series(np.zeros(A.IncomeData.shape[0]))  #初始化
for i in range(A.KeyWord2.shape[0]):
    TempData = list(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].index)
    if len(TempData) == 1: #子类字段只有一个
        A.KeyWord2[i] = TempData[0]
    else:
        bFindResult = False
        for j in TempData:
            if j in A.KeyWord2[i]:
                A.KeyWord2[i] = j
                bFindResult = True
                break
        if not bFindResult:
            A.KeyWord2[i] = '无'
    if A.KeyWord2[i] in TempData:                
        A.ClassifyResult[i] = ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].ix[A.KeyWord2[i]]
    else:
        A.ClassifyResult[i] = '分类错误'

i+=1
C = list(set(list(zip(*list(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].index)))[0]))