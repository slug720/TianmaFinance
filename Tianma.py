# -*- coding: utf-8 -*-
"""
Created on Tue Dec  8 22:49:49 2015

@author: Administrator
"""

from pandas import DataFrame, Series
import pandas as pd
import numpy as np
import math
import string
import os
from numpy import nan as NA
from os.path import join, getsize

WorkPath = 'D:/_Projects/Personal/SVN/_Projects/Python/TianmaFinance'
os.chdir(WorkPath)

StatRule1 = pd.read_excel('静态表.xlsx',index_col = 0)
StatRule2 = pd.read_excel('静态表.xlsx',1)
ClassifyRuleDF = StatRule2.最终结果
ClassifyRuleDF.index = [StatRule2.银行,StatRule2.收支,StatRule2.大类,StatRule2.关键字]
ClassifyRuleDF.ix['招商银行'].ix['支出'].ix['对公售汇'].ix['无']

class StatFileClass:
    def __init__(self,temp=0):
        self.FileName = temp
        self.BankName = self.FileName[0:self.FileName.find('银行')+2]
        if '美元' in temp:
            self.Currency = 'USD'
        elif '日币' in temp:
            self.Currency = 'JPY'
        else:
            self.Currency = 'CNY'
        self.IncomeLable = StatRule1.ix[self.BankName,'收入字段']
        self.PayLable = StatRule1.ix[self.BankName,'支出字段']
        self.BalanceLable = StatRule1.ix[self.BankName,'当日余额字段']
        self.KeyLable1 = StatRule1.ix[self.BankName,'大类字段']
        self.KeyLable2 = StatRule1.ix[self.BankName,'子类字段']
        self.SkipRows = StatRule1.ix[self.BankName,'数据开始行']-1
        
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
        self.KeyWord1 = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        if self.KeyLable1 == '无':
            self.KeyWord1[:] = '无'
        else:
            self.KeyWord1 = self.RawData.ix[:,self.KeyLable1]
        self.KeyWord2 = self.RawData.ix[:,self.KeyLable2]
        self.KeyWord2.fillna('无',inplace = True)
        self.ClassifyResult = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        for i in range(self.KeyWord2.shape[0]):
            TempData = list(ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].index)
            if len(TempData) == 1: #子类字段只有一个
                self.KeyWord2[i] = TempData[0]
            else:
                bFindResult = False
                for j in TempData:
                    if j in self.KeyWord2[i]:
                        self.KeyWord2[i] = j
                        bFindResult = True
                        break
                if not bFindResult:
                    self.KeyWord2[i] = '无'
            if self.KeyWord2[i] in TempData:                
                self.ClassifyResult[i] = ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].ix[self.KeyWord2[i]]
            else:
                self.ClassifyResult[i] = '分类错误'
        
        #if StatRule1.ix[self.BankName,'收入字段'] != StatRule1.ix[self.BankName,'支出字段']:       

index = 0
A = locals()
for FileName in os.listdir(WorkPath):
    if ('银行'  in FileName) and ('xls' in FileName):
        index += 1
        print(FileName)
        A['StatFile' + str(index)] = StatFileClass(FileName);
        
B = pd.read_excel('农业银行.xls',skiprows = StatRule1.ix['农业银行','数据开始行']-1, converters = {'收入金额' : str})

A = StatFileClass('中国银行.xls')
A = StatFileClass('农业银行.xls')
A = StatFileClass('招商银行.xlsx')
A.KeyWord2[:] = '无'
i = 0
TempData = list(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].index)
if len(TempData) == 1:
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
A.ClassifyResult[i] = ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]].ix[A.KeyWord2[i]]
i += 1

            
for i in range(A.KeyWord2.shape[0]):        
    print(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[i]].ix[A.KeyWord1[i]])
ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[1]].ix[A.KeyWord1[1]]
ClassifyRuleDF.ix[A.BankName]
B = ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[1]]
C = list(ClassifyRuleDF.ix[A.BankName].ix[A.IncomeType[1]].ix[A.KeyWord1[1]].index)
for j in C:
    print(j)
    
