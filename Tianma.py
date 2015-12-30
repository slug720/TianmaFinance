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
from datetime import * 
from numpy import nan as NA
from dateutil.parser import parse

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
ClassifyRuleDF = StatRule2['最终结果']
ClassifyRuleDF.index = [StatRule2.银行,StatRule2.收支,StatRule2.大类,StatRule2.对方户名,StatRule2.关键字]
FinalResultRule = StatRule2.reindex(columns = ['收支','最终结果']).drop_duplicates()
FinalResultRule.index = FinalResultRule['收支']
FinalResultRule = FinalResultRule.reindex(columns = ['最终结果'])
FinalResultIncome = FinalResultRule.ix['收入',:]
FinalResultPayment = FinalResultRule.ix['支出',:]
FinalDetail = DataFrame([NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA],index = ['交易日期', '交易时间', '收入', '本币收入','收支类型', '大类', '对方户名', '子类', '分类结果',
       '银行名称', '账户类型', '币种']).T
FinalBalance = DataFrame([NA,NA,NA,NA,NA,NA],index = ['交易日期','余额', '本币余额', '银行名称', '账户类型', '币种']).T
#FinalSummary = DataFrame([NA,NA,NA,NA,NA,NA,NA,NA],index = ['原币收入', '本币收入','银行名称', '账户类型', '币种', '交易日期', '收支类型', '分类结果']).T



Writer = pd.ExcelWriter(ResultPath  + '/分类结果.xlsx')
os.chdir(DataPath)

class StatFileClass:
    global FinalSummary
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
        if '待核查' in temp:
            self.CountType = '待核查'
        elif '专户' in temp:
            self.CountType = '专户'
        else:
            self.CountType = '一般户'
        self.DateLable = StatRule1.ix[self.BankName,'交易日期字段']
        self.TimeLable = StatRule1.ix[self.BankName,'交易时间字段']
        self.TimeFormat = StatRule1.ix[self.BankName,'时间格式'].split(',')
        self.IncomeLable = StatRule1.ix[self.BankName,'收入字段']
        self.PayLable = StatRule1.ix[self.BankName,'支出字段']
        self.BalanceLable = StatRule1.ix[self.BankName,'当日余额字段']
        self.KeyLable1 = StatRule1.ix[self.BankName,'大类字段']
        self.KeyLable2 = StatRule1.ix[self.BankName,'子类字段'].split('+') #字符串list
        self.CountLable = StatRule1.ix[self.BankName,'户名字段']
        self.SkipRows = StatRule1.ix[self.BankName,'数据开始行']-1
        self.ERateUSD = StatRule1.ix[self.BankName,'美元汇率']
        self.ERateJPY = StatRule1.ix[self.BankName,'日元汇率']
        
        #判断收入支出类型
        if self.IncomeLable != self.PayLable: #非中国银行，将收入和支出合并
            self.RawData = pd.read_excel(self.FileName,skiprows = self.SkipRows,converters = {self.IncomeLable : str, self.PayLable : str,self.BalanceLable : str})
            self.IncomeData =  self.RawData[self.IncomeLable].astype(float).fillna(0)
            self.PayData =  self.RawData[self.PayLable].astype(float).fillna(0)
            self.IncomeData = self.IncomeData -  self.PayData
        else:
            self.RawData = pd.read_excel(self.FileName,skiprows = self.SkipRows,converters = {self.IncomeLable : str,self.BalanceLable : str})
            self.IncomeData =  self.RawData[self.IncomeLable].astype(float)
        self.IncomeType = Series(np.zeros(self.IncomeData.shape[0])) #初始化
        for i in range(self.IncomeData.shape[0]):
            if self.IncomeData[i] > 0:
                self.IncomeType[i] = '收入'
            else:
                self.IncomeType[i] = '支出'  
        
        #计算本币收入及余额
        self.IncomeDataLocal = self.IncomeData
        if self.Currency == 'USD':
            self.IncomeDataLocal = self.IncomeData * self.ERateUSD
        elif self.Currency == 'JPY':
            self.IncomeDataLocal = self.IncomeData * self.ERateJPY

                    
        #处理日期数据
        self.Date =  self.RawData[self.DateLable].astype(str)        
        self.Time =  self.RawData[self.TimeLable].astype(str) 
        if len(self.TimeFormat) == 3: # 如果有字符串长度参数则截取
            for i in self.Time.index:
                self.Time[i] = self.Time[i][int(self.TimeFormat[1])-1:int(self.TimeFormat[2])]
        for i in self.Time.index:
            self.Time[i] = datetime.strptime(self.Time[i],self.TimeFormat[0]).strftime('%H:%M:%S') #按照格式处理为时间数据,再转化为格式化的字符串        
        #获取大类数据
        self.KeyWord1 = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        if self.KeyLable1 == '无':
            self.KeyWord1[:] = '无'
        else:
            self.KeyWord1 = self.RawData[self.KeyLable1]
        
        #获取交易户名数据
        self.CountName = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        if self.CountLable == '无':
            self.CountName[:] = '无'
        else:
            self.CountData = self.RawData[self.CountLable]
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
        self.KeyWord2 = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        self.KeyData2 = self.RawData[self.KeyLable2]
        self.KeyData2.fillna(' ',inplace = True) #由于有多个关键字段，空值不能赋为'无'，而是空格
        self.ClassifyResult = Series(np.zeros(self.IncomeData.shape[0]))  #初始化
        for i in range(self.KeyData2.shape[0]):
            TempData = list(set(list(ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].ix[self.CountName[i]].index)))
            if len(TempData) == 1: #子类字段只有一个
                self.KeyWord2[i] = TempData[0]
            else:
                TempKeyWord = [m.split('+') for m in TempData]   #按分隔符分割关键字
                TempKeyWord.sort(key = lambda x:len(x),reverse = True) #按关键字个数排序；关键字越多，排序越靠前
                bFindResult = False
                for j in TempKeyWord:
                    bFindResult2 = True
                    for k in j:
                        if k not in str(list(self.KeyData2.ix[i])):   #只要有一个关键字不匹配，则放弃搜索该关键字           
                            bFindResult2 = False
                            break
                    if bFindResult2:    #全部关键字匹配，则认为匹配成功
                       self.KeyWord2[i] = '+'.join(j) #用+号重新连接为表中的关键字
                       bFindResult = True                      
                       break
                if not bFindResult:
                    self.KeyWord2[i] = '无'
            if self.KeyWord2[i] in TempData:                
                self.ClassifyResult[i] = ClassifyRuleDF.ix[self.BankName].ix[self.IncomeType[i]].ix[self.KeyWord1[i]].ix[self.CountName[i]].ix[self.KeyWord2[i]]
                if type(self.ClassifyResult[i]) != str: #如果出现多个分类结果，取第一个;
                    self.ClassifyResult[i] = self.ClassifyResult[i].ix[0]
            else:
                self.ClassifyResult[i] = '分类错误'

#==============================================================================
#         #产生当日分类汇总
#         self.Summary = self.IncomeData.copy()
#         self.Summary.index = [self.Date,self.IncomeType,self.ClassifyResult]
#         TempIndex = set(list(self.Summary.index)) #合并日期、收入类型、分类结果都相同的项
#         TempIncome = np.array([self.Summary.ix[i].sum() for i in TempIndex])
#         self.Summary = DataFrame(TempIncome, columns  = ['原币收入'])
#         TempIncomeLocal = TempIncome
#         if self.Currency == 'USD':
#             TempIncomeLocal = TempIncome * self.ERateUSD
#         elif self.Currency == 'JPY':
#             TempIncomeLocal *= TempIncome * self.ERateJPY
#         self.Summary['本币收入'] = TempIncomeLocal
#         self.Summary['银行名称'] = self.BankName
#         self.Summary['账户类型'] = self.CountType
#         self.Summary['币种'] = self.Currency
#         TempIndex = list(zip(*list(TempIndex)))
#         self.Summary['交易日期'] = TempIndex[0]
#         self.Summary['收支类型'] = TempIndex[1]
#         self.Summary['分类结果'] = TempIndex[2]
# 
#==============================================================================
        #结果输出
        self.ResultDF = pd.concat([self.Date,self.Time,self.IncomeData,self.IncomeDataLocal,self.IncomeType,self.KeyWord1,self.CountName,self.KeyWord2,self.ClassifyResult],axis = 1)
        self.ResultDF.columns = ['交易日期','交易时间','收入','本币收入','收支类型','大类','对方户名','子类','分类结果'] 
        self.ResultDF['银行名称'] = self.BankName
        self.ResultDF['账户类型'] = self.CountType
        self.ResultDF['币种'] = self.Currency
        
        #单日余额汇总
        #余额数据导入
        self.BalanceData =  self.RawData[self.BalanceLable]
        self.BalanceData = DataFrame([self.Date,self.BalanceData],index = ['交易日期','余额']).T
        self.BalanceData = self.BalanceData.groupby(['交易日期']).last() #取每天的最后一笔交易的余额数据
        self.BalanceData = self.BalanceData.applymap(RemoveComma) #对每个元素去除逗号
        self.BalanceData = self.BalanceData.astype(float)
        #计算本币余额
        self.BalanceDataLocal = self.BalanceData
        if self.Currency == 'USD':
            self.BalanceDataLocal = self.BalanceData * self.ERateUSD
        elif self.Currency == 'JPY':
            self.BalanceDataLocal = self.BalanceData * self.ERateJPY
        self.BalanceData['本币余额'] = self.BalanceDataLocal
        self.BalanceData['银行名称'] = self.BankName
        self.BalanceData['账户类型'] = self.CountType
        self.BalanceData['币种'] = self.Currency
        self.BalanceData['交易日期'] = self.BalanceData.index
        self.BalanceData.reindex(columns = ['交易日期','余额', '本币余额', '银行名称', '账户类型', '币种'])
            
        #self.ResultFileName = ResultPath + '/' +self.FileNameHead + '分类结果.xlsx'       
        #self.ResultDF.to_excel(Writer,self.FileNameHead,index = False)

def ProcessFiles(DataPath,Writer):
    global FinalDetail,FinalSummary,FinalBalance
    for FileName in os.listdir(DataPath):
        if ('银行'  in FileName) and ('xls' in FileName):
            print('正在处理:' + FileName)
            TempClass = StatFileClass(FileName);  
            FinalDetail = pd.concat([FinalDetail,TempClass.ResultDF])
            FinalBalance = pd.concat([FinalBalance,TempClass.BalanceData])
            #FinalSummary = pd.concat([FinalSummary,TempClass.Summary])
    FinalDetail.dropna(how = 'all',inplace = True)   #去掉第一行
    FinalBalance.dropna(how = 'all',inplace = True)   #去掉第一行
    #FinalSummary.dropna(how = 'all',inplace = True)   #去掉第一行
    #DaySummary = FinalDetail.groupby(['交易日期','收支类型','分类结果'])['本币收入'].sum()
    DaySummary = FinalDetail.groupby(['交易日期','收支类型','分类结果']).sum()
    DaySummary.drop(['收入'],axis = 1,inplace = True)
    BankSummary = FinalDetail.groupby(['交易日期','账户类型','银行名称','币种']).sum()
    BalanceSummary = FinalBalance.groupby(['交易日期','账户类型','银行名称','币种']).sum()
    FinalDetail.to_excel(Writer,'明细汇总')
    FinalBalance.to_excel(Writer,'余额明细')
    #FinalSummary.to_excel(Writer,'分类汇总')
    DaySummary.to_excel(Writer,'当日汇总')
    BankSummary.to_excel(Writer,'银行汇总')
    BalanceSummary.to_excel(Writer,'余额汇总')
    Writer.save()
    print('处理完毕!')

#去除中国银行余额字符串中的逗号
def RemoveComma(sValue):
    return sValue.replace(',','')
 
    
ProcessFiles(DataPath,Writer)
#A = StatFileClass('中国银行.xls'); 
#A = StatFileClass('中国银行美元待核查户.xls'); 
#A = StatFileClass('农业银行.xls'); 
#B = A.Summary
#FinalSummary.to_excel(Writer,'分类汇总',header = False)
#Writer.save()
#==============================================================================
# A = StatFileClass('农业银行.xls'); 
# B = A.IncomeData.copy()
# B.index = [A.Date,A.IncomeType,A.ClassifyResult]
# C = set(list(B.index))
# D = [B.ix[i].sum() for i in C]
# E = DataFrame(D,index = C)
# E.ix[:,'银行名称'] = '农业银行'
#==============================================================================



