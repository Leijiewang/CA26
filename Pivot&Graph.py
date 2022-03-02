# -*- coding: utf-8 -*-
"""
Created on Fri Feb 25 18:41:33 2022

@author: ljwang
"""
from matplotlib.ticker import MultipleLocator
import math
import pandas as pd
import matplotlib.pyplot as plt
#讓圖可以顯示繁體字
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False
#自由帶入圖片設計類型
plt.style.use('ggplot')
strsize = 22 
two = pd.read_excel(r".\inputdata\109年+110年事故資料(整合).xlsx", dtype = {" 發生時間" : str})
def ToHour(string):
    str1 = string.zfill(6)
    str1 = str1[0:2]
    if str1 == "00":
        str1 = "24"
    return str1
two["發生時段"] = two[" 發生時間"].apply(lambda x : ToHour(x))
two.rename(columns={'總編號(案件編號)': '案件數'}, inplace=True)
#刪除重複項目
one = two.loc[two[' 當事者順位'] == 1]
print(f"選擇第一當事人後表一剩餘:{len(one)}筆")
one = one.drop_duplicates(subset=['案件數']).reset_index()
print(f"刪除重複案件編號後表一剩餘:{len(one)}筆")
#選出109年資料
one = one.loc[one[' 發生年度'] == 2020].reset_index()
print(f"選出109年事件表一剩餘:{len(one)}筆")
#選出A1A2資料
one = one.loc[one[' 事故類別名稱'] != "A3"]
one.index = range(len(one))
print(f"選出A1A2事故表一剩餘:{len(one)}筆")

#樞紐分析-分月
month = one.pivot_table(index = " 發生月份", values = "案件數", aggfunc = "count")
month[' 發生月份'] = month.index
month = month.reindex(columns=[' 發生月份', "案件數"])
print("月份樞紐分析完成")
#樞紐分析-分時
hour = one.pivot_table(index = "發生時段", values = "案件數", aggfunc = "count")
hour['發生時段'] = hour.index
hour = hour.reindex(columns=['發生時段', "案件數"])
print("時段樞紐分析完成")
plt.figure(figsize=(20,8))
plt.plot(hour["發生時段"], hour["案件數"], marker = "o")
plt.ylabel("案件數", fontsize=strsize, color = 'k') # y label
plt.xlabel("發生時段", fontsize=strsize, color = 'k') # x label
plt.xticks(fontsize=strsize, color = 'k')
plt.yticks(fontsize=strsize, color = 'k')
for a,b in zip(hour["發生時段"], hour["案件數"]):
    plt.text(a, b+0.3, '%.0f' % b, ha='center', va= 'bottom',fontsize=strsize-4)
plt.savefig('.\outputdata\月份.png', dpi=1000, bbox_inches = "tight")
plt.show()
print("時段作圖完成")
#樞紐分析-肇因
cause = one.pivot_table(index = " 肇因研判子類別名稱-主要", values = "案件數", aggfunc = "count")
cause[' 肇因研判子類別名稱-主要'] = cause.index
cause = cause.reindex(columns=[" 肇因研判子類別名稱-主要", "案件數"])

total = cause["案件數"].sum()
def ratio(counting):
    str1 = round((counting/total)*100,2)
    return str1
cause["佔比"] = cause["案件數"].apply(lambda x : ratio(x))
cause = cause.sort_values(by='案件數', ascending = False)
cause = cause[0:6]
print("肇因樞紐分析完成")
#樞紐分析-年齡
age = two.pivot_table(index = " 當事者事故發生時年齡", values = "案件數", aggfunc = "count")
age[' 當事者事故發生時年齡'] = age.index
age = age.reindex(columns=[' 當事者事故發生時年齡', "案件數"])
age = age.loc[age[' 當事者事故發生時年齡'] >=0]
age = age.loc[age[' 當事者事故發生時年齡'] <=100]
print("年齡樞紐分析完成")
plt.figure(figsize=(20,8))
plt.fill_between(age[' 當事者事故發生時年齡'], age["案件數"])
plt.ylabel("案件數", fontsize=strsize, color = 'k') # y label
plt.xlabel("年齡", fontsize=strsize, color = 'k') # x label
plt.xticks(rotation='vertical')
x_major_locator = MultipleLocator(4)
ax = plt.gca()
ax.xaxis.set_major_locator(x_major_locator)
plt.xlim(0,100)
plt.ylim(0, max(age["案件數"])+500)
plt.axvline(18, 0, max(age["案件數"]), label='18歲', color = "k")
plt.axvline(24, 0, max(age["案件數"]), label='24歲', color = "g")
plt.axvline(65, 0, max(age["案件數"]), label='65歲', color = "y")
plt.legend(fontsize=strsize)
plt.xticks(fontsize=strsize, color = 'k')
plt.yticks(fontsize=strsize, color = 'k')
plt.savefig('.\outputdata\年齡.png', dpi=1000, bbox_inches = "tight")
plt.show()
print("年齡作圖完成")
#樞紐分析-鄉鎮
town = one.pivot_table(index = " 發生市區鄉鎮名稱", values = "案件數", aggfunc = "count")
town[' 發生市區鄉鎮名稱'] = town.index
town = town.reindex(columns=[' 發生市區鄉鎮名稱', "案件數"])
town = town.sort_values(by='案件數', ascending = False)
print("鄉鎮樞紐分析完成")
plt.figure(figsize=(20,8))
plt.bar(town[" 發生市區鄉鎮名稱"],
        town['案件數'], 
        width = 0.5, 
        bottom = None, 
        align = 'center',
        label='案件數'
        )
plt.ylabel("案件數", fontsize=strsize, color = 'k') # y label
plt.xlabel("鄉鎮", fontsize=strsize, color = 'k') # x label
plt.xticks(fontsize=strsize, rotation='vertical', color = 'k')
plt.yticks(fontsize=strsize, color = 'k')
plt.ylim(0, max(town['案件數'])+0.2) 
for a,b in zip(town[" 發生市區鄉鎮名稱"],town['案件數']):
    plt.text(a, b+0.01, '%.0f' % b, ha='center', va= 'bottom',fontsize=strsize)
plt.savefig('.\outputdata\鄉鎮.png', dpi=1000, bbox_inches = "tight")
plt.show()
print("鄉鎮作圖完成")

# color = "mistyrose"

#輸出成excel
with pd.ExcelWriter('.\outputdata\表一分析檔.xlsx') as writer:
    month.to_excel(writer, sheet_name='月份', index = False)
    hour.to_excel(writer, sheet_name='時段', index = False)
    cause.to_excel(writer, sheet_name='肇因', index = False)
    age.to_excel(writer, sheet_name='年齡', index = False)
    town.to_excel(writer, sheet_name='鄉鎮', index = False)