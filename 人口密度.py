# -*- coding: utf-8 -*-
"""
Created on Fri Aug 28 10:56:25 2020

@author: 123lab
"""

import os 

import xlrd 

import xlwt 

import csv 


  
#第一區先整理人口資料
Folder1=r'D:\網路管理\爬蟲\opendata-10308_age_cityname'   #資料檔案夾記得改 
os.chdir(Folder1) #修改工作目錄  

read=[]
dataall=[]
File_list1= os.listdir() #列檔案之長度 
for i in range(len(File_list1)) :
    file1 =open(Folder1+'\\'+File_list1[i],'r',encoding="utf-8")
    #↑開檔轉文字轉UTF-8開CSV ，每個檔案放置資料夾不同
    read=csv.DictReader(file1)    #CSV字典功能，讀取需要的列欄 
    data=[ [col['鄉鎮別村里'],col['人口數']] for col in read] #讀取列欄之語法
    dataall+=data   #對每個data開啟檔案疊加

village=[o[0] for o in dataall] 
people=[p[1] for p in dataall] 


#第二區整理面積資料
Folder2=r'D:\網路管理\爬蟲\村面積'   #資料檔案夾記得改
os.chdir(Folder2) #修改工作目錄
File_list2= os.listdir() #列檔案之長度
file2=xlrd.open_workbook(Folder2+'\\'+File_list2[0],'r',encoding_override = "utf-8") 
#↑開檔轉文字轉UTF-8開excel寫法 ，每個檔案放置資料夾不同 
table=file2.sheet_by_index(0)
#用索引取第一個sheet 

a=table.row_values(0)
#抓第一行當列表頭

findit1=a.index('TOWNNAME')
findit2=a.index('VILLNAME')
findit3=a.index('TOWNNAMEVILLNAME')
findit4=a.index('Shape_Area')
#↑上方為尋找表頭，利用表頭回傳欄位

find1=table.col_values(findit1)
find2=table.col_values(findit2)
find3=table.col_values(findit3)
find4=table.col_values(findit4)
#對回傳欄位進行抓取欄位資料

b=[]
for j in range(len(find1)):
    c=[find1[j],find2[j],find3[j],find4[j]]
    b.append(c)
#將資料整理成一個LIST


#第三區整理村里資料
Folder3=r'D:\網路管理\爬蟲\村里屆'   #資料檔案夾記得改
os.chdir(Folder3) #修改工作目錄
File_list3= os.listdir() #列檔案之長度
file3=xlrd.open_workbook(Folder3+'\\'+File_list3[0],'r',encoding_override = "utf-8")  
#↑開檔轉文字轉UTF-8開excel寫法 ，每個檔案放置資料夾不同 
sheet=file3.sheet_by_index(0)
#用索引取第一個sheet 

title=sheet.row_values(0)
main1=title.index('TOWNNAME')
main2=title.index('VILLNAME')
main3=title.index('TOWNNAMEVILLNAME')
#↑上方為尋找表頭，利用表頭回傳欄位

m1=sheet.col_values(main3)
m=m1[1:]
#上方行關鍵字LIST

z=[]

Save_xls=xlwt.Workbook(encoding='utf-8')#Workbook指令大寫 
sheet1=Save_xls.add_sheet("人口面積",cell_overwrite_ok=True) 

for k in range(len(m)):#關鍵字當作成長度起始FOR迴圈
    key=m[k]
    word1=m[k]               #本行回應第三區資料
    
    matsu1=find3.index(key)  #在LIST(find3)當中尋找關鍵字
    word2=find3[matsu1]      #本行回應第二區資料
    word3=find4[matsu1]      #本行回應第二區資料
    
    matsu2=village.index(key)  #在LIST(village)當中尋找關鍵字
    word4=village[matsu2]
    word5=people[matsu2]
    
    wordall=[word1,word2,word3,word4,word5]
    z.append(wordall)
    
    sheet1.write(k,1,word1) 
    sheet1.write(k,2,word2)
    sheet1.write(k,3,word3)
    sheet1.write(k,4,word4)
    sheet1.write(k,5,word5)

Save_xls.save(r'D:\網路管理\test.xls')   #儲存語法  