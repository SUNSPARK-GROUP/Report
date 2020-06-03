# -*- encoding: UTF-8 -*-
'''
def hello(request):
    return HttpResponse("Hello world ! ")'''
from django.http import HttpResponse
from django.shortcuts import render,HttpResponseRedirect
from datetime import date
from datetime import timedelta,datetime
import cx_Oracle
from siteapp.views import gettotaldata
from siteapp.views import getodate
from siteapp.views import mainmenu
from siteapp.views import submenu
from graphos.sources.simple import SimpleDataSource
from graphos.sources.model import ModelDataSource
from graphos.renderers.gchart import ColumnChart
from django.urls import reverse
import openpyxl
import win32com.client as win32
import pyodbc
import pymysql
import time
import datetime
from flask import request
from flask import jsonify

import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8' 

depts=[]
accls=[]

def cratetabs(tabc,recount):#每頁列數,資料筆數
  tabs=[]
  rel=0
  tc=1
  while rel<recount :
    
    if tc==1:
      tabs.append('<a href="#tab'+str(tc)+'" onclick="jsTabs(event,'+"'tab"+str(tc)+"')"+';return false" class="tabs-menu tabs-menu-active">'+'第 '+str(tc)+' 頁'+"</a>")
    else:
      tabs.append('<a href="#tab'+str(tc)+'" onclick="jsTabs(event,'+"'tab"+str(tc)+"')"+';return false" class="tabs-menu">'+'第 '+str(tc)+' 頁'+"</a>")
    
    #tabs.append('<a href="#tab'+str(tc)+'" onclick="jsTabs(event,'+"'tab"+str(tc)+"')"+';return false" class="tabs-menu">'+'第 '+str(tc)+' 頁'+"</a>")
    tc=tc+1
    rel=rel+tabc
  return tabs
  
def tabsdata(tabc,Sfl):#每頁列數,原資料list
  tfl=[]
  fl=[]
  tc=1
  for t in range(len(Sfl)):
    if tc % tabc !=0:
      tfl.append(Sfl[t])
    else:
      tfl.append(Sfl[t])
      fl.append(tfl)
      tfl=[]
    if tc==len(Sfl):
      fl.append(tfl)
    tc=tc+1
  return fl
def CONMYSQL(sqlstr):
 #  f=open(r'C:\Users\Administrator\Desktop\txtlog\CONMYSQL.txt','w')
  db = pymysql.connect(host='192.168.0.210', port=3306, user='apuser', passwd='0920799339', db='main_eipplus_standard',charset='utf8')
  
  cursor = db.cursor()
  cursor.execute(sqlstr)	

  result = cursor.fetchall()
  urls = [row[0] for row in result]
  return result
def CONORACLE(SqlStr):
  hostname='192.168.0.230'
  sid='E910'
  username='PRODDTA'
  #username='CRPDTA'
  password='E910Jde'
  port='1521'
  dsn = cx_Oracle.makedsn(hostname, port, sid)
  conn = cx_Oracle.connect(username+'/'+password+'@' + dsn,encoding='UTF-8')
  cursor = conn.cursor()
  cursor.execute(SqlStr)
  SQLSTRS = SqlStr[0:6].upper()
  if SQLSTRS=="SELECT":
    TotalSession = cursor.fetchall()
    #f.write('TotalSession')
    return TotalSession
    cursor.close()
  else: conn.commit()
def CONMYSQL218(sqlstr):
  #  f=open(r'C:\Users\Administrator\Desktop\txtlog\CONMYSQL.txt','w')
  db = pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  
  cursor = db.cursor()
  cursor.execute(sqlstr)	
  
  # result = cursor.fetchall()
  # urls = [row[0] for row in result]
  # return result

  SQLSTRS = sqlstr[0:6].upper()
  if SQLSTRS=="SELECT":
    TotalSession = cursor.fetchall()
    return TotalSession
    cursor.close()
  else: db.commit()
  #  f.close()
def CONMSSQL(sqlstr):
  connection211=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.211;DATABASE=SUNBERPS;UID=apuser;PWD=0920799339')
  cursor = connection211.cursor()
  cursor.execute(sqlstr)
  SQLSTRS = sqlstr[0:6].upper()
  if SQLSTRS=="SELECT":
    TotalSession = cursor.fetchall()
    return TotalSession
    cursor.close()
  else: connection211.commit()
def CONMSSQL214(sqlstr):
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339')
  cursor = connection214.cursor()
  cursor.execute(sqlstr)
  SQLSTRS = sqlstr[0:6].upper()
  if SQLSTRS=="SELECT":
    TotalSession = cursor.fetchall()
    return TotalSession
    cursor.close()
  else: connection214.commit()
def CONMSSQL206(sqlstr):
  connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
  cursor = connection206.cursor()
  cursor.execute(sqlstr)
  SQLSTRS = sqlstr[0:6].upper()
  if SQLSTRS=="SELECT":
    TotalSession = cursor.fetchall()
    return TotalSession
    cursor.close()
  else: connection206.commit()
def showday(wd,sp,dt):#wd->0 today,wd->1 yesterday,wd->-1 tomorrow sp->Divider dt datetype
  t1 = 0-wd
  d=date.today()-timedelta(t1)
  if dt==1 :
    return d.strftime('%d'+sp+'%m'+sp+'%Y')#%Y->2015 %y->15
  else :
    return d.strftime('%Y'+sp+'%m'+sp+'%d')#%Y->2015 %y->15
def GOART(request): #門市進貨單表頭
  f = open(r'C:\Users\Edward\Desktop\txtlog\GOART.txt','w')

  f.write(str(request))
  context= {}
  try:
    product=[]
    sday=request.GET['Sday']#get values
    eday=request.GET['Eday']
    shop=request.GET['shop01']
    context['shop01'] = shop
    shops=shop[:shop.find('_')]
    shopst=shop[shop.find('_')+1:]
    context['Sday'] = sday
    context['Eday'] = eday
    context['mess'] = shopst
    sd=(sday[:4]+sday[5:7]+sday[8:10])
    ed=(eday[:4]+eday[5:7]+eday[8:10])
    eday=request.GET['Eday']

    #f.write("SELECT [GO_NO],[CUST_NAME],convert(float,[FL_TOTO]) total,[sdate]  FROM [SUNBERPS].[dbo].[GOARTHD] where sa_no='"+shops+"' and sdate BETWEEN '"+sd+"' and  '"+ed+"' AND GO_NO LIKE 'PS%'")

    temp211 = CONMSSQL("SELECT [GO_NO],[CUST_NAME],convert(float,[FL_TOTO]) total,[sdate]  FROM [SUNBERPS].[dbo].[GOARTHD] where sa_no='"+shops+"' and sdate BETWEEN '"+sd+"' and  '"+ed+"' AND GO_NO LIKE 'PS%'")
    
    r=[]
    for d in temp211:
      r.append(d)
    #頁籤    
    
    context['tabs']=cratetabs(15,len(r))
    #頁籤
    #頁籤內容
    #f.write(str(tabsdata(15,r)))
    context['weborder']=tabsdata(15,r)
    #頁籤內容	  
    RPDCT=request.GET['RPDCT']
    context['sRPDCT']=RPDCT
    func=request.GET['funcname']
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    context['funcname'] = ''    

  sls=[]
  temp211 = CONMSSQL("SELECT [COMP_NO]+'_'+[COMP_NAME] as cnm  FROM [SUNBERPS].[dbo].[COMPANY] where AREACODE='web'")
  for a in temp211:
    sls.append(a[0])
  context['shop0']=sls

  return render(request, 'shop//GOART.html',context ,)#傳入參數
  f.close()
def GOARTlist(request):
  show =[]
  context={}
  f = open(r'C:\Users\Edward\Desktop\txtlog\GOARTlist.txt','w')
  GO_NO=request.GET['GO_NO']
   #----------------------------------------------------------------
   #表頭
  table_t=CONMSSQL("SELECT GO_NO,MO_NO,CUST_NO,CUST_NAME,SDATE,APDATE,FL_TOTO FROM [SUNBERPS].[dbo].[GOARTHD] WHERE GO_NO ='"+GO_NO+"'")
  for title in table_t:
    context['GO_NO']    = title[0]
    context['MO_NO']    = title[1]
    context['CUST_NO']  = title[2]
    context['CUST_NAME']= title[3]
    context['SDATE']    = title[4]
    context['APDATE']   = title[5]
    context['FL_TOTO']  = title[6]

   #----------------------------------------------------------------
  
   #明細
  table_s=CONMSSQL("SELECT ITEM_NO,ART_NO,ART_NAME,FQTY,UNITBUY,QTY,UNIT,SIN,IN_TOTO,CHK_OK FROM GOARTFL WHERE GO_NO ='"+GO_NO+"'")

  for data in table_s:
      t1=[]
      # t1.append(str(data[0]))
      t1.append(str(data[0]))
      t1.append(str(data[1]))
      t1.append(str(data[2]))
      t1.append(format(data[3],','))
      t1.append(str(data[4]))
      t1.append(format(data[5],','))
      t1.append(str(data[6]))
      t1.append(format(data[7],','))
      t1.append(format(data[8],','))
      t1.append(str(data[9]))
      show.append(t1)
  context['show']  = show

   #頁籤
  context['tabs']=cratetabs(5,len(show))
  
   #頁籤內容
  context['weborder']=tabsdata(5,show)
  
  return render(request, 'shop//GOARTlist.html',context)
  f.close()
def GOARTInsert(request):
  f = open(r'C:\Users\Edward\Desktop\txtlog\GOARTInsert.txt','w')
  f.write(str(request)+'\n')
  SerialNumber = '' #流水號
  GO_NO =''
  context= {}
  context['title']="新增"
    #門市資料
  sls=[]
  temp211 = CONMSSQL("SELECT [COMP_NO]+'_'+[COMP_NAME] FROM [SUNBERPS].[dbo].[COMPANY] where AREACODE='web'")
  for a in temp211:
    sls.append(a[0])
  context['shop0']=sls
   #進貨號碼
  temp211 = CONMSSQL("SELECT 'PS'+substring(YEAR,3,4)+NO FROM SerialNumber WHERE TABLE_NAME ='GOARTHD' AND YEAR = DATEPART(year,GETDATE())")
  for Number in temp211:
    context['GO_NO']= "進貨單"
    SerialNumber = Number[0]
   #廠商資料
  vender = []
  temp211 =CONMSSQL("SELECT CUST_NO+'_'+CUST_NAME FROM MAKHD")
  for s in temp211:
    vender.append(s[0])
    context['vender0']= vender 
  try:
    Sday   = request.GET['Sday']
    Eday   = request.GET['Eday']
    MO_NO  = request.GET['MO_NO']
    vender = request.GET['vender']
    shop   = request.GET['shop']
    ART_NO = request.GET['ART_NO']
    status = request.GET['status']
    GO_NOITEM   = request.GET['GO_NO']
    
     #進貨單沒值就給她
    if str(GO_NOITEM) == '進貨單' :
      context['GO_NO']= SerialNumber
      GO_NO = SerialNumber
    else:
      context['GO_NO']= GO_NOITEM
      GO_NO = GO_NOITEM
    #廠商號碼空的就帶入進貨號碼
    if str(MO_NO) == '':
      context['MO_NO']= GO_NO
    else:
      context['MO_NO']= MO_NO

    context['Sday'] = Sday
    context['Eday'] = Eday    
    
    context['vender'] = vender
    context['shop'] = shop
    context['ART_NO'] = ART_NO

    ARTNO = ART_NO[:ART_NO.find('_')]
    CUST_NO = vender[:vender.find('_')]
     #廠商商品
    sls=[]
    temp211 =CONMSSQL("SELECT ART_NO+'_'+ART_NAME,UNITUSEQ,UNIT,UNITQ,UNITUSE,SIN FROM ART WHERE CUST_NO ='"+ CUST_NO +"'")
    for a in temp211:
      sls.append(a[0])
    context['ART_NO0']=sls

     #重新跑廠商 清除明細資料
    if status == 'venderList': 
      temp211 = CONMSSQL("DELETE TEMP_ART ")
      
     #取得下拉資料後新增到TABLE
    
    
    if status == "add" :#如果是add才執行 新增動作
      temp211 =CONMSSQL("INSERT INTO TEMP_ART SELECT (SELECT ISNULL(MAX(ITEM)+1,1) FROM TEMP_ART),ART_NO,ART_NAME,0 UNITUSEQ,UNITSTK,0 UNITQ,UNIT,SIN FROM ART WHERE  ART_NO = '"+ ARTNO +"'")

    if status == "alter" :#如果是alter才執行 修改動作
      UNITUSEQ = request.GET['UNITUSEQ']
      list  = request.GET['list']
      ITEM = list[:list.find(',')]
      ARTNO = list[list.find(',')+1:]
       #計算進貨量
      #f.write(str("SELECT (UNITQ/(CASE WHEN QTY IS NULL THEN 1 ELSE QTY END)*"+UNITUSEQ+") XX,ART_NO,ART_NAME,CASE WHEN QTY IS NULL THEN 1 ELSE QTY END QTY ,UNIT,UNITQ,UNITSTK FROM( SELECT ART_NO,ART_NAME, (select SUBSTRING(UNIT,1,1) QTY From ART X Where UNIT  Like '%[0-9]%' AND X.ART_NO = A.ART_NO AND X.UNIT <> X.UNITSTK) QTY, UNIT,UNITQ,UNITSTK FROM ART A ) A WHERE ART_NO ='"+ARTNO+"' "))
      temp211 =CONMSSQL("SELECT (UNITQ/(CASE WHEN QTY IS NULL THEN 1 ELSE QTY END)*"+UNITUSEQ+") XX,ART_NO,ART_NAME,CASE WHEN QTY IS NULL THEN 1 ELSE QTY END QTY ,UNIT,UNITQ,UNITSTK FROM( SELECT ART_NO,ART_NAME, (select SUBSTRING(UNIT,1,1) QTY From ART X Where UNIT  Like '%[0-9]%' AND X.ART_NO = A.ART_NO AND X.UNIT <> X.UNITSTK) QTY, UNIT,UNITQ,UNITSTK FROM ART A ) A WHERE ART_NO ='"+ARTNO+"' ")
      for unitq in temp211:
        UNITQ =unitq[0]
        CONMSSQL("UPDATE TEMP_ART  SET UNITUSEQ = '"+str(UNITUSEQ)+"',UNITQ ='"+str(UNITQ)+"' WHERE ITEM ='"+str(ITEM)+"'")  
     
    if status == "delete" :#如果是delete才執行 刪除動作
      ITEM  = request.GET['ITEM']
      CONMSSQL("DELETE TEMP_ART WHERE ITEM ='"+str(ITEM)+"'")
       #重新計算項次
      CONMSSQL("UPDATE TEMP_ART SET TEMP_ART.ITEM = ROW_ITEM.NEW_ITEM FROM (SELECT ROW_NUMBER() OVER(ORDER BY ITEM ASC) NEW_ITEM,ITEM FROM TEMP_ART) ROW_ITEM  WHERE ROW_ITEM.ITEM = TEMP_ART.ITEM")
    
    #抓取資料
    temp211 =CONMSSQL("SELECT ITEM,ART_NO,ART_NAME,UNITUSEQ,UNITUSE,UNITQ,UNIT,SIN,ROUND(UNITQ*SIN,5)  FROM TEMP_ART ORDER BY ITEM ASC")
    r=[]
    total =[]
    for artdata in temp211:
      r.append(artdata)
      total.append(artdata[8])
    context['r'] = r
     #貨單金額
    FL_TOTO = round(sum(total),5)
    context['FL_TOTO'] = FL_TOTO

    if status == "insert": #如果是insert執行  到資料庫
        Sday = Sday.replace('-','')
        Eday = Eday.replace('-','')
        shopID = shop[:shop.find('_')]
        shopName = shop[shop.find('_')+1:]
        #門市名稱過常會INSERT錯誤 如果過長 就改抓
        CUST_NO = vender[:vender.find('_')]
        CUST_NAME = vender[vender.find('_')+1:]
        NewDateTime =time.strftime("%Y%m%d-%H:%M", time.localtime()) 
        NewTime =time.strftime("%H:%M", time.localtime()) 
        
         #避免 重複新增 先刪除已經進入 資料庫的資料
         
        # f.write("DELETE FROM GOARTHD WHERE GO_NO ='"+str(GO_NO)+"' \n")
        CONMSSQL("DELETE FROM GOARTHD WHERE GO_NO ='"+str(GO_NO)+"'")
        # f.write("DELETE FROM GOARTFL WHERE GO_NO ='"+str(GO_NO)+"' \n")
        CONMSSQL("DELETE FROM GOARTFL WHERE GO_NO ='"+str(GO_NO)+"'")
         #表頭
        #f.write("INSERT INTO  [SUNBERPS].[dbo].[GOARTHD] (MO_NO,GO_NO,BUY_NO,SDATE,CUST_NO,CUST_NAME,FL_TOTO,OUTOK,SA_NO,USER_APPE,USER_MODI,APPE_DATE,MODI_DATE,FLAG_1,APDATE,NEWCUST_NO) VALUES ('"+str(MO_NO)+"','"+str(GO_NO)+"','"+str(GO_NO)+"','"+str(Sday)+"','"+str(CUST_NO)+"','"+str(CUST_NAME)+"','"+str(FL_TOTO)+"','1','"+str(shopID)+"','"+str(shopName)+"','"+str(shopName)+"','"+str(NewDateTime)+"','"+str(NewDateTime)+"','','"+str(Eday)+"','"+str(CUST_NO)+"') \n")

        CONMSSQL("INSERT INTO  [SUNBERPS].[dbo].[GOARTHD] (MO_NO,GO_NO,BUY_NO,SDATE,CUST_NO,CUST_NAME,FL_TOTO,OUTOK,SA_NO,USER_APPE,USER_MODI,APPE_DATE,MODI_DATE,FLAG_1,APDATE,NEWCUST_NO) VALUES ('"+str(MO_NO)+"','"+str(GO_NO)+"','"+str(GO_NO)+"','"+str(Sday)+"','"+str(CUST_NO)+"','"+str(CUST_NAME)+"','"+str(FL_TOTO)+"','1','"+str(shopID)+"','"+str(shopName)+"','"+str(shopName)+"','"+str(NewDateTime)+"','"+str(NewDateTime)+"','','"+str(Eday)+"','"+str(CUST_NO)+"')")
         #明細
        #f.write("INSERT GOARTFL (GO_NO,ITEM_NO,CUST_NO,SDATE,STIME,ART_NO,ART_NAME,FQTY,UNIT,QTY,UNITBUY,SIN,IN_TOTO,CHK_OK,WHERE_DBF,APPE_DATE) SELECT '"+str(GO_NO)+"', RIGHT(REPLICATE('0', 3) + CAST(ITEM as NVARCHAR),3) ITEM , '"+str(CUST_NO)+"', '"+str(Sday)+"', '"+str(NewTime)+"', ART_NO, ART_NAME, UNITUSEQ, UNIT, UNITQ, UNITUSE, SIN, ROUND(UNITQ*SIN,5), ''CHK_OK, '"+str(shopID)+"', '"+str(NewDateTime)+"' FROM TEMP_ART \n")

        CONMSSQL("INSERT GOARTFL (GO_NO,ITEM_NO,CUST_NO,SDATE,STIME,ART_NO,ART_NAME,FQTY,UNIT,QTY,UNITBUY,SIN,IN_TOTO,CHK_OK,WHERE_DBF,APPE_DATE) SELECT '"+str(GO_NO)+"', RIGHT(REPLICATE('0', 3) + CAST(ITEM as NVARCHAR),3) ITEM , '"+str(CUST_NO)+"', '"+str(Sday)+"', '"+str(NewTime)+"', ART_NO, ART_NAME, UNITUSEQ, UNIT, UNITQ, UNITUSE, SIN, ROUND(UNITQ*SIN,5), ''CHK_OK, '"+str(shopID)+"', '"+str(NewDateTime)+"' FROM TEMP_ART")
         #進貨號碼更新
        PSNumber = GO_NO[-5:]
        #f.write("UPDATE SerialNumber SET NO = RIGHT(REPLICATE('0', 5) + CAST(CONVERT(int,'"+PSNumber+"')+1 as NVARCHAR), 5) WHERE TABLE_NAME ='GOARTHD' AND YEAR = DATEPART(year,GETDATE()) \n")
        CONMSSQL("UPDATE SerialNumber SET NO = RIGHT(REPLICATE('0', 5) + CAST(CONVERT(int,'"+PSNumber+"')+1 as NVARCHAR), 5) WHERE TABLE_NAME ='GOARTHD' AND YEAR = DATEPART(year,GETDATE())")
        
        context['mess'] = "進貨號碼: "+str(GO_NO)+" 新增完畢"
        


  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday
    context['Eday'] = nday

  f.close()
  return render(request, 'shop//GOARTInsert.html',context)
def GOARTAlter(request):
  f = open(r'C:\Users\Edward\Desktop\txtlog\GOARTAlter.txt','w')
  f.write(str(request)+'\n')
  SerialNumber = '' #流水號
  GO_NO = request.GET['GO_NO']
  context= {}
  context['title']="修改"
  context['GO_NO'] = GO_NO
  # shop = ''
  # vender =''
   #修改前 表頭資料
  temp211 = CONMSSQL("SELECT GO_NO,MO_NO,CUST_NO+'_'+CUST_NAME,SUBSTRING(SDATE,1,4)+'-'+SUBSTRING(SDATE,5,2)+'-'+SUBSTRING(SDATE,7,2),SUBSTRING(APDATE,1,4)+'-'+SUBSTRING(APDATE,5,2)+'-'+SUBSTRING(APDATE,7,2),FL_TOTO ,COMP_NO+'_'+COMP_NAME FROM GOARTHD A LEFT JOIN COMPANY B ON B.COMP_NO = A.SA_NO WHERE GO_NO ='"+GO_NO+"'")
  #f.write(str(temp211)+'\n')
  for dataList in temp211:
    context['MO_NO'] = dataList[1] #廠商號碼
    vender = dataList[2]
    context['CUST'] = vender #廠商名稱
    context['Sday'] = dataList[3] #進貨日期
    context['Eday'] = dataList[4] #帳款日期
    context['FL_TOTO'] = dataList[5] #貨單金額
    shop = dataList[6]
    context['shop'] = shop  #門市資料
    #廠商商品
  CUST_NO = vender[:vender.find('_')]
  sls=[]
  temp211 =CONMSSQL("SELECT ART_NO+'_'+ART_NAME,UNITUSEQ,UNIT,UNITQ,UNITUSE,SIN FROM ART WHERE CUST_NO ='"+ CUST_NO +"'")
  for a in temp211:
    sls.append(a[0])
  context['ART_NO0']=sls  
   #廠商資料
  vender = []
  temp211 =CONMSSQL("SELECT CUST_NO+'_'+CUST_NAME FROM MAKHD")
  for s in temp211:
    vender.append(s[0])
    context['vender0']= vender 
  try:
    Sday   = request.GET['Sday']
    Eday   = request.GET['Eday']
    MO_NO  = request.GET['MO_NO']
    vender = request.GET['vender']
    ART_NO = request.GET['ART_NO']
    status = request.GET['status']
    GO_NOITEM   = request.GET['GO_NO']
    

    context['Sday'] = Sday
    context['Eday'] = Eday    
    
    context['MO_NO'] = MO_NO
    context['vender'] = vender
    context['ART_NO'] = ART_NO

    ARTNO = ART_NO[:ART_NO.find('_')]
    CUST_NO = vender[:vender.find('_')]
     #廠商商品
    sls=[]
    # f.write(str("SELECT ART_NO+'_'+ART_NAME,UNITUSEQ,UNIT,UNITQ,UNITUSE,SIN FROM ART WHERE CUST_NO ='"+ CUST_NO +"'"))
    temp211 =CONMSSQL("SELECT ART_NO+'_'+ART_NAME,UNITUSEQ,UNIT,UNITQ,UNITUSE,SIN FROM ART WHERE CUST_NO ='"+ CUST_NO +"'")
    for a in temp211:
      sls.append(a[0])
    context['ART_NO0']=sls

     #重新跑廠商 清除明細資料
    if status == 'venderList': 
      temp211 = CONMSSQL("DELETE TEMP_ART ")
      
     #取得下拉資料後新增到TABLE
    
    
    if status == "add" :#如果是add才執行 新增動作
      temp211 =CONMSSQL("INSERT INTO TEMP_ART SELECT (SELECT ISNULL(MAX(ITEM)+1,1) FROM TEMP_ART),ART_NO,ART_NAME,0 UNITUSEQ,UNITSTK,0 UNITQ,UNIT,SIN FROM ART WHERE  ART_NO = '"+ ARTNO +"'")

    if status == "alter" :#如果是alter才執行 修改動作
      UNITUSEQ = request.GET['UNITUSEQ']
      list  = request.GET['list']
      ITEM = list[:list.find(',')]
      ARTNO = list[list.find(',')+1:]
       #計算進貨量
      #f.write(str("SELECT (UNITQ/(CASE WHEN QTY IS NULL THEN 1 ELSE QTY END)*"+UNITUSEQ+") XX,ART_NO,ART_NAME,CASE WHEN QTY IS NULL THEN 1 ELSE QTY END QTY ,UNIT,UNITQ,UNITSTK FROM( SELECT ART_NO,ART_NAME, (select SUBSTRING(UNIT,1,1) QTY From ART X Where UNIT  Like '%[0-9]%' AND X.ART_NO = A.ART_NO AND X.UNIT <> X.UNITSTK) QTY, UNIT,UNITQ,UNITSTK FROM ART A ) A WHERE ART_NO ='"+ARTNO+"' "))
      temp211 =CONMSSQL("SELECT (UNITQ/(CASE WHEN QTY IS NULL THEN 1 ELSE QTY END)*"+UNITUSEQ+") XX,ART_NO,ART_NAME,CASE WHEN QTY IS NULL THEN 1 ELSE QTY END QTY ,UNIT,UNITQ,UNITSTK FROM( SELECT ART_NO,ART_NAME, (select SUBSTRING(UNIT,1,1) QTY From ART X Where UNIT  Like '%[0-9]%' AND X.ART_NO = A.ART_NO AND X.UNIT <> X.UNITSTK) QTY, UNIT,UNITQ,UNITSTK FROM ART A ) A WHERE ART_NO ='"+ARTNO+"' ")
      for unitq in temp211:
        UNITQ =unitq[0]
        #f.write(str("UPDATE TEMP_ART  SET UNITUSEQ = '"+str(UNITUSEQ)+"',UNITQ ='"+str(UNITQ)+"' WHERE ITEM ='"+str(ITEM)+"'"))
        CONMSSQL("UPDATE TEMP_ART  SET UNITUSEQ = '"+str(UNITUSEQ)+"',UNITQ ='"+str(UNITQ)+"' WHERE ITEM ='"+str(ITEM)+"'")  
     
    if status == "delete" :#如果是delete才執行 刪除動作
      ITEM  = request.GET['ITEM']
      CONMSSQL("DELETE TEMP_ART WHERE ITEM ='"+str(ITEM)+"'")
       #重新計算項次
      CONMSSQL("UPDATE TEMP_ART SET TEMP_ART.ITEM = ROW_ITEM.NEW_ITEM FROM (SELECT ROW_NUMBER() OVER(ORDER BY ITEM ASC) NEW_ITEM,ITEM FROM TEMP_ART) ROW_ITEM  WHERE ROW_ITEM.ITEM = TEMP_ART.ITEM")
    
    #抓取資料
    temp211 =CONMSSQL("SELECT ITEM,ART_NO,ART_NAME,UNITUSEQ,UNITUSE,UNITQ,UNIT,SIN,ROUND(UNITQ*SIN,5)  FROM TEMP_ART ORDER BY ITEM ASC")
    r=[]
    total =[]
    for artdata in temp211:
      r.append(artdata)
      total.append(artdata[8])
    context['r'] = r
     #貨單金額
    FL_TOTO = round(sum(total),5)
    context['FL_TOTO'] = FL_TOTO

    if status == "insert": #如果是insert執行  到資料庫
        Sday = Sday.replace('-','')
        Eday = Eday.replace('-','')
        shopID = shop[:shop.find('_')]
        shopName = shop[shop.find('_')+1:]
        #門市名稱過常會INSERT錯誤 如果過長 就改抓
        CUST_NO = vender[:vender.find('_')]
        CUST_NAME = vender[vender.find('_')+1:]
        NewDateTime =time.strftime("%Y%m%d-%H:%M", time.localtime()) 
        NewTime =time.strftime("%H:%M", time.localtime()) 
        
         #避免 重複新增 先刪除已經進入 資料庫的資料
         
        # f.write("DELETE FROM GOARTHD WHERE GO_NO ='"+str(GO_NO)+"' \n")
        CONMSSQL("DELETE FROM GOARTHD WHERE GO_NO ='"+str(GO_NO)+"'")
        # f.write("DELETE FROM GOARTFL WHERE GO_NO ='"+str(GO_NO)+"' \n")
        CONMSSQL("DELETE FROM GOARTFL WHERE GO_NO ='"+str(GO_NO)+"'")
         #表頭
        #f.write("INSERT INTO  [SUNBERPS].[dbo].[GOARTHD] (MO_NO,GO_NO,BUY_NO,SDATE,CUST_NO,CUST_NAME,FL_TOTO,OUTOK,SA_NO,USER_APPE,USER_MODI,APPE_DATE,MODI_DATE,FLAG_1,APDATE,NEWCUST_NO) VALUES ('"+str(MO_NO)+"','"+str(GO_NO)+"','"+str(GO_NO)+"','"+str(Sday)+"','"+str(CUST_NO)+"','"+str(CUST_NAME)+"','"+str(FL_TOTO)+"','1','"+str(shopID)+"','"+str(shopName)+"','"+str(shopName)+"','"+str(NewDateTime)+"','"+str(NewDateTime)+"','','"+str(Eday)+"','"+str(CUST_NO)+"') \n")

        CONMSSQL("INSERT INTO  [SUNBERPS].[dbo].[GOARTHD] (MO_NO,GO_NO,BUY_NO,SDATE,CUST_NO,CUST_NAME,FL_TOTO,OUTOK,SA_NO,USER_APPE,USER_MODI,APPE_DATE,MODI_DATE,FLAG_1,APDATE,NEWCUST_NO) VALUES ('"+str(MO_NO)+"','"+str(GO_NO)+"','"+str(GO_NO)+"','"+str(Sday)+"','"+str(CUST_NO)+"','"+str(CUST_NAME)+"','"+str(FL_TOTO)+"','1','"+str(shopID)+"','"+str(shopName)+"','"+str(shopName)+"','"+str(NewDateTime)+"','"+str(NewDateTime)+"','','"+str(Eday)+"','"+str(CUST_NO)+"')")
         #明細
        #f.write("INSERT GOARTFL (GO_NO,ITEM_NO,CUST_NO,SDATE,STIME,ART_NO,ART_NAME,FQTY,UNIT,QTY,UNITBUY,SIN,IN_TOTO,CHK_OK,WHERE_DBF,APPE_DATE) SELECT '"+str(GO_NO)+"', RIGHT(REPLICATE('0', 3) + CAST(ITEM as NVARCHAR),3) ITEM , '"+str(CUST_NO)+"', '"+str(Sday)+"', '"+str(NewTime)+"', ART_NO, ART_NAME, UNITUSEQ, UNIT, UNITQ, UNITUSE, SIN, ROUND(UNITQ*SIN,5), ''CHK_OK, '"+str(shopID)+"', '"+str(NewDateTime)+"' FROM TEMP_ART \n")

        CONMSSQL("INSERT GOARTFL (GO_NO,ITEM_NO,CUST_NO,SDATE,STIME,ART_NO,ART_NAME,FQTY,UNIT,QTY,UNITBUY,SIN,IN_TOTO,CHK_OK,WHERE_DBF,APPE_DATE) SELECT '"+str(GO_NO)+"', RIGHT(REPLICATE('0', 3) + CAST(ITEM as NVARCHAR),3) ITEM , '"+str(CUST_NO)+"', '"+str(Sday)+"', '"+str(NewTime)+"', ART_NO, ART_NAME, UNITUSEQ, UNIT, UNITQ, UNITUSE, SIN, ROUND(UNITQ*SIN,5), ''CHK_OK, '"+str(shopID)+"', '"+str(NewDateTime)+"' FROM TEMP_ART")
         #進貨號碼更新
        PSNumber = GO_NO[-5:]
        #f.write("UPDATE SerialNumber SET NO = RIGHT(REPLICATE('0', 5) + CAST(CONVERT(int,'"+PSNumber+"')+1 as NVARCHAR), 5) WHERE TABLE_NAME ='GOARTHD' AND YEAR = DATEPART(year,GETDATE()) \n")
        CONMSSQL("UPDATE SerialNumber SET NO = RIGHT(REPLICATE('0', 5) + CAST(CONVERT(int,'"+PSNumber+"')+1 as NVARCHAR), 5) WHERE TABLE_NAME ='GOARTHD' AND YEAR = DATEPART(year,GETDATE())")
        
        context['mess'] = "進貨號碼: "+str(GO_NO)+" 修改完畢"

  except:
    #修改前 明細資料
    #先清除 TEMP的資料   
    CONMSSQL("DELETE TEMP_ART ")
    CONMSSQL("INSERT INTO TEMP_ART SELECT CONVERT(int,ITEM_NO),ART_NO,ART_NAME,FQTY,UNIT,QTY,UNITBUY,SIN FROM GOARTFL WHERE GO_NO = '"+ GO_NO +"'")
    #抓取資料
    temp211 =CONMSSQL("SELECT ITEM,ART_NO,ART_NAME,UNITUSEQ,UNITUSE,UNITQ,UNIT,SIN,ROUND(UNITQ*SIN,5)  FROM TEMP_ART ORDER BY ITEM ASC")
    r=[]
    total =[]
    for artdata in temp211:
      r.append(artdata)
      total.append(artdata[8])
    context['r'] = r 
  
    #  #門市資料
    # sls=[]
    # temp211 = CONMSSQL("SELECT [COMP_NO]+'_'+[COMP_NAME] FROM [SUNBERPS].[dbo].[COMPANY] where AREACODE='web'")
    # for a in temp211:
    #   sls.append(a[0])
    # context['shop0']=sls
  f.close()
  return render(request, 'shop//GOARTAlter.html',context)
def SALEAREA(request): #營業區資料
  # f = open(r'C:\Users\Edward\Desktop\txtlog\SALEAREA.txt','w')
  # f.write(str(request)+'\n')

  context= {}
  Bigarea =[] #大區
  Marea =[]   #中區
  Sarea =[]   #小區
  Shop = []   # 門市
  # 大區資料
  Bigareadata = CONMSSQL214("SELECT DISTINCT SUBSTRING(AREANO,1,1)+' | '+SUBSTRING(AREANAME,1,1) FROM SALESAREA WHERE SUBSTRING(AREANO,1,1) IN ('M','N','S')")
  for barea in Bigareadata:
    Bigarea.append(barea[0])
    context['area0'] = Bigarea
  try :
    getbarea = request.GET['barea']
    context['barea'] = getbarea
    getmarea = request.GET['marea']
    context['marea'] = getmarea
    
    barea = request.GET['barea'] #取得大區資料
    barea_NO = barea[:barea.find(' |')]

    marea_NO = getmarea[:getmarea.find(' |')] 
    marea_NAME = getmarea[getmarea.find(' |')+2:]
    context['AREA_NAME'] = marea_NAME.strip()
    sareadata = CONMSSQL214("SELECT SUBAREANO,AREANAME FROM SALESAREAFL WHERE AREANO ='"+marea_NO+"' ORDER BY SUBAREANO")
    for sarea in sareadata:
      sa = []
      sa.append(sarea[0])
      sa.append(sarea[1])
      Sarea.append(sa)
      context['area2'] = Sarea

    # 負責人
    pldata = CONMSSQL214("SELECT EMPNAME1,EMPNAME2 FROM SALESAREA WHERE AREANO ='"+marea_NO+"'")
    for pl in pldata:
      context['principal1'] = pl[0]
      context['principal2'] = pl[1]     

    status = request.GET['status']
    if status =='Alter':
      context['status'] = "Alter"
      getprincipal1 = request.GET['principal1'].strip()
      context['principal1'] = getprincipal1
      getprincipal2 = request.GET['principal2'].strip()
      context['principal2'] = getprincipal2
      AREA_NAME = request.GET['AREA_NAME'].strip()
      context['AREA_NAME'] = AREA_NAME 

      # f.write("UPDATE SALESAREA SET AREANAME ='"+str(AREA_NAME)+"',EMPNAME1 = '"+str(getprincipal1)+"',EMPNAME2 = '"+str(getprincipal2)+"' WHERE AREANO = '"+str(marea_NO)+"'")
      CONMSSQL214("UPDATE SALESAREA SET AREANAME ='"+str(AREA_NAME)+"',EMPNAME1 = '"+str(getprincipal1)+"',EMPNAME2 = '"+str(getprincipal2)+"' WHERE AREANO = '"+str(marea_NO)+"'")
      context['mess'] = "變更完成"

      #變更營業中區
      mareadata = CONMSSQL214("SELECT AREANO+' | '+AREANAME FROM SALESAREA WHERE SUBSTRING(AREANO,1,1) = '"+str(barea_NO)+"' ORDER BY AREANO,AREANAME")
      for marea in mareadata:
        Marea.append(marea[0])
        context['area1'] = Marea
      
      # f.write(str(status)+'\n')
      newMeara = str(marea_NO) +" | "+ str(AREA_NAME)
      # f.write(str(newMeara))
      context['newMeara'] = newMeara
    else:
      mareadata = CONMSSQL214("SELECT AREANO+' | '+AREANAME FROM SALESAREA WHERE SUBSTRING(AREANO,1,1) = '"+str(barea_NO)+"' ORDER BY AREANO,AREANAME")
      for marea in mareadata:
        Marea.append(marea[0])
        context['area1'] = Marea

  except:
    s =''
  # f.close()
  return render(request, 'shop//SALEAREA.html',context)
def AREASHOP(request): #營業區門市
  # f = open(r'C:\Users\Edward\Desktop\txtlog\AREASHOP.txt','w')
  # f.write(str(request)+'\n')
  context= {}
  Shop =[]
  getsareaNO = request.GET['sareaNO']
  #小區資料
  sareadata = CONMSSQL214("SELECT SUBAREANO,AREANAME FROM SALESAREAFL WHERE SUBAREANO = '"+getsareaNO+"'")
  for sdata in sareadata :
    context['sareaNO'] = sdata[0]
    context['sareaNAME'] = sdata[1]
  try:
    status = request.GET['status']
    
    
    #加入
    if status == "in":
      selected = request.GET['selected']
      # f.write(str("UPDATE basicstoreinfo SET STRING_500_3  = '"+str(getsareaNO)+"' WHERE String_20_1 IN ("+str(selected)+")")+'\n')
      CONMYSQL218("UPDATE basicstoreinfo SET STRING_500_3  = '"+str(getsareaNO)+"' WHERE String_20_1 IN ("+str(selected)+")")
      status = "Checkbox1:true"
      context['mess'] = "加入到" + str(getsareaNO)

    #移出
    if status == "out":
      selected = request.GET['selected']
      # f.write(str("UPDATE basicstoreinfo SET STRING_500_3  = NULL WHERE String_20_1 IN ("+str(selected)+")")+'\n')
      CONMYSQL218("UPDATE basicstoreinfo SET STRING_500_3  = NULL WHERE String_20_1 IN ("+str(selected)+")")
      context['mess'] = str(getsareaNO) + "已移出"

    if status == "Checkbox1:true":
      context['ck'] = 'true'
      #未分配的門市
      shopdata = CONMYSQL218("SELECT STRING_20_1,STRING_50_1  FROM basicstoreinfo WHERE STRING_500_3 IS NULL")
      for shdata in shopdata:
        sh = []
        sh.append(shdata[0])
        sh.append(shdata[1])
        Shop.append(sh)
        context['shop'] = Shop
    else:   
      #門市資料
      # f.write(str("SELECT STRING_20_1,STRING_50_1 FROM basicstoreinfo WHERE STRING_500_3 ='"+getsareaNO+"'"))
      shopdata = CONMYSQL218("SELECT STRING_20_1,STRING_50_1 FROM basicstoreinfo WHERE STRING_500_3 ='"+getsareaNO+"'")
      for shdata in shopdata :
        sh = []
        sh.append(shdata[0])
        sh.append(shdata[1])
        Shop.append(sh)
        context['shop'] = Shop
        # f.write(str('OK'))
    
  except:
    x=''
  # f.close()
  return render(request, 'shop//AREASHOP.html',context)
