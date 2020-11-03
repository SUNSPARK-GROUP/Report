# -*- encoding: UTF-8 -*-
'''
def hello(request):
    return HttpResponse("Hello world ! ")'''
from django.http import HttpResponse
from django.shortcuts import render,HttpResponseRedirect
from datetime import date
from datetime import timedelta
import cx_Oracle
from siteapp.views import gettotaldata
from siteapp.views import getodate
from siteapp.views import mainmenu
from siteapp.views import submenu
from graphos.sources.simple import SimpleDataSource
from graphos.sources.model import ModelDataSource
from graphos.renderers.gchart import ColumnChart
from django.urls import reverse
from openpyxl import Workbook
import openpyxl
import win32com.client as win32
import pyodbc

import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8' 

depts=[]
accls=[]
global prodls 
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
def showday(wd,sp,dt):#wd->0 today,wd->1 yesterday,wd->-1 tomorrow sp->Divider dt datetype
  t1 = 0-wd
  d=date.today()-timedelta(t1)
  if dt==1 :
    return d.strftime('%d'+sp+'%m'+sp+'%Y')#%Y->2015 %y->15
  else :
    return d.strftime('%Y'+sp+'%m'+sp+'%d')#%Y->2015 %y->15
def cper (slist,sno,n):#slist原data sno指定欄位list,n分母數在原data取值欄位
  for tt in range(len(sno)) :        
    if slist[sno[tt]-1]==0:
      slist[sno[tt]]=0
    else:
      slist[sno[tt]]=round(slist[sno[tt]-1]/slist[n]*100,1)       
      #f.write(str(slist[sno[tt]-1])+'/'+ str(slist[sno[tt]])+'\n')
  return(slist)
def layashopdata(request):
  context= {}
  connection211=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=POSSA;UID=apuser;PWD=0920799339') 
  temp211=connection211.cursor() 
  uno=format(request.COOKIES['userno'])  
  try:
    f=open(r'C:\Users\chris\chrisdjango\shoperror.txt','w')
    product=[]
    sday=request.GET['Sday']#get values
    eday=request.GET['Eday']
    context['Sday'] = sday
    context['Eday'] = eday
    '''
    sd=(sday[:4]+sday[5:7]+sday[8:10])
    ed=(eday[:4]+eday[5:7]+eday[8:10])
    '''
    f.write(sday+'/'+eday+'\n')
    sd=sday
    ed=eday
    f.write(sd+'/'+ed+'\n')
    shop=request.GET['shop01']
    context['shop01'] = shop
    shop=shop[:shop.find('_')]
    try:
      ck1=request.GET['Checkbox1']
      context['ck'] ='ck1'
    except:
      ck1='off'
    try:
      ck2=request.GET['Checkbox2']
      context['ck'] ='ck2'
    except:
      ck2='off'
    try:
      ck3=request.GET['Checkbox3']
      context['ck'] ='ck3'
    except:
      ck3='off'
    saledata=[]#web data
    saledatae=[]#excel data
    wb = Workbook()	
    ws1 = wb.active	
    ws1.title = "data"
    #營業額
    if ck1=='on':
      ws1.append(['拉亞直營營業額('+sday+'~'+eday+')'])
      ws1.append(['日期','星期','門市','營業額','內用筆數','內用金額','佔比','外帶筆數','外帶金額','佔比','電話筆數','電話金額','佔比','門市外送筆數','門市外送金額','佔比','Uber筆數','Uber金額','佔比','熊貓筆數','熊貓金額','佔比','其他筆數','其他金額','佔比','筆數小計','平均客單價'])
      title=['<th style="width:60px;"  align="center" >日期</th>','<th style="width:60px;"  align="center" >星期</th>','<th style="width:112px;">門市</th>','<th style="width:40px;">營業額</th>','<th style="width:40px;">內用筆數</th>','<th style="width:40px;">內用金額</th>','<th style="width:40px;">佔比</th>'
             ,'<th style="width:40px;">外帶筆數</th>','<th style="width:40px;" >外帶金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">電話筆數</th>','<th style="width:40px;">電話金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">外送筆數</th>','<th style="width:40px;">外送金額</th>','<th style="width:40px;">佔比</th>'
			 ,'<th style="width:40px;">Uber筆數</th>','<th style="width:40px;">Uber金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">熊貓筆數</th>','<th style="width:40px;">熊貓金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">其他筆數</th>','<th style="width:40px;">其他金額</th>','<th style="width:40px;">佔比</th>'
			 ,'<th style="width:40px;">筆數小計</th>','<th style="width:40px;">平均客單價</th>']
      saletype=['內用','外帶','電話','外送','Uber','foodpanda','其他']
      saledataq=[]#日期、門市、筆數小計、金額小計
      tsale=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]#各欄總計暫存
      if shop=='全部門市':
        shop="and h.sa_no like 'LA%'"
      else:
        shop="and h.sa_no='"+shop+"'"
      for s in range(len(saletype)): 
        
        if 	saletype[s]=='Uber':
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  ")
          '''		
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.COMP_NAME,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent]='"+saletype[s]+"'" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.COMP_NAME order by c.COMP_NAME,sdate")
          '''
        elif 	saletype[s]=='foodpanda':
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('FoodPanda','熊貓')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('FoodPanda','熊貓')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  ")
          
        elif 	saletype[s]=='其他':
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3   and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3   and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  ")
          
        else:
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,SNAME,count(go_no) as cs ,sum(price) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,SNAME  order by sdate,SNAME  ")
        f.write('temp211'+'\n')
          
        for d in temp211.fetchall():
          if s==0:		  
            saledata.append(['<td style="width:60px;"  align="center">'+str(d[0])+'</td>','<td style="width:60px;"  align="center">'+str(d[1]).replace('星期','')+'</td>','<td style="width:112px;">'+str(d[2])+'</td>','<td style="width:40px;" align="right">0</td>'
                            ,'<td style="width:40px;" align="center" >'+str(d[3])+'</td>','<td style="width:40px;" align="right">'+format(int(d[4]),',')+'</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
                            ,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>'])
            f.write(str(s)+'\n')
            saledatae.append([str(d[0]),str(d[1]).replace('星期',''),d[2],'0',d[3],d[4],0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])
            tsale[1]=int(tsale[1])+int(d[3])
            tsale[2]=int(tsale[2])+int(d[4])
            saledataq.append([str(d[0]),str(d[2]),int((d[3])),int(d[4])])
            f.write(str(d[0])+'/'+str(d[2])+'/'+str(d[4])+'\n')
          
          else:
            for i in range(len(saledataq)):
              if str(d[0])==saledataq[i][0] and str(d[2])==saledataq[i][1] :
                saledataq[i][2]=saledataq[i][2]+int(d[3])
                saledataq[i][3]=saledataq[i][3]+int(d[4])
                #f.write('saledata-len:'+str(len(saledata))+str(saledata))
                saledata[i][s*3+4]='<td style="width:40px;" align="center">'+str(d[3])+'</td>'#因每類訂單有數量與金額二筆數字故 s*2 再因跳過表格前4欄再 +4
                saledata[i][s*3+5]='<td style="width:40px;" align="right">'+format(int(d[4]),',')+'</td>'
                saledatae[i][s*3+4]=d[3]
                saledatae[i][s*3+5]=d[4]                
                tsale[s*3+1]=tsale[s*3+1]+int(d[3])
                tsale[s*3+2]=tsale[s*3+2]+int(d[4])
      #f.write('tempf'+'\n')
      for i in range(len(saledataq)):
        saledata[i][3]='<td style="width:40px;" align="right">'+format(int(saledataq[i][3]),',')+'</td>'#該日總營業額
        saledata[i][25]='<td style="width:40px;" align="center">'+format(int(saledataq[i][2]),',')+'</td>'#訂單數
        saledata[i][26]='<td style="width:40px;" align="right">'+format(round(saledataq[i][3]/saledataq[i][2]),',')+'</td>'#客單價
        saledatae[i][3]=saledataq[i][3]
        saledatae[i][25]=saledataq[i][2]
        saledatae[i][26]=round(saledataq[i][3]/saledataq[i][2])
        tsale[0]=tsale[0]+saledataq[i][3]
        tsale[len(tsale)-2]=tsale[len(tsale)-2]+saledataq[i][2]#訂單數
        #tsale[len(tsale)-1]=tsale[len(tsale)-1]+int(saledataq[i][3])
      #f.write(str(saledataq)+'\n')
      tsale[len(tsale)-1]=round(tsale[0]/tsale[len(tsale)-2])#客單價
      tg=[3,6,9,12,15,18,21]
      #f.write(str(tsale))
      tsale=cper(tsale,tg,0)
      #f.write(str(tsale))
      saledata.append(['<td style="width:60px;"></td>','<td style="width:60px;"></td>','<td style="width:112px;">小計</td>','<td style="width:40px;" align="right">'+format(tsale[0],',')+'</td>'
                       ,'<td style="width:40px;" align="center" >'+str(tsale[1])+'</td>','<td style="width:40px;" align="right">'+format(tsale[2],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[3])+'</td>'
					   ,'<td style="width:40px;" align="center">'+str(tsale[4])+'</td>','<td style="width:40px;" align="right">'+format(tsale[5],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[6])+'</td>'
                       ,'<td style="width:40px;" align="center">'+str(tsale[7])+'</td>','<td style="width:40px;" align="right">'+format(tsale[8],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[9])+'</td>'
					   ,'<td style="width:40px;" align="center">'+str(tsale[10])+'</td>','<td style="width:40px;" align="right">'+format(tsale[11],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[12])+'</td>'
                       ,'<td style="width:40px;" align="center">'+str(tsale[13])+'</td>','<td style="width:40px;" align="right">'+format(tsale[14],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[15])+'</td>'
					   ,'<td style="width:40px;" align="center">'+format(int(tsale[16]),',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[17],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[18])+'</td>'
                       ,'<td style="width:40px;" align="center">'+format(tsale[19],',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[20],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[21])+'</td>'
					   ,'<td style="width:40px;" align="center">'+format(tsale[22],',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[23],',')+'</td>'])
      
      try:
        saledatae.append(['','','小計',tsale[0],tsale[1],tsale[2],tsale[3],tsale[4],tsale[5],tsale[6],tsale[7],tsale[8],tsale[9],tsale[10],tsale[11],tsale[12],tsale[13],tsale[14],tsale[15],tsale[16],tsale[17],tsale[18],tsale[19],tsale[20],tsale[21],tsale[22],tsale[23]])
        '''
        saledatae.append(['','','小計',tsale[0],tsale[1],tsale[2],round(tsale[2]/tsale[0]*100,1),tsale[3],tsale[4],round(tsale[4]/tsale[0]*100,1),tsale[5],tsale[6],round(tsale[6]/tsale[0]*100,1)
		           ,tsale[7],tsale[8],round(tsale[8]/tsale[0]*100,1),tsale[9],sale[10],round(tsale[10]/tsale[0]*100,1),tsale[11],tsale[12],round(tsale[12]/tsale[0]*100,1),tsale[13],tsale[14]])
        '''
      except Exception as e:  
        f.write(e)	  
      
      #f.write('test'+'\n')
      for e in range(len(saledatae)):
        tg=[6,9,12,15,18,21,24]
        saledatae[e]=cper(saledatae[e],tg,3)        
        for ht in range(len(tg)):
          saledata[e][tg[ht]]='<td style="width:40px;" align="center">'+str(saledatae[e][tg[ht]])+'</td>'        
        ws1.append(saledatae[e])
      	  
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx'
    #銷售數量
    if ck2=='on':
      ws1.append(['拉亞直營銷售數量('+sday+'~'+eday+')'])
      ws1.append(['門市','大類','名稱','商品編號','商品名稱','銷售數量'])
      title=['<th style="width:10%;">門市</th>','<th style="width:10%;">大類</th>','<th style="width:10%;">名稱</th>','<th style="width:20%;">商品編號</th>','<th style="width:40%;">商品名稱</th>'
	           ,'<th style="width:20%;">銷售數量</th>']
      if shop=='全部門市':
        shop=''
      else:
        shop="and d.sa_no='"+shop+"'"
      f.write("SELECT c.SNAME ,[Category],F.NAME ,[Product_ID] ,[Product_Name] ,sum(convert(int,[Quantity])) as qty  FROM [POSSA].[dbo].[SHOP_detail] d ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].[FKIND] F "
                       +" where SA_NO like 'la%' and c.[SID]=d.sa_no  and d.invoice_date >= '"+sd+" 00:00:00' and d.invoice_date <='"+ed+" 23:59:59'"+shop+" AND f.ID=d.[Category]   group by  c.SNAME ,[Category] ,[Product_ID] ,[Product_Name] ,F.NAME"
					   +"order by [Category]")
      temp211.execute("SELECT c.SNAME ,[Category],F.NAME ,[Product_ID] ,[Product_Name] ,sum(convert(int,[Quantity])) as qty  FROM [POSSA].[dbo].[SHOP_detail] d ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].[FKIND] F "
                       +" where SA_NO like 'la%' and c.[SID]=d.sa_no  and d.invoice_date >= '"+sd+" 00:00:00' and d.invoice_date <='"+ed+" 23:59:59'"+shop+" AND f.ID=d.[Category]   group by  c.SNAME ,[Category] ,[Product_ID] ,[Product_Name] ,F.NAME "
					   +"order by [Category]")
      
      for d in temp211.fetchall():
        saledata.append(['<td style="width:10%;">'+str(d[0])+'</td>','<td style="width:10%;">'+str(d[1])+'</td>','<td style="width:10%;">'+str(d[2])+'</td>','<td style="width:20%;">'+str(d[3])+'</td>','<td style="width:40%;" >'+str(d[4])+'</td>'
                          ,'<td style="width:20%;"  align="right">'+str(d[5])+'</td>'])
        ws1.append([str(d[0]),str(d[1]),str(d[2]),str(d[3]),str(d[4]),str(d[5])])
      
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_直營銷售數量('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_直營銷售數量('+sday+'~'+eday+').xlsx'
    #時段金額
    if ck3=='on':
      ws1.append(['拉亞直營時段營業額('+sday+'~'+eday+')'])
      ws1.append(['門市','時段','營業額','內用筆數','內用金額','佔比','外帶筆數','外帶金額','佔比','電話筆數','電話金額','佔比','門市外送筆數','門市外送金額','佔比','Uber筆數','Uber金額','佔比','熊貓筆數','熊貓金額','佔比','其他筆數','其他金額','佔比','筆數小計','平均客單價'])
      title=['<th style="width:112px;">門市</th>','<th style="width:60px;"  align="center" >時段</th>','<th style="width:40px;">營業額</th>','<th style="width:40px;">內用筆數</th>','<th style="width:40px;">內用金額</th>','<th style="width:40px;">佔比</th>'
             ,'<th style="width:40px;">外帶筆數</th>','<th style="width:40px;" >外帶金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">電話筆數</th>','<th style="width:40px;">電話金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">外送筆數</th>','<th style="width:40px;">外送金額</th>','<th style="width:40px;">佔比</th>'
			 ,'<th style="width:40px;">Uber筆數</th>','<th style="width:40px;">Uber金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">熊貓筆數</th>','<th style="width:40px;">熊貓金額</th>','<th style="width:40px;">佔比</th>','<th style="width:40px;">其他筆數</th>','<th style="width:40px;">其他金額</th>','<th style="width:40px;">佔比</th>'
			 ,'<th style="width:40px;">筆數小計</th>','<th style="width:40px;">平均客單價</th>']
      saletype=['內用','外帶','電話','外送','Uber','foodpanda','其他']
      saledataq=[]#日期、門市、筆數小計、金額小計
      tsale=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]#各欄總計暫存
      if shop=='全部門市':
        shop="and h.sa_no like 'LA%'"
      else:
        shop="and h.sa_no='"+shop+"'"
      for s in range(len(saletype)): 
        f.write(saletype[s]+'\n')
        if 	saletype[s]=='Uber':
          f.write("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) "+'\n')
          temp211.execute("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) ")
          '''		
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.COMP_NAME,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent]='"+saletype[s]+"'" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.COMP_NAME order by c.COMP_NAME,sdate")
          '''
        elif 	saletype[s]=='foodpanda':
          f.write("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('FoodPanda','熊貓')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) "+'\n')
          temp211.execute("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments]  in ('FoodPanda','熊貓')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) ")
          
        elif 	saletype[s]=='其他':
          f.write("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3   and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2)  "+'\n')
          temp211.execute("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3   and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2)  ")
          
        else:
          f.write("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) "+'\n')
          temp211.execute("select SNAME,substring(convert(varchar, invoice_date,108),1,2) as times,count(go_no) as cs ,sum(price) as ts from (SELECT h.invoice_date,c.SNAME,h.go_no ,p.price,p.[payments] " 
                         +" FROM [POSSA].[dbo].[SHOP_orders] h ,[POSSA].[dbo].[SHOP] c,[POSSA].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[SID]=h.sa_no and h.sa_no like 'la%' and p.[payments] not  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no and p.invoice_date=h.invoice_date  and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by substring(convert(varchar, invoice_date,108),1,2),SNAME  order by SNAME ,substring(convert(varchar, invoice_date,108),1,2) ")
        f.write('temp211'+'\n')
          
        for d in temp211.fetchall():
          if s==0:		  
            saledata.append(['<td style="width:60px;"  align="center">'+str(d[0])+'</td>','<td style="width:60px;"  align="center">'+str(d[1])+'</td>','<td style="width:40px;" align="right">0</td>'
                            ,'<td style="width:40px;" align="center" >'+str(d[2])+'</td>','<td style="width:40px;" align="right">'+format(int(d[3]),',')+'</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
                            ,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>','<td style="width:40px;" align="right">0</td>'
							,'<td style="width:40px;" align="center">0</td>','<td style="width:40px;" align="right">0</td>'])
            #f.write(str(s)+'\n')
            saledatae.append([str(d[0]),str(d[1]),'0',d[2],int(d[3]),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0])
            #f.write(str(saledatae))
            tsale[1]=int(tsale[1])+int(d[2])
            tsale[2]=int(tsale[2])+int(d[3])
            saledataq.append([str(d[0]),str(d[1]),int((d[2])),int(d[3])])
            #f.write(str(d[0])+'/'+str(d[1])+'/'+str(d[3])+'\n')
          
          else:
            for i in range(len(saledataq)):
              if str(d[0])==saledataq[i][0] and str(d[1])==saledataq[i][1] :
                saledataq[i][2]=saledataq[i][2]+int(d[2])
                saledataq[i][3]=saledataq[i][3]+int(d[3])
                #f.write('saledata-len:'+str(len(saledata))+str(saledata))
                saledata[i][s*3+3]='<td style="width:40px;" align="center">'+str(d[2])+'</td>'#因每類訂單有數量與金額二筆數字故 s*2 再因跳過表格前3欄再 +3
                saledata[i][s*3+4]='<td style="width:40px;" align="right">'+format(int(d[3]),',')+'</td>'
                saledatae[i][s*3+3]=int(d[2])
                saledatae[i][s*3+4]=int(d[3])                
                tsale[s*3+1]=tsale[s*3+1]+int(d[2])
                tsale[s*3+2]=tsale[s*3+2]+int(d[3])
      
      	  
      for i in range(len(saledataq)):
        saledata[i][2]='<td style="width:40px;" align="right">'+format(int(saledataq[i][3]),',')+'</td>'#該時段總營業額
        saledata[i][24]='<td style="width:40px;" align="center">'+format(int(saledataq[i][2]),',')+'</td>'#訂單數
        saledata[i][25]='<td style="width:40px;" align="right">'+format(round(saledataq[i][3]/saledataq[i][2]),',')+'</td>'#客單價
        saledatae[i][2]=saledataq[i][3]
        saledatae[i][24]=saledataq[i][2]
        saledatae[i][25]=round(saledataq[i][3]/saledataq[i][2])
        tsale[0]=tsale[0]+saledataq[i][3]
        tsale[len(tsale)-2]=tsale[len(tsale)-2]+saledataq[i][2]#訂單數
        #tsale[len(tsale)-1]=tsale[len(tsale)-1]+int(saledataq[i][3])
      #f.write(str(saledataq)+'\n')
      tsale[len(tsale)-1]=round(tsale[0]/tsale[len(tsale)-2])#客單價
      tg=[3,6,9,12,15,18,21]
      #f.write(str(tsale))
      
      
      '''
      tsale=cper(tsale,tg,0)
      #f.write(str(tsale))
      saledata.append(['<td style="width:60px;"></td>','<td style="width:60px;"></td>','<td style="width:112px;">小計</td>','<td style="width:40px;" align="right">'+format(tsale[0],',')+'</td>'
                       ,'<td style="width:40px;" align="center" >'+str(tsale[1])+'</td>','<td style="width:40px;" align="right">'+format(tsale[2],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[3])+'</td>'
					   ,'<td style="width:40px;" align="center">'+str(tsale[4])+'</td>','<td style="width:40px;" align="right">'+format(tsale[5],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[6])+'</td>'
                       ,'<td style="width:40px;" align="center">'+str(tsale[7])+'</td>','<td style="width:40px;" align="right">'+format(tsale[8],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[9])+'</td>'
					   ,'<td style="width:40px;" align="center">'+str(tsale[10])+'</td>','<td style="width:40px;" align="right">'+format(tsale[11],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[12])+'</td>'
                       ,'<td style="width:40px;" align="center">'+str(tsale[13])+'</td>','<td style="width:40px;" align="right">'+format(tsale[14],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[15])+'</td>'
					   ,'<td style="width:40px;" align="center">'+format(int(tsale[16]),',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[17],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[18])+'</td>'
                       ,'<td style="width:40px;" align="center">'+format(tsale[19],',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[20],',')+'</td>','<td style="width:40px;" align="right">'+str(tsale[21])+'</td>'
					   ,'<td style="width:40px;" align="center">'+format(tsale[22],',')+'</td>','<td style="width:40px;" align="right">'+format(tsale[23],',')+'</td>'])
      
      try:
        saledatae.append(['','','小計',tsale[0],tsale[1],tsale[2],tsale[3],tsale[4],tsale[5],tsale[6],tsale[7],tsale[8],tsale[9],tsale[10],tsale[11],tsale[12],tsale[13],tsale[14],tsale[15],tsale[16],tsale[17],tsale[18],tsale[19],tsale[20],tsale[21],tsale[22],tsale[23]])
        
      except Exception as e:  
        f.write(e)	  
      '''
      #f.write(str(saledatae)+'\n')
      for e in range(len(saledatae)):#計算佔比
        tg=[5,8,11,14,17,20,23]
        #f.write(str(saledatae[e])+'\n')
        saledatae[e]=cper(saledatae[e],tg,2)
        #f.write(str(saledatae[e])+'\n')        
        for ht in range(len(tg)):
          saledata[e][tg[ht]]='<td style="width:40px;" align="center">'+str(saledatae[e][tg[ht]])+'</td>'        
        ws1.append(saledatae[e])
      #f.write(str(saledatae)+'\n')  
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_拉亞直營時段營業額('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_拉亞直營時段營業額('+sday+'~'+eday+').xlsx'


	
    context['sdata'] = saledata				
    context['title'] = title
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    context['funcname'] = ''    
  sls=[]
  temp211.execute("SELECT [SID]+'_'+[SNAME] FROM [POSSA].[dbo].[SHOP]  where SID like 'LA%' ")
  sls.append('全部門市_')
  for a in temp211:
    sls.append(a[0])
  context['shop0']=sls
  return render(request, 'shop//layashopdata.html',context )#傳入參數
  f.close()
def OracleSalesCalc(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
  area=connection214.cursor()
  #f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  try:
    sc=request.GET['sc']
    sday=request.GET['Sday']#get values
    eday=request.GET['Eday']
    context['Sday'] = sday
    context['Eday'] = eday
    sd=(sday[:4]+sday[5:7]+sday[8:10])
    ed=(eday[:4]+eday[5:7]+eday[8:10])
    context['CK1']= request.GET['CK1']
    ck1=request.GET['CK1']
    if ck1=='ON':
      firstr=" and  (shdel1 not like '%首批%' and shdel2  not like '%首批%') "
    else:
      firstr=""
    #f.write(firstr)
    context['area01']= request.GET['area01']
    area.execute("Select DISTINCT LEFT(AREANO,1)AREANO,LEFT(AREANAME,1) AREANAME from SalesArea where areano<>'TINO'")
    area0=[]
    for a in area.fetchall():
      area0.append(str(a[0])+'|'+str(a[1]))
    context['area0']= area0#大區list
    an1=request.GET['area01']#大區名稱
    area1=[]
    if an1!='':
      area.execute("Select AREANO+'|'+areaname from  SalesArea where Left(AreaNo,1)='"+an1[:an1.find('|')]+"'")        
      for a in area.fetchall():
        area1.append(str(a[0]))
    context['area1']= area1#中區list
    context['area11']= request.GET['area11']
    an11=request.GET['area11']#中區名稱
    area2=[]
    if an11!='':
      area.execute("SELECT  [SUBAREANO]+'|'+[AREANAME]  FROM [ERPS].[dbo].[SALESAREAFL] where AREANO='"+an11[:an11.find('|')]+"'")       
      for a in area.fetchall():
        area2.append(str(a[0]))
    context['area2']= area2#小區list
    context['area21']= request.GET['area21']
    an21=request.GET['area21']#小區名稱
    shop=[]
    if an21!='':
      area.execute("Select UserID+'|'+UserName from OPENQUERY(WEB206,'SELECT * FROM UserGroup where  AreaNo=''"+an21[:an21.find('|')]+"''')")      
      for a in area.fetchall():
        shop.append(str(a[0]))
    context['shop0']= shop#門市list
    shopname=request.GET['shop'][:request.GET['shop'].find('|')]#門市名稱 
    context['shop01']=request.GET['shop']		
    if sc=='1':
      osaledata=[]
      if shopname!='':  
        title=['<th style="width:10%;">門市代碼</th>','<th style="width:15%;">門市名稱</th>','<th style="width:10%;">負責業務</th>','<th style="width:20%;">產品名稱</th>','<th style="width:15%;">產品代碼</th>','<th style="width:20%;">銷售金額</th>','<th style="width:10%;">數量</th>']
        
        area.execute("select o.abalky as '門市代碼',o.ABALPH as '門市名稱',lv.EMPNAME1 as '負責業務',o.SDDSC1 as '產品名稱',o.SDLITM as '產品代碼',sum(o.DTTAEXP) as '銷售金額',SUM(o.qty) as '數量' from "
            +" (select distinct(USERID) as USERID,EMPNAME1 from LayaViewAreaUser  ) lv,(select h.shshan,h.abalky,h.ABALPH,f.SDLITM,f.sddsc1,CAST(f.DTTAEXP as decimal(18,0)) as DTTAEXP,CAST(f.qty as decimal(18,0)) as qty from "
            +" (select * from vs_salehd where    (shdcto='S2' OR shdcto='S3' OR shdcto='S4') and SHDRQJ >= '"+sd+"' and SHDRQJ <= '"+ed+"') h,(select * from vs_salefl where sdlitm>='100000' and sdlitm<='399999' "
			#20190927 取消 +" and sdlitm not in ('100004')" 
            +" and   (sddcto='S2' OR sddcto='S3' OR sddcto='S4') AND SDDRQJ >= '"+sd+"' and SDDRQJ <= '"+ed+"') f where f.sddoco=h.shdoco and f.sdan8=h.shshan "+firstr+"  ) o "
            +" where o.abalky=lv.USERID  and o.abalky='"+shopname+"' group by o.abalky,o.ABALPH,lv.EMPNAME1,o.SDLITM,o.SDDSC1  order by o.abalky ")      
        for a in area.fetchall():
          tl=[]
          tl.append('<td style="width:10%;">'+str(a[0])+'</td>')
          tl.append('<td style="width:15%;">'+str(a[1])+'</td>')
          tl.append('<td style="width:10%;">'+str(a[2])+'</td>')
          tl.append('<td style="width:20%;">'+str(a[3])+'</td>')
          tl.append('<td style="width:15%;">'+str(a[4])+'</td>')
          tl.append('<td style="width:20%;">'+format(float(a[5]),',')+'</td>')
          tl.append('<td style="width:10%;">'+str(a[6])+'</td>')	
          osaledata.append(tl)		  
      elif an21!='': 
        title=['<th style="width:100;">門市代碼</th>','<th style="width:150;">門市名稱</th>','<th style="width:100;">負責業務</th>','<th style="width:100;">銷售金額</th>']
        area.execute("select o.abalky as '門市代碼',o.ABALPH as '門市名稱',lv.EMPNAME1 as '負責業務',(o.DTTAEXP) as '銷售金額' from "
              +" (select distinct(USERID) as USERID,EMPNAME1 from LayaViewAreaUser  where subareano='"+an21[:an21.find('|')]+"'  ) lv,(select h.shshan,h.abalky,h.ABALPH,sum(CAST(f.DTTAEXP as decimal(18,0))) as DTTAEXP from "
              +" (select * from VS_SALEHD where   SHDRQJ >= '"+sd+"' and SHDRQJ <= '"+ed+"') h"
              +",(select * from VS_SALEFL where sdlitm>='100000' and sdlitm<='399999' "
			  #20190927 取消 +"and sdlitm not in ('100004')"
              +"  AND SDDRQJ >= '"+sd+"'  and SDDRQJ <= '"+ed+"') f where f.sddoco=h.shdoco and f.sdan8=h.shshan  "+firstr 
			  +" group by h.shshan,h.abalky,h.ABALPH ) o  where o.abalky=lv.USERID   order by o.abalky ")      
        for a in area.fetchall():
          tl=[]
          tl.append('<td style="width:100;">'+str(a[0])+'</td>')
          tl.append('<td style="width:150;">'+str(a[1])+'</td>')
          tl.append('<td style="width:100;">'+str(a[2])+'</td>')
          tl.append('<td style="width:100;">'+format(float(a[3]),',')+'</td>')	
          osaledata.append(tl)
      elif an11!='':  #中區名稱
        title=['<th style="width:100;">區碼(小)</th>','<th style="width:150;">負責區</th>','<th style="width:100;">負責業務</th>','<th style="width:100;">訂單數</th>','<th style="width:100;">店家數</th>','<th style="width:100;">銷售金額</th>','<th style="width:100;">均銷貨金額</th>']
        
        area.execute("select lv.SUBAREANO as '區碼(小)',lv.SUBAREANAME,lv.EMPNAME1 as '負責業務',count(lv.USERID) as '訂單數',count(distinct(lv.USERID)) as '店家數',sum(o.DTTAEXP) as '銷售金額', cast(round(sum(o.DTTAEXP)/count(distinct(lv.USERID)),0) as numeric(10,2))  as '均銷貨金額' from "
            +" (select distinct(USERID) as USERID,EMPNAME1,SUBAREANO,SUBAREANAME from LayaViewAreaUser  where areano='"+an11[:an11.find('|')]+"') lv,(select h.shshan,h.abalky,h.ABALPH,sum(CAST(f.DTTAEXP as decimal(18,0))) as DTTAEXP from "
            +" (select * from vs_salehd where   (shdcto='S2' OR shdcto='S3' OR shdcto='S4') and SHDRQJ >= '"+sd+"' and SHDRQJ <= '"+ed+"') h,(select * from vs_salefl where sdlitm>='100000' and sdlitm<='399999' "
			#20190927 取消 +"and sdlitm not in ('100004') "
			+" and (sddcto='S2' OR sddcto='S3' OR sddcto='S4') AND SDDRQJ >= '"+sd+"' and SDDRQJ <= '"+ed+"' ) f where f.sddoco=h.shdoco and f.sdan8=h.shshan "+firstr+" group by h.shshan,h.abalky,h.ABALPH ) o "
            +" where o.abalky=lv.USERID   group by lv.SUBAREANO,lv.EMPNAME1,lv.SUBAREANAME ORDER BY lv.SUBAREANO,lv.EMPNAME1   ")      
        for a in area.fetchall():
          tl=[]
          tl.append('<td style="width:100;">'+str(a[0])+'</td>')
          tl.append('<td style="width:150;">'+str(a[1])+'</td>')
          tl.append('<td style="width:100;">'+str(a[2])+'</td>')
          tl.append('<td style="width:100;">'+str(a[3])+'</td>')
          tl.append('<td style="width:100;">'+str(a[4])+'</td>')
          tl.append('<td style="width:100;">'+format(float(a[5]),',')+'</td>')
          tl.append('<td style="width:100;">'+format(float(a[6]),',')+'</td>')	
          osaledata.append(tl)
      elif an1!='': #大區名稱
        title=['<th style="width:100;">區碼(中)</th>','<th style="width:150;">負責業務</th>','<th style="width:100;">訂單數</th>','<th style="width:100;">店家數</th>','<th style="width:100;">銷售金額</th>','<th style="width:100;">均銷貨金額</th>']
        
        area.execute("select lv.AREANO as '區碼(中)',lv.EMPNAME1 as '負責業務',count(lv.USERID) as '訂單數',count(distinct(lv.USERID)) as '店家數',sum(o.DTTAEXP) as '銷售金額', cast(round(sum(o.DTTAEXP)/count(distinct(lv.USERID)),0) as numeric(10,2)) as '均銷貨金額' from "
              +" (select distinct(USERID) as USERID,EMPNAME1,AREANO from LayaViewAreaUser  where areano like '"+an1[:an1.find('|')]+"%' ) lv,(select h.shshan,h.abalky,h.ABALPH,sum(CAST(f.DTTAEXP as decimal(18,0))) as DTTAEXP from "
              +"(select * from vs_salehd where  (shdcto='S2' OR shdcto='S3' OR shdcto='S4') and SHDRQJ >= '"+sd+"' and SHDRQJ <= '"+ed+"') h,(select * from vs_salefl where sdlitm>='100000'"
              +" and sdlitm<='399999' "
			  #20190927 取消 +"and sdlitm not in ('100004')"
              +" and (sddcto='S2' OR sddcto='S3' OR sddcto='S4') AND SDDRQJ >= '" +sd+'" and SDDRQJ <= "'
              +ed+"' ) f where f.sddoco=h.shdoco and f.sdan8=h.shshan "+firstr+" group by h.shshan,h.abalky,h.ABALPH  ) o  where o.abalky=lv.USERID  group by lv.AREANO,lv.EMPNAME1 ORDER BY lv.AREANO,lv.EMPNAME1    ")      
        for a in area.fetchall():
          tl=[]
          tl.append('<td style="width:100;">'+str(a[0])+'</td>')
          tl.append('<td style="width:150;">'+str(a[1])+'</td>')
          tl.append('<td style="width:100;">'+str(a[2])+'</td>')
          tl.append('<td style="width:100;">'+str(a[3])+'</td>')
          tl.append('<td style="width:100;">'+format(float(a[4]),',')+'</td>')
          tl.append('<td style="width:100;">'+format(float(a[5]),',')+'</td>')
          
		  #tl.append('<td style="width:100;">'+str(a[6])+'</td>')	
          osaledata.append(tl)
      context['title']=title
      context['osaledata']=osaledata	  
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    area.execute("Select DISTINCT LEFT(AREANO,1)AREANO,LEFT(AREANAME,1) AREANAME from SalesArea where areano<>'TINO'")
    area0=[]
    for a in area.fetchall():
      area0.append(str(a[0])+'|'+str(a[1]))
    context['CK1']= 'OFF'
    context['area0']= area0
  #f.close()
  return render(request, 'product//OracleSalesCalc.html',context )#傳入參數

def pccss(request):  
  return render(request, 'pc.css', )
def tablecss(request):  
  return render(request, 'table.css', )
def spcss(request):  
  return render(request, 'sp.css', )
  
