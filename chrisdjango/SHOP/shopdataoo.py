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
    saledata=[]
    saledatae=[]
    wb = Workbook()	
    ws1 = wb.active	
    ws1.title = "data"
    #營業額
    if ck1=='on':
      ws1.append(['拉亞直營營業額('+sday+'~'+eday+')'])
      ws1.append(['日期','星期','門市','營業額','內用筆數','內用金額','外帶筆數','外帶金額','電話筆數','電話金額','門市外送筆數','門市外送金額','外送網筆數','外送網金額','其他筆數','其他金額','筆數小計','平均客單價'])
      title=['<th style="width:9%;">日期</th>','<th style="width:6%;">星期</th>','<th style="width:9%;">門市</th>','<th style="width:5%;">營業額</th>','<th style="width:5%;">內用筆數</th>','<th style="width:5%;">內用金額</th>'
             ,'<th style="width:5%;">外帶筆數</th>','<th style="width:5%;" >外帶金額</th>','<th style="width:5%;">電話筆數</th>','<th style="width:5%;">電話金額</th>'
    		 ,'<th style="width:5%;">外送筆數</th>','<th style="width:5%;">外送金額</th>','<th style="width:5%;">外送網筆數</th>','<th style="width:5%;">外送網金額</th>','<th style="width:5%;">其他筆數</th>','<th style="width:5%;">其他金額</th>','<th style="width:5%;">筆數小計</th>','<th style="width:5%;">平均客單價</th>']
      saletype=['內用','外帶','電話','外送','外送網','其他']
      saledataq=[]#日期、門市、筆數小計、金額小計
      tsale=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]#各欄總計暫存
      if shop=='全部門市':
        shop=" and h.sa_no like 'la%'"
      else:
        shop=" and h.sa_no='"+shop+"'"
      for s in range(len(saletype)): 
        
        if 	saletype[s]=='外送網':
          f.write('外送網'+'\n')
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('FoodPanda','熊貓','Uber','UberEats')"
                         +" and h.go_no=p.go_no  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  ")
          '''		
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.sname,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent]='"+saletype[s]+"'" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.sname order by c.sname,sdate")
          '''
        elif 	saletype[s]=='其他':
          f.write('其他'+'\n')
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('現金','刷卡','悠遊卡','easycard','')"
                         +" and h.go_no=p.go_no and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('現金','刷卡','悠遊卡','easycard','')"
                         +" and h.go_no=p.go_no and h.Serve_Type not in('內用','外帶','電話','外送')  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  ")
          
        else:
          f.write('XXXX'+'\n')
          f.write("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('現金','刷卡','悠遊卡','easycard')"
                         +" and h.go_no=p.go_no and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  "+'\n')
          temp211.execute("select sdate,DATENAME(Weekday, sdate) as wd,sname,convert(varchar,count(go_no)) as cs ,sum(sales_amount) as ts from (SELECT convert(varchar(10),h.invoice_date,120) as sdate,c.sname,h.go_no ,sales_amount,p.[payments] " 
                         +" FROM [possa].[dbo].[SHOP_orders] h ,[possa].[dbo].[shop] c,[possa].[dbo].SHOP_PAYMENTS p where h.status=3  and c.[sid]=h.sa_no and h.sa_no like 'l%' and p.[payments]  in ('現金','刷卡','悠遊卡','easycard')"
                         +" and h.go_no=p.go_no and h.Serve_Type='"+saletype[s]+"'  and h.invoice_date >= '"+sd+" 00:00:00' and h.invoice_date <='"+ed+" 23:00:00' "+shop+") a  group by sdate,sname  order by sdate,sname  ")
          '''
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.sname,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent] not in ('內用','外帶','電話','外送','外賣')" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.sname order by c.sname,sdate ")
          '''
        for d in temp211.fetchall():
          if s==0:		  
            saledata.append(['<td style="width:9%;">'+str(d[0])+'</td>','<td style="width:6%;">'+str(d[1])+'</td>','<td style="width:9%;">'+str(d[2])+'</td>','<td style="width:5%;" align="right">0</td>'
                            ,'<td style="width:5%;" align="center" >'+str(d[3])+'</td>','<td style="width:5%;" align="right">'+format(int(d[4]),',')+'</td>','<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="right">0</td>'
                            ,'<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="center">0</td>','<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="right">0</td>'
							,'<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="center">0</td>','<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="right">0</td>','<td style="width:5%;" align="right">0</td>'])
            f.write(str(s)+'\n')
            saledatae.append([str(d[0]),str(d[1]),str(d[2]),'0',str(d[3]),format(int(d[4]),','),'0','0','0','0','0','0','0','0','0','0','0','0'])
            tsale[1]=int(tsale[1])+int(d[3])
            tsale[2]=int(tsale[2])+int(d[4])
            saledataq.append([str(d[0]),str(d[2]),int((d[3])),int(d[4])])
            f.write(str(d[0])+'/'+str(d[2])+'/'+str(d[4])+'\n')
          
          else:
            for i in range(len(saledataq)):
              if str(d[0])==saledataq[i][0] and str(d[2])==saledataq[i][1] :
                saledataq[i][2]=saledataq[i][2]+int(d[3])
                saledataq[i][3]=saledataq[i][3]+int(d[4])
                saledata[i][s*2+4]='<td style="width:5%;" align="center">'+str(d[3])+'</td>'#因每類訂單有數量與金額二筆數字故 s*2 再因跳過表格前4欄再 +4
                saledata[i][s*2+5]='<td style="width:5%;" align="right">'+format(int(d[4]),',')+'</td>'
                saledatae[i][s*2+4]=str(d[3])
                saledatae[i][s*2+5]=format(int(d[4]),',')
                '''
                saledata[i][s+5+s-1]='<td style="width:5%;" align="center">'+str(d[3])+'</td>'
                saledata[i][s+6+s-1]='<td style="width:5%;" align="right">'+format(int(d[4]),',')+'</td>'
                saledatae[i][s+5+s-1]=str(d[3])
                saledatae[i][s+6+s-1]=format(int(d[4]),',')
                '''
                tsale[s*2+1]=tsale[s*2+1]+int(d[3])
                tsale[s*2+2]=tsale[s*2+2]+int(d[4])
      f.write(str((saledataq))+'\n')
      for i in range(len(saledataq)):
        saledata[i][3]='<td style="width:5%;" align="right">'+format(int(saledataq[i][3]),',')+'</td>'#該日總營業額
        saledata[i][16]='<td style="width:5%;" align="center">'+format(int(saledataq[i][2]),',')+'</td>'#訂單數
        saledata[i][17]='<td style="width:5%;" align="right">'+format(round(saledataq[i][3]/saledataq[i][2]),',')+'</td>'#客單價
        saledatae[i][3]=format(int(saledataq[i][3]),',')
        saledatae[i][16]=format(int(saledataq[i][2]),',')
        saledatae[i][17]=format(round(saledataq[i][3]/saledataq[i][2]),',')
        tsale[0]=tsale[0]+int(saledataq[i][3])
        tsale[len(tsale)-2]=tsale[len(tsale)-2]+int(saledataq[i][2])#訂單數
        #tsale[len(tsale)-1]=tsale[len(tsale)-1]+int(saledataq[i][3])
      f.write(str(saledataq)+'\n')
      tsale[len(tsale)-1]=round(tsale[0]/tsale[len(tsale)-2])#客單價
      saledata.append(['<td style="width:9%;"></td>','<td style="width:6%;"></td>','<td style="width:9%;">小計</td>','<td style="width:5%;" align="right">'+format(tsale[0],',')+'</td>'
                            ,'<td style="width:5%;" align="center" >'+str(tsale[1])+'</td>','<td style="width:5%;" align="right">'+format(tsale[2],',')+'</td>','<td style="width:5%;" align="center">'+str(tsale[3])+'</td>','<td style="width:5%;" align="right">'+format(tsale[4],',')+'</td>'
                            ,'<td style="width:5%;" align="center">'+str(tsale[5])+'</td>','<td style="width:5%;" align="right">'+format(tsale[6],',')+'</td>','<td style="width:5%;" align="center">'+str(tsale[7])+'</td>','<td style="width:5%;" align="right">'+format(tsale[8],',')+'</td>'
							,'<td style="width:5%;" align="center">'+str(tsale[9])+'</td>','<td style="width:5%;" align="right">'+format(tsale[10],',')+'</td>','<td style="width:5%;" align="center">'+format(int(tsale[11]),',')+'</td>','<td style="width:5%;" align="right">'+format(tsale[12],',')+'</td>'
							,'<td style="width:5%;" align="center">'+format(tsale[13],',')+'</td>','<td style="width:5%;" align="right">'+format(tsale[14],',')+'</td>'])
      
      try:
        
        saledatae.append(['','','小計',format(tsale[0],','),str(tsale[1]),format(tsale[2],','),str(tsale[3]),format(tsale[4],','),str(tsale[5]),format(tsale[6],','),str(tsale[7]),format(tsale[8],','),str(tsale[9])
	                   ,format(tsale[10],','),str(tsale[11]),format(tsale[12],','),str(tsale[13]),format(tsale[14],',')])
        
      except Exception as e:
        f.write(e)	  
      
      f.write('test'+'\n')
      for e in range(len(saledatae)):
        ws1.append(saledatae[e])
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx'
    #銷售數量
    if ck2=='on':
      ws1.append(['拉亞直營銷售數量('+sday+'~'+eday+')'])
      ws1.append(['門市','大類','商品編號','商品名稱','銷售數量'])
      title=['<th style="width:20%;">門市</th>','<th style="width:10%;">大類</th>','<th style="width:20%;">商品編號</th>','<th style="width:35%;">商品名稱</th>'
	           ,'<th style="width:15%;">銷售數量</th>']
      if shop=='全部門市':
        shop=''
      else:
        shop="and d.sa_no='"+shop+"'"
      temp211.execute("SELECT c.sname ,[Category] ,[Product_ID] ,[Product_Name] ,sum(convert(int,[Quantity])) as qty  FROM [possa].[dbo].[SHOP_detail] d ,[possa].[dbo].[shop] c "
                       +" where  c.[sid]=d.sa_no  and d.invoice_date >= '"+sd+" 00:00:00' and d.invoice_date <='"+ed+" 23:59:59'"+shop+"  group by  c.sname ,[Category] ,[Product_ID] ,[Product_Name] "
					   +"order by [Category]")
      '''
      temp211.execute("SELECT c.sname,[KindID],[ART_NO],[ART_NAME],sum([QTY]) qty FROM [LayaPos].[dbo].[OUTARTFL] f,[ERPSPOS].[dbo].[COMPANY] c "
                       +"where f.go_no not like '%*' and f.sdate >='"+sd+"' and f.sdate<= '"+ed+"' and  c.[COMP_NO]=f.[ShoopID] and f.[KindID] not like 's%'  "+shop
                       +" group by f.[KindID],c.sname,f.[ART_NO],f.[ART_NAME]  order by f.[ART_NO]")
      '''
      for d in temp211.fetchall():
        saledata.append(['<td style="width:20%;">'+str(d[0])+'</td>','<td style="width:10%;">'+str(d[1])+'</td>','<td style="width:20%;">'+str(d[2])+'</td>','<td style="width:35%;" >'+str(d[3])+'</td>'
                          ,'<td style="width:15%;"  align="right">'+str(d[4])+'</td>'])
        ws1.append([str(d[0]),str(d[1]),str(d[2]),str(d[3]),str(d[4])])
      
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_直營銷售數量('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_直營銷售數量('+sday+'~'+eday+').xlsx'
    #銷售量佔比
    	
    context['sdata'] = saledata				
    context['title'] = title
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    context['funcname'] = ''    
  sls=[]
  temp211.execute("SELECT [SID]+'_'+[SNAME] FROM [POSSA].[dbo].[SHOP]  where SID like 'la%' ")
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
  
