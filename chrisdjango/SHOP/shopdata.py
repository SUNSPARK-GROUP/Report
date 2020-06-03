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
  connection211=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.211;DATABASE=ERPSPOS;UID=apuser;PWD=0920799339') 
  temp211=connection211.cursor() 
  uno=format(request.COOKIES['userno'])  
  try:
    f=open(r'C:\Users\chris\chrisdjango\shoperror.txt','w')
    product=[]
    sday=request.GET['Sday']#get values
    eday=request.GET['Eday']
    context['Sday'] = sday
    context['Eday'] = eday
    sd=(sday[:4]+sday[5:7]+sday[8:10])
    ed=(eday[:4]+eday[5:7]+eday[8:10])
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
    if ck1=='on':
      ws1.append(['拉亞直營營業額('+sday+'~'+eday+')'])
      ws1.append(['日期','星期','門市','營業額','內用筆數','內用金額','外帶筆數','外帶金額','電話筆數','電話金額','外送筆數','外送金額','外賣筆數','外賣金額'])
      title=['<th style="width:5%;">日期</th>','<th style="width:6%;">星期</th>','<th style="width:9%;">門市</th>','<th style="width:9%;">營業額</th>','<th style="width:6%;">內用筆數</th>','<th style="width:6%;">內用金額</th>'
             ,'<th style="width:6%;">外帶筆數</th>','<th style="width:6%;">外帶金額</th>','<th style="width:6%;">電話筆數</th>','<th style="width:6%;">電話金額</th>'
    		 ,'<th style="width:6%;">外送筆數</th>','<th style="width:6%;">外送金額</th>','<th style="width:6%;">外賣筆數</th>','<th style="width:6%;">外賣金額</th>','<th style="width:6%;">其他筆數</th>','<th style="width:6%;">其他金額</th>']
      saletype=['內用','外帶','電話','外送','外賣','其他']
      saledata1=[]
      if shop=='全部門市':
        shop=''
      else:
        shop="and h.sa_no='"+shop+"'"
      for s in range(len(saletype)): 
        
        if 	saletype[s]!='其他':					
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.COMP_NAME,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent]='"+saletype[s]+"'" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.COMP_NAME order by c.COMP_NAME,sdate")
        else:
          
          temp211.execute("SELECT h.sdate,DATENAME(Weekday, h.sdate) as wd,c.COMP_NAME,count(go_no) as cs ,sum([TOTO1]) as ts FROM [LayaPos].[dbo].[OUTARTHD] h ,[ERPSPOS].[dbo].[COMPANY] c where go_no not like '%*'  and [deskparent] not in ('內用','外帶','電話','外送','外賣')" 
                        +"  and sdate >='"+sd+"' and sdate<= '"+ed+"'  and c.[COMP_NO]=h.sa_no "+shop+" group by sdate,c.COMP_NAME order by c.COMP_NAME,sdate")
        for d in temp211.fetchall():
          if s==0:                        
            saledata.append(['<td style="width:5%;">'+str(d[0])+'</td>','<td style="width:6%;">'+str(d[1])+'</td>','<td style="width:9%;">'+str(d[2])+'</td>','<td style="width:9%;">0</td>'
                            ,'<td style="width:6%;">'+str(d[3])+'</td>','<td style="width:6%;">'+str(d[4])+'</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>'
                            ,'<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>','<td style="width:6%;">0</td>'])
            saledatae.append([str(d[0]),str(d[1]),str(d[2]),'0',str(d[3]),str(d[4]),'0','0','0','0','0','0','0','0','0','0'])
            saledata1.append([str(d[0]),str(d[2]),int(d[4])])
            f.write(str(d[0])+'/'+str(d[2])+'/'+str(d[4])+'\n')
          
          else:
            for i in range(len(saledata1)):
              if str(d[0])==saledata1[i][0] and str(d[2])==saledata1[i][1] :
                saledata1[i][2]=saledata1[i][2]+int(d[4])
                saledata[i][s+5+s-1]='<td style="width:6%;">'+str(d[3])+'</td>'
                saledata[i][s+6+s-1]='<td style="width:6%;">'+str(d[4])+'</td>'
                saledatae[i][s+5+s-1]=str(d[3])
                saledatae[i][s+6+s-1]=str(d[4])
      for i in range(len(saledata1)):
        saledata[i][3]='<td style="width:9%;">'+str(saledata1[i][2])+'</td>'
        saledatae[i][3]=str(saledata1[i][2])
      #f.write(saledata+'\n')
      for e in range(len(saledatae)):
        ws1.append(saledatae[e])
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_拉亞直營營業額('+sday+'~'+eday+').xlsx'
    if ck2=='on':
      ws1.append(['拉亞直營銷售數量('+sday+'~'+eday+')'])
      ws1.append(['門市','大類','商品編號','商品名稱','銷售數量'])
      title=['<th style="width:20%;">門市</th>','<th style="width:10%;">大類</th>','<th style="width:20%;">商品編號</th>','<th style="width:35%;">商品名稱</th>'
	           ,'<th style="width:15%;">銷售數量</th>']
      if shop=='全部門市':
        shop=''
      else:
        shop="and f.ShoopID='"+shop+"'"
      temp211.execute("SELECT c.COMP_NAME,[KindID],[ART_NO],[ART_NAME],sum([QTY]) qty FROM [LayaPos].[dbo].[OUTARTFL] f,[ERPSPOS].[dbo].[COMPANY] c "
                       +"where f.go_no not like '%*' and f.sdate >='"+sd+"' and f.sdate<= '"+ed+"' and  c.[COMP_NO]=f.[ShoopID] and f.[KindID] not like 's%'  "+shop
                       +" group by f.[KindID],c.COMP_NAME,f.[ART_NO],f.[ART_NAME]  order by f.[ART_NO]")
      for d in temp211.fetchall():
        saledata.append(['<td style="width:20%;">'+str(d[0])+'</td>','<td style="width:10%;">'+str(d[1])+'</td>','<td style="width:20%;">'+str(d[2])+'</td>','<td style="width:35%;">'+str(d[3])+'</td>'
                          ,'<td style="width:15%;">'+str(d[4])+'</td>'])
        ws1.append([str(d[0]),str(d[1]),str(d[2]),str(d[3]),str(d[4])])
      
      wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\'+uno+'_直營銷售數量('+sday+'~'+eday+').xlsx')
      context['efilename']=uno+'_直營銷售數量('+sday+'~'+eday+').xlsx'    
    context['sdata'] = saledata				
    context['title'] = title
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    context['funcname'] = ''    
  sls=[]
  temp211.execute("SELECT [COMP_NO]+'_'+[COMP_NAME] FROM [ERPSPOS].[dbo].[COMPANY]  where comp_no like 'la%' and groupid='LA'")
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
  
