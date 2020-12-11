# -*- encoding: UTF-8 -*-
'''
def hello(request):
    return HttpResponse("Hello world ! ")'''
from django.http import HttpResponse
from django.shortcuts import render,HttpResponseRedirect
from datetime import date
from datetime import timedelta
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

def showday(wd,sp,dt):#wd->0 today,wd->1 yesterday,wd->-1 tomorrow sp->Divider dt datetype
  t1 = 0-wd
  d=date.today()-timedelta(t1)
  if dt==1 :
    return d.strftime('%d'+sp+'%m'+sp+'%Y')#%Y->2015 %y->15
  else :
    return d.strftime('%Y'+sp+'%m'+sp+'%d')#%Y->2015 %y->15
def stockmanag(request):
  return render(request, 'stock//STOCKmanag.html' )#傳入參數
def stockpub(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
  temp214=connection214.cursor() 
  temp214d=connection214.cursor() 
  uno=format(request.COOKIES['userno']) 
  f=open(r'C:\Users\chris\chrisdjango\shoperror.txt','w')  
  stype=''
  sstype=''
  tday=showday(0,'',0)
  try: 
    tmess=''
    try:  
      ck1=request.GET['Checkbox1']
      stype='N'
    except:
      ck1='off'
    try:  
      ck2=request.GET['Checkbox2']
      stype='S'
    except:
      ck2='off'
    try:  
      ck3=request.GET['Checkbox3']
      sstype='D' #壹仟股
    except:
      ck3='off'
    try:  
      ck4=request.GET['Checkbox4']
      sstype='E' #壹萬股
    except:
      ck4='off'
    try:  
      ck5=request.GET['Checkbox5']
      sstype='F' #拾萬股
    except:
      ck5='off'
    
    if (ck3=='off' and ck4=='off' and ck5=='off') :
      tmess=tmess+'請選擇股數類型；'	
    sp=request.GET['sp']
    syear=request.GET['syear']
    sno=request.GET['sno']
    eno=request.GET['eno']
    if (len(sno)!=7 or len(eno)!=7 or (int(sno)>int(eno))):
      tmess=tmess+'股號長度不符7碼或股號起迄錯誤；'
    #tmess=tmess+"select count(*) from Stock_sp_mem where tsno>='"+syear+stype+sstype+sno+"' and tsno<='"+syear+stype+sstype+eno+"'"
    temp214.execute("select count(*) from Stock_sp_mem where tsno>='"+syear+stype+sstype+sno+"' and tsno<='"+syear+stype+sstype+eno+"'") 
    for s in temp214.fetchall():
      if int(s[0])>0:
        tmess=tmess+'股號:'+syear+stype+sstype+sno+'~'+syear+stype+sstype+eno+'已發行，請查明。'
    if tmess=='':
      no=int(sno)
      while no<=int(eno):
        f.write("insert into Stock_sp_mem (tsno,tskind,tsdate,shares,perval,account) VALUES ('"+syear+stype+sstype+str(no).zfill(7)+"','"+stype+"','"+tday+"','"+sstype+"','"+sp+"','SP00000')")
        temp214d.execute("insert into Stock_sp_mem (tsno,tskind,tsdate,shares,perval,account) VALUES ('"+syear+stype+sstype+str(no).zfill(7)+"','"+stype+"','"+tday+"','"+sstype+"','"+sp+"','SP00000')")
        temp214d.commit()
        no=no+1 	
    context['ck1']=ck1
    context['ck2']=ck2
    context['ck3']=ck3
    context['ck4']=ck4
    context['ck5']=ck5
    context['sp']=sp
    context['syear']=syear
    context['sno']=sno
    context['eno']=eno
    context['tmess']=tmess
  except:
    nday=showday(0,'-',0) #今天日期context['sp']	='10'
    context['syear']=(int(nday[:4])-1911)
    context['sno']=''
    context['eno']=''
    context['sp']='10'
  nday=showday(0,'-',0) #今天日期
  pubdata=''#發行狀況
  #未移轉
  temp214.execute("SELECT rtrim([tskind]) t,rtrim([shares]) s,rtrim((select min(tsno) from [ERPS].[dbo].[Stock_sp_mem] where account ='SP00000' )) as minno"
           +",rtrim((select max(tsno) from [ERPS].[dbo].[Stock_sp_mem] where account ='SP00000' )) as maxno,COUNT( [tsno]) as pic "
           +" FROM [ERPS].[dbo].[Stock_sp_mem] where account ='SP00000' group by [tskind],[shares],account")
  for s in temp214.fetchall():
    if pubdata=='':
      pubdata='1.未移轉:'+str(s[2])+' ~ '+str(s[3])+'\n'
    if s[0]=='N':
      if s[1]=='D':
        pubdata=pubdata+'  壹仟股(普通股): '+str(s[4])+' 張'+'\n'
      if s[1]=='E':
        pubdata=pubdata+'  壹萬股(普通股): '+str(s[4])+' 張'+'\n'
      if s[1]=='F':
        pubdata=pubdata+'  拾萬股(普通股): '+str(s[4])+' 張'+'\n'
    if s[0]=='S':
      if s[1]=='D':
        pubdata=pubdata+'  壹仟股(特別股): '+str(s[4])+' 張'+'\n'
      if s[1]=='E':
        pubdata=pubdata+'  壹萬股(特別股): '+str(s[4])+' 張'+'\n'
      if s[1]=='F':
        pubdata=pubdata+'  拾萬股(特別股): '+str(s[4])+' 張'+'\n'
  if pubdata=='':
    pubdata='1.無未移轉股'+'\n'
  pubdata1=''
  temp214.execute("SELECT rtrim([tskind]) t,rtrim([shares]) s,rtrim((select min(tsno) from [ERPS].[dbo].[Stock_sp_mem] where account <> 'SP00000' )) as minno"
           +",rtrim((select max(tsno) from [ERPS].[dbo].[Stock_sp_mem] where account  <> 'SP00000' )) as maxno,COUNT( [tsno]) as pic " 
           +" FROM [ERPS].[dbo].[Stock_sp_mem] where account  <> 'SP00000' group by [tskind],[shares]")
  for s in temp214.fetchall():
    if pubdata1=='':
      pubdata1='2.已移轉:'+str(s[2])+' ~ '+str(s[3])+'\n'
    if s[0]=='N':
      if s[1]=='D':
        pubdata1=pubdata1+'  壹仟股(普通股): '+str(s[4])+' 張'+'\n'
      if s[1]=='E':
        pubdata1=pubdata1+'  壹萬股(普通股): '+str(s[4])+' 張'+'\n'
      if s[1]=='F':
        pubdata1=pubdata1+'  拾萬股(普通股): '+str(s[4])+' 張'+'\n'
    if s[0]=='S':
      if s[1]=='D':
        pubdata1=pubdata1+'  壹仟股(特別股): '+str(s[4])+' 張'+'\n'
      if s[1]=='E':
        pubdata1=pubdata1+'  壹萬股(特別股): '+str(s[4])+' 張'+'\n'
      if s[1]=='F':
        pubdata1=pubdata1+'  拾萬股(特別股): '+str(s[4])+' 張'+'\n'
  context['pubdata']=pubdata+pubdata1
  return render(request, 'stock//stockpub.html',context )#傳入參數
  f.close()
def stockholder(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
  temp214=connection214.cursor() 
  uno=format(request.COOKIES['userno']) 
  f=open(r'C:\Users\chris\chrisdjango\shoperror.txt','w')  
  try:    
    stockh=[]    
    account=request.GET['saccount']
    context['saccount'] = account
    shname=request.GET['sshname']
    context['sshname'] = shname
    sqlc=''
    if account!='':
      sqlc=" and account='"+account+"' "
    if shname!='':
      sqlc=sqlc+" and shname like '%"+shname+"%' "    
    '''
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
    '''
    f.write("SELECT [account],[shname] from [ERPS].[dbo].[Stock_sh_infor] where 1=1  "+sqlc)
    temp214.execute("SELECT [account],[shname] from [ERPS].[dbo].[Stock_sh_infor] where 1=1  "+sqlc)
    for d in temp214.fetchall():
      tstockh=[]
      tstockh.append(str(d[0]))
      tstockh.append(str(d[1]))
      stockh.append(tstockh)	 
    context['stockh'] = stockh	
    #頁籤    
    context['tabs']=cratetabs(10,len(stockh))
    #頁籤
    #頁籤內容
    context['condata']=tabsdata(10,stockh)
  except:
    nday=showday(0,'-',0) #今天日期  
  return render(request, 'stock//STOCKholder.html',context )#傳入參數
  f.close()
def stockdetel(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
  temp214=connection214.cursor() 
  f=open(r'C:\Users\chris\chrisdjango\shoperror.txt','w')  
  try:
    cid=request.GET['cid']#戶號
    try:  
      sdir=request.GET['sdir']
      idir='1'
    except:
      sdir='off'
      idir='0'
    sname=request.GET['sname']#姓名
    bday=request.GET['bday'].replace('-','/')#生日
    scid2=request.GET['scid2']#身分證
    scadd=request.GET['scadd']#地址
    sbank=request.GET['sbank']#銀行
    sbranches=request.GET['sbranches']#分行
    sremittance=request.GET['sremittance']#帳號
    sb=request.GET['Submit']
    if sb=='修改資料':
      temp214.execute("update [ERPS].[dbo].[Stock_sh_infor] set shname='"+sname+"',pid='"+scid2+"',bdate='"+bday+"',address='"+scadd+"',bank='"+sbank+"',branches='"+sbranches
                    +"',remittance='"+sremittance+"',dir="+idir+" where account='"+cid+"' ")
      temp214.commit()
      context['smess'] ='修改完成'
    elif sb=='新增股東':
      temp214.execute("select max(convert(int,account))+1 from [ERPS].[dbo].[Stock_sh_infor] where account<>'SP00000' ")
      for m in temp214.fetchall():
        cid=str(m[0])
      temp214.execute("insert into [ERPS].[dbo].[Stock_sh_infor] ([account],[shname],[pid],[bdate],[address],[bank],[branches],[remittance],[dir] )"
	                 +"VALUES ('"+cid+"','"+sname+"','"+scid2+"','"+bday+"','"+scadd+"','"+sbank+"','"+sbranches+"','"+sremittance+"','"+idir+"')")
      temp214.commit()					 
      context['smess'] ='新增完成'
  except:
    nday=showday(0,'-',0) #今天日期
    context['smess'] =''
  temp214.execute("SELECT i.[account],i.[shname],i.[sqty],i.[pid],i.[bdate],i.[address],i.[bank],i.[branches],i.[remittance],i.[remark],i.[dir],d.tss,e.tts,f.hts "
                  +" from [ERPS].[dbo].[Stock_sh_infor] i left join (select account, count(tsno) as tss   from   [ERPS].[dbo].Stock_sp_mem where account='"+cid+"' "
				  +" and shares='D' group by account) d on d. account=i.account "
                  +" left join (select account, count(tsno) as tts   from   [ERPS].[dbo].Stock_sp_mem where account='"+cid+"' and shares='E' group by account) e on e. account=i.account "
                  +" left join  (select account, count(tsno) as hts   from   [ERPS].[dbo].Stock_sp_mem where account='"+cid+"' and shares='F' group by account) f on f. account=i.account "
                  +" where i.account='"+cid+"' ")
  for d in temp214.fetchall():
    context['cid'] = cid
    context['scid2'] = str(d[3])	  
    context['sname'] = str(d[1])
    context['bday'] = str(d[4]).replace('/','-')
    context['scadd'] = str(d[5])
    context['sbank'] = str(d[6])
    context['sbranches'] = str(d[7])
    context['sremittance'] = str(d[8])
    stockc=0
    if str(d[13])!='None':
      context['ht'] = str(d[13])
      stockc=stockc+int(d[13])*100000
    else:
      context['ht'] ='0'
    if str(d[12])!='None':
      context['tt'] = str(d[12])
      stockc=stockc+int(d[12])*10000
    else:
      context['tt'] ='0'
    if str(d[11])!='None':
      context['ts'] = str(d[11])
      stockc=stockc+int(d[11])*1000
    else:
      context['ts'] ='0'
    if str(d[10])=='True': 
      context['sdir'] = 'on'
    else:
      context['sdir'] = 'off'
    context['stockc'] = format(stockc,',')
  return render(request, 'stock//detelstock.html',context )#傳入參數    
  f.close()
def stockshift(request):#股票移轉
  context= {}
  sday=showday(0,'-',0) #今天日期
  context['sday']=sday  
  f=open(r'C:\Users\chris\chrisdjango\stockshifterror.txt','w')
  try:
    f.write('acc11'+'\n')       
    saccount=request.GET['saccount'] #移出者
    context['saccount']	=saccount
    f.write("saccount:"+saccount+'\n')
    saccount1=request.GET['saccount1'] #轉入者
    context['saccount1']=saccount1
    f.write("saccount1:"+saccount1+'\n')
    stype=request.GET['stocktype']
    f.write("stocktype:"+stype+'\n')
    context['stype']=stype
    connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
    temp214=connection214.cursor()
    stockd=[]
    if stype=='壹仟股':
      temp214.execute("select a.tsno,a.shname,a.tsdate,b.shname,a.repdate from (select m.tsno,RTrim(i.shname) as shname,m.tsdate,m.remark,m.repdate FROM Stock_sp_mem m,(SELECT account,shname  FROM Stock_sh_infor where account='"+saccount[:saccount.index('_')]+"') i "
                      +" where m.shares='D' AND m.account='"+saccount[:saccount.index('_')]+"') a left join (SELECT account,shname  FROM Stock_sh_infor) b on a.remark=b.account")
      for d in temp214.fetchall():
        td=[str(d[0]),str(d[1]),str(d[2]),str(d[3]),str(d[4])]
        stockd.append(td)
      f.write(str(stockd))		
    if stype=='壹萬股':
      temp214.execute("select m.tsno,i.shname,m.tsdate,m.repdate,m.remark FROM Stock_sp_mem m,(SELECT account,shname  FROM Stock_sh_infor where account='"+saccount[:saccount.index('_')]+"') i where m.shares='E' AND m.account='"+saccount[:saccount.index('_')]+"'")
    if stype=='拾萬股':
      temp214.execute("select m.tsno,i.shname,m.tsdate,m.repdate,m.remark FROM Stock_sp_mem m,(SELECT account,shname  FROM Stock_sh_infor where account='"+saccount[:saccount.index('_')]+"') i where m.shares='F' AND m.account='"+saccount[:saccount.index('_')]+"'")
    sb=request.GET['sb']
    f.write("sb:"+sb+'\n')
    context['stockd']=stockd #持有股票
  except:
    f.write("error123"+'\n')
  return render(request, 'stock//STOCKshift.html' ,context)#傳入參數
  f.close()
def stockhsearch(request):#股東搜尋
  context= {}
  stockh=[]
  input=request.GET['input']
  context['input']=input
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
  temp214=connection214.cursor()
  f=open('C:\\Users\\chris\\chrisdjango\\comerror.txt','w')
  f.write('get'+'\n')
  try:
    cstr=request.GET['cstr']
    f.write(cstr+'\n')
    temp214.execute("select [account],[shname] from  [ERPS].[dbo].[Stock_sh_infor] where  shname like '%"+cstr+"%'  order by account")
  except:
    temp214.execute('select [account],[shname] from  [ERPS].[dbo].[Stock_sh_infor]  order by account')
  for data in temp214:
    tcomno=[]
    tcomno.append(str(data[0]))
    tcomno.append(str(data[1]))
    stockh.append(tcomno)
  context['stockh']=stockh
  
  return render(request, 'stock//stockhsearch.html' ,context)#傳入參數
  f.close()
def pccss(request):  
  return render(request, 'pc.css', )
def tablecss(request):  
  return render(request, 'table.css', )
def spcss(request):  
  return render(request, 'sp.css', )
  
