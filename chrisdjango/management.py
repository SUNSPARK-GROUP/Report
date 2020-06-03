from django.http import HttpResponse
from django.shortcuts import render,HttpResponseRedirect
from django.db import connection, transaction
from django.shortcuts import render_to_response
from django.template import RequestContext
from datetime import date
from datetime import timedelta
from siteapp.models import ViewMs211Daytotal
from siteapp.views import gettotaldata
from siteapp.views import checkuser
from siteapp.views import mainmenu
from siteapp.views import submenu
from graphos.sources.simple import SimpleDataSource
from graphos.sources.model import ModelDataSource
from graphos.renderers.gchart import ColumnChart
from django.urls import reverse
import cx_Oracle
import os
import io
import xlsxwriter
from . import Saletojde_aw
import pyodbc
import re
#from docx import Document

from django.shortcuts import render
from docxtpl import DocxTemplate
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

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
  #f=open(r'C:\Users\chris\chrisdjango\error_o.txt','w')
  if SQLSTRS=="SELECT":
    #f.write(SqlStr)
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
	
def worklog(request):#工作日誌主檔
  context= {}
  context_wt= {}
  # f=open(r'D:\chrisdjango\worklog.txt','w')
  try:
    connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
    cur214hd = connection214.cursor()
    orderhd=[]
    sdp=request.GET['Department']#get values
    ssp=request.GET['Sponsor']
    sbj=request.GET['Subject']
    sps=request.GET['Person']
    sta=request.GET['status']
    sday=request.GET['Sday']
    eday=request.GET['Eday']
    sd=sday[:4]+sday[5:7]+sday[8:10]
    ed=eday[:4]+eday[5:7]+eday[8:10]
    # f.write(sday+'~'+eday)
    context['Sday'] = sday
    context['Eday'] = eday
    # f.write(str(sta)+'\n')
    if sta == '':
      star = ''
    else:
      star = 'and misson_status = \'' + sta + '\''
    # f.write(str(sbj)+'\n')
    # f.write(str(sbjs)+'\n')
    try:
      ck1=request.GET['Checkbox1']
    except:
      ck1='off'
    # f.write(ck1)
    if ck1 == 'off':
      context['CK1'] = 'off'
      sbjs = re.sub(",","%\' or Subject like \'%",sbj)
      sbjr = re.sub(",","%\' or Detail like \'%",sbj)
      # f.write(sbjs+'\n')
      # f.write(sbjr+'\n')
      if eday == '':
        cur214hd.execute("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                    +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                    +""+str(star)+" and (Subject like '%"+str(sbjs)+"%' or Detail like '%"+str(sbjr)+"%') order by Date desc")
        # f.write("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      # +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      # +""+str(star)+" and (Subject like '%"+str(sbjs)+"%' or Detail like '%"+str(sbjr)+"%') order by Date desc")
      else:
        ed = int(ed)
        ed += 1
        ed = str(ed)
        cur214hd.execute("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      +" WHERE ([Date]>='"+sd+"' and [Date]<='"+ed+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      +""+str(star)+" and (Subject like '%"+str(sbjs)+"%' or Detail like '%"+str(sbjr)+"%') order by Date desc")
        # f.write("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      # +" WHERE ([Date]>='"+sd+"' and [Date]<='"+ed+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      # +""+str(star)+" and (Subject like '%"+str(sbjs)+"%' or Detail like '%"+str(sbjr)+"%') order by Date desc")
    else:
      remark=" and remark<>'' "
      context['CK1'] = 'on'
      sbjs = re.sub(",","%\' and Subject like \'%",sbj)
      sbjr = re.sub(",","%\' and Detail like \'%",sbj)
      #f.write(str(sbj))
      if eday == '':
        cur214hd.execute("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      +""+str(star)+" and (Subject like '%"+str(sbjr)+"%') or (Detail like '%"+str(sbjs)+"%') order by Date desc")
        # f.write("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      # +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      # +""+str(star)+" and (Subject like '%"+str(sbjr)+"%') or (Detail like '%"+str(sbjs)+"%') order by Date desc")
      else:
        ed = int(ed)
        ed += 1
        ed = str(ed)
        cur214hd.execute("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      +""+str(star)+" and (Subject like '%"+str(sbjr)+"%') or (Detail like '%"+str(sbjs)+"%') order by Date desc")
        # f.write("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog] "
                      # +" WHERE ([Date]>='"+sd+"') and Person like '%"+sps+"%' and Department like '%"+sdp+"%' and Sponsor like'%"+ssp+"%'"
                      # +""+str(star)+" and (Subject like '%"+str(sbjr)+"%') or (Detail like '%"+str(sbjs)+"%') order by Date desc")

        
    for data in cur214hd:
      torderhd=[]
      torderhd.append(str(data[0]))
      torderhd.append(str(data[1]))
      torderhd.append(str(data[2]))
      torderhd.append(str(data[3]))
      torderhd.append(str(data[4]))
      torderhd.append(str(data[5]))
      torderhd.append(str(data[7]))
      orderhd.append(torderhd)
  
    context['tabs']=cratetabs(15,len(orderhd))
    context['webWorklog']=tabsdata(15,orderhd)
    context['sDepartment']=sdp
    context['sSponsor']=ssp
    context['sSubject']=sbj
    context['sPerson']=sps
    context['mstatus']=sta
    if len(orderhd)==0:
      context['mess']='查無資料'
    else:
      context['mess']="共 "+str(len(orderhd))+" 筆資料"
  except:
    s='' 
    nday=showday(-14,'-',0) #2星期前
    eday=showday(0,'-',0) #今天
    context['Sday'] = nday
    context['Eday'] = eday	
  #f.close()  
  return render(request, 'worklog.html',context )
def logcheck(request):#工作日誌明細查看/新增/修改
  context= {}
  # f=open(r'D:\chrisdjango\log.txt','w')
  try:
    Sweborderfl=[]
    weborderfl=[]#頁簽內容	
    #f.write('gonoaa'+'\n')
    logno=request.GET['logno']#get values
    try:
      # f.write('1'+'\n') 
      connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
      cur214order = connection214.cursor()
      cur214order.execute("select [No],[Date],[Department],[Sponsor],[Subject],[Person],[Detail],[misson_status] from [ERPS].[dbo].[webWorklog]  where No='"+logno+"'")
      for h in cur214order:
        context['sNo']=str(h[0])
        context['sDate']=str(h[1])
        context['sSubject']=str(h[4])
        context['sSponsor']=str(h[3])
        context['sDepartment']=str(h[2])
        context['sPerson']=str(h[5])
        context['sDetail']=str(h[6])
        st = re.sub('\s','',str(h[7]))
        context['mstatus']=str(st)
        # f.write(str(h[0]) + str(h[1]) + str(h[4]) + str(h[3]) + str(h[2]) + str(h[5]) + str(h[6]) + str(st))
    except:
      # f.write('2'+'\n')
      connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
      cur214add = connection214.cursor()
      if request.method == "POST":#如果是以POST的方式才處理
        ck=request.POST['checkbox']
        sno=request.POST['sNo']
        day=request.POST['sDate']
        ssj=request.POST['sSubject']
        ssp=request.POST['sSponsor']
        sdp=request.POST['sDepartment']
        sps=request.POST['sPerson']
        sdt=request.POST['sDetail']
        sts=request.POST['status']
        # f.write(ck+'\n'+sno+'\n'+day+'\n'+ssj+'\n'+ssp+'\n'+sdp+'\n'+sps+'\n'+sdt+'\n'+sts+'\n')
        if sno == '':
          # f.write("INSERT INTO webWorklog(Date,Department,Sponsor,Subject,Person,Detail) VALUES ('"+day+"','"+sdp+"','"+ssp+"','"+ssj+"','"+sps+"','"+sdt+"')")
          cur214add.execute("INSERT INTO webWorklog(Date,Department,Sponsor,Subject,Person,Detail,misson_status) VALUES ('"+day+"','"+sdp+"','"+ssp+"','"+ssj+"','"+sps+"','"+sdt+"','"+sts+"')")
          cur214add.commit()
        else:
          # f.write("UPDATE [ERPS].[dbo].[webWorklog] set Detail ='"+sdt+"', misson_status = '"+sts+"' where No ='"+sno+"'")
          cur214add.execute("UPDATE [ERPS].[dbo].[webWorklog] set Detail ='"+sdt+"', misson_status = '"+sts+"' where No ='"+sno+"'")
          cur214add.commit()
          context['rt']="OK"
    try:
      # f.write('2'+'\n')
      connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
      cur214add = connection214.cursor()
      if request.method == "POST":#如果是以POST的方式才處理
        ck=request.POST['checkbox']
        sno=request.POST['sNo']
        day=request.POST['sDate']
        ssj=request.POST['sSubject']
        ssp=request.POST['sSponsor']
        sdp=request.POST['sDepartment']
        sps=request.POST['sPerson']
        sdt=request.POST['sDetail']
        sts=request.POST['status']
        # f.write(ck+'\n'+sno+'\n'+day+'\n'+ssj+'\n'+ssp+'\n'+sdp+'\n'+sps+'\n'+sdt+'\n'+sts+'\n')
        if sno == '':
          # f.write("INSERT INTO webWorklog(Date,Department,Sponsor,Subject,Person,Detail) VALUES ('"+day+"','"+sdp+"','"+ssp+"','"+ssj+"','"+sps+"','"+sdt+"')")
          cur214add.execute("INSERT INTO webWorklog(Date,Department,Sponsor,Subject,Person,Detail,misson_status) VALUES ('"+day+"','"+sdp+"','"+ssp+"','"+ssj+"','"+sps+"','"+sdt+"','"+sts+"')")
          cur214add.commit()
          context['rt']="OK"
        else:
          # f.write("UPDATE [ERPS].[dbo].[webWorklog] set Detail ='"+sdt+"'and misson_status = '"+sts+"' where No ='"+sno+"'")
          cur214add.execute("UPDATE [ERPS].[dbo].[webWorklog] set Detail ='"+sdt+"' , misson_status = '"+sts+"' where No ='"+sno+"'")
          cur214add.commit()
          context['rt']="OK"
    except:
      cur214add.close()
  except:
    if request.method == "POST":#如果是以POST的方式才處理
        ck=request.POST['checkbox']
        sno=request.POST['sNo']
        day=request.POST['sDate']
        ssj=request.POST['sSubject']
        ssp=request.POST['sSponsor']
        sdp=request.POST['sDepartment']
        sps=request.POST['sPerson']
        sdt=request.POST['sDetail']
        sts=request.POST['status']
        #f.write(ck+'\n'+sno+'\n'+day+'\n'+ssj+'\n'+ssp+'\n'+sdp+'\n'+sps+'\n'+sdt+'\n')
        connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
        cur214add = connection214.cursor()
        cur214add.execute("INSERT INTO webWorklog(Date,Department,Sponsor,Subject,Person,Detail,misson_status) VALUES ('"+day+"','"+sdp+"','"+ssp+"','"+ssj+"','"+sps+"','"+sdt+"','"+sts+"')")
        cur214add.commit()
        context['rt']="OK"
        cur214add.close()
  return render(request, 'logcheck.html',context )