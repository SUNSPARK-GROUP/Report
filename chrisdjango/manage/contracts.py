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
from ftplib import FTP
from django.http import HttpRequest
import os
import sys
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
def getflist(cft,cno):
  ftpl = FTP()
  ftpl.connect('192.168.0.214',21)
  ftpl.login(user='chris',passwd='chrisk123')	
  ftpl.encoding='utf-8'
  ftpl.cwd('type'+str(cft)+'/'+str(cno)) #更改遠端目錄
  fl=ftpl.nlst()
  fls=''
  #f=open(r'C:\Users\chris\chrisdjango\ftplist.txt','w',encoding="utf-8")
  
  return fl
  ftpl.quit
  #f.close()
def ftpupload(fn,cft,cno):
  try:
    ftpu = FTP()
    ftpu.connect('192.168.0.214',21)
    ftpu.login(user='chris',passwd='chrisk123')	    	
    ftpu.encoding='utf-8'
    ftpu.cwd('type'+str(cft)+'/'+str(cno)) #更改遠端目錄
    bufsize=2048 #設定的緩衝區大小
    ftext=open('C:\\Users\\Administrator\\chrisdjango\\ftpftemp\\'+fn,'rb') #以寫模式在本地開啟檔案
    ftpu.storbinary('STOR '+fn,ftext)
    ftext.close()
    ftpu.quit #退出ft
    return fn+'上傳成功'
  except :    
    return  fn+'上傳失敗'
def contracts(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
  temp214=connection214.cursor() 
  
  uno=format(request.COOKIES['userno'])  
  try:
    f=open(r'C:\Users\chris\chrisdjango\contractrror.txt','w')
    cno=request.GET['cno']
    context['scno'] = cno
    cname=request.GET['cname']
    context['scname'] = cname
    tp=request.GET['tp']
    context['tp'] =tp 
    conts=[]	
    try:#過期合約
      ck1=request.GET['Checkbox1']
      context['CK1'] =ck1
    except:
      ck1='off'
    try:#所有合約
      ck2=request.GET['Checkbox2']
      context['CK2'] =ck2
    except:
      ck2='off'
    f.write(ck1+'\n')
    if tp!='全部':
      ctype="and contracttype='"+tp+"'"
    else:
      ctype=''
    f.write('ctype:'+tp+'\n')

    #目錄
    
    temp214.execute("select * from (select  Law_no as '合約編號',Contract_name as '合約名稱',Signb_name as '客戶/廠商' "
             +",Contract_fdate as '起始日',Contract_edate  as '迄止日' FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()>=0 "+ctype+"  union all "
             +"select  '過期_'+Law_no as '合約編號',Contract_name as '合約名稱',Signb_name as '客戶/廠商'"
             +",Contract_fdate as '起始日',Contract_edate  as '迄止日' FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()<0 "+ctype+" ) a "
             +" order by SUBSTRING(合約編號,CHARINDEX('L',合約編號),11)")
    wb = Workbook()	
    ws1 = wb.active	
    ws1.title = "data"
    ws1.append(['合約編號','合約名稱','客戶/廠商','起始日','迄止日'])
    f.write('wb'+'\n') 
    for t in temp214.fetchall():
      tb=[]
      tb.append(str(t[0]))
      tb.append(str(t[1]))
      tb.append(str(t[2]))
      tb.append(str(t[3]))
      tb.append(str(t[4]))
      ws1.append(tb)
    wb.save('C:\\Users\\Administrator\\chrisdjango\\MEDIA\\合約目錄.xlsx')
    context['efilename']='合約目錄.xlsx'

    if ck2=='on' and ck1=='off':
      f.write("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()<0  and Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
      temp214.execute("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()<0  and Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
    elif ck1=='on' and ck2=='off':#未過期合約
      f.write("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()>=0  and Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
      temp214.execute("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where Contract_edate-getdate()>=0  and Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
      
    else:#不分過期所有的合約
      f.write("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
      temp214.execute("select [Law_no],[Contract_name],[Contract_fdate],[Contract_edate] FROM ERPS.dbo.CONTRACTS where  Law_no like '"
                       +cno+"%' and Contract_name like '"+cname+"%' "+ctype+" Order by Law_no")
    for data in temp214:
      tconts=[]
      tconts.append(str(data[0]))
      tconts.append(str(data[1]))
      tconts.append(str(data[2]))
      tconts.append(str(data[3]))
      conts.append(tconts)
    #f.write(str(conts))
    #頁籤    
    context['tabs']=cratetabs(10,len(conts))
    #頁籤
    #頁籤內容
    context['condata']=tabsdata(10,conts)
    if len(conts)==0:
      context['mess']='查無資料'
    else:
      context['mess']="共 "+str(len(conts))+" 筆資料"
    
    context['typels'] = ['全部','工程暨維保類(含資訊)合約','行銷暨合作類合約','委任暨服務類合約','保密類合約','保全類合約','租賃類合約','海外授權代理合約','合作類合約','智財權類合約','買賣採購暨經銷類合約'
	                      ,'設計規劃類合約','備忘錄(包含意向書)','營運暨門市類合約','其他類合約','事業一處加盟合約','事業二處加盟合約','承攬類合約'] 
  except:
    #nday=showday(0,'-',0) #今天日期
    #context['Sday'] = nday[:8]+'01'
    #context['Eday'] = nday
    context['CK1'] ='on'  
    context['typels'] = ['全部','工程暨維保類(含資訊)合約','行銷暨合作類合約','委任暨服務類合約','保密類合約','保全類合約','租賃類合約','海外授權代理合約','合作類合約','智財權類合約','買賣採購暨經銷類合約'
	                      ,'設計規劃類合約','備忘錄(包含意向書)','營運暨門市類合約','其他類合約','事業一處加盟合約','事業二處加盟合約','承攬類合約'] 
  return render(request, 'manage//contracts.html',context )#傳入參數
  f.close()
def contd(request):
  context={}
  f=open(r'C:\Users\chris\chrisdjango\contractrror.txt','w')
  try:
    cno=request.GET['cno']    
    connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=ERPS;UID=apuser;PWD=0920799339') 
    temp214=connection214.cursor()
    try:
      context['Sday'] = request.GET['Sday']
      sday=request.GET['Sday']
      f.write(sday)
      sd=(sday[:4]+'/'+sday[5:7]+'/'+sday[8:10])
      f.write("update ERPS.dbo.CONTRACTS set Contract_edate='"+sd+"' where Law_no='"+cno+"'")
      temp214.execute("update ERPS.dbo.CONTRACTS set Contract_edate='"+sd+"' where Law_no='"+cno+"'")
      temp214.commit()
    except:
      cno=request.GET['cno']
    temp214.execute("select Law_no,Contract_no,Contract_cft,Contracttype,Contract_lab,keeper,Signa_name,Signb_name,Work_dep,Work_emp,Contract_deposit,Contract_amount,Contract_fdate,Contract_edate"
	               +" from ERPS.dbo.CONTRACTS where Law_no='"+cno+"'")    
    ms='合約編號:'+cno+' '
    for c in temp214:
      Contract_no=c[1]
      Contractcft=str(c[2])
      if c[2]==0 :
        ms=ms+'/普通機密'
      elif c[2]==5 :
        ms=ms+'/最高機密'
      elif c[2]==9 :
        ms=ms+'/極機密'
      if c[3]:
        ms=ms+'/制式'
      else:
        ms=ms+'/非制式'
      if c[4]:
        ms=ms+'/重大'+'\n'
      else:
        ms=ms+'/非重大'+'\n'
      
      ms=ms+'保管人: '+c[5]+'  我方:'+c[6]+'   客戶/廠商:'+c[7]+'\n'
      ms=ms+'承辦部門:'+c[8]+'   承辦人:'+c[9]+'\n'
      
      ms=ms+'合約保證金(含稅）:'+c[10]+'\n'
      ms=ms+'合約總金額(含稅）:'+c[11]+'\n'
      ms=ms+'合約起迄日期:'+c[12]+' ~ '+c[13]+'\n'
      context['Sday'] =c[13][:4]+'-'+c[13][5:7]+'-'+c[13][8:10]
    ms=ms+'******審核意見與申請人說明*********'
    '''
    ftp=FTP("192.168.0.203") #設定變數    
    ftp.login("chrisk600","chrisk123")#連線的使用者名稱，密碼 
    '''
    ftp = FTP()
    ftp.connect('192.168.0.214',21)
    ftp.login(user='chris',passwd='chrisk123')	
    ftp.cwd('type'+str(Contractcft)+'/'+str(Contract_no)) #更改遠端目錄
    bufsize=1024 #設定的緩衝區大小    
    filename=str(Contract_no)+".txt" #需要下載的檔案
    ftext=open('c:\\contd.txt',"wb") #以寫模式在本地開啟檔案
    ftp.retrbinary("RETR "+filename,ftext.write,bufsize)
    ftext.close()
    ftxt=open("c:\\contd.txt",'r')
    line = ftxt.readline()
    while line:
      line = ftxt.readline() 
      ms=ms+line
    ftxt.close()
    ftp.quit #退出ft    
    context['contd']=ms
    context['cno']=cno
    context['ctno']=str(Contractcft)
    context['ceno']=str(Contract_no)
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday
  return render(request, 'manage//contsdetel.html',context )#傳入參數
  f.close()
def upload(request):
  context={}
  f=open(r'C:\Users\chris\chrisdjango\uploadrror.txt','w')
  eno=request.GET['eno']
  ctno=request.GET['ctno']
  try:
    f.write('test'+'\n')    
    if request.POST:
      ss=request.POST['Submit1']
      if ss=='上傳檔案':
        fname=request.FILES['uf']
        f.write(str(ss)+'\n')
        uf= open('C:\\Users\\Administrator\\chrisdjango\\ftpftemp\\'+fname.name, 'wb+')
        with uf as destination:
          for chunk in fname.chunks():
            destination.write(chunk)
        uf.close()
        f.write(fname.name+','+str(ctno)+','+str(eno))
        context['fmess'] = ftpupload(fname.name,ctno,eno)
        
      #if ss=='下載檔案':
    f.write(eno+'\n')
  except:
    nday=showday(0,'-',0) #今天日期
  context['eno'] = eno
  context['ctno'] = ctno
  ftxt= getflist(ctno,eno)
  context['ftxt'] =ftxt
  fls=''
  for lf in ftxt:
    fls=fls+lf+'\n'
  context['flist'] = fls
  return render(request, 'manage//upload.html',context )#傳入參數
  f.close()
def pccss(request):  
  return render(request, 'pc.css', )
def tablecss(request):  
  return render(request, 'table.css', )
def spcss(request):  
  return render(request, 'sp.css', )
  
