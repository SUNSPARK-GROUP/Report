'''
def hello(request):
    return HttpResponse("Hello world ! ")'''
from django.http import HttpResponse
from django.shortcuts import render,HttpResponseRedirect
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
from . import Saletojde_aw_cc
import pyodbc
import pymysql
import time
#from docx import Document

from django.shortcuts import render
from docxtpl import DocxTemplate
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
db206_erps=pyodbc.connect('DRIVER={SQL Server}; SERVER=192.168.0.206,1433; DATABASE=TGSalary; UID=apuser; PWD=0920799339')
def excel(request):
  output = io.BytesIO()  #用BytesIO 來存我們的資料  
  workbook = xlsxwriter.Workbook(output)
  worksheet = workbook.add_worksheet()  #新增一個sheet
  row = 0
  col = 0
  worksheet.write(row, col,     '說明' )
  worksheet.write(row, col + 1, '2' )
  worksheet.write(row, col + 2, '3' )  #在某行某列加入資料
  workbook.close()  #把workbook關閉  
  output.seek(0)
  response = HttpResponse(output.read(),content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
  response['Content-Disposition'] = "attachment; filename=excel.xlsx"
  return response
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
def ordermerge(custid,yday):
  #f=open(r'C:\Users\chris\chrisdjango\ordermerge.txt','w')
  chaincodehd=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccdmodhd = chaincodehd.cursor()
  ccdmodhd2= chaincodehd.cursor()
  chaincodem=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccmod=chaincodem.cursor()
  mos='0'#合併單數
  mtotal='0'#合併總金額
  nremark=''#備註
  ccdmodhd.execute("select * from (sELECT accountid,arrivaltime,count(order_no) as orders FROM orderformpos  where  no_sm='未轉單' and arrivaltime like '"
                 +yday+"%' and accountid='"+custid+"'  group by accountid,arrivaltime)a where orders>1")
  for c in ccdmodhd.fetchall():
    mos=str(c[2])  
    ccdmodhd2.execute("select order_no,double_1 from orderformpos where accountid='"+str(c[0])+"' and no_sm='未轉單' and arrivaltime like '"+yday+"%' order by order_no ")
    ors=''
    drs=''
    ford=''
    toto=0
    for h in ccdmodhd2.fetchall():
      if ors=='':
        ors="'"+str(h[0])+"'"
        ford=str(h[0])
      else:
        ors=ors+",'"+str(h[0])+"'"
        if drs=='':
          drs="'"+str(h[0])+"'"
        else:
          drs=drs+",'"+str(h[0])+"'"
      toto=toto+int(h[1])
    if ors!='':
      ccdmodhd2.execute("select AccountID,AccountName,AccountUser,order_no,DateTime_1,Double_1,StoreName,StoreAddr,RelevantPerson,"
                    +"RelevantTel,ArrivalTime,Remark,modifynum,modifydate,modifyman,modifyID,isImport,Attr_Brand,NO_SM"
                    +",NormalDelivery from orderformpos where  order_no in("+ors+")")
      for h in ccdmodhd2.fetchall():      
        ccmod.execute("insert into orderformpos_t (AccountID,AccountName,AccountUser,order_no,DateTime_1,Double_1,StoreName,StoreAddr,RelevantPerson,"
                    +"RelevantTel,ArrivalTime,Remark,modifynum,modifydate,modifyman,modifyID,isImport,Attr_Brand,NO_SM"
                    +",NormalDelivery) values('"+str(h[0])+"','"+str(h[1])+"','"+str(h[2])+"','"+str(h[3])+"','"+str(h[4])+"','"+str(h[5])+"','"+str(h[6])+"','"+str(h[7])
                    +"','"+str(h[8])+"','"+str(h[9])+"','"+str(h[10])+"','"+str(h[11])+"','"+str(h[12])+"','"+str(h[13])+"','"+str(h[14])+"','"+str(h[15])+"','"+str(h[16])
                    +"','"+str(h[17])+"','"+str(h[18])+"','"+str(h[19])+"')")
        chaincodem.commit()
        nremark=nremark+str(h[11])
      ccdmodhd2.execute("select order_no,AccountID,ProdID,ProdName,ProdGroup,Amount,Double_1,Double_2,Img_Name from orderformpos_prod_sub where  order_no in("+ors+")")
      for h in ccdmodhd2.fetchall():
        pname=str(h[3]).replace("'",'')
        ccmod.execute("insert into orderformpos_prod_sub_t (order_no,AccountID,ProdID,ProdName,ProdGroup,Amount,Double_1,Double_2,Img_Name) values('"
                    +str(h[0])+"','"+str(h[1])+"','"+str(h[2])+"','"+pname+"','"+str(h[4])+"','"+str(h[5])+"','"+str(h[6])+"','"+str(h[7])
                    +"','"+str(h[8])+"')")
        chaincodem.commit()
    		
    ccmod.execute("update orderformpos_prod_sub set order_no='"+ford+"' where order_no in("+ors+")")
    chaincodem.commit()
    ccmod.execute("update orderformpos set double_1='"+str(toto)+"',remark='"+nremark+"' where order_no='"+ford+"'")
    chaincodem.commit()
    mtotal=str(toto)
    ccmod.execute("delete from orderformpos  where order_no in("+drs+")")
    chaincodem.commit()
  #f.write('共合併: '+mos+'張 金額: '+mtotal+' 元'+'\n')
  #f.close()	
  return('共合併: '+mos+'張 金額: '+mtotal+' 元') 
  
def passchang(request):
  context={}
  f=open('pass.txt','w')
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
  try:
    tuno=request.GET['uno'] 
    f.write(tuno+'\n')
    ty=request.GET['ty']
    if ty=='BEGINCHANGPS':
      context['uno'] =uno
      f.write(ty+'\n')
      return render(request, 'passchang.html',context)#傳入參數
    else:
      f.write(ty+'\n')
      mod214=connection214.cursor()
      mod214.execute("update [ERPS].[dbo].[WebSPusers] set password='"+ty+"' where [userid]='"+tuno+"'")
      mod214.commit()
      context['MSG'] ='YES'
      return render(request, 'passchang.html',context)#
  except:
    context['uno'] ='工號' 
    f.write('工號'+'\n')	
    return render(request, 'passchang.html',context)#傳入參數
  f.close()	
def admincheck(request):
  global uno
  uno=''
  global uname
  uname=''
  lmes='請輸入帳密'  
  f=open('c:\\admincheck.txt','w')
  f.write('yyy'+'\n')
  try:
    eno=request.GET['sempno']	
    context= {}
    context['sempno'] = eno
    dataset=db206_erps.cursor()
    dataset1=db206_erps.cursor()
    dataset.execute("SELECT COUNT(*)  FROM [TGSalary].[dbo].[EMPLOYEE] WHERE EMPNO='"+eno+"' AND (LDATE IS NULL OR LDATE='')")
    for d in dataset.fetchall():
      if d[0]!=0:
        f.write(str(eno)+'\n')
        ndate=showday(0,'',0) #今天日期
        f.write(ndate+'\n')
        ntime=time.strftime("%X")
        dataset1.execute("SELECT COUNT(*)  FROM [TGSalary].[dbo].[POPINOUT] WHERE EMPNO='"+eno+"' AND SDATE ='"+ndate+"'")
        for i in dataset1.fetchall():
          ind=int(i[0])
        if ind%2==0:
          stp='I'
        else:
          stp='O'
        f.write(ndate+ntime[:5]+'\n')
        f.write("insert into [TGSalary].[dbo].[POPINOUT] ([EMPNO],[SDATE],[STIME],[STYPE]) values('"+eno+"','"+ndate+"','"+ntime[:5]+"','"+stp+"')")
        dataset.execute("insert into [TGSalary].[dbo].[POPINOUT] ([EMPNO],[SDATE],[STIME],[STYPE]) values('"+eno+"','"+ndate+"','"+ntime[:5]+"','"+stp+"')")
        dataset.commit()
        context['wmess']=eno+'_'+ndate[4:]+'_'+ntime[:5]+'成功!!'
      else:
        context['wmess']='工號錯誤!!'
    return render(request, 'admincheck.html',context )#傳入參數
  except:
    if request.POST:
      uno=request.POST['username']
      uname=checkuser(request.POST['username'],request.POST['passw'])
      if uname=='':	
        lmes='帳密錯誤'
    #f.write(uno+':'+uname)
    if uname=='':
      return render(request, 'admincheck.html',{'loginmes':lmes} )  
    else:
      response =HttpResponseRedirect('/sysadmin')
        #f.write('HttpResponse''/sysadmin''')
      response.set_cookie('userno',uno)
        #f.write('response.set_cookie(''userno'',uno)''')
        #response.set_cookie('username',uname)
        #f.write('response.set_cookie(''username'',uname)''')
      return response
        #aurl = reverse('sysadmin',kwargs={'uno':uno,'uname':uname}) # 在urls.py 相關路經在加個name
        #return HttpResponseRedirect(aurl)    
  f.write('yes')
  f.close()
    
def sysadmin(request): 
  #global uno
  #global uname
  submlist={}
  #f=open('sysadmincheck.txt','w')
  if 'userno' in request.COOKIES:
    #f.write('userno')
    uno=format(request.COOKIES['userno'])
    #f.write(uno)  
  try: 
    ml=mainmenu(uno)
    crm=''
    try:
      s=request.GET['mmenu']
      crm=request.GET['crm']
	  #suburl=request.GET['surl']
      submlist=submenu(s,uno)
    #return HttpResponse(s)
    except: print('cc')
    if uname!='':
      return render(request, 'sysadmin.html',{'uno':uno,'uname':uname,'menum':ml,'menus':submlist,'crm':crm} )
    else:
      return HttpResponseRedirect('/admincheck')
  except:
    return HttpResponseRedirect('/admincheck')
  
  #f.close()    
  #return HttpResponse(request, 'test1.html', )
  #return render_to_response('test1.html',locals())
def second(request): 
  context= {}
  try:
    s=request.GET['SHOP']#get values
    context['nowday'] = request.GET['date']
    context['shop'] = "'"+s+"_"+request.GET['date']+"'"
    #context['ttl'] = s
    if s in ('TINOS','VASA','LAYA','FANI','SelFish'):
      d=request.GET['date']
      td=d[:4]+d[5:7]+d[8:10]
      context['ttl'] = gettotaldata(s,td,0) 
      gadata=gettotaldata(s,td,1)
      #data_source = ModelDataSource(gettotaldata(s,td,0),fields=['shop', 'sales'])
      data_source = SimpleDataSource(data=gadata)
      gchart = ColumnChart(data_source)
      context['gchart'] =gadata
  except:
    s=''    
    nday=showday(0,'-',0) #今天日期
    context['nowday'] = nday
    context['ttl'] = s
  return render(request, 'shopsale.html',context )#傳入參數

def weborder2jde(request):
  context= {}
  #f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  
  try:
    set=request.GET['Submit1']
    if set=='明日訂單轉入Jde':
      #t=Saletojde_aw.set2jde
      #f.write(set+'(1)\n')
      #try:
      context['weborder']=Saletojde_aw.set2jde(1)
      #except Exception as e: 
        #f.write(e+'/n')
    elif set=='今日訂單轉入Jde':
      #t=Saletojde_aw.set2jde
      #f.write(set+'(0)\n')
      context['weborder']=Saletojde_aw.set2jde(0)
    context['macct']="共轉入 "+str(len(context['weborder']))+" 筆資料"
    '''
      try:	  
        context['macct']="共轉入 "+str(len(context['weborder']))+" 筆資料" 
      except Exception as e: f.write(str(e))
    '''	  
  except:
    s=''    
  #f.close()  
  return render(request, 'weborder2jde.html',context )#傳入參數
def weborder2jde_cc(request):
  context= {}
  #f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  
  try:
    set=request.GET['Submit1']
    if set=='明日訂單轉入Jde':
      #t=Saletojde_aw.set2jde
      #f.write(set+'(1)\n')
      #try:
      context['weborder']=Saletojde_aw_cc.set2jde(1)
      #except Exception as e: 
        #f.write(e+'/n')
    elif set=='今日訂單轉入Jde':
      #t=Saletojde_aw.set2jde
      #f.write(set+'(0)\n')
      context['weborder']=Saletojde_aw_cc.set2jde(0)
    context['macct']="共轉入 "+str(len(context['weborder']))+" 筆資料"
    '''
      try:	  
        context['macct']="共轉入 "+str(len(context['weborder']))+" 筆資料" 
      except Exception as e: f.write(str(e))
    '''	  
  except:
    s=''    
  #f.close()  
  return render(request, 'weborder2jde_cc.html',context )#傳入參數
def saledetelf4211(request): 
  context= {}
  context_wt= {}
  f=open(r'C:\Users\chris\chrisdjango\error_4211.txt','w')
  try:
    f4211d=[]
    f4211dw=[]
    cid=request.GET['an8']#get values
    doco=request.GET['doco']
    dcto=request.GET['dcto']
    '''
    f.write("select F01.aban8, F01.abalky,F01.abalph,F16.ALADD1, replace(f15.wpar1,' ')||'-'||replace(f15.wpph1,' '),f11.wwalph,F01.ABTX2 from f0101 F01,F0116 F16,f0115 f15,f0111 f11 "
             +" WHERE f16.alan8=f01.aban8 and f01.aban8='"+cid+"' and f15.wpan8=f01.aban8 and f15.WPphtp='TEL' and f11.wwalph like replace(abac30,' ')||'-%'"+'\n')
    '''
    '''
    serdata=CONORACLE("select F01.aban8, F01.abalky,F01.abalph,F16.ALADD1, replace(f15.wpar1,' ')||'-'||replace(f15.wpph1,' '),f11.wwalph,F01.ABTX2,t009.dv75tsun,f112.WWMLNM from f0101 F01,F0116 F16,f0115 f15,f0111 f11,F0111 F112,F75T009 T009 "
                      +" WHERE f16.alan8=f01.aban8 and f01.aban8='"+cid+"' and f15.wpan8=f01.aban8 and f15.WPphtp='TEL' and t009.dvan8=f01.aban8 and (t009.dvco='00000' or t009.dvco='00100') and f11.wwalph like replace(abac30,' ')||'-%' AND F112.wwan8=  f01.aban8 AND F112.WWIDLN=0")
    '''
    serdata=CONORACLE("select F01.aban8, F01.abalky,F01.abalph,F16.ALADD1, replace(f15.wpar1,' ')||'-'||replace(f15.wpph1,' '),f11.wwalph,F01.ABTX2,t009.dv75tsun,f112.WWMLNM from f0101 F01,F0116 F16,f0115 f15,f0111 f11,F0111 F112,F75T009 T009 "
                      +" WHERE f16.alan8=f01.aban8 and f01.aban8='"+cid+"' and f15.wpan8=f01.aban8 and f15.WPphtp='TEL' and t009.dvan8=f01.aban8 and f11.wwan8=f01.aban8   and (t009.dvco='00000' or t009.dvco='00100')  AND F112.wwan8=  f01.aban8 AND F112.WWIDLN=0")
    for data in serdata:
      context['cid'] = str(data[0])#門市碼
      context['cid2'] = str(data[1])#門市碼2
      context['cname'] = str(data[2])#店名
      context['cadd'] = str(data[3])#地址
      context['ctel'] = str(data[4])#電話"
      context['cdriver'] = str(data[5])#路線
      context['cust'] = str(data[6])#聯絡人
      context_wt['cid'] = str(data[0])#門市碼
      context_wt['cid2'] = str(data[1])#門市碼2
      context_wt['cname'] = str(data[2])#店名
      context_wt['cadd'] = str(data[3])#地址
      context_wt['ctel'] = str(data[4])#電話"
      context_wt['cdriver'] = str(data[5])#路線
      context_wt['cust'] = str(data[6])#聯絡人
      context_wt['ino'] = str(data[7])#客戶統編
      context_wt['cname2'] = str(data[8])#客戶抬頭
      cn=str(data[2]).replace(' ','')
    serdata=CONORACLE("select F01.aban8, F01.abalky,F01.abalph,F16.ALADD1, replace(f15.wpar1,' ')||'-'||replace(f15.wpph1,' '),f11.wwalph,F01.ABTX2 from f0101 F01,F0116 F16,f0115 f15,f0111 f11 "
                      +" WHERE f16.alan8=f01.aban8 and f01.aban8='"+cid+"' and f15.wpan8=f01.aban8 and f15.WPphtp='CEL' and f15.WPrck7='2' and f11.wwalph like replace(abac30,' ')||'-%'")
    context_wt['cctel'] =''
    context_wt['doco'] =doco
    for data in serdata:
      context_wt['cctel'] = str(data[4])#手機"
    
    f.write("select sdlnid/1000 as sdlnid,sddsc1,sddsc2,sdsoqs/10000,sduom1,round(sduprc/10000*1.05,0),round(sduprc/10000*1.05,0)*sdsoqs/10000 as sdaexp ,TO_CHAR(TO_NUMBER(TO_CHAR(TO_DATE('20'||SUBSTR(SDTRDJ,2,2)||'-01-01','yyyy-mm-dd')+SUBSTR(SDTRDJ,4,3)-1,'yyyymmdd'))) AS SDTRDJ"
	                 +" from f4211  where sddoco='"+doco+"'   AND SDAN8='"+cid+"' and sddcto='"+dcto+"' ")
    
    serdata=CONORACLE("select shdoco,shdel1,shdel2 from f4201 where shdoco='"+doco+"' and shan8='"+cid+"'  ")
    for data in serdata:
      context_wt['memo'] = str(data[1])#memo
      context_wt['invno'] = str(data[2])#發票"
    serdata=CONORACLE("select sdlnid/1000 as sdlnid,sddsc1,sddsc2,sdsoqs/10000,sduom1,round(sduprc/10000*1.05,0),round(sduprc/10000*1.05,0)*sdsoqs/10000 as sdaexp ,TO_CHAR(TO_NUMBER(TO_CHAR(TO_DATE('20'||SUBSTR(SDTRDJ,2,2)||'-01-01','yyyy-mm-dd')+SUBSTR(SDTRDJ,4,3)-1,'yyyymmdd'))) AS SDTRDJ"
	                 +",sdlitm from f4211  where sddoco='"+doco+"'   AND SDAN8='"+cid+"' and sddcto='"+dcto+"' order by sdlnid ")#AND SDNXTR='620'
    asum=0
    for data in serdata:
      tf4211d=[]
      tf4211d.append(str(data[0]))
      tf4211d.append(str(data[1]))
      tf4211d.append(str(data[2]))
      tf4211d.append(str(data[3]))
      tf4211d.append(str(data[4]))
      tf4211d.append(str(data[5]))
      tf4211d.append(str(data[6]))
      asum=asum+int(data[6])
      odate=str(data[7])
      f4211d.append(tf4211d)
      tf4211dw={'no':str(data[0]),'litm':str(data[8]),'property':str(data[1]),'unit':str(data[4]),'qty':str(data[3]),'sprice':str(data[5]),'sum':str(data[6])}
      '''
      tf4211dw.append(str(data[0]))
      tf4211dw.append(str(data[7]))
      tf4211dw.append(str(data[1]))
      tf4211dw.append(str(data[3]))
      tf4211dw.append(str(data[4]))
      tf4211dw.append(str(data[5]))
      tf4211dw.append(str(data[6]))
      '''
      f4211dw.append(tf4211dw)
    #f.write(str(f4211d)+'\n')
    #f.write(str(context)+'\n')
    #f.write(str(context_wt)+'\n')
    #context_wt['sale_labels']=['序號','品號','品名規格','數量','單位','含稅單價','小計','客戶簽收']
    context['F4211d']=f4211d
    context['doco']=doco
    context_wt['asum']=str(format(asum,','))
    context_wt['atax']=str(format(asum-round(asum/1.05,0),','))
    context_wt['bsum']=str(format(round(asum/1.05,0),','))
    context_wt['sale_list']=f4211dw
    context_wt['odate']=odate
    context['reportmes']='<B>森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')</B>'
    context_wt['CAPTION']='森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')'
    #f.write(r'C:\Users\CHRIS\chrisdjango\docxt\saledetelf4211_tp.docx')
    #tpl = DocxTemplate(r'C:\Users\CHRIS\chrisdjango\docxt\saledetelf4211_tp.docx')#測試路徑
    tpl = DocxTemplate(r'C:\Users\Administrator\chrisdjango\docxt\saledetelf4211_tp.docx')#實際路徑
    f.write(str(context_wt))
    tpl.render(context_wt)
    #f.write(r'C:\Users\CHRIS\chrisdjango\MEDIA\saledetel'+doco+'.docx')	
    #tpl.save(r'C:\Users\CHRIS\chrisdjango\MEDIA\saledetel'+doco+'.docx')#測試路徑
    tpl.save(r'C:\Users\Administrator\chrisdjango\MEDIA\saledetel'+doco+r'.docx')#實際路徑
    #tpl.save(r'C:\Users\CHRIS\chrisdjango\static\saledetel'+doco+r'.docx')
  except:
    s=''    
  f.close()  
  return render(request, 'saledetelf4211.html',context )#傳入參數
def weborder(request): 
  context= {}
  context_wt= {}
  f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  try:
    connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
    cur206hd = connection206.cursor()
    orderhd=[]
    #f4211dw=[]
    #f.write('cid')
    cid=request.GET['IDCUST']#get values
    sday=request.GET['Sday']
    eday=request.GET['Eday']
    sd=sday[:4]+sday[5:7]+sday[8:10]
    ed=eday[:4]+eday[5:7]+eday[8:10]
    #f.write(sday+'~'+eday)
    context['Sday'] = sday
    context['Eday'] = eday
    
    try:
      ck1=request.GET['Checkbox1']
    except:
      ck1='off'
    try:
      ck2=request.GET['Checkbox2']
    except:
      ck2='off'
    try:
      ck3=request.GET['Checkbox3']
    except:
      ck3='off'
    try:
      ck4=request.GET['Checkbox4']
    except:
      ck4='off'
    f.write(ck1+'\n')
    f.write(ck2+'\n')
    f.write(ck3+'\n')
    f.write(ck4+'\n')
    '''
    f.write("select [GO_NO],[nm_c],[TOTAMT],[APPLYDATE] from [TGSalary].[dbo].[WEBORDERHD] "
                      +" WHERE [APPLYDATE]>='"+sd+"' and [APPLYDATE]<='"+ed+"' and ID_CUST like '"+cid+"%'")
    '''
    remark=''
    edi=''
    canc=''
    unnomo=''
    if ck1=='on':
      remark=" and remark<>'' "
      context['CK1'] = 'on'
    else:
      context['CK1'] = 'off'
    if ck2=='on':
      edi=" AND (len(NO_SM)<1 or  NO_SM  is null)  "
      context['CK2'] = 'on'
    else:
      context['CK2'] = 'off'
    if ck3=='on':
      canc=" and NO_SM='取消' "
      edi=""
      context['CK3'] = 'on'
    else:
      context['CK3'] = 'off'
    if ck4=='on':
      unnomo=" and remark like '%非正配單%' "
      remark=""
      context['CK4'] = 'on'
    else:
      context['CK4'] = 'off'
    
    f.write("select [GO_NO],[nm_c],[TOTAMT],[APPLYDATE] from [TGSalary].[dbo].[WEBORDERHD] "
                      +" WHERE [APPLYDATE]>='"+sd+"' and [APPLYDATE]<='"+ed+"' and ID_CUST like '"+cid+"%' "+remark+edi+canc+unnomo+" order by APPLYDATE desc ")
    
    cur206hd.execute("select [GO_NO],[nm_c],[TOTAMT],[APPLYDATE] from [TGSalary].[dbo].[WEBORDERHD] "
                      +" WHERE [APPLYDATE]>='"+sd+"' and [APPLYDATE]<='"+ed+"' and ID_CUST like '"+cid+"%' "+remark+edi+canc+unnomo+" order by APPLYDATE desc ")
					  
    
    for data in cur206hd:
      torderhd=[]
      torderhd.append(str(data[0]))
      torderhd.append(str(data[1]))
      torderhd.append(str(data[2]))
      torderhd.append(str(data[3]))
      orderhd.append(torderhd)
    #頁籤    
    context['tabs']=cratetabs(10,len(orderhd))
    #頁籤
    #頁籤內容
    context['weborder']=tabsdata(10,orderhd)
    #頁籤內容
    #context['weborder']=orderhd
    context['sIDCUST']=cid
    if len(orderhd)==0:
      context['mess']='查無資料'
    else:
      context['mess']="共 "+str(len(orderhd))+" 筆資料"
    #context['reportmes']='<B>森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')</B>'
    '''
    context['doco']=doco
    context_wt['asum']=str(format(asum,','))
    context_wt['atax']=str(format(asum-round(asum/1.05,0),','))
    context_wt['bsum']=str(format(round(asum/1.05,0),','))
    context_wt['sale_list']=f4211dw
    context_wt['odate']=odate
    
    #context_wt['CAPTION']='森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')'
    '''
  except:
    s='' 
    nday=showday(0,'-',0) #今天日期
    eday=showday(13,'-',0) #今天日期
    context['Sday'] = nday
    context['Eday'] = eday	
    context['CK1'] = 'off'
    context['CK2'] = 'on'
    context['CK3'] = 'off'
  f.close()  
  return render(request, 'weborder.html',context )#傳入參數
def weborder_cc(request): 
  context= {}
  context_wt= {}
  f=open(r'C:\Users\chris\chrisdjango\error11.txt','w')
  try:
    #chaincodedb=pymysql.connect(host='192.168.0.218', port=3306, user='sunsparkit', passwd='28813605', db='ccerp_tw001114hq',charset='utf8')
    chaincodedb=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
    #f.write('11')
    ccdmod=chaincodedb.cursor()
    orderhd=[]
    f.write('11')
    #f4211dw=[]
    #f.write('cid')
    cid=request.GET['IDCUST']#get values
    sday=request.GET['Sday']
    eday=request.GET['Eday']
    sd=sday[:4]+sday[5:7]+sday[8:10]
    ed=eday[:4]+eday[5:7]+eday[8:10]
    f.write(sday+'~'+eday)
    context['Sday'] = sday
    context['Eday'] = eday
    
    try:
      ck1=request.GET['Checkbox1']
    except:
      ck1='off'
    try:
      ck2=request.GET['Checkbox2']
    except:
      ck2='off'
    try:
      ck3=request.GET['Checkbox3']
    except:
      ck3='off'
    try:
      ck4=request.GET['Checkbox4']
    except:
      ck4='off'
    '''
    f.write(ck1+'\n')
    f.write(ck2+'\n')
    f.write(ck3+'\n')
    f.write(ck4+'\n')
    '''
    '''
    f.write("select [GO_NO],[nm_c],[TOTAMT],[APPLYDATE] from [TGSalary].[dbo].[WEBORDERHD] "
                      +" WHERE [APPLYDATE]>='"+sd+"' and [APPLYDATE]<='"+ed+"' and ID_CUST like '"+cid+"%'")
    '''
    remark=''
    edi=''
    canc=''
    unnomo=''
    if ck1=='on':
      remark=" and Remark<>'' "
      context['CK1'] = 'on'
    else:
      context['CK1'] = 'off'
    if ck2=='on':
      edi=" AND (NO_SM='未轉單')  "
      context['CK2'] = 'on'
    else:
      context['CK2'] = 'off'
    if ck3=='on':
      canc=" and NO_SM='已取消' "
      edi=""
      context['CK3'] = 'on'
    else:
      context['CK3'] = 'off'
    if ck4=='on':
      unnomo=" and NormalDelivery like '非正配單' "
      remark=""
      context['CK4'] = 'on'
    else:
      context['CK4'] = 'off'
    '''
    f.write("select order_no,StoreName,Double_1,ArrivalTime from orderformpos "
                      +" WHERE DateTime_1>='"+sday+" 00:00:00' and DateTime_1<='"+eday+" 00:00:00' and AccountID like '"+cid+"%' "+remark+edi+canc+unnomo+" order by ArrivalTime desc ")
    '''
    f.write("select order_no,StoreName,Double_1,date(ArrivalTime) from orderformpos "
                      +" WHERE ArrivalTime>='"+sday+" 00:00:00' and ArrivalTime<='"+eday+" 00:00:00' and AccountID like '"+cid+"%' "+remark+edi+canc+unnomo+" order by ArrivalTime desc,AccountID ASC ")
					
    ccdmod.execute("select order_no,StoreName,Double_1,date(ArrivalTime) from orderformpos "
                      +" WHERE ArrivalTime>='"+sday+" 00:00:00' and ArrivalTime<='"+eday+" 00:00:00' and AccountID like '"+cid+"%' "+remark+edi+canc+unnomo+" order by ArrivalTime desc,AccountID ASC ")
					  
    
    for data in ccdmod:
      torderhd=[]
      torderhd.append(str(data[0]))
      torderhd.append(str(data[1]))
      torderhd.append(str(data[2]))
      torderhd.append(str(data[3]))
      orderhd.append(torderhd)
    #頁籤    
    context['tabs']=cratetabs(10,len(orderhd))
    #頁籤
    #頁籤內容
    context['weborder']=tabsdata(10,orderhd)
    #頁籤內容
    #context['weborder']=orderhd
    context['sIDCUST']=cid
    if len(orderhd)==0:
      context['mess']='查無資料'
    else:
      context['mess']="共 "+str(len(orderhd))+" 筆資料"
    #context['reportmes']='<B>森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')</B>'
    '''
    context['doco']=doco
    context_wt['asum']=str(format(asum,','))
    context_wt['atax']=str(format(asum-round(asum/1.05,0),','))
    context_wt['bsum']=str(format(round(asum/1.05,0),','))
    context_wt['sale_list']=f4211dw
    context_wt['odate']=odate
    
    #context_wt['CAPTION']='森邦(股) 銷貨明細('+cn+'/單號:'+doco+'/日期:'+odate+')'
    '''
  except:
    s='' 
    nday=showday(0,'-',0) #今天日期
    eday=showday(13,'-',0) #今天日期
    context['Sday'] = nday
    context['Eday'] = eday	
    context['CK1'] = 'off'
    context['CK2'] = 'on'
    context['CK3'] = 'off'
  f.close()  
  return render(request, 'weborder_cc.html',context )#傳入參數
def weborderdetel(request):
  context= {}
  f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  try:
    Sweborderfl=[]
    weborderfl=[]#頁簽內容	
    #f.write('gonoaa'+'\n')
    #try:
    gono=request.GET['gono']#get values
    #except  Exception as e: f.write(str(e))
    #f.write('go_no:'+'\n')
    try:
      editb=request.GET['edit']
      aday=request.GET['aday']
      canc=request.GET['canc']
      umo=request.GET['umo']
      f.write(umo)
      aday=aday[:4]+aday[5:7]+aday[8:10]
    except:  editb='no'
    connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
    cur206order = connection206.cursor()
    if editb=='yes':
      if canc=='true':
        cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',NO_SM='取消' where go_no='"+gono+"'")
      else:
        cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"' where go_no='"+gono+"'")
      if umo=='true':
        cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',REMARK=REMARK+'*非正配單*' where go_no='"+gono+"'")
      else:
        cur206order.execute("select REMARK from [TGSalary].[dbo].[WEBORDERHD] where go_no='"+gono+"'")
        for r in cur206order.fetchall():
          f.write(str(r[0]))
          rk=str(r[0]).replace('非正配單','')
          f.write(rk)
        cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',REMARK='"+rk+"' where go_no='"+gono+"'")
      cur206order.commit()
    recount=0
    cur206order.execute("select [GO_NO],[id_cust],[nm_c],[TOTAMT],[APPLYDATE],[remark],[SDATETIME],[no_sm] from [TGSalary].[dbo].[WEBORDERHD] where go_no='"+gono+"'")
    for h in cur206order:
      context['gono']=str(h[0])
      context['cid']=str(h[1])
      context['cname']=str(h[2])
      context['cmark']=str(h[5])
      context['sdate']=str(h[6])
      aday=str(h[4])
      context['Aday']=aday[:4]+'-'+aday[4:6]+'-'+aday[6:8]
      #f.write(str(h[7]))
      if str(h[7])=='取消':
        context['CK3']='on'
      else:
        context['CK3']='off'
      if str(h[5]).find('非正配單')>-1:
        context['CK4']='on'
      else:
        context['CK4']='off'		
    
    cur206order.execute("select a.new_iditem,a.nm_item,fl.[QTY],fl.[UPRICE],fl.[SUBTOT]  FROM [TGSalary].[dbo].[WEBORDERFL] fl,[TGSalary].[dbo].[webart] a where fl.go_no='"+gono+"' and a.id_item=fl.id_item  order by  a.new_iditem ") 
    #f.write("select a.new_iditem,a.nm_item,fl.[QTY],fl.[UPRICE],fl.[SUBTOT]  FROM [TGSalary].[dbo].[WEBORDERFL] fl,[TGSalary].[dbo].[webart] a where fl.go_no='"+gono+"' and a.id_item=fl.id_item  order by  a.new_iditem ")
    #f.write(str(recount))
    for o in cur206order:
      tweborderfl=[]
      tweborderfl.append(str(o[0]))      
      tweborderfl.append(str(o[1]))
      tweborderfl.append(str(o[2]))
      tweborderfl.append(str(o[3]))
      tweborderfl.append(str(o[4]))
      #TF4211.append(str(o[4]))
      #TF4211.append(str(o[5]))
      #TF4211.append(str(o[5]))
      #TF4211.append(str(o[6]))
      #TF4211.append(str(o[7]))
      Sweborderfl.append(tweborderfl)
      recount=recount+1
    #f.write(str(Sweborderfl)+'\n')
    #頁籤    
    context['tabs']=cratetabs(12,recount)
    #頁籤
    #頁籤內容
    context['weborderfl']=tabsdata(12,Sweborderfl)
    #頁籤內容
    #f.write(str(weborderfl)) 
        
    context['reportmes']='<B>'+context['cname']+' 訂貨明細</B>'
    	
  except:
    s=''    
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday
    context['Eday'] = nday
  f.close()
  return render(request, 'weborderdetel.html',context )#傳入參數
def weborderdetel_cc(request):
  context= {}
  #f=open(r'C:\Users\chris\chrisdjango\error.txt','w')
  try:
    Sweborderfl=[]
    weborderfl=[]#頁簽內容	
    #f.write('gonoaa'+'\n')
    #try:
    gono=request.GET['gono']#get values
    
    #except  Exception as e: f.write(str(e))
    #f.write('go_no:'+gono+'\n')
    try:
      editb=request.GET['edit']
      aday=request.GET['aday']
      canc=request.GET['canc']
      umo=request.GET['umo']
      cid=request.GET['cid']#20200520
      '''
      f.write(umo)	  	  
      f.write('editb:'+editb+'\n')
      f.write('canc:'+canc+'\n')
      f.write('umo:'+umo+'\n')
      '''
      #aday=aday[:4]+aday[5:7]+aday[8:10]
    except:  editb='no'
    #connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
    #cur206order = connection206.cursor()
    chaincodedb=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
    ccdmod = chaincodedb.cursor()	
    #f.write('ccdmod ok'+'\n')
    if editb=='merge':
      context['mergomess']=ordermerge(cid,aday)
    if editb=='yes':
      if canc=='true':
        #cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',NO_SM='取消' where go_no='"+gono+"'")
        #f.write("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NO_SM='已取消' where order_no='"+gono+"' and  NO_SM='未轉單'"+'\n')
        ccdmod.execute("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NO_SM='已取消' where order_no='"+gono+"' and  NO_SM='未轉單'")
        chaincodedb.commit()
      else:
        #cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"' where go_no='"+gono+"'")
        #f.write("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NO_SM='未轉單' where order_no='"+gono+"'"+'\n')
        ccdmod.execute("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NO_SM='未轉單' where order_no='"+gono+"' ")
        chaincodedb.commit()
      if umo=='true':
        #cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',REMARK=REMARK+'*非正配單*' where go_no='"+gono+"'")
        #f.write("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NormalDelivery='非正配單' where order_no='"+gono+"'"+'\n')
        ccdmod.execute("update orderformpos set ArrivalTime='"+aday+" 00:00:00',NormalDelivery='非正配單' where order_no='"+gono+"'")
        chaincodedb.commit()
      else:
        '''
        cur206order.execute("select REMARK from [TGSalary].[dbo].[WEBORDERHD] where go_no='"+gono+"'")+"'")
        for r in cur206order.fetchall():
          f.write(str(r[0]))
          rk=str(r[0]).replace('非正配單','')
          f.write(rk)
        cur206order.execute("update [TGSalary].[dbo].[WEBORDERHD] set applydate='"+aday+"',REMARK='"+rk+"' where go_no='"+gono+"'")
        '''
        ccdmod.execute("update orderformpos set NormalDelivery='正配單' where order_no='"+gono+"'")
        chaincodedb.commit()
      #f.write("chaincodedb.commit()"+'\n')
      chaincodedb.commit()
    recount=0
    #cur206order.execute("select [GO_NO],[id_cust],[nm_c],[TOTAMT],[APPLYDATE],[remark],[SDATETIME],[no_sm] from [TGSalary].[dbo].[WEBORDERHD] where go_no='"+gono+"'") 
    #f.write("select order_no,AccountID,StoreName,Double_1,ArrivalTime,Remark,DateTime_1,no_sm from orderformpos where order_no='"+gono+"'"+'\n')	
    ccdmod.execute("select order_no,AccountID,StoreName,Double_1,ArrivalTime,Remark,DateTime_1,no_sm,NormalDelivery from orderformpos where order_no='"+gono+"'")
    for h in ccdmod:
      context['gono']=str(h[0])
      context['cid']=str(h[1])
      context['cname']=str(h[2])
      context['cmark']=str(h[5])
      context['sdate']=str(h[6])
      aday=str(h[4])
      context['Aday']=aday
      #f.write(str(h[7]))
      if str(h[7])=='已取消':
        context['CK3']='on'
      else:
        context['CK3']='off'
      if str(h[8]).find('非正配單')>-1:
        context['CK4']='on'
      else:
        context['CK4']='off'		
    #cur206order.execute("select a.new_iditem,a.nm_item,fl.[QTY],fl.[UPRICE],fl.[SUBTOT]  FROM [TGSalary].[dbo].[WEBORDERFL] fl,[TGSalary].[dbo].[webart] a where fl.go_no='"+gono+"' and a.id_item=fl.id_item  order by  a.new_iditem ") 
    ccdmod.execute("select ProdID,ProdName,Amount,Double_1,Double_2 from orderformpos_prod_sub where  order_no='"+gono+"' order by ProdID")
    #f.write("select ProdID,ProdName,Amount,Double_1,Double_2 from orderformpos_prod_sub where  order_no='"+gono+"'"+'\n')
    #f.write(str(recount))
    for o in ccdmod:
      tweborderfl=[]
      tweborderfl.append(str(o[0]))      
      tweborderfl.append(str(o[1]))
      tweborderfl.append(str(o[2]))
      tweborderfl.append(str(o[3]))
      tweborderfl.append(str(o[4]))
      #TF4211.append(str(o[4]))
      #TF4211.append(str(o[5]))
      #TF4211.append(str(o[5]))
      #TF4211.append(str(o[6]))
      #TF4211.append(str(o[7]))
      Sweborderfl.append(tweborderfl)
      recount=recount+1
    #f.write(str(Sweborderfl)+'\n')
    #頁籤    
    context['tabs']=cratetabs(12,recount)
    #頁籤
    #頁籤內容
    context['weborderfl']=tabsdata(12,Sweborderfl)
    #頁籤內容
    #f.write(str(weborderfl)) 
        
    context['reportmes']='<B>'+context['cname']+' 訂貨明細</B>'
    	
  except:
    s=''    
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday
    context['Eday'] = nday
  #f.close()
  return render(request, 'weborderdetel_cc.html',context )#傳入參數 
def pccss(request):  
  return render(request, 'pc.css', )
def tablecss(request):  
  return render(request, 'table.css', )
def spcss(request):  
  return render(request, 'sp.css', )
  
