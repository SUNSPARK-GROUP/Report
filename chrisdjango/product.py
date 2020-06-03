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
import time

import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8' 

depts=[]
accls=[]
global prodls 

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
def productsts(request):
  context= {}
  global prodls
    
  try:
    f=open(r'C:\Users\chris\chrisdjango\Perror.txt','w')
    product=[]
    sday=request.GET['Sday']#get values
    eday=request.GET['Eday']
    #RPDCT=request.GET['RPDCT']
    #context['sRPDCT']=RPDCT
    func=request.GET['funcname']
    imlitm1=request.GET['prods1']
    imlitm2=request.GET['prods2']
    context['Sday'] = sday
    context['Eday'] = eday
    context['funcname'] = func
    sd=getodate(sday[:4]+sday[5:7]+sday[8:10])
    ed=getodate(eday[:4]+eday[5:7]+eday[8:10])
    sday=sday[:4]+sday[5:7]+sday[8:10]
    eday=eday[:4]+eday[5:7]+eday[8:10]
    #f.write(sday+'~'+eday)
    connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
    connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
    context['imlitm1'] = imlitm1
    context['imlitm2'] = imlitm2
    #f.write(func+'\n')
    context['prodls']=prodls
    wb = Workbook()	
    ws1 = wb.active	
    ws1.title = "data"
    #ws1.append(['森邦(股)會計科目餘額明細('+mcu+'_'+macct+'_'+subacct+'_'+sday+'~'+eday+')'])
    
    title=[]
    nday=showday(0,'',0) #今天日期
    #f.write(nday)
    if func=='品牌產品目錄':
      title=['<th style="width:5%;">第二料號</th>','<th style="width:15%;">名稱</th>','<th style="width:15%;">供應商</th>','<th style="width:5%;">成本</th>','<th style="width:5%;">售價</th>','<th style="width:5%;">毛利%</th>','<th style="width:5%;">廠商編號</th>']
      #f.write(func)
      titlee=['第二料號','名稱','供應商','成本','售價','毛利%','廠商編號']	  
      getbra=connection206.cursor()
      #getbra2=connection206.cursor()
      prodl=connection214.cursor()
      #f.write('SELECT distinct(kindname) as kindname  FROM TGSalary.dbo.VIEWARTKIND')
      getbra.execute('SELECT distinct(kindname) as kindname  FROM TGSalary.dbo.VIEWARTKIND') 
      ds=0	
      dl=[]	  
      for b in getbra.fetchall():
        ds=ds+1
        title.append('<th style="width:5%;">'+b[0]+'</th>')
        titlee.append(b[0])
        dl.append(b[0])
        #f.write('<th style="width:5%;">'+b[0]+'</th>')
      #f.write(str(dl))
      ws1.append(titlee)
      prodl.execute("SELECT W.*,F6.abalph,f6.cban8,F6.CBPRRC,CASE WHEN F6.CBPRRC>'0' THEN ROUND(((W.PR_I-F6.CBPRRC)/W.PR_I*100),1)"
             +" ELSE '0' END AS SALEGP  FROM OPENQUERY(WEB206,'SELECT distinct(new_iditem) as new_iditem,nm_item,ID_ITEM,pr_i FROM TGSalary.dbo.VIEWARTKIND "
             +" group by id_item,nm_item,pr_i,PR_CSTD,new_iditem  order by new_iditem' ) W,"
             +" OPENQUERY(e910,'SELECT abalph,CBAN8,CBLITM,CBPRRC*1.05 as CBPRRC FROM PRODDTA.VS_F41061 WHERE  CBEXDJ>=''"+nday+"''') F6 "
             +" WHERE F6.CBLITM= W.NEW_IDITEM  ORDER BY W.NEW_IDITEM  ") 
      
      for p in prodl.fetchall():
        tproduct=[]
        tproduct.append(str(p[0]))
        tproduct.append(str(p[1]))
        tproduct.append(str(p[4]))
        tproduct.append(str(p[6]))
        tproduct.append(str(p[3]))        
        tproduct.append(str(p[7]))
        tproduct.append(str(p[5]))
        #f.write("SELECT kindname FROM TGSalary.dbo.VIEWARTKIND where new_iditem = '"+str(p[0])+"'")
        getbra.execute("SELECT kindname FROM TGSalary.dbo.VIEWARTKIND where new_iditem = '"+str(p[0])+"'")
        dlt=[]
        for b2 in getbra.fetchall():
          #f.write(b2[0])
          dlt.append(dl.index(str(b2[0])))
          #f.write(dl.index(str(b2[0]))+'\n')
        l=0
        #f.write(str(dlt)+'\n')		  
        #f.write(str(ds))
        while l < ds:
          try:
            if dlt.index(l)>-1:
              tproduct.append('*')
          except :
            tproduct.append('')
          l=l+1			
        product.append(tproduct)
        ws1.append(tproduct)
    if func=='日統計':
      prodl=connection214.cursor()
      title=['<th style="width:13%;">第二料號</th>','<th style="width:10%;">日期</th>','<th style="width:48%;">名稱</th>','<th style="width:10%;" >數量</th>']
      #ws1.append(['第二料號','日期','短料號','名稱','數量'])
      ws1.append(['第二料號','日期','名稱','數量'])
      f.write("Select f.sdlitm as '第二料號',h.SHDRQJ as '日期',f41.IMDSC1 as '名稱',sum(CAST(f.qty as decimal(18,0)) )as '數量' " 
                    +"from (Select sdlitm,sditm, SDDSC1,sddoco,sddcto,sdan8,qty from vs_salefl) f,(select * from vs_salehd where  SHDRQJ>='"+sday+"' and SHDRQJ<='"+eday+"') h " 
                    +",(SELECT NEW_IDITEM AS IMLITM, NM_ITEM as IMDSC1 ,litm as IMITM  FROM [ERPS].[dbo].[LayaOracleArt]) f41 "
                    +" where f.sddoco=h.shdoco and f.sddcto=h.shdcto and f.sdan8=h.SHPA8 AND F41.IMLITM=F.sdlitm and f.sdlitm>='"+imlitm1+"' and f.sdlitm<='"+imlitm2+"' "
                    +" group by h.SHDRQJ,f.sditm,f.SDDSC1,f.sdlitm,f41.IMDSC1   order by f.sdlitm ,h.SHDRQJ ")
      prodl.execute("Select f.sdlitm as '第二料號',h.SHDRQJ as '日期',f.sditm as '短料號',f41.IMDSC1 as '名稱',sum(CAST(f.qty as decimal(18,0)) )as '數量' " 
                    +"from (Select sdlitm,sditm, SDDSC1,sddoco,sddcto,sdan8,qty from vs_salefl) f,(select * from vs_salehd where  SHDRQJ>='"+sday+"' and SHDRQJ<='"+eday+"') h " 
                    +",(SELECT NEW_IDITEM AS IMLITM, NM_ITEM as IMDSC1 ,litm as IMITM  FROM [ERPS].[dbo].[LayaOracleArt]) f41 "
                    +" where f.sddoco=h.shdoco and f.sddcto=h.shdcto and f.sdan8=h.SHPA8 AND F41.IMLITM=F.sdlitm and f.sdlitm>='"+imlitm1+"' and f.sdlitm<='"+imlitm2+"' "
                    +" group by h.SHDRQJ,f.sditm,f.SDDSC1,f.sdlitm,f41.IMDSC1   order by f.sdlitm ,h.SHDRQJ ")	
      	  
      for p in prodl.fetchall():
        tproduct=[]
        tproduct.append(str(p[0]))
        tproduct.append(str(p[1]))
        #tproduct.append(str(p[2]))
        tproduct.append(str(p[3]))
        tproduct.append(str(p[4]))
        product.append(tproduct)
        ws1.append(tproduct)
    if func=='週統計':
      prodl=connection214.cursor()
      title=['<th style="width:13%;">第二料號</th>','<th style="width:20%;">日期</th>','<th style="width:40%;">名稱</th>','<th style="width:10%;">數量</th>']	  
      ws1.append(['第二料號','日期','名稱','數量'])
      sy=sday[:4]
      ey=eday[:4]
      prodl.execute("SELECT '"+sy+"' as myear,Datepart(wk, DATE) AS wk,CONVERT(varchar(100), Min(DATE), 111)mindate,CONVERT(varchar(100),Max(DATE), 111) maxdate FROM   (SELECT Dateadd(dd, NUMBER, Cast(Ltrim('"+sy+"') + '0101' "
                    +" AS DATETIME)) date FROM   master..spt_values WHERE  [TYPE] = 'p' AND NUMBER BETWEEN 0 AND 364) s   where  Year(DATE) = '"+sy+"' GROUP  BY Datepart(wk, DATE) " 
                    +" union "
                    +" SELECT '"+ey+"' as myear,Datepart(wk, DATE) AS wk,CONVERT(varchar(100), Min(DATE), 111)mindate,CONVERT(varchar(100),Max(DATE), 111) maxdate FROM   (SELECT Dateadd(dd, NUMBER, Cast(Ltrim('"+ey+"') + '0101' "
                    +" AS DATETIME)) date FROM   master..spt_values WHERE  [TYPE] = 'p' AND NUMBER BETWEEN 0 AND 364) s   where  Year(DATE) =  '"+ey+"' GROUP  BY Datepart(wk, DATE) order by mindate ")
      dl=[]
      for i in prodl.fetchall():
        tdl=[]
        tdl.append(str(i[0]))
        tdl.append(str(i[1]))
        tdl.append(str(i[2]))
        tdl.append(str(i[3]))
        dl.append(tdl)
      prodl.execute("select y.myear,y.第二料號,y.日期,y.短料號,f41.IMDSC1 as '名稱',y.數量 from (select myear,sdlitm as '第二料號',SHDRQJ as '日期',sditm as '短料號' "
                    +",sum(qty)as '數量' from (Select substring(h.SHDRQJ,1,4) as myear,f.sdlitm as 'sdlitm',datepart(week,h.SHDRQJ) as 'SHDRQJ',f.sditm as 'sditm',"
                    +"sum(CAST(f.qty as decimal(18,0)) )as 'qty' from (Select sdlitm,sditm, SDDSC1,sddoco,sddcto,sdan8,qty from vs_salefl "
                    +" where SdDRQJ>='"+sday+"' and SdDRQJ<='"+eday+"' and sdlitm>='"+imlitm1+"' and sdlitm<='"+imlitm2+"' ) f,"
                    +" (select * from vs_salehd where   SHDRQJ>='"+sday+"' and SHDRQJ<='"+eday+"') h  where f.sddoco=h.shdoco and f.sddcto=h.shdcto and f.sdan8=h.SHPA8 "
                    +" group by h.SHDRQJ,f.sditm,f.SDDSC1,f.sdlitm ) a group by myear,SHDRQJ,sditm,sdlitm) y, (SELECT NEW_IDITEM AS IMLITM, NM_ITEM as IMDSC1 ,litm as IMITM "
                    +" FROM [ERPS].[dbo].[LayaOracleArt])  f41 where  F41.IMLITM=y.第二料號  order by y.第二料號 ,y.myear ,y.日期 ")
      for p in prodl.fetchall():
        tproduct=[]
        tproduct.append(str(p[1]))
        for l in range(len(dl)):
          if str(p[0])==dl[l][0] and str(p[2])==dl[l][1]:
            tproduct.append(dl[l][2]+"~"+dl[l][3])
        #tproduct.append(str(p[3]))
        tproduct.append(str(p[4]))
        tproduct.append(str(p[5]))
        product.append(tproduct)        
        ws1.append(tproduct)
    if func=='月統計':
      pname=[]
      pds=CONORACLE("select distinct(TRIM(imlitm)) as imlitm,imitm,TRIM(imdsc1) from proddta.VS_21UNCS where TRIM(imlitm)>='"+imlitm1+"' and TRIM(imlitm)<='"+imlitm2+"'")
      for l in pds:
        tpname=[]
        tpname.append(str(l[0]))
        tpname.append(str(l[1]))
        tpname.append(str(l[2]))
        pname.append(tpname)
      prodl=connection214.cursor()
      title=['<th style="width:15%;">第二料號</th>','<th style="width:12%;">短料號</th>','<th style="width:65%;">名稱</th>']
      titlee=['第二料號','短料號','名稱']
      #f.write(str(titlee)+'\n')
      prodl.execute("select distinct(a.日期) from (Select f.sdlitm as '第二料號',substring(F.SDDRQJ,1,6) as '日期',f.sditm as '短料號',sum(CAST(f.qty as decimal(18,0)) )as '數量' "
              +"from (Select SDDRQJ,sdlitm,sditm, SDDSC1,sddoco,sddcto,sdan8,qty from vs_salefl where  SDDRQJ>='"+sday+"' and SDDRQJ<='"+eday+"') f "
              +" where   f.sdlitm>='"+imlitm1+"' and f.sdlitm<='"+imlitm2+"' group by substring(F.SDDRQJ,1,6),f.sditm,f.sdlitm   ) a order by a.日期 ")
      mls=[]
      for d in prodl.fetchall():
        mls.append(str(d[0]))
      gp=int(50/len(mls))
      for t in range(len(mls)):
        title.append('<th style="width:'+str(gp)+'%;">'+str(mls[t])+'</th>')
        titlee.append(str(mls[t]))
      ws1.append(titlee)
            
      prodl.execute("Select f.sdlitm as '第二料號',substring(F.SDDRQJ,1,6) as '日期',f.sditm as '短料號',sum(CAST(f.qty as decimal(18,0)) )as '數量' "
              +"from (Select SDDRQJ,sdlitm,sditm, SDDSC1,sddoco,sddcto,sdan8,qty from vs_salefl where  SDDRQJ>='"+sday+"' and SDDRQJ<='"+eday+"') f "
              +" where   f.sdlitm>='"+imlitm1+"' and f.sdlitm<='"+imlitm2+"' group by substring(F.SDDRQJ,1,6),f.sditm,f.sdlitm   order by 第二料號,日期 ")
      product=[]
      sls=[]
      lpname=''
      #f.write(str(pname)+'\n')
      for p in prodl.fetchall():
        tp=[]
        tsls=[]
        tp.append(str(p[0]))
        tsls.append(str(p[0]))
        tsls.append(str(p[1]))
        tsls.append(str(p[3]))
        tp.append(str(p[2]))
        for n in range(len(pname)):
          if pname[n][0]==str(p[0]):
            tp.append(pname[n][2])
            break
        for m in range(len(mls)):
          tp.append('0')
        if lpname!=str(p[0]):
          product.append(tp)
        lpname=str(p[0])
        sls.append(tsls)
      #f.write(str(product)+'\n')
      #f.write(str(sls)+'\n')
      for s in range(len(sls)):
        lno=sls[s][0]
        ldate=sls[s][1]
        lq=sls[s][2]
        #f.write(lno+','+ldate+','+lq+'\n')
        for lp in range(len(product)):
          if product[lp][0]==lno:
            mi=0
            for lm in range(len(mls)):
              if mls[lm]==ldate:
                try:
                  product[lp][mi+3]=lq #推移三欄
                except:
                  product[lp].append(lq)
              mi=mi+1
      #f.write(str(product)+'\n')
      for e in range(len(product)):
        ws1.append(product[e])
      #f.write('end pro'+'\n')
    if func=='平均用量':
      
      wd=CONORACLE("select WED, count(days)   as cc from (select days,TO_CHAR(to_date(days,'yyyymmDD'), 'D') AS WED  from dual,(SELECT TO_CHAR(to_date('"
	               +sday+"', 'yyyymmdd') + (level - 1), 'yyyymmdd')  as days  FROM dual   CONNECT BY TRUNC(to_date('"+sday
				   +"', 'yyyymmdd')) + level - 1 <= TRUNC(to_date('"+eday+"', 'yyyymmdd')))  ) group by wed order by wed ")
      SSD=0 #週日天數
      AD=0 #總天數
      RD=0 #有效天數
      for w in wd:
        AD=AD+int(w[1])
        if w[0]==1 :
          SSD=SSD+int(w[1])
        else:
          RD=RD+int(w[1])
      prodl=connection214.cursor()
      
      prodl.execute(" select sdlitm,sditm,art.NM_ITEM,'',sum(CAST(f.qty as decimal(18,0)) )as '總用量' , round(sum(CAST(f.qty as decimal(18,0)) )/"+str(RD)+",1) as 日用量  from "
                  +"( Select sdlitm,sditm,SDDSC1, sddoco,sddcto,sdan8,qty from vs_salefl f,(select shdoco,SHPA8 from vs_salehd where  SHDRQJ>='"+sday+"' and SHDRQJ<='"+eday+"') h"
				  +" where f.sdlitm>='"+imlitm1+"' and f.sdlitm<='"+imlitm2+"' and f.SDDOCO =h.shdoco and f.sdan8=h.SHPA8) f , (select * from LayaOracleArt) art"
				  +"  where  art.NEW_IDITEM=f.sdlitm group by f.sditm,f.sdlitm,art.NM_ITEM order by f.sdlitm ")
      title=['<th style="width:15%;">第二料號</th>','<th style="width:55%;">名稱</th>','<th style="width:15%;">總用量</th>','<th style="width:15%;">日用量</th>']				  
      ws1.append(['第二料號','名稱','總用量','日用量'])
      #f.write(str(title))
      for p in prodl.fetchall():
        tp=[]
        tp.append(str(p[0]))
        #tp.append(str(p[1]))
        tp.append(str(p[2]))
        '''
        if p[3]==None:
          tp.append('')
        else:
          tp.append(str(p[3]))
        '''
        tp.append(str(p[4]))
        tp.append(str(round(p[5],1)))
        product.append(tp)	
        ws1.append(tp)	
    context['title'] = title		
    context['product'] = product
    ntime=str(time.strftime("%X"))
    ntime=ntime[:2]+ntime[3:5]+ntime[6:8]
    #f.write(ntime)
    context['efilename']='商品查詢結果'+ntime+'.xlsx'
    #f.write("C:\\Users\\Administrator\\chrisdjango\\MEDIA\\商品查詢結果"+ntime+".xlsx")
    wb.save("C:\\Users\\Administrator\\chrisdjango\MEDIA\\商品查詢結果"+ntime+".xlsx")#正式路徑
    
  except:
    nday=showday(0,'-',0) #今天日期
    context['Sday'] = nday[:8]+'01'
    context['Eday'] = nday
    context['funcname'] = ''    
    prodls=[]
    jdepro=CONORACLE("select distinct(imlitm) as imlitm,imdsc1 from proddta.VS_21UNCS order by imlitm")
    for a in jdepro:
      iml=str(a[0]).replace(' ','')
      imd=str(a[1]).replace(' ','')
      tpl={'imlitm':iml,'imdsc1':imd}
      prodls.append(tpl)
    context['prodls']=prodls
  return render(request, 'product//productsts.html',context )#傳入參數
  f.close()
def OracleSalesCalc(request):
  context= {}
  connection214=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.214;DATABASE=erps;UID=apuser;PWD=0920799339')
  area=connection214.cursor()
  f=open(r'C:\Users\chris\chrisdjango\OOerror.txt','w')
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
      area.execute("Select String_20_1+'|'+String_50_1 from OPENQUERY(MYSQL,'SELECT * FROM basicstoreinfo where  String_500_3=''"+an21[:an21.find('|')]+"''')")      
      for a in area.fetchall():
        shop.append(str(a[0]))
    context['shop0']= shop#門市list
    shopname=request.GET['shop'][:request.GET['shop'].find('|')]#門市名稱 
    context['shop01']=request.GET['shop']		
    if sc=='1':
      osaledata=[]
      if shopname!='':  
        title=['<th style="width:10%;">門市代碼</th>','<th style="width:15%;">門市名稱</th>','<th style="width:10%;">負責業務</th>','<th style="width:20%;">產品名稱</th>','<th style="width:15%;">產品代碼</th>','<th style="width:20%;">銷售金額</th>','<th style="width:10%;">數量</th>']
        f.write("select o.abalky as '門市代碼',o.ABALPH as '門市名稱',lv.EMPNAME1 as '負責業務',o.SDDSC1 as '產品名稱',o.SDLITM as '產品代碼',sum(o.DTTAEXP) as '銷售金額',SUM(o.qty) as '數量' from "
            +" (select distinct(USERID) as USERID,EMPNAME1 from LayaViewAreaUser  ) lv,(select h.shshan,h.abalky,h.ABALPH,f.SDLITM,f.sddsc1,CAST(f.DTTAEXP as decimal(18,0)) as DTTAEXP,CAST(f.qty as decimal(18,0)) as qty from "
            +" (select * from vs_salehd where    (shdcto='S2' OR shdcto='S3' OR shdcto='S4') and SHDRQJ >= '"+sd+"' and SHDRQJ <= '"+ed+"') h,(select * from vs_salefl where sdlitm>='100000' and sdlitm<='399999' "
			#20190927 取消 +" and sdlitm not in ('100004')" 
            +" and   (sddcto='S2' OR sddcto='S3' OR sddcto='S4') AND SDDRQJ >= '"+sd+"' and SDDRQJ <= '"+ed+"') f where f.sddoco=h.shdoco and f.sdan8=h.shshan "+firstr+"  ) o "
            +" where o.abalky=lv.USERID  and o.abalky='"+shopname+"' group by o.abalky,o.ABALPH,lv.EMPNAME1,o.SDLITM,o.SDDSC1  order by o.abalky ")
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
  f.close()
  return render(request, 'product//OracleSalesCalc.html',context )#傳入參數

def pccss(request):  
  return render(request, 'pc.css', )
def tablecss(request):  
  return render(request, 'table.css', )
def spcss(request):  
  return render(request, 'sp.css', )
  
