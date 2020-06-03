# -*- encoding: UTF-8 -*-
import cx_Oracle
import pyodbc
from datetime import date
from datetime import timedelta
import datetime
import string
import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
def showday(wd,sp,dt):#wd->0 today,wd->1 yesterday,wd->-1 tomorrow sp->Divider dt datetype
    t1 = 0-wd
    d=date.today()-timedelta(t1)
    if dt==1 :
      return d.strftime('%d'+sp+'%m'+sp+'%Y')#%Y->2015 %y->15
    else :
      return d.strftime('%Y'+sp+'%m'+sp+'%d')#%Y->2015 %y->15
def CONORACLE(SqlStr):
	hostname='192.168.0.230'
	sid='E910'
	username='PRODDTA'
	password='E910Jde'
	port='1521'
	dsn = cx_Oracle.makedsn(hostname, port, sid)
	conn = cx_Oracle.connect(username+'/'+password+'@' + dsn)
	cursor = conn.cursor()
	#print (SqlStr)
	cursor.execute(SqlStr)
	SQLSTRS = SqlStr[0:6].upper()
	if SQLSTRS=="SELECT":
	  TotalSession = cursor.fetchall()
	  return TotalSession
	  cursor.close()
	else: conn.commit()
def set2jde(tday):
  #OR_DATAID='CRPDTA'#測試區
  OR_DATAID='PRODDTA'#正式區
  #1.搜尋當天的單
  sdate = showday(tday,'',0)
  ndate = showday(0,'',0)
  #sdate='20190225'
  #print (sdate)
  #f=open('set2jde.txt','w')
  connection206=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
  connection2061=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
  connection2062=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
  connection206t=pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339')
  cur206hd = connection206.cursor()
  cur206temp = connection206t.cursor()
  cur2061hd = connection2061.cursor()
  cur2061fl = connection2062.cursor()
  cur206Bhd = connection2061.cursor()#BOOKING_HD 
  cur206Bfl = connection2062.cursor()#BOOKING_FL
  #f.write('DELETE'+'\n')
  #ORADB = CONORACLE("DELETE  "+OR_DATAID+".F47012 ")#手動轉單不能刪除，會影響其他人轉單
  #ORADB = CONORACLE("DELETE  "+OR_DATAID+".F47011 ")
  salel=[]
  #f.write("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust<>'appdv' ORDER BY ID_CUST,GO_NO ") 
  #cur206hd.execute("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust in ('appdv','chrisk') ORDER BY ID_CUST,GO_NO ") 
  cur206hd.execute("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust<>'appdv' ORDER BY ID_CUST,GO_NO ") 
  
  for sales in cur206hd :
    hdl=[]
    hdl.append(str(sales[0]))
    hdl.append(str(sales[1]))
    hdl.append(str(sales[2]))
    hdl.append(str(sales[3]))
    hdl.append(str(sales[4]))
    salel.append(hdl)
    go_no=str(sales[0])
    #print (go_no)
    #print(str(sales[2]))
    cur2061hd.execute("SELECT A.go_no,a.no_sm,a.id_cust,a.applydate,a.sdatetime,convert(nvarchar(200),a.remark) as remark,B.USERID,B.TYPENO,"
                    +"B.SALETYPE,IsNull(B.BONNER,0) AS USERBON,ISNULL(B.CARNO,'')CARNO,B.USERIDNEW ,a.aDATE "
                    +"FROM (SELECT GO_NO,NO_SM,ID_CUST, '1'+substring(APPLYDATE,3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,APPLYDATE) "
                    +"as NVARCHAR), 3) as APPLYDATE,'1'+substring('"+ndate+"',3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,'"
                    +ndate+"') as NVARCHAR), 3) as aDATE,SDATETIME,REMARK FROM WEBORDERHD WHERE GO_NO='"
                    +str(sales[0])+"')A,USERGROUP B WHERE A.ID_CUST=B.USERID")  
    #print "SELECT count(A.*) as q FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+str(sales[0])+"')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ORDER BY B.ID_ITEM"
    cur2061fl.execute("SELECT count(*) as q FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ")
    for flcount in cur2061fl:
      dnos=int(flcount[0])
    for sale in cur2061hd:
      if str(sale[1])=='None':#開始轉單
        cur206temp.execute("select count(*) from WEBBookingOrder where ID_CUST='"+str(sale[2])+"' and applydate ='"+str(sale[3])+"'")#20191031 查詢預購商品清單，加入總數量
        for bcount in cur206temp:
          dnos=dnos+int(bcount [0])
        #####sale head#####
        cid=str(sale[2])#20191031
        nowdate=str(sale[12])
        ordate=str(sale[3])
        apdate=ordate
        yy=sales[0][4:6]
        #print yy
        #print ndate
        #備註
        remarkt=(sale[5])
        remark=''
        for r in range(len(remarkt)):
          if ord(remarkt[r])<62995:
            remark=remark+remarkt[r] 
        remark1=''
        remark2=''
        if len(remark)>15:
          remark1=remark[:15]
          remark2=remark[15:len(remark)]      
        else: remark1=remark      
        #print remark1+'/'+remark2
        #user_id
        if str(sale[11])=='None' : uid=str(sale[2])
        else: uid=str(sale[11])
        #jde an8
        ORADB = CONORACLE("Select TO_CHAR(ABAN8) as ID_Cust, TO_CHAR(ABALPH) as NM_C  FROM "+OR_DATAID+".F0101 WHERE ABAT1='C'  and ABALKY ='"+uid+"'")
        adid=''
        for ds in ORADB:
          adid=str(ds[0])
        #print ouserid
        if dnos>10:
          zon=' '     #全聯
        else:
          zon='1';    #半聯
        cur206temp.execute("SELECT count(go_no) as tnos FROM  WEBORDERHD WHERE GO_NO like 'WS20"+yy+"%' AND len(NO_SM) >0")     
        for gono in cur206temp:
          n = str(gono[0]+1)
          TNOS = yy+n.zfill(6)        
        #print TNOS
        #print "INSERT INTO "+OR_DATAID+".F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOO"+",SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) "+" VALUES ('1','1','00100','"+TNOS+"','E1','1000','850','"+nowdate+"','R','"+str(dnos)+"','N','00','00100','S2','        A001','00100','00100','"+TNOS+"','E1','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+remark1+"','"+remark2+"','"+TNOS+"','"+zon+"')"
          
        ORADB=CONORACLE("INSERT INTO "+OR_DATAID+".F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOO"
                      +",SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) "
                      +" VALUES ('1','1','00100','"+TNOS+"','E1','1000','850','"+nowdate+"','R','"+str(dnos)+"','N','00','00100','S2','        A001','00100','00100','"
                      +TNOS+"','E1','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+remark1+"','"+remark2+"','"+TNOS+"','"+zon+"')")      
        cur206temp.execute("update WEBORDERHD set NO_SM='"+TNOS+"' where GO_NO ='"+go_no+"'")
        cur206temp.commit()
        #查出下次出貨日
        cur206temp.execute("SELECT h.[GO_NO],h.[ID_CUST],h.[NM_C],substring(h.[APPLYDATE],1,8) as odate "
                           +",w.dldate1-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1) as gwd1"
                           +",w.dldate2-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1) as gwd2"
                           +",w.dldate3-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1) as gwd3"
                           +",w.dldate1-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1)+7 as gwd4"
                           +",w.dldate2-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1)+7 as gwd5"
                           +",w.dldate3-DATEPART(WEEKDAY, convert(datetime, substring(h.APPLYDATE,1,8), 101)-1)+7 as gwd6"
                           +"  ,'1'+substring(APPLYDATE,3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,h.APPLYDATE)" 
                           +"as NVARCHAR), 3) as applydate FROM [TGSalary].[dbo].[WEBORDERHD] h,"
                           +"  ( SELECT userid,dldate1,dldate2,dldate3  FROM [TGSalary].[dbo].[USERGROUP] where [FG_ACTIVE]='1' ) w"
                           +"    where h.go_no='"+go_no+"'  and h.id_cust in (SELECT [USERID]" 
                           +"FROM [TGSalary].[dbo].[USERGROUP] where [FG_ACTIVE]='1'  )  and w.userid=h.id_cust")
        for t in cur206temp.fetchall():
          if t[5]!=None and int(t[5])>0:
            adays=int(t[5])
          elif t[6]!=None and int(t[6])>0:
            adays=int(t[6])
          elif t[7]!=None and int(t[7])>0:
            adays=int(t[7])
          elif t[8]!=None and int(t[8])>0:
            adays=int(t[8])          
          elif t[9]!=None and int(t[9])>0:
            adays=int(t[9])
          #apdateb=int(t[10])+int(t[9])
        t_str =str(t[3])#本次出貨日
        d = datetime.datetime.strptime(t_str, '%Y%m%d')
        delta = datetime.timedelta(days=adays)#下一次出貨日天數
        n_days = d + delta
        apdateb=n_days.strftime('%Y%m%d')#下一次出貨日
		#查出下次出貨日
        #####sale items######
        #cur2061fl.execute("SELECT A.*,B.NEW_IDITEM,B.IFSTOCK FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ORDER BY B.ID_ITEM")
        cur2061fl.execute("SELECT A.*,B.NEW_IDITEM,B.IFSTOCK FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A,WEBART B " 
                        +"WHERE A.ID_ITEM=B.ID_ITEM and A.id_item  not in (SELECT  [id_item]  FROM [TGSalary].[dbo].[WEBBookingART]) "
                        +" union SELECT '"+go_no+"',A.[ID_ITEM],A.[NM_ITEM],A.[QTY],A.[UPRICE],A.[SUBTOT],B.NEW_IDITEM,B.IFSTOCK "
                        +"FROM (SELECT * FROM [WEBBookingOrder] WHERE [applydate]='"+t_str+"' and [ID_CUST]='"+cid+"' and resale='N' )A,WEBART B  "
                        +" WHERE A.ID_ITEM=B.ID_ITEM ORDER BY ID_ITEM")
        I12=0
        for sitem in cur2061fl:
          I12+=1
          if sitem[6]=='None':
            NewItemID='A'
          else:
            NewItemID=str(sitem[6]).upper()
          if sitem[7]=='N':
            Branch_Plan=' '
          else:
            Branch_Plan='        A001'
          qty=int(sitem[3])*10000
          #print (qty)
          #print NewItemID
          #print str(sitem[3])
          #print "INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"+",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "+" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"+",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+NewItemID+"','','','','','','"+str(sitem[3])+"0000','"+TNOS+"'"
          #print "INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"+",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "+" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"+",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+NewItemID+"','','','','','',"+str(qty)+",'"+TNOS+"')"
          
          ORADB=CONORACLE("INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"
                        +",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "
                        +" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"
                        +",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate
                        +"','"+NewItemID
                        +"','','','','','',"+str(qty)+",'"+TNOS+"')")
        #查詢是否有預購商品有則新增到 WEBBookingOrder
        #print("SELECT A.* FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A WHERE  A.id_item  in (SELECT  [id_item]  FROM [TGSalary].[dbo].[WEBBookingART])  ")
        cur2061fl.execute("SELECT A.* FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A WHERE  A.id_item  in (SELECT  [id_item]  FROM [TGSalary].[dbo].[WEBBookingART])  ")
        for sitem in cur2061fl:
          itemid=str(sitem[1])
          itemnm=str(sitem[2])
          iqty=str(sitem[3])
          iuprice=str(sitem[4])
          isubtot=str(sitem[5])
          cur206temp.execute("insert into [TGSalary].[dbo].[WEBBookingOrder] ([ID_ITEM],[NM_ITEM],[QTY],[UPRICE],[SUBTOT],[resale],[applydate],[ID_CUST]) VALUES "
                            +"('"+itemid+"','"+itemnm+"','"+iqty+"','"+iuprice+"','"+isubtot+"','N','"+str(apdateb)+"','"+cid+"')")
          cur206temp.commit()
        #查詢是否有預購商品有則新增到 WEBBookingOrder
        cur206temp.execute(" UPDATE [WEBBookingOrder] set  resale='Y' WHERE [applydate]='"+t_str+"' and [ID_CUST]='"+cid+"' and resale='N' ")#20191031
        cur206temp.commit()
  #f.close()
  return salel
        
