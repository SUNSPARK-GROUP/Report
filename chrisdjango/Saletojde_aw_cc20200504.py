# -*- encoding: UTF-8 -*-
import cx_Oracle
import pyodbc
from datetime import date
from datetime import timedelta
import datetime
import string
import os
import pymysql
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
def getday(y,m,d,n):
  the_date = datetime.datetime(y,m,d)
  result_date = the_date + datetime.timedelta(days=n)
  d = result_date.strftime('%Y-%m-%d')
  return d
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
  sdate = showday(tday,'-',0)
  ndate = showday(0,'',0)
  #sdate='20190509'
  #print (sdate)
  f=open('set2jde.txt','w')
  f.write(sdate+'\n')
  '''
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
  '''
  #f.write('ccdmodhd'+'\n')
  chaincodehd=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccdmodhd = chaincodehd.cursor()  
  chaincodefl=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccdmodfl = chaincodefl.cursor()
  chaincodedb=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccdmod = chaincodedb.cursor()
  chaincodet=pymysql.connect(host='192.168.0.218', port=3306, user='root', passwd='TYGHBNujm', db='ccerp_tw001114hq',charset='utf8')
  ccdmodt = chaincodet.cursor()
  #f.write('DELETE'+'\n')
  #ORADB = CONORACLE("DELETE  "+OR_DATAID+".F47012 ")#手動轉單不能刪除，會影響其他人轉單
  #ORADB = CONORACLE("DELETE  "+OR_DATAID+".F47011 ")
  salel=[]
  #f.write("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust in ('appdv','chrisk') ORDER BY ID_CUST,GO_NO ") 
  #cur206hd.execute("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust in ('appdv','chrisk') ORDER BY ID_CUST,GO_NO ") 
  #cur206hd.execute("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='"+sdate+"'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust<>'appdv' ORDER BY ID_CUST,GO_NO ") 
  #cur206hd.execute("SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE GO_NO='WS20191200289' ORDER BY ID_CUST,GO_NO ")
  #f.write("1. select order_no,AccountID,StoreName,Double_1,Remark from  orderformpos  where  NO_SM='未轉單' and ArrivalTime like '"+sdate+"%'"+'\n')
  ccdmod.execute("select order_no,AccountID,StoreName,Double_1,Remark from  orderformpos  where  NO_SM='未轉單' and ArrivalTime like '"+sdate+"%'")

  for sales in ccdmod :
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
    '''
    cur2061hd.execute("SELECT A.go_no,a.no_sm,a.id_cust,a.applydate,a.sdatetime,convert(nvarchar(200),a.remark) as remark,B.USERID,B.TYPENO,"
                    +"B.SALETYPE,IsNull(B.BONNER,0) AS USERBON,ISNULL(B.CARNO,'')CARNO,B.USERIDNEW ,a.aDATE "
                    +"FROM (SELECT GO_NO,NO_SM,ID_CUST, '1'+substring(APPLYDATE,3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,APPLYDATE) "
                    +"as NVARCHAR), 3) as APPLYDATE,'1'+substring('"+ndate+"',3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,'"
                    +ndate+"') as NVARCHAR), 3) as aDATE,SDATETIME,REMARK FROM WEBORDERHD WHERE GO_NO='"
                    +str(sales[0])+"')A,USERGROUP B WHERE A.ID_CUST=B.USERID")
    '''					
    ccdmodhd.execute("select order_no,no_sm,AccountID,CONCAT('1',substring(ArrivalTime,3,2),LPAD(LTRIM(CAST(DAYOFYEAR(ArrivalTime) AS CHAR)),3,'0')) as apdate,DateTime_1,NormalDelivery,AccountID"
                     +",AccountType,'','','',AccountID,CONCAT('1',substring('"+ndate+"',3,2),LPAD(LTRIM(CAST(DAYOFYEAR('"+ndate+"') AS CHAR)),3,'0')) as ndate"
                     +",substring(ArrivalTime,1,10) ArrivalTime,remark from orderformpos where  order_no='"+str(sales[0])+"'")
    
    f.write("2. select order_no,no_sm,AccountID,CONCAT('1',substring(ArrivalTime,3,2),LPAD(LTRIM(CAST(DAYOFYEAR(ArrivalTime) AS CHAR)),3,'0')) as apdate,DateTime_1,NormalDelivery,AccountID"
                     +",AccountType,'','','',AccountID,CONCAT('1',substring('"+ndate+"',3,2),LPAD(LTRIM(CAST(DAYOFYEAR('"+ndate+"') AS CHAR)),3,'0')) as ndate"
                     +",substring(ArrivalTime,1,10) ArrivalTime,remark from orderformpos where  order_no='"+str(sales[0])+"'"+'\n')
    
    #print "SELECT count(A.*) as q FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+str(sales[0])+"')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ORDER BY B.ID_ITEM"
    ccdmodfl.execute("SELECT count(*) as q FROM orderformpos_prod_sub WHERE order_no='"+str(sales[0])+"' ")
    #f.write("3. SELECT count(*) as q FROM orderformpos_prod_sub WHERE order_no='"+str(sales[0])+"' "+'\n')
    for flcount in ccdmodfl:
      odnos=int(flcount[0])
    for sale in ccdmodhd:
      f.write(str(sale[1]))
      #f.write(str(odnos+' : '+indate+' : '+str(sale[5])+'\n'))
      if str(sale[1])=='未轉單':#開始轉單
        dnos=odnos
        indate=str(sale[13])#20191212
        f.write(str(dnos)+' : '+indate+' : '+str(sale[5])+'\n')
        if str(sale[5]).find('非正配單')==-1:#20191115
          #cur206temp.execute("select count(*) from WEBBookingOrder where ID_CUST='"+str(sale[2])+"' and applydate <='"+str(sale[3])+"'  and resale='N' ")#20191031 查詢預購商品清單，加入總數量 20191112 and applydate ='"+str(sale[3])+"'
          f.write("4. select count(*) from WEBBookingOrder where ID_CUST='"+str(sale[2])+"' and applydate <='"+str(sale[13])+"'  and resale='N' "+'\n')
          ccdmodt.execute("select count(*) from WEBBookingOrder where ID_CUST='"+str(sale[2])+"' and applydate <='"+str(sale[13])+"'  and resale='N' ")#20191031 查詢預購商品清單，加入總數量 20191112 and applydate ='"+str(sale[3])+"'

          for bcount in ccdmodt:
            dnos=odnos+int(bcount [0])
        #####sale head#####
        cid=str(sale[2])#20191031
        nowdate=str(sale[12])
        ordate=str(sale[3])
        apdate=ordate
        yy=sales[0][4:6]
        f.write(cid+' : '+nowdate+' : '+ordate+' : '+apdate+' : '+yy+'\n')
        #print yy
        #print ndate
        #備註
        remarkt=(sale[14])
        f.write(cid+' : '+nowdate+' : '+ordate+' : '+apdate+' : '+remarkt+'\n')
        remark=''
        for r in range(len(remarkt)):
          if ord(remarkt[r])<62995:
            remark=remark+remarkt[r]         
        remark1=''
        remark2=''
        if len(remark)>12 and len(remark)<=24:
          remark1=remark[:12]
          remark2=remark[12:len(remark)]
        elif len(remark)>24:
          remark1=remark[:12]
          remark2=remark[12:24]
        else: remark1=remark		
        #print remark1+'/'+remark2
        #user_id
        f.write(cid+' : '+nowdate+' : '+ordate+' : '+apdate+' : '+yy+' : '+' : '+remark1+'\n')
        if str(sale[11])=='None' : uid=str(sale[2])
        else: uid=str(sale[11])
        #jde an8
        ORADB = CONORACLE("Select TO_CHAR(ABAN8) as ID_Cust, TO_CHAR(ABALPH) as NM_C  FROM "+OR_DATAID+".F0101 WHERE ABAT1='C'  and ABALKY ='"+uid+"'")
        f.write("5. Select TO_CHAR(ABAN8) as ID_Cust, TO_CHAR(ABALPH) as NM_C  FROM "+OR_DATAID+".F0101 WHERE ABAT1='C'  and ABALKY ='"+uid+"'"+'\n')
        adid=''
        for ds in ORADB:
          adid=str(ds[0])
        #print ouserid
        if dnos>10:
          zon=' '     #全聯
        else:
          zon='1';    #半聯
        '''
        cur206temp.execute("SELECT count(go_no) as tnos FROM  WEBORDERHD WHERE GO_NO like 'WS20"+yy+"%' AND len(NO_SM) >0")     
        for gono in cur206temp:
          n = str(gono[0]+1)
          TNOS = yy+n.zfill(6)
        '''		  
        #print TNOS
        #print "INSERT INTO "+OR_DATAID+".F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOO"+",SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) "+" VALUES ('1','1','00100','"+TNOS+"','E1','1000','850','"+nowdate+"','R','"+str(dnos)+"','N','00','00100','S2','        A001','00100','00100','"+TNOS+"','E1','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+remark1+"','"+remark2+"','"+TNOS+"','"+zon+"')"
        TNOS =  str(int(str(sale[0])[8:14])+int(str(sale[0])[14:]))#20200212值太大要8位
        f.write(adid+' : '+zon+' : '+str(sale[0])[8:14]+'+'+str(sale[0])[14:]+'\n')
        f.write("6. INSERT INTO "+OR_DATAID+".F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOO"
                      +",SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) "
                      +" VALUES ('1','1','00100','"+TNOS+"','E1','1000','850','"+nowdate+"','R','"+str(dnos)+"','N','00','00100','S2','        A001','00100','00100','"
                      +TNOS+"','E1','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+remark1+"','"+remark2+"','"+TNOS+"','"+zon+"')"+'\n')
        ORADB=CONORACLE("INSERT INTO "+OR_DATAID+".F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOO"
                      +",SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) "
                      +" VALUES ('1','1','00100','"+TNOS+"','E1','1000','850','"+nowdate+"','R','"+str(dnos)+"','N','00','00100','S2','        A001','00100','00100','"
                      +TNOS+"','E1','"+adid+"','"+adid+"','"+ordate+"','"+apdate+"','"+remark1+"','"+remark2+"','"+TNOS+"','"+zon+"')")					  
        ccdmodt.execute("update orderformpos set NO_SM='已轉單' where order_no ='"+go_no+"'"+'\n')
        f.write("7. update orderformpos set NO_SM='已轉單' where order_no ='"+go_no+"'"+'\n')
        chaincodet.commit()
        #查出下次出貨日
        ccdmodt.execute("select h.order_no,h.AccountID,h.AccountName,substring(h.ArrivalTime,1,10) as applydate,u.DeliveryDate,WEEKDAY(h.ArrivalTime)+1 as wk from orderformpos h,basicstoreinfo u "
		                  +"where  h.order_no='"+go_no+"' and h.AccountID=u.String_20_1")
        f.write("8. select h.order_no,h.AccountID,h.AccountName,substring(h.ArrivalTime,1,10) as applydate,u.DeliveryDate,WEEKDAY(h.ArrivalTime)+1 as wk from orderformpos h,basicstoreinfo u "
		                  +"where  h.order_no='"+go_no+"' and h.AccountID=u.String_20_1"+'\n')
        #wdt=[]
        for t in ccdmodt.fetchall():
          f.write(str(t[4])+'\n')
          wdt=str(t[4]).split(',')
          ad=int(t[5])
          t_str =str(t[3])#本次出貨日
        f.write(str(wdt)+':'+str(ad)+':'+t_str+'\n')
        wd=[]		  
        for w in range(len(wdt)):#將中文轉成數字
          if wdt[w]=='一':
            wd.append(1)
          elif wdt[w]=='二':
            wd.append(2)
          elif wdt[w]=='三':
            wd.append(3)
          elif wdt[w]=='四':
            wd.append(4)
          elif wdt[w]=='五':
            wd.append(5)
          elif wdt[w]=='六':
            wd.append(6)
        f.write(str(wd)+'\n')
        for d in range(len(wd)):
          dt=wd[d]-ad
          if dt<=0:
            dt=dt+7
          wd[d]=getday(int(t_str[:4]),int(t_str[5:7]),int(t_str[8:10]),dt)
        wd=sorted(wd)
        '''
        t_str =str(t[3])#本次出貨日
        d = datetime.datetime.strptime(t_str, '%Y%m%d')
        delta = datetime.timedelta(days=adays)#下一次出貨日天數
		n_days = d + delta
        apdateb=n_days.strftime('%Y%m%d')#下一次出貨日
        ''' 
        f.write(str(wd)+'\n')		
        apdateb=wd[0]#下一次出貨日
		#20200211階段
		#查出下次出貨日
        #####sale items######
        #cur2061fl.execute("SELECT A.*,B.NEW_IDITEM,B.IFSTOCK FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ORDER BY B.ID_ITEM")
        f.write(str(sale[5])+'\n')
        if str(sale[5]).find('非正配單')==-1:#20191115
          '''
          cur2061fl.execute("SELECT A.*,B.NEW_IDITEM,B.IFSTOCK FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A,WEBART B " 
                        +"WHERE A.ID_ITEM=B.ID_ITEM and A.id_item  not in (SELECT  [id_item]  FROM [TGSalary].[dbo].[WEBBookingART]) "
                        +" union all SELECT '"+go_no+"',A.[ID_ITEM],A.[NM_ITEM],sum(A.[QTY]),A.[UPRICE],sum(A.[SUBTOT]),B.NEW_IDITEM,B.IFSTOCK "
                        +"FROM (SELECT * FROM [WEBBookingOrder] WHERE [applydate]<='"+t_str+"' and [ID_CUST]='"+cid+"' and resale='N' )A,WEBART B  "
                        +" WHERE A.ID_ITEM=B.ID_ITEM group by A.[ID_ITEM],A.[NM_ITEM],A.[UPRICE],B.NEW_IDITEM,B.IFSTOCK ORDER BY ID_ITEM") #20191112 [applydate]='"+t_str+"'
          '''
          f.write("9. SELECT order_no,ProdID,ProdName,Amount,Double_1,Double_2,ProdID1,'' FROM orderformpos_prod_sub  where order_no='"+go_no+"'"
                          +" union all "
                          +" select 'order_no',a.ID_ITEM,a.NM_ITEM,sum(a.QTY) as QTY,a.UPRICE,sum(a.SUBTOT) as SUBTOT ,a.ID_ITEM ,''  from  (SELECT * FROM WEBBookingOrder "
                          +" WHERE applydate<='"+apdateb+"' and ID_CUST='"+cid+"' and resale='N' )A group by  a.ID_ITEM,a.NM_ITEM,a.UPRICE   order by ProdID"+'\n')
          ccdmodfl.execute("SELECT order_no,ProdID,ProdName,Amount,Double_1,Double_2,ProdID1,'' FROM orderformpos_prod_sub  where order_no='"+go_no+"' and ProdID not in (SELECT  id_item  FROM WEBBookingART) "
                          +" union all "
                          +" select 'order_no',a.ID_ITEM,a.NM_ITEM,sum(a.QTY) as QTY,a.UPRICE,sum(a.SUBTOT) as SUBTOT ,a.ID_ITEM ,''  from  (SELECT * FROM WEBBookingOrder "
                          +" WHERE applydate<='"+apdateb+"' and ID_CUST='"+cid+"' and resale='N' )A group by  a.ID_ITEM,a.NM_ITEM,a.UPRICE   order by ProdID")
          inbooking=1
        else:
          f.write("9-1. SELECT order_no,ProdID,ProdName,Amount,Double_1,Double_2,ProdID1,'' FROM orderformpos_prod_sub  where order_no='"+go_no+"'"+'\n')
          ccdmodfl.execute("SELECT order_no,ProdID,ProdName,Amount,Double_1,Double_2,ProdID1,'' FROM orderformpos_prod_sub  where order_no='"+go_no+"'  order by ProdID")
          inbooking=0
        I12=0
        for sitem in ccdmodfl:
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
          '''
          ORADB=CONORACLE("INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"
                        +",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "
                        +" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"
                        +",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate
                        +"','"+NewItemID
                        +"','','','','','',"+str(qty)+",'"+TNOS+"')")
          '''
          f.write("10. INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"
                        +",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "
                        +" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"
                        +",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate
                        +"','"+NewItemID
                        +"','','','','','',"+str(qty)+",'"+TNOS+"')"+'\n')
          ORADB=CONORACLE("INSERT INTO "+OR_DATAID+".F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNO"
                        +",SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) "
                        +" VALUES ('2','"+str(I12)+"','00100','"+TNOS+"','E1','"+str(I12)+"000','850','"+nowdate+"','R','N','00100','S2'"
                        +",'"+str(I12)+"000','"+Branch_Plan+"','00100','00100','"+TNOS+"','E1','"+str(I12)+"000','"+adid+"','"+adid+"','"+ordate+"','"+apdate
                        +"','"+NewItemID
                        +"','','','','','',"+str(qty)+",'"+TNOS+"')")
        #查詢是否有預購商品有則新增到 WEBBookingOrder
        #20200213先把商品更新再測預購
        if inbooking==1: #20191125
          #print("SELECT A.* FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='"+go_no+"')A WHERE  A.id_item  in (SELECT  [id_item]  FROM [TGSalary].[dbo].[WEBBookingART])  ")
          ccdmodfl.execute("SELECT A.order_no,A.ProdID,A.ProdName,A.Amount,A.Double_1,A.Double_2 FROM (SELECT * FROM orderformpos_prod_sub WHERE order_no='"+go_no+"')A WHERE  A.ProdID  in (SELECT  id_item  FROM WEBBookingART)  ")
          f.write("SELECT A.order_no,A.ProdID,A.ProdName,A.Amount,A.Double_1,A.Double_2 FROM (SELECT * FROM orderformpos_prod_sub WHERE order_no='"+go_no+"')A WHERE  A.ProdID  in (SELECT  id_item  FROM WEBBookingART)  "+'\n')
          for sitem in ccdmodfl:
            itemid=str(sitem[1])
            itemnm=str(sitem[2])
            iqty=str(sitem[3])
            iuprice=str(sitem[4])
            isubtot=str(sitem[5])
            '''
            cur206temp.execute("insert into [TGSalary].[dbo].[WEBBookingOrder] ([ID_ITEM],[NM_ITEM],[QTY],[UPRICE],[SUBTOT],[resale],[applydate],[ID_CUST],[indate],[go_no]) VALUES "
                              +"('"+itemid+"','"+itemnm+"','"+iqty+"','"+iuprice+"','"+isubtot+"','N','"+str(apdateb)+"','"+cid+"','"+indate+"','"+go_no+"')")#20191212
            cur206temp.commit()
            '''
            f.write("11. insert into WEBBookingOrder (ID_ITEM,NM_ITEM,QTY,UPRICE,SUBTOT,resale,applydate,ID_CUST,indate,go_no) VALUES "
                              +"('"+itemid+"','"+itemnm+"','"+iqty+"','"+iuprice+"','"+isubtot+"','N','"+str(apdateb)+"','"+cid+"','"+indate+"','"+go_no+"')"+'\n')
            ccdmodt.execute("insert into WEBBookingOrder (ID_ITEM,NM_ITEM,QTY,UPRICE,SUBTOT,resale,applydate,ID_CUST,indate,go_no) VALUES "
                              +"('"+itemid+"','"+itemnm+"','"+iqty+"','"+iuprice+"','"+isubtot+"','N','"+str(apdateb)+"','"+cid+"','"+indate+"','"+go_no+"')")
            
            chaincodet.commit()
        #查詢是否有預購商品有則新增到 WEBBookingOrder
        if inbooking==1:#20191115
          f.write(" UPDATE WEBBookingOrder set  resale='Y' WHERE applydate<='"+t_str+"' and ID_CUST='"+cid+"' and resale='N' "+'\n')
          ccdmodt.execute(" UPDATE WEBBookingOrder set  resale='Y' WHERE applydate<='"+t_str+"' and ID_CUST='"+cid+"' and resale='N' ")
          chaincodet.commit()
          
  f.close()
  return salel
        
