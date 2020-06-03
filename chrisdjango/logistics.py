# -*- encoding: UTF-8 -*-
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
import openpyxl
import win32com.client as win32
import pyodbc
import pymysql
import os
import xlsxwriter
import re
import time
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
def CONMYSQL(sqlstr):
  db = pymysql.connect(host='192.168.0.210', port=3306, user='apuser', passwd='0920799339', db='main_eipplus_standard',charset='utf8')
  cursor = db.cursor()
  cursor.execute(sqlstr)	
  result = cursor.fetchall()
  urls = [row[0] for row in result]
  return result
def CONORACLE(SqlStr):
  # cf = open(r'D:\chrisdjango\ora.txt','w')
  hostname='192.168.0.230'
  sid='E910'
  username='PRODDTA'
  password='E910Jde'
  port='1521'
  dsn = cx_Oracle.makedsn(hostname, port, sid)
  conn = cx_Oracle.connect(username+'/'+password+'@' + dsn)
  cursor = conn.cursor()
  
  
  
  SQLSTRS = SqlStr[0:6].upper()
  # cf.write(SQLSTRS)
  
  if SQLSTRS=="SELECT":
    cursor.execute(SqlStr)
    TotalSession = cursor.fetchall()
    return TotalSession
    cursor.close()
  else:
    # cf.write(SqlStr + '\n')
    # cf.close()
    cursor.execute(SqlStr)
    conn.commit()
  
def f47121m(request):#物料領用/歸還查詢
  context= {}
  f47121m=[]
  data=CONMYSQL("SELECT formsflow_id,formsflow_display,formsflow_version FROM hplus_formsflow h where formsflow_state=2 and formsflow_content=''"+
  " and formsflow_title = '物流部表單03.產品物料領用/歸還申請表'")
  # data=CONMYSQL("SELECT formsflow_id,formsflow_display,formsflow_version FROM hplus_formsflow h where formsflow_id='81157' or formsflow_id='81167'")
  # data=CONMYSQL("SELECT formsflow_id,formsflow_display,formsflow_version FROM hplus_formsflow h where formsflow_state=2 and formsflow_content='N'"+
  # " and formsflow_title = '物流部表單03.產品物料領用/歸還申請表'")
  for d in data:
    dver=CONMYSQL("select distinct(formsflow_field_version) FROM hplus_formsflow_field where formsflow_field_formsflow='"+str(d[0])+"' order by formsflow_field_version desc limit 1")
    dev=''
    for dv in dver:
      dev=dv[0]
    group=''#申請部門
    name=''#使用人
    why=''#用途
    take=''#歸或領 take->領 re->歸還
    dverm=CONMYSQL("select  f.formsflow_field_name,f.formsflow_field_value FROM hplus_formsflow_field f where formsflow_field_formsflow='"
                 +str(d[0])+"' and formsflow_field_version='"+str(dev)+"'  ")
    tf47121=[]	
    for dt in dverm:
      if dt[0]=='__group':
        getd=CONMYSQL("SELECT account_lid FROM hplus_accounts h where account_id='"+str(dt[1])+"'")
        for gd in getd:
          group=gd[0]
      if dt[0]=='__name':
        getname=CONMYSQL("SELECT empname FROM hplus_special h WHERE account_id='"+str(dt[1])+"'")
        for n in getname:
          name=n[0]
      if dt[0]=='__why':
        why=str(dt[1])
      if dt[0]=='__take':
        if str(dt[1])=='take':
          take='領'
        if str(dt[1])=='re':
          take='歸'
    tf47121.append(group)
    tf47121.append(name)
    tf47121.append(why)
    tf47121.append(take)
    tf47121.append(str(d[0]))
    f47121m.append(tf47121)
  context['f47121m'] = f47121m
  if len(f47121m)==0:
    context['mess']='查無資料'
  else:
    context['mess']="共 "+str(len(f47121m))+" 筆資料"
  context['reportmes']='<B>森邦(股)領用/歸還申請表</B>'
  return render(request, 'F47121m.html',context )#傳入參數
def f47121d(request):#物料領用/歸還項目明細
  # f=open(r'D:\chrisdjango\error.txt','w')
  context= {}
  f47121d=[]
  try:
    eid=request.GET['eipid']
    context['eid'] = eid
    dver=CONMYSQL("select  distinct(formsflow_field_version) FROM hplus_formsflow_field where formsflow_field_formsflow='"+eid+"' order by formsflow_field_version desc limit 1 ")
    dev=''
    for dv in dver:
      dev=dv[0]
    try:
      wtype=request.GET['wtype']
    except:
      wtype='' 
    dverm=CONMYSQL("select  f.formsflow_field_name,f.formsflow_field_value FROM hplus_formsflow_field f where formsflow_field_formsflow='"+eid+"' and"
    +" formsflow_field_version='"+str(dev)+"'  ")
    context['vs'] = str(dev)
    for dt in dverm:
      if dt[0]=='__group':
        getd=CONMYSQL("SELECT account_lid FROM hplus_accounts h where account_id='"+str(dt[1])+"'")
        for gd in getd:
          context['group']=gd[0]#申請部門
      if dt[0]=='__name':
        getname=CONMYSQL("SELECT empname FROM hplus_special h WHERE account_id='"+str(dt[1])+"'")
        for n in getname:
          context['name']=n[0]#使用人
      if dt[0]=='__mcu':
        context['mcu']=str(dt[1])#使用單位
      if dt[0]=='__why':
        context['why']=str(dt[1])#用途
      if dt[0]=='__TD':
        context['TD']=str(dt[1])#需求日期
      if dt[0]=='__AK1':
        context['AK1']=str(dt[1])#需求時間1
      if dt[0]=='__AH1':
        context['AH1']=str(dt[1])#需求時間2
      if dt[0]=='__AM1':
        context['AM1']=str(dt[1])#需求時間3
      if dt[0]=='__txt':
        context['txt']=str(dt[1])#說明
    deverd=CONMYSQL("SELECT formsflow_field_name,formsflow_field_value FROM hplus_formsflow_field h where formsflow_field_formsflow='"+eid+"' and formsflow_field_version='"+str(dev)+"' "
                   +" and (formsflow_field_name like '__unit-%' or formsflow_field_name like '__whatN-%' or formsflow_field_name like '__what-%' or formsflow_field_name like '__amount-A%')"
                   +" order by LPAD(substring_index(formsflow_field_name,'1_',-1),3,0),formsflow_field_name") 
    amounts=[]
    units=[]
    whats=[]
    whatns=[]
    for dl in deverd:
      if dl[0].find('__amount')==0:
        amounts.append(str(dl[1]))
      if dl[0].find('__unit')==0:
        units.append(str(dl[1]))
      if dl[0].find('__what-')==0:
        whats.append(str(dl[1]))
      if dl[0].find('__whatN-')==0:
        whatns.append(str(dl[1]))
    for l in range(len(amounts)):
      tf47121d=[]	  
      tf47121d.append(whats[l])  
      tf47121d.append(whatns[l])  
      tf47121d.append(amounts[l])  
      tf47121d.append(units[l])
      f47121d.append(tf47121d)
    context['f47121d'] = f47121d
    
  except:
    eipid = request.GET['eid']
    vs = request.GET['vs']
    set=request.GET['Submit1']
    if set == '轉入JDE':
      taketojde(eipid,vs)
  return render(request, 'F47121detel.html',context )#傳入參數
def taketojde(eipno,version):#物料領用/歸還項目轉入JDE
  # OR_DATAID='CRPDTA'#測試區
  OR_DATAID='PRODDTA'#正式區
  context= {}
  #f = open(r'C:\Users\Administrator\chrisdjango\jde.txt','w')
  depa = ''
  DGL = ''
  name = ''
  EDCT = ''
  DL01 = ''
  LITM = ''
  TRQT = ''
  URRF = ''
  ANI = ''
  RCD = ''
  DSC1 = ''  
  dverm = CONMYSQL("select f.formsflow_field_name,f.formsflow_field_value FROM  hplus_formsflow_field f where formsflow_field_formsflow='"+eipno+"' and"
  +" formsflow_field_version='"+version+"'")
  for dt in dverm:
    if dt[0] == '__mcu':#使用單位
      depa = dt[1]
    why = CONMYSQL("select f.formsflow_field_value FROM  hplus_formsflow_field f where formsflow_field_formsflow='"+eipno+"' and"
    +" formsflow_field_version='"+version+"' and f.formsflow_field_name = '__why'")
    for w in why:
      if w[0] == '食材打樣、研發、檢測':
        ANI = '      ' + depa + '.6220'
        RCD = '031'
        break
      elif w[0] == '新品.樣品拍照':
        ANI = '      ' + depa + '.6220'
        RCD = '022'
        break
      elif w[0] == '打樣借用.留樣(包材類)':
        ANI = '      ' + depa + '.6290'
        RCD = '092'
        break
      elif w[0] == '教室領料(台灣)':
        ANI = '      ' + depa + '.6222.002'
        RCD = '036'
        break
      elif w[0] == '茶水間.廁所':
        ANI = '      ' + depa + '.6219.003'
        RCD = '001'
        break
      elif w[0] == '領用製服':
        ANI = '      ' + depa + '.6219.003'
        RCD = '002'
        break
      elif w[0] == '六兩袋.洗潔精.口罩':
        ANI = '      ' + depa + '.6290'
        RCD = '091'
        break
      elif w[0] == '企業參訪(來訪贈品)':
        ANI = '      ' + depa + '.6211'
        RCD = '012'
        break
      elif w[0] == '企業活動(標籤貼.加盟展)':
        ANI = '      ' + depa + '.6211'
        RCD = '011'
        break
      elif w[0] == '記者會':
        ANI = '      ' + depa + '.6211'
        RCD = '021'
        break
      elif w[0] == '加盟展贈送(台灣)':
        ANI = '      ' + depa + '.6253.001'
        RCD = '011'
        break
      elif w[0] == '活動推廣':
        ANI = '      ' + depa + '.6208.004'
        RCD = ' '
        break
      elif w[0] == '加盟展領用(製服.食材)':
        ANI = '      ' + depa + '.6253.001'
        RCD = '043'
        break
      elif w[0] == '加盟商大會(研習會)':
        ANI = '      ' + depa + '.6252'
        RCD = '044'
        break
      elif w[0] == '直營-內部員工教學':
        ANI = '      ' + depa + '.6222.001'
        RCD = '036'
        break
      elif w[0] == '業務部-說明會':
        ANI = '      ' + depa + '.6251'
        RCD = '041'
        break
    if dt[0] == '__TD':#需求日期
      day = CONMYSQL("select DATE_FORMAT(f.formsflow_field_value,+'%y%j') FROM  hplus_formsflow_field f where f.formsflow_field_name = '"+dt[0]+"'and"
      +" formsflow_field_formsflow='"+eipno+"' and formsflow_field_version='"+version+"'")
      for d in day:
        DGL = '1'+str(d[0])
    
    if dt[0] == '__name':#使用人
      wname = CONMYSQL("SELECT empname FROM hplus_special h WHERE account_id='"+str(dt[1])+"'")
      for n in wname:
        name = str(n[0])
    
    if dt[0] == '__take':#領or還
      if dt[1] == 'take':
        EDCT = "II"#領
      else:
        EDCT = "IR"#還
    if dt[0] == '__txt':#說明
      txt = str(dt[1])
      DL01 = txt[0:30]
  item = CONMYSQL("SELECT formsflow_field_name,formsflow_field_value FROM hplus_formsflow_field h where formsflow_field_formsflow='"+eipno+"' and"
  +" formsflow_field_version='"+version+"' and (formsflow_field_name like '__unit-%' or formsflow_field_name like '__whatN-%' or formsflow_field_name like '__what-%' or"
  +" formsflow_field_name like '__amount-A%') order by LPAD(substring_index(formsflow_field_name,'1_',-1),3,0),formsflow_field_name")
  amounts=[]
  units=[]
  whats=[]
  whatns=[]
  nums=[]
  nums2=[]
  DSC=[]
  OUM=[]
  EDSPS=[]
  HEDSPS=[]
  s = 0
  for it in item:
    if it[0].find('__what-')==0:#物品名稱
      its = str(it[1])
      itsn = re.sub("\'","''",its)
      whats.append(str(itsn))
    if it[0].find('__whatN-')==0:#料號
      itn = str(it[1])
      itns = re.sub('\s','',itn)
      whatns.append(str(itns))
      PQOH = '-1'
      LOCN = ' '
      wlocn = CONORACLE("Select IMITM,IMLITM,LIITM,LIMCU,LILOCN,LIPQOH,IMUOM1 FROM PRODDTA.F4101,PRODDTA.F41021 where IMITM = LIITM and IMLITM = '"+str(itns)+"' and"
      +" LIMCU = '        A001'")
      # f.write(str(wlocn) + '\n')
      # f.write(str(itns)+'\n')
      DSC2 = '料號錯誤'
      EDSP = 'E'
      HEDSP = 'E'
      oum=''
      snums=LOCN      
      for lc in wlocn:
        DSC2 = ' '
        EDSP = '0'
        HEDSP = '0'
        cun = str(lc[5])
        cuns = re.sub('\s','',cun)
        if cuns == '0':
          DSC2 = '庫存為0'
          EDSP = 'E'
          HEDSP = 'E'
          oum=str(lc[6])
          sl = str(lc[4])
          LOCN = sl[0:1]+"."+sl[1:3]+"."+sl[3:5]+"."+sl[5:6]+"."+sl[6:7]          
          PQOH = str(cuns)
        else:
          sl = str(lc[4])
          sln = sl[0:1]+"."+sl[1:3]+"."+sl[3:5]+"."+sl[5:6]+"."+sl[6:7]
          oum=str(lc[6])
          LOCN=str(sln)
          PQOH = str(cuns)
      nums.append(LOCN)
      nums2.append(PQOH)
      OUM.append(oum)
      DSC.append(DSC2)
      EDSPS.append(EDSP)
      HEDSPS.append(HEDSP)
    '''	  
      if len(wlocn) == 0:
        # f.write('error' + '\n')
        DSC2 = '料號錯誤'
        EDSP = 'E'
        HEDSP = 'E'
        DSC.append(DSC2)
        nums.append(str(LOCN))
        nums2.append(PQOH)
        DSC.append(DSC2)
        EDSPS.append(EDSP)
        HEDSPS.append(HEDSP)
      else:
        for lc in wlocn:
          # f.write('start' + '\n')
          DSC2 = ' '
          EDSP = '0'
          HEDSP = '0'
          cun = str(lc[5])
          cuns = re.sub('\s','',cun)
          if cuns == '0':
            DSC2 = '庫存為0'
            EDSP = 'E'
            HEDSP = 'E'
            sl = str(lc[4])
            LOCN = sl[0:1]+"."+sl[1:3]+"."+sl[3:5]+"."+sl[5:6]+"."+sl[6:7]
            nums.append(str(LOCN))
            PQOH = str(cuns)
            nums2.append(PQOH)
            DSC.append(DSC2)
            OUM.append(str(lc[6]))
            EDSPS.append(EDSP)
            HEDSPS.append(HEDSP)
          else:
            sl = str(lc[4])
            sln = sl[0:1]+"."+sl[1:3]+"."+sl[3:5]+"."+sl[5:6]+"."+sl[6:7]
            nums.append(str(sln))
            PQOH = str(cuns)
            nums2.append(PQOH)
            DSC.append(DSC2)
            OUM.append(str(lc[6]))
            EDSPS.append(EDSP)
            HEDSPS.append(HEDSP)
    '''
    if it[0].find('__amount')==0:#物品數量
      if EDCT == 'II':
        total = float(it[1]) * 10000
        stotal = int(total)
        amounts.append(str(stotal))
      else:
        total = float(it[1]) * -10000
        stotal = int(total)
        amounts.append(str(stotal))
    if it[0].find('__unit')==0:#計算的單位
      units.append(str(it[1]))
  Dtime = time.strftime("%H%M%S", time.localtime())
  # f.write(Dtime)
  # f.write(str(HEDSPS)+'\n')
  for e in range(len(HEDSPS)):
    if HEDSPS[e].find('0')==0:
      HEDSPout = '0'
      # f.write(str(HEDSPout)+'\n')
    else:
      HEDSPout = 'E'
      # f.write(str(HEDSPout)+'\n')
      break
  '''
  f.write(str(amounts)+'\n')
  f.write(str(units)+'\n')
  f.write(str(whats)+'\n')
  f.write(str(whatns)+'\n')
  f.write(str(nums)+'\n')
  f.write(str(nums2)+'\n')
  f.write(str(DSC)+'\n')
  f.write(str(OUM)+'\n')
  f.write(str(EDSPS)+'\n')
  f.write(str(HEDSPS)+'\n')
  f.close()
  '''
  for l in range(len(amounts)):
    s += 1
    i = s * 1000
    addf47122 = CONORACLE("INSERT INTO "+OR_DATAID+".F47122(MJEKCO,MJEDOC,MJEDCT,MJEDLN,MJEDDT,MJEDDL,MJPACD,MJKSEQ,MJAN8,MJMCU,MJLITM,MJLOCN,MJSTUN,MJLDSQ,MJTRNO,MJLOTP,MJMMEJ,"
    +"MJDSC1,MJTRDJ,MJTRQT,MJKCO,MJSFX,MJDGL,MJTREX,MJANI,MJTORG,MJUSER,MJPID,MJJOBN,MJUPMJ,MJTDAY,MJSQOR,MJRCD,MJDSC2,MJEDSP,MJPNS,MJDCT,MJTRUM,MJLOTN,MJUOM2,MJPMPN) VALUES ('00100','"+eipno+"','"+EDCT+"'"+
    ",'"+str(i)+"','0','0','"+EDCT+"','0','100002','        A001','"+whatns[l]+"','"+nums[l]+"','0','0','1','0','0','"+whats[l]+"','"+DGL+"','"+amounts[l]+"','00100','000','"+DGL+""+
    "','"+DL01+"','"+ANI+"','"+name+"','0238','ER47121','erpdb','"+DGL+"','"+Dtime+"','0','"+RCD+"','"+DSC[l]+"','"+EDSPS[l]+"','0','"+EDCT+"','"+OUM[l]+"',' ',' ',' ')")
    # f.write("INSERT INTO "+OR_DATAID+".F47122(MJEKCO,MJEDOC,MJEDCT,MJEDLN,MJEDDT,MJEDDL,MJPACD,MJKSEQ,MJAN8,MJMCU,MJLITM,MJLOCN,MJSTUN,MJLDSQ,MJTRNO,MJLOTP,MJMMEJ,"
            # +"MJDSC1,MJTRDJ,MJTRQT,MJKCO,MJSFX,MJDGL,MJTREX,MJANI,MJTORG,MJUSER,MJPID,MJJOBN,MJUPMJ,MJTDAY,MJSQOR,MJRCD,MJDSC2,MJEDSP,MJPNS,MJDCT,MJTRUM,MJLOTN,MJUOM2,MJPMPN) VALUES ('00100','"+eipno+"','"+EDCT+"'"+
            # ",'"+str(i)+"','0','0','"+EDCT+"','0','100002','        A001','"+whatns[l]+"','"+nums[l]+"','0','0','1','0','0','"+whats[l]+"','"+DGL+"','"+amounts[l]+"','00100','000','"+DGL+""+
            # "','"+DL01+"','"+ANI+"','"+name+"','0238','ER47121','erpdb','"+DGL+"','"+Dtime+"','0','"+RCD+"','"+DSC[l]+"','"+EDSPS[l]+"','0','"+EDCT+"','"+OUM[l]+"',' ',' ',' ')"+'\n')
    
  addf47121 = CONORACLE("INSERT INTO "+OR_DATAID+".F47121(M1EKCO,M1EDOC,M1EDCT,M1EDLN,M1EDST,M1EDDT,M1EDDL,M1EDSP,M1THCD,M1AN8,M1DTFR,M1DTTO,M1URDT,M1URAT,M1URAB,M1TORG,M1USER,"
          +"M1PID,M1JOBN,M1UPMJ,M1TDAY) VALUES ('00100','"+eipno+"','"+EDCT+"','0','852','0','0','"+HEDSPout+"','T','100022','0','0','0','0','0','"+name+"','0238','ER47121','erpdb'"
          +",'"+DGL+"','"+Dtime+"')") 
  # f.write("INSERT INTO "+OR_DATAID+".F47121(M1EKCO,M1EDOC,M1EDCT,M1EDLN,M1EDST,M1EDDT,M1EDDL,M1EDSP,M1THCD,M1AN8,M1DTFR,M1DTTO,M1URDT,M1URAT,M1URAB,M1TORG,M1USER,"
          # +"M1PID,M1JOBN,M1UPMJ,M1TDAY) VALUES ('00100','"+eipno+"','"+EDCT+"','0','852','0','0','"+HEDSPout+"','T','100022','0','0','0','0','0','"+name+"','0238','ER47121','erpdb'"
          # +",'"+DGL+"','"+Dtime+"')"+'\n')
  update = pymysql.connect(host='192.168.0.210', port=3306, user='apuser', passwd='0920799339', db='main_eipplus_standard',charset='utf8')
  updatemysql = update.cursor()
  updatemysql.execute("Update hplus_formsflow h set formsflow_content ='Y' where formsflow_id='"+eipno+"'")
  update.commit()
  updatemysql.close()