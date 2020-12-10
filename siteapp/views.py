from django.shortcuts import render
import pyodbc
import json
db211_erps=pyodbc.connect('DRIVER={SQL Server}; SERVER=192.168.0.211,1433; DATABASE=erpspos; UID=apuser; PWD=0920799339')
db214_erps=pyodbc.connect('DRIVER={SQL Server}; SERVER=192.168.0.214,1433; DATABASE=erps; UID=apuser; PWD=0920799339')
def getodate(yyyymmdd):
  dataset=db214_erps.cursor()
  yy=yyyymmdd[2:4]
  dataset.execute("select datepart(dayofyear,'"+yyyymmdd+"') [days]")
  for (days) in dataset:
    ofds = list(days)
    ofds = ofds[0]
  i=1
  ofdsl = str(ofds)
  for i in range(len(ofdsl), 3, 1):
    ofdsl = "0"+ofdsl
  return("1"+yy+ofdsl)
def gettotaldata(sano,tdate,dtype):
  dataset=db211_erps.cursor()
  if sano=='TINOS' and dtype==0:
    dataset.execute("select sdate,sa_no,sname,Replace(Convert(Varchar(12),CONVERT(money,total),1),'.00','') as atotal from daytotal where sdate='"+tdate+"' and sa_no like 'TN%' order by total desc")
  elif  sano=='TINOS':
    dataset.execute("select sdate,sa_no,sname,total from daytotal where sdate='"+tdate+"' and sa_no like 'TN%' order by total desc")
  if sano=='VASA' and dtype==0:
    dataset.execute("select sdate,sa_no,sname,Replace(Convert(Varchar(12),CONVERT(money,total),1),'.00','') as atotal from daytotal where sdate='"+tdate+"' and sa_no like 'VA%' order by total desc")
  elif sano=='VASA':
    dataset.execute("select sdate,sa_no,sname,total from daytotal where sdate='"+tdate+"' and sa_no like 'VA%' order by total desc")
  if sano=='LAYA' and dtype==0:
    dataset.execute("select sdate,sa_no,sname,Replace(Convert(Varchar(12),CONVERT(money,total),1),'.00','') as atotal from daytotal where sdate='"+tdate+"' and (sa_no like 'LA%' or sa_no like 'CN%') order by total desc")
  elif sano=='LAYA':
    dataset.execute("select sdate,sa_no,sname,total from daytotal where sdate='"+tdate+"' and (sa_no like 'LA%' or sa_no like 'CN%') order by total desc")
  if sano=='FANI' and dtype==0:
    dataset.execute("select sdate,sa_no,sname,Replace(Convert(Varchar(12),CONVERT(money,total),1),'.00','') as atotal from daytotal where sdate='"+tdate+"' and (sa_no like 'FA%' or sa_no like 'CF%') order by total desc")
  elif sano=='FANI':
    dataset.execute("select sdate,sa_no,sname,total from daytotal where sdate='"+tdate+"' and (sa_no like 'FA%' or sa_no like 'CF%') order by total desc")
  if sano=='SelFish' and dtype==0:
    dataset.execute("select sdate,sa_no,sname,Replace(Convert(Varchar(12),CONVERT(money,total),1),'.00','') as atotal from daytotal where sdate='"+tdate+"' and sa_no like 'FK%' order by total desc")
  elif sano=='SelFish':
    dataset.execute("select sdate,sa_no,sname,total from daytotal where sdate='"+tdate+"' and sa_no like 'FK%' order by total desc")
  #dbs="select * from daytotal where sdate='"+tdate+"' and sa_no like '"+sano[:2]+"%'"
  sdb=[]
  djs=[]
  ccolor=['#0188a9','#cccc00','#759e00']
  cs=0
  if dtype==1:
    gpstr="[['店名','業績',{ role: 'style' }]"
    '''t.append('店名')
    t.append('業績')
    sdb.append(t)'''  
  for t in dataset.fetchall():
    tdb=[]
    #,str(t[1]),str(t[2]),str(t[3])
    if dtype==1:
      '''tdb.append((t[2]))
      tdb.append(float(t[3]))'''
      csc=cs%3
      gpstr=gpstr+",['"+str(t[2])+"',"+str(t[3])+",'"+ccolor[csc]+"']"
    else:
      tdb.append(str(t[2]))	
      tdb.append(str(t[3]))	
    sdb.append(tdb)
    cs=cs+1
  if dtype==1:
    return(gpstr+"]")  
  elif dtype==0: return (sdb)
def checkuser(uid,pwd):
  dataset=db214_erps.cursor()
  dataset.execute("select username from WebSPusers where userid='"+uid+"' and password='"+pwd+"' collate Chinese_PRC_CS_AI")# collate Chinese_PRC_CS_AI 可區分大小寫區別 
  uname=''
  for t in dataset.fetchall():
    uname=str(t[0])
  return uname
def mainmenu(uid):
  dataset=db214_erps.cursor()
  #dataset.execute("SELECT MID, MNAME FROM WebSPmainMenu order by mid")
  dataset.execute("SELECT MID, MNAME FROM WebSPmainMenu WHERE MID IN (SELECT distinct([MID])  FROM [ERPS].[dbo].WebSPsubMenu where fid in (SELECT [fid]  FROM [ERPS].[dbo].[WebSPusersunc] where userid='"+uid+"')) order by mid")
  mlist=[]
  for mn in dataset.fetchall():
    mll={'mid':str(mn[0]),'mname':str(mn[1])}
    mlist.append(mll)
  return mlist
def submenu(sid,uid):
  dataset=db214_erps.cursor()
  dataset.execute("SELECT  fid, fname, url FROM WebSPsubMenu  where mid='"+sid+"' and fid in (SELECT fid FROM  WebSPusersunc  WHERE userid = '"+uid+"' ) order by fid")
  sublist=[]
  for sm in dataset.fetchall():
    subll={'fid':str(sm[0]),'sname':str(sm[1]),'url':str(sm[2])}
    sublist.append(subll)
  return sublist
'''
def handler404(request):
    response = render(request, 'ERROR404.html')
    response.status_code = 404
    return response


def handler500(request):
    response = render(request, 'ERROR404.html')
    response.status_code = 500
    return response  
'''
# Create your views here.
