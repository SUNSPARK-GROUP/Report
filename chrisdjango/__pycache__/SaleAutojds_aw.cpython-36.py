3
Y]�\�"  �               @   s^   d dl Z d dlZd dlmZ d dlmZ d dlZd dlZdejd< dd� Zdd	� Z	d
d� Z
dS )�    N)�date)�	timedeltazSIMPLIFIED CHINESE_CHINA.UTF8�NLS_LANGc             C   sX   d|  }t j� t|� }|dkr:|jd| d | d �S |jd| d | d �S d S )Nr   �   z%dz%mz%Y)r   �todayr   �strftime)�wd�sp�dt�t1�d� r   �8C:\Users\CHRIS\chrisdjango\chrisdjango\SaleAutojds_aw.py�showday	   s
    r   c             C   s~   d}d}d}d}d}t j|||�}t j|d | d | �}|j� }|j| � | dd	� j� }	|	d
krr|j� }
|
S |j�  d S )Nz192.168.0.230�E910�PRODDTA�E910Jde�1521�/�@r   �   �SELECT)	�	cx_Oracle�makedsn�connect�cursor�execute�upper�fetchall�close�commit)�SqlStr�hostname�sid�username�password�port�dsn�connr   �SQLSTRS�TotalSessionr   r   r   �	CONORACLE   s    
 r+   c       )      C   s�  d}t | dd�}t ddd�}tjd�}tjd�}tjd�}tjd�}|j� }|j� }	|j� }
|j� }g }|jd| d � �x.|D �]$}g }|jt|d �� |jt|d �� |jt|d �� |jt|d	 �� |jt|d
 �� |j|� t|d �}|
jdd d d | d | d t|d � d � |jd| d � x|D ]}t|d �}�qJW �xD|
D �]:}t|d �dk�rht|d �}t|d	 �}|}|d d
d� }|d }d}x4tt	|��D ]$}t
|| �dk �r�|||  }�q�W d}d}t	|�dk�r"|d d� }|dt	|�� }n|}t|d �dk�rFt|d �}nt|d �}td| d | d �}d}x|D ]}t|d �}�qtW |dk�r�d} nd } |	jd!| d" � x,|	D ]$}!t|!d d �}"||"jd� }#�q�W td#| d$ d% d& |# d' | d( t|� d) |# d* | d+ | d+ | d+ | d+ | d+ | d+ |# d+ |  d, �}|	jd-|# d. | d � |	j�  |jd/| d0 � d}$�x|D �]}%|$d7 }$|%d dk�r�d1}&nt|%d �j� }&|%d2 d3k�r�d}'nd4}'t|%d	 �d5 }(td#| d6 d7 d8 t|$� d9 |# d* t|$� d: | d; d< t|$� d= |' d> |# d* t|$� d= | d+ | d+ | d+ | d+ |& d? t|(� d< |# d, �}�q�W �qhW q�W |S )@Nr   � r   zTDRIVER={SQL Server};SERVER=192.168.0.206;DATABASE=TGSalary;UID=apuser;PWD=0920799339z^SELECT [GO_NO],[ID_CUST],[NM_C],[TOTAMT],[REMARK] FROM WEBORDERHD WHERE ISNULL(APPLYDATE,'')='zU'  AND (len(NO_SM)<1 or  NO_SM  is null) and id_cust<>'appdv' ORDER BY ID_CUST,GO_NO r   �   �   �   zuSELECT A.go_no,a.no_sm,a.id_cust,a.applydate,a.sdatetime,convert(nvarchar(200),a.remark) as remark,B.USERID,B.TYPENO,zVB.SALETYPE,IsNull(B.BONNER,0) AS USERBON,ISNULL(B.CARNO,'')CARNO,B.USERIDNEW ,a.aDATE z|FROM (SELECT GO_NO,NO_SM,ID_CUST, '1'+substring(APPLYDATE,3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,APPLYDATE) z-as NVARCHAR), 3) as APPLYDATE,'1'+substring('z:',3,2)+RIGHT(REPLICATE('0', 3) + CAST(datepart(dayofyear,'zK') as NVARCHAR), 3) as aDATE,SDATETIME,REMARK FROM WEBORDERHD WHERE GO_NO='z(')A,USERGROUP B WHERE A.ID_CUST=B.USERIDzASELECT count(*) as q FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='z'')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM �None�   r   �   i�  �   �   z@Select TO_CHAR(ABAN8) as ID_Cust, TO_CHAR(ABALPH) as NM_C  FROM z%.F0101 WHERE ABAT1='C'  and ABALKY ='�'�
   � �1zCSELECT count(go_no) as tnos FROM  WEBORDERHD WHERE GO_NO like 'WS20z%' AND len(NO_SM) >0zINSERT INTO zc.F47011 (SYEDTY,SYEDSQ,SYEKCO,SYEDOC,SYEDCT,SYEDLN,SYEDST,SYEDDT,SYEDER,SYEDDL,SYEDSP,SYTPUR,SYKCOOz_,SYDCTO,SYMCU,SYCO,SYOKCO,SYOORN,SYOCTO,SYAN8,SYSHAN,SYTRDJ,SYPPDJ,SYDEL1,SYDEL2,SYVR01,SYZON) z VALUES ('1','1','00100','z','E1','1000','850','z','R','z8','N','00','00100','S2','        A001','00100','00100','z','E1','z','z')zupdate WEBORDERHD set NO_SM='z' where GO_NO ='zNSELECT A.*,B.NEW_IDITEM,B.IFSTOCK FROM (SELECT * FROM WEBORDERFL WHERE GO_NO='z9')A,WEBART B WHERE A.ID_ITEM=B.ID_ITEM ORDER BY B.ID_ITEM�A�   �Nz        A001i'  z�.F47012(SZEDTY,SZEDSQ,SZEKCO,SZEDOC,SZEDCT,SZEDLN,SZEDST,SZEDDT,SZEDER,SZEDSP,SZKCOO,SZDCTO,SZLNID,SZMCU,SZCO,SZOKCO,SZOORN,SZOCTO,SZOGNOzU,SZAN8,SZSHAN,SZTRDJ,SZRSDJ,SZLITM,SZLOCN,SZLNTY,SZNXTR,SZLTTR,SZUOM,SZUORG,SZVR01 ) z VALUES ('2','z','00100','z000','850','z','R','N','00100','S2'z,'z000','z','00100','00100','z','','','','','',)r   �pyodbcr   r   r   �append�str�int�range�len�ordr+   �zfillr    r   ))ZtdayZ	OR_DATAID�sdateZndate�connection206Zconnection2061Zconnection2062Zconnection206t�cur206hdZ
cur206tempZ	cur2061hdZ	cur2061flZsalelZsales�hdlZgo_noZflcountZdnosZsaleZnowdateZordateZapdate�yyZremarkt�remark�rZremark1Zremark2�uidZORADBZadid�dsZzon�gono�nZTNOSZI12ZsitemZ	NewItemIDZBranch_Plan�qtyr   r   r   �set2jde!   s�    




2
 


z�rP   )r   r<   �datetimer   r   �string�os�environr   r+   rP   r   r   r   r   �<module>   s   
