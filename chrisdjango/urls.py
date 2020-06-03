"""chrisdjango URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib.staticfiles.urls import staticfiles_urlpatterns
from django.contrib import staticfiles
from django.contrib import admin
from django.urls import path
from django.conf.urls import url
from django.conf.urls.static import static
from . import first #載入專案 chrisdjango下first.py
from . import Finance #載入專案
from . import product #載入專案
from . import logistics #載入專案
from . import management #載入專案
from .SHOP import shopdata 
from .SHOP import shopstock
from .manage import contracts
from django.conf import settings
from siteapp import views

'''
from django.conf import settings 
if settings.DEBUG: 
  pass 
else: 
  from django.views.static import serve 
  from project.settings import STATIC_ROOT 
  urlpatterns.append(url(r'^static/(?P<path>.*)$', serve, {'document_root': STATIC_ROOT})) 
'''

urlpatterns = [
    #url(r'^$', first.admin),
	#path('', siteviews.index),  # 来自服务器的请求为网站根目录时，由视图中的index函数进行处理。
	#url('first', first.admin),
	url('second', first.second),
	url('pccss', first.pccss),
	url('tablecss', first.tablecss),
	url('spcss', first.spcss),
    url('admincheck', first.admincheck),
	url('sysadmin', first.sysadmin,name='sysadmin'),
	url('F0911', Finance.F0911),#會計科目餘額明細
	url('F4211saleitem', Finance.F4211saleitem),
	url('F4211', Finance.F4211),
	url('F03B11', Finance.F03B11),
	url('F43121item', Finance.F43121item),#應付帳款查詢
	url('F43121', Finance.F43121),
	url('saledetelf4211', first.saledetelf4211),
	url('weborderdetel_cc', first.weborderdetel_cc),
	url('weborderdetel', first.weborderdetel),
    url('weborder2jde_cc', first.weborder2jde_cc),
	url('weborder2jde', first.weborder2jde),
    url('weborder_cc', first.weborder_cc),	
	url('weborder', first.weborder),    		
	url('ReceivableItem', Finance.ReceivableItem),
    url('productsts', product.productsts),	
	url('excel', first.excel),
	url('company', Finance.company),
    url('client', Finance.client),
	url('OracleSalesCalc', product.OracleSalesCalc),    
    url('worklog', management.worklog),
    url('logcheck', management.logcheck),
	url('F47121m', logistics.f47121m),
	url('F47121detel', logistics.f47121d),
    url('layashopdata', shopdata.layashopdata),
	url('invoiceCheck', Finance.invoiceCheck), # 每月發票
	url('saleTotal', Finance.saleTotal), # 銷售總表
    url('SALEAREA',shopstock.SALEAREA), #營業區資料
    url('AREASHOP', shopstock.AREASHOP), #營業區門市
    #url('userpage', authority.userpage),
    #url('userlevel', authority.userlevel),
    url('pettycash', Finance.pettycash),
    url('paylist', Finance.paylist),
    url('passchang', first.passchang),
    url('contracts', contracts.contracts),
    url('contsdetel', contracts.contd),
    url('upload', contracts.upload),
	url('F4311item', Finance.F4311item),#未驗收應付帳款查詢
	#url('static', STATICFILES_DIR),
	#path('second', views.retest),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
#urlpatterns += static('/static/', document_root=media_root)
#urlpatterns += staticfiles_urlpatterns
'''urlpatterns = [
    path('admin/', admin.site.urls),
]'''
