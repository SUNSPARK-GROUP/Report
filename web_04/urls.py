"""MyWeb URL Configuration

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
from django.contrib import admin
from django.urls import path
from django.urls import re_path  # 导入通过正则表达式处理路径的方法
from MySite import views as siteviews  # 导入模块并创建别名

urlpatterns = [
    path('', siteviews.index),
    path('home.html', siteviews.home),
    path('trans/', siteviews.translate),  # 通过?传递参数进行处理
    path('trans/<str:from_lang>/<str:to_lang>/<str:words>', siteviews.translate2),  # 通过定义url指定部分为参数进行处理
    path('trans&<str:from_lang>&<str:to_lang>&<str:words>', siteviews.translate2),  # 通过定义url指定部分为参数进行处理
    re_path('trans/(.+)&(.+)-(.+)$', siteviews.translate2),  # 通过正则表达式获取url匹配部分为参数进行处理
    path('admin/', admin.site.urls),
]
