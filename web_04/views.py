from django.shortcuts import render  # 暂时没有作用
from django.http import HttpResponse  # 从http模块中导入HttpResponse类
from MyWeb import trans # 导入翻译模块


# Create your views here.
def index(request):  # 定义站点首页视图函数
    return HttpResponse('啊！~~这是我的第一次！')  # 返回响应内容对象


def translate(request):  # 定义视图函数
    from_lang = request.GET['from_lang']  # 获取URL中的参数
    to_lang = request.GET['to_lang']  # 获取URL中的参数
    text = request.GET['words']  # 获取URL中的参数
    return HttpResponse(trans.trans(text, from_lang, to_lang))  # 返回响应内容对象


def translate2(request, words, from_lang, to_lang):  # 定义视图函数
    return HttpResponse(trans.trans(words, from_lang, to_lang))  # 返回响应内容对象