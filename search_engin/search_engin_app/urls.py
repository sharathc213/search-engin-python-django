
from django.contrib import admin
from django.urls import path,include
from . import views

urlpatterns = [
    path('', views.index,name='index'),
    path('result.html', views.result,name='result'),
    path('getresults', views.getresults,name='getresults'),
]
