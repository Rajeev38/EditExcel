from django.contrib import admin
from django.urls import path
from . import views


urlpatterns = [
    path('',views.index,name='index'),
    path('save',views.savefile,name='savefile'),
    path('create',views.create,name='create'),
    path('add',views.add,name='add'),
    path('update',views.update,name='update'),
    path('delete',views.delete,name='delete'),
    path('download',views.download,name='download')

]
