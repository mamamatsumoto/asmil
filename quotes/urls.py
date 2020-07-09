from django.contrib import admin
from django.urls import path, include
from . import views

app_name = "quotes"

urlpatterns = [
    path('top/', views.top, name='top'),
    path('quotelists/', views.UploadView.as_view(), name='quotelists'),
    path('quotedetail/', views.quotedetail, name='quotedetail'),
    #path('upload/', views.UploadView.as_view(), name='upload'),
]