from django.shortcuts import render
from django.views import generic
from django.contrib.auth.mixins import LoginRequiredMixin
from django.core.files.storage import default_storage
from django.views.generic import TemplateView
import shutil, os

def top(request):
    return render(request, 'quotes/top.html')

def quotelist(request):
    template_name = 'quotes/quote_index.html'
    return render(request, template_name)

def quotedetail(request):
    template_name = 'quotes/quote_details.html'
    return render(request, template_name)

