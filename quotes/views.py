from django.shortcuts import render
from django.views import generic
from django.contrib.auth.mixins import LoginRequiredMixin
from django.core.files.storage import default_storage
from django.views.generic import TemplateView
import shutil, os
import re
import openpyxl

def top(request):
    return render(request, 'quotes/top.html')

def quotelist(request):
    template_name = 'quotes/quote_index.html'
    return render(request, template_name)

def quotedetail(request):
    template_name = 'quotes/quote_details.html'

    ## excel data 取得

    wb = openpyxl.load_workbook("quotes/E2004076A01見積書_税抜.xlsx")
    ws = wb["template1"]

    quote_header = []

    #宛先取得
    for i in ws.iter_rows(min_col=3, min_row=3, max_col=3, max_row=4):
        values_original = [cell.value for cell in i]
        quote_header.append(values_original)
    # 発行日/見積書NO.取得
    for i in ws.iter_rows(min_col=65, min_row=1, max_col=65, max_row=2):
        values_original = [cell.value for cell in i]
        quote_header.append(values_original)
    # 部署名取得    
    for i in ws.iter_rows(min_col=59, min_row=6, max_col=59, max_row=7):
        values_original = [cell.value for cell in i]
        quote_header.append(values_original)
    # PIC/TEL/FAX NO取得
    for i in ws.iter_rows(min_col=65, min_row=8, max_col=65, max_row=10):
        values_original = [cell.value for cell in i]
        quote_header.append(values_original)

    #文字削除
    delete_moji = r"[ '\[\]\u3000 ]"


    issued_date = re.sub(delete_moji, "", str(quote_header[2]))
    quote_no = re.sub(delete_moji, "", str(quote_header[3]))
    section_1 = re.sub(delete_moji, "", str(quote_header[4]))
    section_2 = re.sub(delete_moji, "", str(quote_header[5]))
    pic = re.sub(delete_moji, "", str(quote_header[6]))
    tel_no = re.sub(delete_moji, "", str(quote_header[7]))
    fax_no = re.sub(delete_moji, "", str(quote_header[8]))

    wb.close()

    context = {
        'address' : quote_header[0:2],
        'issued_date': issued_date,
        'quote_no': quote_no,
        'section_1': section_1,
        'section_2': section_2,
        'pic': pic,
        'tel_no': tel_no,
        'fax_no': fax_no,
        }
    return render(request, template_name, context)

