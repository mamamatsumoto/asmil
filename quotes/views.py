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

    # Header情報取得
    for row in ws.iter_rows(min_col=3, max_col=78, max_row=29):
        for cell in row:
            cell_value = cell.value
            if cell_value is not None:
                quote_header.append(cell_value)
    
    quote_details = []

    # 見積金額詳細情報取得
    for i in ws.iter_rows(min_col=2, min_row=32, max_col=78, max_row=35):
        values_original = [cell.value for cell in i]
        values_original_filtered = [j for j in values_original if j is not None]
        quote_details.append(values_original_filtered)

    dt_lists = []

    for k in quote_details:

        dt_lists.append({
            "works" : k[0],
            "volume" : '{:,.3f}'.format(k[2]),
            "unit_v" : k[3],
            "currency" : k[4],
            "unit_p" : '{:,.2f}'.format(k[5]),
            "amount" : '{:,.2f}'.format(k[6]),
            "ex_rate" : k[7],
            "amount_j" : '{:,}'.format(k[8]),
            "tax" : k[9],
            })

    wb.close()

    context = {
        'address' : quote_header[4:6],
        'issued_date': quote_header[1],
        'quote_no': quote_header[3],
        'section_1': quote_header[6],
        'section_2': quote_header[8],
        'pic': quote_header[10],
        'tel_no': quote_header[12],
        'fax_no': quote_header[14],
        'class1' : quote_header[19],
        'class2' : quote_header[21],
        'project_name' : quote_header[23],
        'item' : quote_header[25],
        'from' : quote_header[27],
        'to' : quote_header[30],
        'POL' : quote_header[33],
        'POD' : quote_header[36],
        'execute_from' : quote_header[39],
        'execute_upto' : quote_header[41],
        'validity' : quote_header[43],
        'packages' : '{:,}'.format(quote_header[46]),
        'max_l' : '{:,}'.format(quote_header[49]),
        't_gw' : '{:,}'.format(quote_header[52]),
        'max_w' : '{:,}'.format(quote_header[55]),
        't_mm' : '{:,.3f}'.format(quote_header[58]),
        'max_h' : '{:,}'.format(quote_header[61]),
        't_ft' : '{:,.3f}'.format(quote_header[64]),
        'max_wt' : '{:,}'.format(quote_header[67]),
        't_rt' : '{:,.3f}'.format(quote_header[70]),
        'cwt' : '{:,.1f}'.format(quote_header[73]),
        'scope' : quote_header[75:78],
        }
        
    context["dt_lists"] = dt_lists

    return render(request, template_name, context)

