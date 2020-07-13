from django.shortcuts import render
from django.views import generic
from django.contrib.auth.mixins import LoginRequiredMixin
from django.conf import settings
from .forms import UploadForm
from django.core.files.storage import default_storage
from django.views.generic import TemplateView
import shutil, os, re, openpyxl, glob, pickle
from .reading import reading, setting_context


def top(request):
    return render(request, 'quotes/top.html')

class UploadView(LoginRequiredMixin, generic.FormView):
    form_class = UploadForm
    template_name = 'quotes/quote_index.html'

    def form_valid(self, form):
        user_name = self.request.user.username # ログオンユーザー名取得
        user_dir = os.path.join(settings.MEDIA_ROOT, "data", user_name) # ユーザディレクトリパスの生成
        if not os.path.isdir(user_dir): #ユーザーディレクトリの作成
            os.makedirs(user_dir)
        temp_dir = form.save() #Upload一時フォルダの取得
        upload_path = temp_dir + "/*" #アップロードされたファイル名取得
        uploaded_file = glob.glob(upload_path) #アップロードされたリスト作成

        extracted_data = reading(uploaded_file) #Deta抽出
        f = open('quotes/test_list.txt', 'wb')
        pickle.dump(extracted_data, f)
        context = setting_context(extracted_data) #HTML用表示変換

        shutil.rmtree(temp_dir) # Upload一時フォルダの削除
        return render(self.request, 'quotes/quote_details.html', context)

    def from_invalid(self, form):
        return render(self.request, 'quotes/quote_index.html', {'form': form})

def quotedetail(request):
    template_name = 'quotes/quote_details.html'
    uploaded_file = ["quotes/E2004076A01見積書_税抜.xlsx"]

    extracted_data = reading(uploaded_file)
    context = setting_context(extracted_data)

    return render(request, template_name, context)

