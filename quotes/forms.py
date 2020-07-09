from django import forms
from django.conf import settings
from django.core.files.storage import default_storage
import os, random, string

class UploadForm(forms.Form):
    """excel fileアップロード用のフォーム
       saveメソッドはアップロードしたexcelフォームを一時フォルダに保存する。
    """
    document = forms.FileField(label="smil excel見積書データアップロード",
        widget=forms.ClearableFileInput(attrs={'multiple': True})
        )

    def save(self):
        upload_files = self.files.getlist('document')
        temp_dir = os.path.join(settings.MEDIA_ROOT, self.create_dir(10)) # 一時フォルダの生成
        for excel in upload_files:
            default_storage.save(os.path.join(temp_dir, excel.name), excel) # 一時フォルダにExcelファイルを保存
        return temp_dir

    def create_dir(self, n):
        """一時フォルダ名生成関数"""
        return 'excel\\' + ''.join(random.choices(string.ascii_letters + string.digits, k=n))
