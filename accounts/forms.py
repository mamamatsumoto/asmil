from django.contrib.auth.forms import AuthenticationForm

class LoginForm(AuthenticationForm):
    """Sing in forms"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            field.widget.attrs['class'] = 'form-control'
            field.widget.attrs['placeholder'] = field.label
