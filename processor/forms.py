# processor/forms.py

from django import forms

class UploadDocxForm(forms.Form):
    file = forms.FileField(label='Upload DOCX File')
