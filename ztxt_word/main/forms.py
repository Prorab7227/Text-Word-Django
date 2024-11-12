from .models import Inputed_text
from django.forms import ModelForm, Textarea

class Inputed_textForm(ModelForm):
    class Meta:
        model = Inputed_text
        fields = ['text']
        widgets = {
            'text': Textarea(
                attrs={'class':'form-control','placeholder':'Input your text'}),
        }