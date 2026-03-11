from django import forms
from evaluations.models import Program, Topic

class ProgramForm(forms.ModelForm):
    class Meta:
        model = Program
        fields = ['name', 'description']

class TopicForm(forms.ModelForm):
    class Meta:
        model = Topic
        fields = ['name', 'class_level']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'class_level': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Например: 1 класс'}),
        }

class ImportForm(forms.Form):
    file = forms.FileField()