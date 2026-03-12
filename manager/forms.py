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


class AIAnalyticsRunForm(forms.Form):
    methodology_file = forms.FileField(
        label="Файл методички",
        help_text="Поддерживаются текстовые файлы и Excel. Для остальных форматов файл будет сохранен без извлечения текста.",
        widget=forms.ClearableFileInput(
            attrs={
                "class": "form-control",
                "accept": ".txt,.md,.csv,.json,.xlsx,.pdf,.doc,.docx",
            }
        ),
    )