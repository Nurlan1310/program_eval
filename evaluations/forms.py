from django import forms
from .models import EvaluatorSession
import re

class StartForm(forms.ModelForm):
    class Meta:
        model = EvaluatorSession
        fields = ['full_name', 'phone']

    def clean(self):
        cleaned = super().clean()
        phone = cleaned.get('phone')

        if not phone:
            return cleaned

        # Если телефон существует — НЕ ошибка, это повторный вход
        existing = EvaluatorSession.objects.filter(phone=phone).first()
        if existing:
            # не создаём ошибку!
            self.existing_user = existing
        return cleaned


class EvaluatorForm(forms.ModelForm):
    """Форма для ввода ФИО и телефона (основная)."""
    class Meta:
        model = EvaluatorSession
        fields = ["full_name", "phone"]
        widgets = {
            "full_name": forms.TextInput(attrs={"class": "form-control", "placeholder": "Ваше ФИО"}),
            "phone": forms.TextInput(attrs={"class": "form-control", "placeholder": "+7..."}),
        }

    def clean_phone(self):
        """Проверка телефона: формат +7 и 10 цифр, если существует - это нормально, это повторный вход"""
        phone = self.cleaned_data.get("phone", "").strip()
        
        if not phone:
            raise forms.ValidationError("Номер телефона обязателен для заполнения.")
        
        # Нормализуем телефон: убираем все кроме цифр и +
        phone_normalized = re.sub(r'[^\d+]', '', phone)
        
        # Проверяем формат: должен начинаться с +7 и содержать ровно 10 цифр после +7
        if not phone_normalized.startswith('+7'):
            # Пытаемся исправить: если начинается с 7 или 8, заменяем на +7
            if phone_normalized.startswith('7'):
                phone_normalized = '+7' + phone_normalized[1:]
            elif phone_normalized.startswith('8'):
                phone_normalized = '+7' + phone_normalized[1:]
            else:
                phone_normalized = '+7' + phone_normalized
        
        # Проверяем, что после +7 идет ровно 10 цифр
        digits_after_plus7 = phone_normalized[2:]  # Все после +7
        if not digits_after_plus7.isdigit() or len(digits_after_plus7) != 10:
            raise forms.ValidationError(
                "Номер телефона должен быть в формате +7XXXXXXXXXX (10 цифр после +7). "
                "Например: +79123456789"
            )
        
        # Форматируем в стандартный вид: +7XXXXXXXXXX
        phone = '+7' + digits_after_plus7
        
        # Проверяем, существует ли пользователь с таким телефоном
        existing = EvaluatorSession.objects.filter(phone=phone).first()
        if existing:
            # Сохраняем существующего пользователя для использования в clean() и save()
            self.existing_user = existing
            # НЕ выдаем ошибку уникальности - это нормальный повторный вход
        
        return phone

    def validate_unique(self):
        """Переопределяем проверку уникальности: для существующего телефона не выдаем ошибку"""
        # Получаем телефон из cleaned_data
        phone = self.cleaned_data.get('phone', '').strip()
        
        if phone:
            existing = EvaluatorSession.objects.filter(phone=phone).first()
            if existing:
                # Пользователь существует - это нормально, не проверяем уникальность
                self.existing_user = existing
                # Пропускаем стандартную проверку уникальности
                return
        
        # Для нового пользователя проверяем уникальность как обычно
        super().validate_unique()
    
    def clean(self):
        """Дополнительная валидация формы"""
        cleaned = super().clean()
        
        # Если пользователь уже существует, это нормально
        # existing_user уже установлен в clean_phone() или validate_unique()
        if hasattr(self, 'existing_user') and self.existing_user:
            # Пользователь существует - это повторный вход, не ошибка
            pass
        
        return cleaned
    
    def save(self, commit=True):
        """Переопределяем save: если пользователь существует, возвращаем его вместо создания нового"""
        if hasattr(self, 'existing_user') and self.existing_user:
            # Пользователь уже существует - обновляем имя если нужно
            if self.existing_user.full_name != self.cleaned_data.get('full_name'):
                self.existing_user.full_name = self.cleaned_data.get('full_name')
                if commit:
                    self.existing_user.save(update_fields=['full_name'])
            return self.existing_user
        
        # Пользователя нет - создаем нового
        return super().save(commit=commit)

