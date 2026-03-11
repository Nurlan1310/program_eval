from .models import EvaluatorSession
from django.utils import translation
from django.conf import settings

def evaluator_context(request):
    """Context processor для передачи evaluator и языка во все шаблоны"""
    evaluator = None
    evaluator_id = request.session.get('evaluator_id')
    if evaluator_id:
        try:
            evaluator = EvaluatorSession.objects.get(id=evaluator_id)
        except EvaluatorSession.DoesNotExist:
            pass
    
    # Получаем язык из сессии или из cookie, или используем текущий активный язык
    current_language = request.session.get(settings.LANGUAGE_COOKIE_NAME)
    if not current_language:
        # Проверяем cookie
        current_language = request.COOKIES.get(settings.LANGUAGE_COOKIE_NAME)
    if not current_language:
        # Используем текущий активный язык
        current_language = translation.get_language()
    
    # Берем только код (ru или kk), без локали
    if current_language:
        current_language = current_language.split('-')[0].split('_')[0]
    
    # Убеждаемся, что язык активирован
    if current_language and current_language in ['ru', 'kk']:
        translation.activate(current_language)
    
    return {
        'evaluator': evaluator,
        'current_language': current_language or 'ru',
    }

