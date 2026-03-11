from django.urls import path
from . import views

urlpatterns = [
    # Шаг 1. Ввод личных данных (корень сайта)
    path('', views.index, name='index'),
    
    # Выход из системы
    path('logout/', views.logout, name='logout'),
    
    # Шаг 2. Выбор образовательной программы
    path('programs/', views.programs, name='programs'),
    
    # Шаг 3. Список тем выбранной программы
    path('program/<int:program_id>/topics/', views.topics, name='topics'),
    
    # Шаг 4. Оценка темы: критерии
    path('topic/<int:topic_id>/evaluate/', views.evaluate_topic, name='evaluate_topic'),
    
    # Шаг 5. Просмотр результатов по программе
    path('program/<int:program_id>/results/', views.program_results, name='program_results'),
    
    # Экспорт CSV результатов программы
    path('program/<int:program_id>/export-csv/', views.export_program_csv, name='export_program_csv'),
    
    # Переключение языка (кастомный)
    path('set-language/', views.set_language_custom, name='set_language_custom'),
]
