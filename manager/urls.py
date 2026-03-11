from django.urls import path
from . import views

urlpatterns = [
    path('login/', views.manager_login, name='manager_login'),
    path('', views.manager_index, name='manager_index'),
    path('program/add/', views.program_add, name='program_add'),
    path('program/<int:pk>/', views.program_detail, name='program_detail'),
    path('program/<int:pk>/statistics/', views.program_statistics, name='program_statistics'),
    path('program/<int:pk>/delete/', views.program_delete, name='program_delete'),
    path('program/<int:pk>/topic/add/', views.topic_add, name='topic_add'),
    path('topic/<int:pk>/delete/', views.topic_delete, name='topic_delete'),
    path('import/', views.import_programs, name='manager_import'),
    path('import/confirm/', views.import_programs, name='manager_import_confirm'),
    path('export/program/<int:program_id>/xlsx/', views.export_program_xlsx, name='export_program_xlsx'),
    path('export/program/<int:program_id>/modal/', views.export_modal_xlsx, name='export_modal_xlsx'),
    path('export/program/<int:program_id>/csv/', views.export_program_csv, name='export_program_csv'),
    path('export/all/xlsx/', views.export_all_xlsx, name='export_all_xlsx'),
    path('export/evaluators/xlsx/', views.export_evaluators_xlsx, name='export_evaluators_xlsx'),
    path('export/evaluations/xlsx/', views.export_evaluations_xlsx, name='export_evaluations_xlsx'),
    path('export/evaluation-time/xlsx/', views.export_evaluation_time_xlsx, name='export_evaluation_time_xlsx'),
    path('program/<int:pk>/assign-evaluators/', views.assign_evaluators, name='assign_evaluators'),
    path('grant-re-evaluation/', views.grant_re_evaluation_access, name='grant_re_evaluation_access'),
    path('grant-re-evaluation/program/', views.grant_re_evaluation_access_program, name='grant_re_evaluation_access_program'),
]