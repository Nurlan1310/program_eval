from django.contrib import admin
from .models import Program, Topic, CriterionBlock, Criterion, EvaluatorSession, ProgramEvaluation, TopicEvaluation, Answer, TopicCompletion, ProgramCompletion

class TopicInline(admin.TabularInline):
    model = Topic
    extra = 0

@admin.register(Program)
class ProgramAdmin(admin.ModelAdmin):
    list_display = ('name',)
    inlines = [TopicInline]

@admin.register(Topic)
class TopicAdmin(admin.ModelAdmin):
    list_display = ('name','program','class_level')
    list_filter = ('program','class_level')

@admin.register(CriterionBlock)
class CriterionAdmin(admin.ModelAdmin):
    list_display = ('name',)

@admin.register(Criterion)
class CriterionAdmin(admin.ModelAdmin):
    list_display = ('name','type')
    search_fields = ('name',)
    list_filter = ('type',)

@admin.register(EvaluatorSession)
class EvaluatorAdmin(admin.ModelAdmin):
    list_display = ('full_name','phone','created_at')
    search_fields = ('full_name','phone')

@admin.register(ProgramEvaluation)
class ProgramEvaluationAdmin(admin.ModelAdmin):
    list_display = ('program','evaluator','created_at')

@admin.register(TopicEvaluation)
class TopicEvaluationAdmin(admin.ModelAdmin):
    list_display = ('program_evaluation','topic','completed_at')

@admin.register(Answer)
class AnswerAdmin(admin.ModelAdmin):
    list_display = ('topic_evaluation','criterion','value','yes_no')

@admin.register(TopicCompletion)
class TopicCompletionAdmin(admin.ModelAdmin):
    list_display = ('evaluator', 'topic', 'completed_at')
    list_filter = ('completed_at', 'topic__program')
    search_fields = ('evaluator__full_name', 'topic__name')

@admin.register(ProgramCompletion)
class ProgramCompletionAdmin(admin.ModelAdmin):
    list_display = ('evaluator', 'program', 'completed_at')
    list_filter = ('completed_at', 'program')
    search_fields = ('evaluator__full_name', 'program__name')

from .models import ActionLog

@admin.register(ActionLog)
class ActionLogAdmin(admin.ModelAdmin):
    list_display = ('created_at','user','action','object_type','object_id')
    list_filter = ('action','object_type','created_at','user')
    search_fields = ('description','user__username')