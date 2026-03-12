from django.conf import settings
from django.db import models

from evaluations.models import Program


class ProgramAIAnalysisRun(models.Model):
    STATUS_PENDING = "pending"
    STATUS_COMPLETED = "completed"
    STATUS_FAILED = "failed"
    STATUS_CHOICES = [
        (STATUS_PENDING, "Ожидает"),
        (STATUS_COMPLETED, "Завершено"),
        (STATUS_FAILED, "Ошибка"),
    ]

    program = models.ForeignKey(
        Program,
        on_delete=models.CASCADE,
        related_name="ai_analysis_runs",
    )
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="program_ai_analysis_runs",
    )
    provider_key = models.CharField(max_length=50, default="stub")
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default=STATUS_PENDING,
    )
    methodology_file = models.FileField(upload_to="ai_analytics/%Y/%m/%d/")
    methodology_filename = models.CharField(max_length=255)
    methodology_excerpt = models.TextField(blank=True)
    context_json = models.JSONField(default=dict, blank=True)
    result_text = models.TextField(blank=True)
    error_message = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-created_at", "-id"]
        verbose_name = "Запуск AI-аналитики"
        verbose_name_plural = "Запуски AI-аналитики"

    def __str__(self):
        return f"AI analytics for {self.program.name} ({self.created_at:%Y-%m-%d %H:%M})"
