from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ("evaluations", "0009_alter_topic_name"),
    ]

    operations = [
        migrations.CreateModel(
            name="ProgramAIAnalysisRun",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("provider_key", models.CharField(default="stub", max_length=50)),
                (
                    "status",
                    models.CharField(
                        choices=[("pending", "Ожидает"), ("completed", "Завершено"), ("failed", "Ошибка")],
                        default="pending",
                        max_length=20,
                    ),
                ),
                ("methodology_file", models.FileField(upload_to="ai_analytics/%Y/%m/%d/")),
                ("methodology_filename", models.CharField(max_length=255)),
                ("methodology_excerpt", models.TextField(blank=True)),
                ("context_json", models.JSONField(blank=True, default=dict)),
                ("result_text", models.TextField(blank=True)),
                ("error_message", models.TextField(blank=True)),
                ("created_at", models.DateTimeField(auto_now_add=True)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                (
                    "created_by",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="program_ai_analysis_runs",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
                (
                    "program",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="ai_analysis_runs",
                        to="evaluations.program",
                    ),
                ),
            ],
            options={
                "verbose_name": "Запуск AI-аналитики",
                "verbose_name_plural": "Запуски AI-аналитики",
                "ordering": ["-created_at", "-id"],
            },
        ),
    ]
