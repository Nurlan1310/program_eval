from django.db import models
from django.utils import translation


class Program(models.Model):
    name = models.CharField(max_length=1000)
    description = models.TextField(blank=True)

    def __str__(self):
        return self.name

class Topic(models.Model):
    program = models.ForeignKey(Program, on_delete=models.CASCADE, related_name='topics')
    name = models.CharField(max_length=2700)
    class_level = models.CharField(max_length=100, default="Без класса")
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return f"{self.program.name} — {self.name}"

class CriterionBlock(models.Model):
    """Блок критериев (один блок может включать несколько критериев)."""
    name = models.CharField(max_length=255)
    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def __str__(self):
        return self.name


class Criterion(models.Model):
    MAIN = 'main'
    EXTRA = 'extra'
    TYPE_CHOICES = [(MAIN, 'Основной'), (EXTRA, 'Дополнительный')]

    name = models.CharField(max_length=255)
    name_kz = models.CharField(max_length=255, blank=True, null=True)

    type = models.CharField(max_length=10, choices=TYPE_CHOICES, default=MAIN)

    description = models.TextField(blank=True)  # ← новое поле
    description_kz = models.TextField(blank=True, null=True)  # если нужно

    # Блок, к которому относится критерий (опционально)
    block = models.ForeignKey(
        CriterionBlock,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='criteria',
        verbose_name='Блок'
    )

    order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['order', 'id']

    def display_name(self):
        lang = translation.get_language()
        if lang.startswith("kk") and self.name_kz:
            return self.name_kz
        return self.name

    def display_description(self):
        lang = translation.get_language()
        if lang.startswith("kk") and self.name_kz:
            return self.description_kz
        return self.description

    def __str__(self):
        block_part = f" [{self.block.name}]" if self.block else ""
        return f"{self.name}{block_part} ({self.type})"

class EvaluatorSession(models.Model):
    full_name = models.CharField(max_length=255)
    phone = models.CharField(max_length=50, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)


    def __str__(self):
        return f"{self.full_name} — {self.phone}"

class ProgramEvaluation(models.Model):
    evaluator = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='program_evals')
    program = models.ForeignKey(Program, on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)
    is_completed = models.BooleanField(default=False)

    def __str__(self):
        return f"Eval {self.program.name} by {self.evaluator.full_name}"

class TopicEvaluation(models.Model):
    program_evaluation = models.ForeignKey(ProgramEvaluation, on_delete=models.CASCADE, related_name='topic_evals')
    topic = models.ForeignKey(Topic, on_delete=models.CASCADE)
    comment = models.TextField(blank=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    can_re_evaluate = models.BooleanField(default=False, help_text="Разрешение на переоценку темы")

    class Meta:
        unique_together = ('program_evaluation', 'topic')

    def __str__(self):
        return f"{self.program_evaluation} — {self.topic.name}"

class Answer(models.Model):
    topic_evaluation = models.ForeignKey(TopicEvaluation, on_delete=models.CASCADE, related_name='answers')
    criterion = models.ForeignKey(Criterion, on_delete=models.CASCADE)
    # for main criteria: value 1/2/3. for extra: yes_no
    value = models.IntegerField(null=True, blank=True)
    yes_no = models.BooleanField(null=True, blank=True)
    comment = models.TextField(blank=True)

    class Meta:
        unique_together = ('topic_evaluation', 'criterion')

    def __str__(self):
        return f"Answer to {self.criterion.name}"


from django.conf import settings

class TopicCompletion(models.Model):
    """Вспомогательная модель для отслеживания завершенности тем."""
    evaluator = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='topic_completions')
    topic = models.ForeignKey(Topic, on_delete=models.CASCADE, related_name='completions')
    completed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('evaluator', 'topic')

    def __str__(self):
        return f"{self.evaluator.full_name} — {self.topic.name}"


class ProgramCompletion(models.Model):
    """Вспомогательная модель для отслеживания завершенности программ."""
    evaluator = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='program_completions')
    program = models.ForeignKey(Program, on_delete=models.CASCADE, related_name='completions')
    completed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('evaluator', 'program')

    def __str__(self):
        return f"{self.evaluator.full_name} — {self.program.name}"


class ActionLog(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, null=True, on_delete=models.SET_NULL)
    action = models.CharField(max_length=100)
    object_type = models.CharField(max_length=100, blank=True)
    object_id = models.IntegerField(null=True, blank=True)
    description = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.created_at:%Y-%m-%d %H:%M} — {self.user} — {self.action}"


class EvaluatorAssignment(models.Model):
    """Закрепление 3 оценщиков для класса тем в программе для расчета модальных значений."""
    program = models.ForeignKey(Program, on_delete=models.CASCADE, related_name='evaluator_assignments')
    class_level = models.CharField(max_length=50)
    evaluator1 = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='assignments_as_first')
    evaluator2 = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='assignments_as_second')
    evaluator3 = models.ForeignKey(EvaluatorSession, on_delete=models.CASCADE, related_name='assignments_as_third')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('program', 'class_level')

    def __str__(self):
        return f"{self.program.name} - {self.class_level}: {self.evaluator1.full_name}, {self.evaluator2.full_name}, {self.evaluator3.full_name}"