import shutil
import tempfile

from django.contrib.auth.models import Group, User
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.urls import reverse
from django.utils import timezone

from evaluations.models import (
    Answer,
    Criterion,
    EvaluatorAssignment,
    EvaluatorSession,
    Program,
    ProgramEvaluation,
    Topic,
    TopicEvaluation,
)
from manager.models import ProgramAIAnalysisRun


TEST_MEDIA_ROOT = tempfile.mkdtemp()


@override_settings(MEDIA_ROOT=TEST_MEDIA_ROOT, MEDIA_URL="/media/")
class ProgramAIAnalyticsViewTests(TestCase):
    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        shutil.rmtree(TEST_MEDIA_ROOT, ignore_errors=True)

    def setUp(self):
        self.program = Program.objects.create(name="Тестовая программа", description="Описание")
        self.topic = Topic.objects.create(
            program=self.program,
            name="Тема 1",
            class_level="1 класс",
            order=1,
        )
        self.criterion = Criterion.objects.create(name="Критерий 1", type=Criterion.MAIN, order=1)

        self.manager_group = Group.objects.create(name="SubAdmins")
        self.manager_user = User.objects.create_user(username="manager", password="secret123")
        self.manager_user.groups.add(self.manager_group)

        self.plain_user = User.objects.create_user(username="plain", password="secret123")

        self.evaluators = [
            EvaluatorSession.objects.create(full_name=f"Оценщик {index}", phone=f"+7700000000{index}")
            for index in range(1, 4)
        ]

        for evaluator in self.evaluators:
            program_evaluation = ProgramEvaluation.objects.create(
                evaluator=evaluator,
                program=self.program,
            )
            topic_evaluation = TopicEvaluation.objects.create(
                program_evaluation=program_evaluation,
                topic=self.topic,
                comment=f"Комментарий {evaluator.full_name}",
            )
            topic_evaluation.completed_at = timezone.now()
            topic_evaluation.save(update_fields=["completed_at"])
            Answer.objects.create(
                topic_evaluation=topic_evaluation,
                criterion=self.criterion,
                value=3,
            )

        EvaluatorAssignment.objects.create(
            program=self.program,
            class_level="1 класс",
            evaluator1=self.evaluators[0],
            evaluator2=self.evaluators[1],
            evaluator3=self.evaluators[2],
        )

    def test_ai_analytics_page_requires_subadmin_permissions(self):
        self.client.force_login(self.plain_user)

        response = self.client.get(reverse("program_ai_analytics", args=[self.program.id]))

        self.assertEqual(response.status_code, 403)

    def test_ai_analytics_page_renders_for_subadmin(self):
        self.client.force_login(self.manager_user)

        response = self.client.get(reverse("program_ai_analytics", args=[self.program.id]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Аналитика AI")
        self.assertContains(response, "Запуск аналитики")
        self.assertContains(response, "Сохранить в PDF")

    def test_post_creates_completed_ai_analysis_run(self):
        self.client.force_login(self.manager_user)

        response = self.client.post(
            reverse("program_ai_analytics", args=[self.program.id]),
            {
                "methodology_file": SimpleUploadedFile(
                    "guide.txt",
                    b"\xd0\x9c\xd0\xb5\xd1\x82\xd0\xbe\xd0\xb4\xd0\xb8\xd1\x87\xd0\xba\xd0\xb0\n\xd0\xa6\xd0\xb5\xd0\xbb\xd0\xb8 \xd0\xbf\xd1\x80\xd0\xbe\xd0\xb3\xd1\x80\xd0\xb0\xd0\xbc\xd0\xbc\xd1\x8b",
                    content_type="text/plain",
                )
            },
            follow=True,
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(ProgramAIAnalysisRun.objects.count(), 1)

        run = ProgramAIAnalysisRun.objects.get()
        self.assertEqual(run.status, ProgramAIAnalysisRun.STATUS_COMPLETED)
        self.assertEqual(run.created_by, self.manager_user)
        self.assertIn("Тестовая программа", run.result_text)
        self.assertContains(response, "Аналитический отчет")
        self.assertTrue(run.methodology_file.name.endswith("guide.txt"))

    def test_pdf_download_returns_demo_file(self):
        self.client.force_login(self.manager_user)

        response = self.client.get(reverse("download_program_ai_pdf", args=[self.program.id]))

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/pdf")
        self.assertIn("ai_analytics_demo_report.pdf", response["Content-Disposition"])
