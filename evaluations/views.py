from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.utils import timezone
from django.http import HttpResponse, JsonResponse
from django.db import transaction
from django.utils import translation
from django.utils.translation import activate, get_language
from django.views.decorators.http import require_POST
from django.conf import settings
from collections import defaultdict
from .models import (
    Program, Topic, Criterion, EvaluatorSession, ProgramEvaluation, 
    TopicEvaluation, Answer, TopicCompletion, ProgramCompletion
)
from .forms import EvaluatorForm
from .utils import log_action
import csv


def index(request):
    """Шаг 1. Ввод личных данных (корень сайта /)"""
    # Если пользователь уже авторизован, перенаправляем на программы
    if request.session.get('evaluator_id'):
        return redirect('programs')
    
    if request.method == 'POST':
        form = EvaluatorForm(request.POST)
        if form.is_valid():
            # form.save() автоматически вернет существующего пользователя или создаст нового
            evaluator = form.save()

            # Сохраняем evaluator_id в session
            request.session['evaluator_id'] = evaluator.id
            return redirect('programs')
    else:
        form = EvaluatorForm()

    return render(request, "evaluations/start.html", {"form": form})


def logout(request):
    """Выход из системы - очистка session"""
    if 'evaluator_id' in request.session:
        del request.session['evaluator_id']
    return redirect('index')


def get_evaluator_from_session(request):
    """Получить evaluator из session или вернуть None"""
    evaluator_id = request.session.get('evaluator_id')
    if evaluator_id:
        try:
            return EvaluatorSession.objects.get(id=evaluator_id)
        except EvaluatorSession.DoesNotExist:
            del request.session['evaluator_id']
    return None


def programs(request):
    """Шаг 2. Выбор образовательной программы (/programs/)"""
    evaluator = get_evaluator_from_session(request)
    if not evaluator:
        return redirect('index')

    program_stats = []

    for program in Program.objects.all():
        total_topics = program.topics.count()

        pe = ProgramEvaluation.objects.filter(
            evaluator=evaluator,
            program=program
        ).first()

        if not pe:
            status = "not_started"
            evaluated_topics = 0
        else:
            evaluated_topics = pe.topic_evals.filter(completed_at__isnull=False).count()

            if evaluated_topics == 0:
                status = "not_started"
            elif evaluated_topics < total_topics:
                status = "partial"
            else:
                status = "completed"
                if not pe.is_completed:
                    pe.is_completed = True
                    pe.save(update_fields=["is_completed"])
                    # Создаём ProgramCompletion
                    ProgramCompletion.objects.get_or_create(
                        evaluator=evaluator,
                        program=program
                    )

        program_stats.append({
            "program": program,
            "total": total_topics,
            "evaluated": evaluated_topics,
            "status": status,
        })

    # Группируем программы по статусам и сортируем по названию
    completed_programs = sorted(
        [p for p in program_stats if p["status"] == "completed"],
        key=lambda x: x["program"].name.lower()
    )
    partial_programs = sorted(
        [p for p in program_stats if p["status"] == "partial"],
        key=lambda x: x["program"].name.lower()
    )
    not_started_programs = sorted(
        [p for p in program_stats if p["status"] == "not_started"],
        key=lambda x: x["program"].name.lower()
    )

    return render(
        request,
        "evaluations/programs.html",
        {
            "evaluator": evaluator,
            "completed_programs": completed_programs,
            "partial_programs": partial_programs,
            "not_started_programs": not_started_programs,
        }
    )


def topics(request, program_id):
    """Шаг 3. Список тем выбранной программы (/program/<id>/topics/)"""
    evaluator = get_evaluator_from_session(request)
    if not evaluator:
        return redirect('index')

    program = get_object_or_404(Program, id=program_id)
    
    # Создаём ProgramEvaluation если его нет
    pe, _ = ProgramEvaluation.objects.get_or_create(evaluator=evaluator, program=program)

    # Получаем все темы программы
    all_topics = program.topics.all()
    
    # Группируем темы по классам
    grouped = defaultdict(list)
    for topic in all_topics:
        klass = topic.class_level or "Без класса"
        
        # Проверяем статус темы
        te = TopicEvaluation.objects.filter(
            program_evaluation=pe,
            topic=topic
        ).first()
        
        if te and te.completed_at:
            if te.can_re_evaluate:
                status = 'can_re_evaluate'
            else:
                status = 'completed'
        else:
            status = 'not_evaluated'
        
        grouped[klass].append((topic, status, te))
    
    # Сортируем классы
    sorted_classes = sorted(grouped.keys())
    grouped_sorted = {klass: grouped[klass] for klass in sorted_classes}
    
    # Проверяем, все ли темы оценены
    total_topics = all_topics.count()
    evaluated_topics = pe.topic_evals.filter(completed_at__isnull=False).count()
    all_completed = (evaluated_topics == total_topics and total_topics > 0)

    return render(request, "evaluations/topics.html", {
        "program": program,
        "grouped": grouped_sorted,
        "evaluator": evaluator,
        "all_completed": all_completed,
    })


@transaction.atomic
def evaluate_topic(request, topic_id):
    """Шаг 4. Оценка темы: критерии (/topic/<id>/evaluate/)"""
    evaluator = get_evaluator_from_session(request)
    if not evaluator:
        return redirect('index')

    topic = get_object_or_404(Topic, id=topic_id)
    program = topic.program

    pe, _ = ProgramEvaluation.objects.get_or_create(evaluator=evaluator, program=program)

    # Проверяем, есть ли уже оценка и разрешена ли переоценка
    existing_te = TopicEvaluation.objects.filter(program_evaluation=pe, topic=topic).first()
    if existing_te and existing_te.completed_at and not existing_te.can_re_evaluate:
        # Тема уже оценена и переоценка не разрешена
        return redirect("topics", program_id=program.id)

    main_criteria = Criterion.objects.filter(type=Criterion.MAIN).order_by('order', 'id')
    extra_criteria = Criterion.objects.filter(type=Criterion.EXTRA).order_by('order', 'id')

    if request.method == "POST":
        te, created = TopicEvaluation.objects.get_or_create(program_evaluation=pe, topic=topic)

        # основные критерии
        for c in main_criteria:
            val = request.POST.get(f"main_{c.id}")
            if val:
                Answer.objects.update_or_create(
                    topic_evaluation=te,
                    criterion=c,
                    defaults={"value": int(val), "yes_no": None},
                )

        # доп критерии - теперь целое число
        for c in extra_criteria:
            val = request.POST.get(f"extra_{c.id}")
            if val:
                try:
                    int_val = int(val)
                    Answer.objects.update_or_create(
                        topic_evaluation=te,
                        criterion=c,
                        defaults={"value": int_val, "yes_no": None},
                    )
                except ValueError:
                    pass  # Игнорируем невалидные значения

        te.comment = request.POST.get("comment", "")
        te.completed_at = timezone.now()
        # Сбрасываем флаг разрешения на переоценку после сохранения
        te.can_re_evaluate = False
        te.save()

        # Создаём TopicCompletion
        TopicCompletion.objects.get_or_create(
            evaluator=evaluator,
            topic=topic
        )

        return redirect("topics", program_id=program.id)

    return render(request, "evaluations/evaluate_topic.html", {
        "topic": topic,
        "evaluator": evaluator,
        "main_criteria": main_criteria,
        "extra_criteria": extra_criteria,
        "existing_evaluation": existing_te,
    })


def program_results(request, program_id):
    """Шаг 5. Просмотр результатов по программе (/program/<id>/results/)"""
    evaluator = get_evaluator_from_session(request)
    if not evaluator:
        return redirect('index')

    program = get_object_or_404(Program, id=program_id)
    pe = ProgramEvaluation.objects.filter(evaluator=evaluator, program=program).first()
    
    if not pe:
        return redirect('topics', program_id=program_id)

    # Получаем все оценки тем для этого evaluator и программы
    topic_evals = TopicEvaluation.objects.filter(
        program_evaluation=pe
    ).select_related('topic').prefetch_related('answers__criterion').order_by('topic__order', 'topic__id')

    # Формируем данные для таблицы
    results_data = []
    main_criteria = list(Criterion.objects.filter(type=Criterion.MAIN).order_by('order', 'id'))
    extra_criteria = list(Criterion.objects.filter(type=Criterion.EXTRA).order_by('order', 'id'))
    
    for te in topic_evals:
        # Создаём словари для ответов, где ключ - это ID критерия
        main_answers_dict = {}
        extra_answers_dict = {}
        
        # Собираем ответы по критериям
        for answer in te.answers.all():
            if answer.criterion.type == Criterion.MAIN:
                main_answers_dict[answer.criterion.id] = answer.value
            else:
                # Для extra критериев теперь используем value (целое число)
                extra_answers_dict[answer.criterion.id] = answer.value
        
        # Формируем списки ответов в порядке критериев
        main_answers_list = []
        for c in main_criteria:
            main_answers_list.append(main_answers_dict.get(c.id, None))
        
        extra_answers_list = []
        for c in extra_criteria:
            extra_answers_list.append(extra_answers_dict.get(c.id, None))
        
        topic_data = {
            'topic': te.topic,
            'comment': te.comment,
            'main_answers': main_answers_list,
            'extra_answers': extra_answers_list,
        }
        
        results_data.append(topic_data)

    return render(request, "evaluations/results.html", {
        "program": program,
        "evaluator": evaluator,
        "results_data": results_data,
        "main_criteria": main_criteria,
        "extra_criteria": extra_criteria,
    })


def export_program_csv(request, program_id):
    """Экспорт результатов программы в CSV"""
    evaluator = get_evaluator_from_session(request)
    if not evaluator:
        return redirect('index')

    program = get_object_or_404(Program, id=program_id)
    pe = ProgramEvaluation.objects.filter(evaluator=evaluator, program=program).first()
    
    if not pe:
        return redirect('topics', program_id=program_id)

    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="program_{program_id}_results.csv"'
    
    # Добавляем BOM для корректного отображения кириллицы в Excel
    response.write('\ufeff')
    
    writer = csv.writer(response)
    
    # Заголовки
    writer.writerow([
        "Программа",
        "Тема",
        "Критерий",
        "Оценка",
        "Да/Нет",
        "Комментарий",
    ])

    topic_evals = TopicEvaluation.objects.filter(
        program_evaluation=pe
    ).select_related('topic').prefetch_related('answers__criterion')

    for te in topic_evals:
        for ans in te.answers.all():
            value_str = str(ans.value) if ans.value is not None else ""
            # yes_no больше не используется для extra критериев, но оставляем для обратной совместимости
            yes_no_str = "Да" if ans.yes_no else "Нет" if ans.yes_no is False else ""
            
            writer.writerow([
                program.name,
                te.topic.name,
                ans.criterion.name,
                value_str,
                yes_no_str,
                te.comment,
            ])

    return response


@require_POST
def set_language_custom(request):
    """Кастомный view для переключения языка"""
    language = request.POST.get('language')
    next_url = request.POST.get('next', '/')
    
    if language and language in ['ru', 'kk']:
        # Активируем язык немедленно
        translation.activate(language)
        # Сохраняем язык в сессию
        request.session[settings.LANGUAGE_COOKIE_NAME] = language
        request.session.save()  # Принудительно сохраняем сессию
        
        # Очищаем префикс языка из URL, если он есть (для i18n_patterns)
        if next_url.startswith('/kk/') or next_url.startswith('/ru/'):
            # Убираем префикс языка
            parts = next_url.split('/')
            if len(parts) > 2:
                next_url = '/' + '/'.join(parts[2:])
            else:
                next_url = '/'
        
        # Сохраняем язык в cookie
        response = redirect(next_url)
        response.set_cookie(settings.LANGUAGE_COOKIE_NAME, language, max_age=365*24*60*60, path='/')
        return response
    
    return redirect(next_url or '/')
