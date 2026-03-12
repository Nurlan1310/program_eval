from collections import Counter, defaultdict
from io import BytesIO

import openpyxl
from django.utils import timezone

from evaluations.models import (
    Answer,
    Criterion,
    EvaluatorAssignment,
    ProgramEvaluation,
    TopicEvaluation,
)


MAX_DOCUMENT_EXCERPT = 4000


def calc_modal(values):
    if not values:
        return ("-", "conditional", 0, 0)

    freq = Counter(values)
    modal_value, max_count = next(item for item in freq.items() if item[1] == max(freq.values()))
    total = len(values)

    if max_count == 1 and total > 1:
        degree = "conditional"
    elif max_count == 2:
        degree = "acceptable"
    else:
        degree = "exact"

    return (modal_value, degree, max_count, total)


def extract_document_text(filename, file_bytes):
    extension = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""

    if extension in {"txt", "md", "csv", "json"}:
        return _decode_text_bytes(file_bytes)

    if extension == "xlsx":
        return _extract_xlsx_text(file_bytes)

    return (
        "Автоизвлечение текста для этого типа файла пока не реализовано. "
        "Файл сохранен и будет доступен для будущей AI-интеграции."
    )


def build_program_overview(program):
    completed_topic_evals = TopicEvaluation.objects.filter(
        program_evaluation__program=program,
        completed_at__isnull=False,
    )
    total_topics = program.topics.count()
    evaluators_count = completed_topic_evals.values("program_evaluation__evaluator").distinct().count()
    completed_evaluations = completed_topic_evals.count()
    completion_percent = 0
    if total_topics > 0:
        completion_percent = round(
            (completed_evaluations / (total_topics * max(evaluators_count, 1))) * 100,
            1,
        )

    class_levels = sorted({topic.class_level or "Без класса" for topic in program.topics.all()})
    assignments = EvaluatorAssignment.objects.filter(program=program)
    assigned_classes = {assignment.class_level for assignment in assignments}
    missing_classes = sorted(set(class_levels) - assigned_classes)

    return {
        "total_topics": total_topics,
        "evaluators_count": evaluators_count,
        "completed_evaluations": completed_evaluations,
        "completion_percent": completion_percent,
        "has_all_assignments": not missing_classes,
        "missing_classes": missing_classes,
        "class_levels": class_levels,
        "assignments_count": assignments.count(),
    }


def build_program_ai_context(program, document_name="", document_text=""):
    overview = build_program_overview(program)
    topics = list(program.topics.order_by("order", "id"))
    criteria = list(Criterion.objects.select_related("block").order_by("order", "id"))
    criteria_by_id = {criterion.id: criterion for criterion in criteria}

    topic_evaluations = list(
        TopicEvaluation.objects.filter(
            program_evaluation__program=program,
            completed_at__isnull=False,
        )
        .select_related("program_evaluation__evaluator", "topic")
        .prefetch_related("answers__criterion")
        .order_by("-completed_at", "id")
    )

    evaluator_summary = {}
    recent_comments = []
    criterion_scores = defaultdict(list)
    modal_answer_values = defaultdict(list)

    assignments_by_class = {
        assignment.class_level: {
            assignment.evaluator1_id,
            assignment.evaluator2_id,
            assignment.evaluator3_id,
        }
        for assignment in EvaluatorAssignment.objects.filter(program=program)
    }

    for topic_eval in topic_evaluations:
        evaluator = topic_eval.program_evaluation.evaluator
        class_level = topic_eval.topic.class_level or "Без класса"
        summary = evaluator_summary.setdefault(
            evaluator.id,
            {
                "id": evaluator.id,
                "full_name": evaluator.full_name,
                "completed_topics": 0,
                "last_completed_at": None,
            },
        )
        summary["completed_topics"] += 1
        if summary["last_completed_at"] is None and topic_eval.completed_at:
            summary["last_completed_at"] = topic_eval.completed_at.isoformat()

        comment = (topic_eval.comment or "").strip()
        if comment and len(recent_comments) < 12:
            recent_comments.append(
                {
                    "topic": topic_eval.topic.name,
                    "class_level": class_level,
                    "evaluator": evaluator.full_name,
                    "comment": comment,
                    "completed_at": topic_eval.completed_at.isoformat() if topic_eval.completed_at else "",
                }
            )

        assigned_evaluators = assignments_by_class.get(class_level, set())
        for answer in topic_eval.answers.all():
            numeric_value = _normalize_answer_value(answer)
            if numeric_value is None:
                continue

            criterion_scores[answer.criterion_id].append(numeric_value)
            if evaluator.id in assigned_evaluators:
                modal_answer_values[(topic_eval.topic_id, answer.criterion_id)].append(numeric_value)

    class_topic_map = defaultdict(list)
    for topic in topics:
        class_topic_map[topic.class_level or "Без класса"].append(topic)

    class_summaries = []
    criterion_modal_degrees = defaultdict(lambda: {"exact": 0, "acceptable": 0, "conditional": 0})
    total_modal_counts = {"exact": 0, "acceptable": 0, "conditional": 0}
    modal_examples = []
    topic_summaries = []

    for class_level in sorted(class_topic_map.keys()):
        class_counts = {"exact": 0, "acceptable": 0, "conditional": 0}
        criteria_with_modal = 0

        for topic in class_topic_map[class_level]:
            topic_degree_counts = {"exact": 0, "acceptable": 0, "conditional": 0}
            topic_modal_items = []
            for criterion in criteria:
                values = modal_answer_values.get((topic.id, criterion.id), [])
                if not values:
                    continue

                modal_value, degree, count_modal, total_answers = calc_modal(values)
                class_counts[degree] += 1
                topic_degree_counts[degree] += 1
                criterion_modal_degrees[criterion.id][degree] += 1
                total_modal_counts[degree] += 1
                criteria_with_modal += 1

                if len(modal_examples) < 15:
                    modal_examples.append(
                        {
                            "topic": topic.name,
                            "class_level": class_level,
                            "criterion": criterion.name,
                            "modal_value": modal_value,
                            "degree": degree,
                            "count_modal": count_modal,
                            "total_answers": total_answers,
                        }
                    )
                topic_modal_items.append(
                    {
                        "criterion": criterion.name,
                        "modal_value": modal_value,
                        "degree": degree,
                        "count_modal": count_modal,
                        "total_answers": total_answers,
                    }
                )

            topic_summaries.append(
                {
                    "topic": topic.name,
                    "class_level": class_level,
                    "modal_summary": topic_degree_counts,
                    "modal_items": topic_modal_items,
                    "comment_count": sum(1 for item in recent_comments if item["topic"] == topic.name),
                }
            )

        class_summaries.append(
            {
                "class_level": class_level,
                "topics_count": len(class_topic_map[class_level]),
                "assigned_evaluators_count": len(assignments_by_class.get(class_level, set())),
                "modal_pairs_count": criteria_with_modal,
                "modal_summary": class_counts,
            }
        )

    criterion_summaries = []
    for criterion in criteria:
        score_distribution = Counter(criterion_scores.get(criterion.id, []))
        criterion_summaries.append(
            {
                "id": criterion.id,
                "name": criterion.name,
                "type": criterion.type,
                "block": criterion.block.name if criterion.block else "Без блока",
                "scores": {str(key): value for key, value in sorted(score_distribution.items())},
                "modal_degrees": criterion_modal_degrees[criterion.id],
            }
        )

    total_modal_pairs = sum(total_modal_counts.values())
    modal_summary = {
        **total_modal_counts,
        "total_pairs": total_modal_pairs,
        "exact_percent": round((total_modal_counts["exact"] / total_modal_pairs) * 100, 1)
        if total_modal_pairs
        else 0,
        "acceptable_percent": round((total_modal_counts["acceptable"] / total_modal_pairs) * 100, 1)
        if total_modal_pairs
        else 0,
        "conditional_percent": round((total_modal_counts["conditional"] / total_modal_pairs) * 100, 1)
        if total_modal_pairs
        else 0,
    }

    return {
        "generated_at": timezone.now().isoformat(),
        "program": {
            "id": program.id,
            "name": program.name,
            "description": program.description or "",
        },
        "document": {
            "name": document_name,
            "excerpt": (document_text or "")[:MAX_DOCUMENT_EXCERPT],
            "characters": len(document_text or ""),
        },
        "overview": overview,
        "counts": {
            "criteria_total": len(criteria),
            "program_evaluations_total": ProgramEvaluation.objects.filter(program=program).count(),
            "topic_evaluations_completed_total": len(topic_evaluations),
            "answers_total": Answer.objects.filter(
                topic_evaluation__program_evaluation__program=program,
                topic_evaluation__completed_at__isnull=False,
            ).count(),
        },
        "class_summaries": class_summaries,
        "evaluator_summaries": list(evaluator_summary.values()),
        "criterion_summaries": criterion_summaries,
        "modal_summary": modal_summary,
        "modal_examples": modal_examples,
        "topic_summaries": topic_summaries,
        "recent_comments": recent_comments,
    }


def _decode_text_bytes(file_bytes):
    for encoding in ("utf-8", "utf-8-sig", "cp1251"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("latin-1", errors="ignore")


def _extract_xlsx_text(file_bytes):
    workbook = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet = workbook.active
    rows = []
    for row in sheet.iter_rows(values_only=True):
        cleaned = [str(cell).strip() for cell in row if cell not in (None, "")]
        if cleaned:
            rows.append(" | ".join(cleaned))
        if len(rows) >= 50:
            break
    if rows:
        return "\n".join(rows)
    return "Файл Excel не содержит читаемых данных в активном листе."


def _normalize_answer_value(answer):
    if answer.value is not None:
        return answer.value
    if answer.yes_no is None:
        return None
    return 1 if answer.yes_no else 0
