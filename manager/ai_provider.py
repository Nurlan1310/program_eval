from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from django.conf import settings


REFERENCE_SECTION_TITLES = [
    "1. Общая характеристика программы",
    "2. Структурный анализ образовательной программы",
    "3. Тематический анализ содержания программы",
    "4. Анализ ожидаемых результатов обучения",
    "5. Сводный результат оценки",
    "6. Анализ соответствия объема программы учебной нагрузке",
    "7. Сильные стороны программы",
    "8. Выявленные недостатки",
    "9. Расширенные рекомендации",
    "10. Детализированный анализ всех тем программы",
    "11. Матрица расхождения экспертных оценок",
]


def generate_program_ai_report(context, provider_key="stub"):
    program = context["program"]
    overview = context["overview"]
    modal_summary = context["modal_summary"]
    document = context["document"]
    reference_titles = _load_reference_report_outline()
    topic_summaries = context.get("topic_summaries", [])
    criterion_summaries = context.get("criterion_summaries", [])
    recent_comments = context.get("recent_comments", [])

    strengths = _build_strengths(context)
    weaknesses = _build_weaknesses(context)
    recommendations = _build_recommendations(context)

    top_topics = sorted(
        topic_summaries,
        key=lambda item: (
            -item["modal_summary"]["conditional"],
            -item["modal_summary"]["acceptable"],
            item["topic"].lower(),
        ),
    )[:8]

    score_rows = []
    for criterion in criterion_summaries[:8]:
        score_rows.append(
            {
                "criterion": criterion["name"],
                "block": criterion["block"],
                "level_1": criterion["scores"].get("1", 0),
                "level_2": criterion["scores"].get("2", 0),
                "level_3": criterion["scores"].get("3", 0),
                "exact": criterion["modal_degrees"].get("exact", 0),
                "acceptable": criterion["modal_degrees"].get("acceptable", 0),
                "conditional": criterion["modal_degrees"].get("conditional", 0),
            }
        )

    sections = [
        {
            "title": reference_titles[0],
            "kind": "paragraphs",
            "paragraphs": [
                f"Образовательная программа «{program['name']}» рассмотрена в рамках демонстрационного AI-анализа по образцу загруженного аналитического отчета.",
                f"В анализ включены данные по {overview['total_topics']} темам, {overview['evaluators_count']} оценщикам и {overview['completed_evaluations']} завершенным оценкам тем.",
                _document_paragraph(document),
            ],
        },
        {
            "title": reference_titles[1],
            "kind": "bullets",
            "lead": "Структурный анализ каркаса программы и экспертного контура показывает следующее:",
            "bullets": [
                f"Классовых групп в программе: {len(overview['class_levels']) or 1}.",
                f"Закреплений оценщиков настроено: {overview['assignments_count']}.",
                (
                    "Закрепления присутствуют по всем классам."
                    if overview["has_all_assignments"]
                    else "Есть классы без закрепленных оценщиков: " + ", ".join(overview["missing_classes"])
                ),
                f"Процент заполнения экспертных оценок составляет {overview['completion_percent']}%.",
            ],
        },
        {
            "title": reference_titles[2],
            "kind": "table",
            "lead": "Ниже показаны темы, где модальные расхождения требуют наибольшего внимания:",
            "columns": ["Тема", "Класс", "Точные", "Допустимые", "Условные", "Комментариев"],
            "rows": [
                [
                    item["topic"],
                    item["class_level"],
                    item["modal_summary"]["exact"],
                    item["modal_summary"]["acceptable"],
                    item["modal_summary"]["conditional"],
                    item["comment_count"],
                ]
                for item in (top_topics or topic_summaries[:5])
            ],
        },
        {
            "title": reference_titles[3],
            "kind": "paragraphs",
            "paragraphs": [
                "Ожидаемые результаты обучения интерпретируются через качество раскрытия критериев, стабильность экспертных совпадений и наличие содержательных комментариев.",
                (
                    "В экспертных комментариях присутствует качественная обратная связь, позволяющая связывать числовые оценки с содержательными замечаниями."
                    if recent_comments
                    else "Качественная интерпретация ограничена, так как комментариев от оценщиков пока немного."
                ),
            ],
        },
        {
            "title": reference_titles[4],
            "kind": "metrics",
            "metrics": [
                {"label": "Точных модальностей", "value": modal_summary["exact"], "suffix": f" / {modal_summary['exact_percent']}%"},
                {"label": "Допустимых модальностей", "value": modal_summary["acceptable"], "suffix": f" / {modal_summary['acceptable_percent']}%"},
                {"label": "Условных модальностей", "value": modal_summary["conditional"], "suffix": f" / {modal_summary['conditional_percent']}%"},
                {"label": "Всего модальных пар", "value": modal_summary["total_pairs"], "suffix": ""},
            ],
        },
        {
            "title": reference_titles[5],
            "kind": "paragraphs",
            "paragraphs": [
                "В демонстрационном каркасе вместо фактической учебной нагрузки используется оценка полноты и согласованности данных.",
                (
                    "Текущий массив данных уже достаточен для предварительного аналитического вывода."
                    if overview["completion_percent"] >= 50
                    else "Объем завершенных оценок пока ограничивает глубину итогового аналитического вывода."
                ),
            ],
        },
        {
            "title": reference_titles[6],
            "kind": "bullets",
            "bullets": strengths,
        },
        {
            "title": reference_titles[7],
            "kind": "bullets",
            "bullets": weaknesses,
        },
        {
            "title": reference_titles[8],
            "kind": "bullets",
            "bullets": recommendations,
        },
        {
            "title": reference_titles[9],
            "kind": "table",
            "lead": "Краткая детализация по критериям и уровням оценивания:",
            "columns": ["Критерий", "Блок", "Уровень 1", "Уровень 2", "Уровень 3", "Точные", "Допустимые", "Условные"],
            "rows": [
                [
                    row["criterion"],
                    row["block"],
                    row["level_1"],
                    row["level_2"],
                    row["level_3"],
                    row["exact"],
                    row["acceptable"],
                    row["conditional"],
                ]
                for row in score_rows
            ],
        },
        {
            "title": reference_titles[10],
            "kind": "table",
            "lead": "Матрица расхождения показывает зоны устойчивости и зоны условной интерпретации:",
            "columns": ["Класс", "Тем", "Точных", "Допустимых", "Условных", "Статус"],
            "rows": [
                [
                    class_item["class_level"],
                    class_item["topics_count"],
                    class_item["modal_summary"]["exact"],
                    class_item["modal_summary"]["acceptable"],
                    class_item["modal_summary"]["conditional"],
                    _class_status_label(class_item["modal_summary"]),
                ]
                for class_item in context.get("class_summaries", [])
            ],
        },
    ]

    report = {
        "title": "Аналитический отчет",
        "subtitle": f"по результатам экспертизы образовательной программы «{program['name']}»",
        "provider_label": f"{provider_key} (демо-режим)",
        "reference_note": "Структура демонстрационного отчета собрана по образцу прикрепленного аналитического документа.",
        "summary_cards": [
            {"label": "Темы", "value": overview["total_topics"]},
            {"label": "Оценщики", "value": overview["evaluators_count"]},
            {"label": "Заполнено оценок", "value": overview["completed_evaluations"]},
            {"label": "Точные модальности", "value": f"{modal_summary['exact_percent']}%"},
        ],
        "sections": sections,
    }

    report["plain_text"] = _build_plain_text_report(report)
    return report


def _build_strengths(context):
    overview = context["overview"]
    modal_summary = context["modal_summary"]
    strengths = []

    if overview["completion_percent"] >= 80:
        strengths.append("Высокий уровень заполнения экспертных оценок по программе.")
    elif overview["completion_percent"] >= 50:
        strengths.append("По программе накоплен достаточный массив данных для первичного аналитического вывода.")

    if modal_summary["exact"] >= modal_summary["conditional"]:
        strengths.append("По значимой части критериев наблюдается устойчивое совпадение мнений закрепленных оценщиков.")

    if not overview["missing_classes"]:
        strengths.append("Закрепление оценщиков по классам настроено полностью.")

    if context.get("recent_comments"):
        strengths.append("В анализ можно включать качественные комментарии оценщиков для интерпретации результатов.")

    return strengths or ["Данных достаточно для демонстрации шаблона аналитического отчета."]


def _build_weaknesses(context):
    overview = context["overview"]
    modal_summary = context["modal_summary"]
    weaknesses = []

    if overview["completion_percent"] < 50:
        weaknesses.append("Низкая доля завершенных оценок ограничивает глубину итогового анализа.")

    if modal_summary["conditional"] > modal_summary["exact"]:
        weaknesses.append("Доля условных модальностей превышает долю точных совпадений, поэтому часть выводов требует осторожной интерпретации.")

    if overview["missing_classes"]:
        weaknesses.append("Не для всех классов настроены закрепленные оценщики: " + ", ".join(overview["missing_classes"]) + ".")

    if not context.get("recent_comments"):
        weaknesses.append("Недостаточно качественных комментариев от оценщиков.")

    return weaknesses or ["Существенных ограничений в демонстрационном наборе данных не выявлено."]


def _build_recommendations(context):
    overview = context["overview"]
    recommendations = [
        "Использовать этот шаблон как основу для подключения реального AI-провайдера с генерацией narrative-отчета.",
        "Добавить в следующий этап разметку отчета по обязательным блокам и тонкую настройку промпта.",
    ]

    if overview["completion_percent"] < 80:
        recommendations.insert(0, "Увеличить количество завершенных оценок по темам до формирования финального заключения.")

    if overview["missing_classes"]:
        recommendations.insert(0, "Завершить закрепление оценщиков по всем классам программы.")

    if context.get("modal_summary", {}).get("conditional", 0):
        recommendations.append("Проверить темы с высокой долей условных модальностей и провести дополнительную калибровку оценщиков.")

    return recommendations[:5]


def _document_paragraph(document):
    if document["characters"]:
        return (
            "Методический документ загружен и использован как текстовая основа для структуры выводов. "
            f"Во входном документе распознано около {document['characters']} символов."
        )
    return "Методический документ сохранен, но его текст не был автоматически извлечен; в демонстрации использована только структура шаблонного отчета."


def _class_status_label(summary):
    if summary["conditional"] > summary["exact"]:
        return "Нужно уточнение"
    if summary["acceptable"] > summary["exact"]:
        return "Умеренная устойчивость"
    return "Стабильная картина"


def _build_plain_text_report(report):
    lines = [report["title"], report["subtitle"], f"Провайдер: {report['provider_label']}", ""]
    for section in report["sections"]:
        lines.append(section["title"])
        if section["kind"] == "paragraphs":
            lines.extend(section.get("paragraphs", []))
        elif section["kind"] == "bullets":
            if section.get("lead"):
                lines.append(section["lead"])
            lines.extend(f"- {item}" for item in section.get("bullets", []))
        elif section["kind"] == "metrics":
            lines.extend(f"- {item['label']}: {item['value']}{item.get('suffix', '')}" for item in section.get("metrics", []))
        elif section["kind"] == "table":
            if section.get("lead"):
                lines.append(section["lead"])
            lines.append(" | ".join(section.get("columns", [])))
            for row in section.get("rows", []):
                lines.append(" | ".join(str(cell) for cell in row))
        lines.append("")
    return "\n".join(lines).strip()


def _load_reference_report_outline():
    reference_path = Path(settings.BASE_DIR) / "Аналитический отчет.docx"
    if not reference_path.exists():
        return REFERENCE_SECTION_TITLES

    try:
        with ZipFile(reference_path) as archive:
            xml_content = archive.read("word/document.xml")
        root = ElementTree.fromstring(xml_content)
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = []
        for paragraph in root.findall(".//w:p", namespace):
            text_parts = [node.text for node in paragraph.findall(".//w:t", namespace) if node.text]
            if text_parts:
                paragraphs.append("".join(text_parts).strip())

        headings = [line for line in paragraphs if _looks_like_heading(line)]
        return headings[: len(REFERENCE_SECTION_TITLES)] or REFERENCE_SECTION_TITLES
    except Exception:
        return REFERENCE_SECTION_TITLES


def _looks_like_heading(value):
    if not value:
        return False
    stripped = value.strip()
    if stripped in {"Аналитический отчет"}:
        return False
    prefix = stripped.split(".", 1)[0]
    return prefix.isdigit()
