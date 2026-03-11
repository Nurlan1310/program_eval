# Файлы переводов

Эта директория содержит файлы переводов для интернационализации проекта.

## Структура

```
locale/
├── ru/
│   └── LC_MESSAGES/
│       ├── django.po (исходный файл переводов)
│       └── django.mo (скомпилированный файл)
└── kk/
    └── LC_MESSAGES/
        ├── django.po (исходный файл переводов)
        └── django.mo (скомпилированный файл)
```

## Редактирование переводов

Файлы переводов уже созданы. Вы можете редактировать их вручную:

- `locale/ru/LC_MESSAGES/django.po` - русский язык
- `locale/kk/LC_MESSAGES/django.po` - казахский язык

### Формат файла .po:

```po
msgid "Русский текст"
msgstr "Казахский перевод"
```

## Компиляция переводов

### Вариант 1: С установленным gettext

```bash
python manage.py compilemessages
```

### Вариант 2: Без gettext (установка)

Установите gettext tools:

**Через Chocolatey:**
```powershell
choco install gettext
```

**Или скачайте вручную:**
https://mlocati.github.io/articles/gettext-iconv-windows.html

После установки используйте:
```bash
python manage.py compilemessages
```

### Вариант 3: Использование скрипта

```bash
python compile_translations.py
```

## Создание новых переводов (требует gettext)

Если нужно добавить новые строки для перевода:

```bash
python manage.py makemessages -l ru
python manage.py makemessages -l kk
```

## Примечание

Данные для критериев (названия, описания) будут заполняться вручную через админ-панель Django, поэтому они не требуют перевода через систему i18n. Вы можете создать отдельные записи критериев для каждого языка в админ-панели.

