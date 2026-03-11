# Установка GNU gettext tools для Windows

Если вы хотите использовать команды `makemessages` и `compilemessages` Django, вам нужно установить GNU gettext tools.

## Вариант 1: Через Chocolatey (рекомендуется)

Если у вас установлен Chocolatey:

```powershell
choco install gettext
```

## Вариант 2: Ручная установка

1. Скачайте gettext для Windows: https://mlocati.github.io/articles/gettext-iconv-windows.html
2. Или используйте готовый установщик: https://github.com/mlocati/gettext-iconv-windows/releases
3. Распакуйте архив
4. Добавьте путь к `bin` в переменную окружения PATH

## Вариант 3: Использование без gettext

Вы можете редактировать файлы `.po` вручную в директории `locale/`:
- `locale/ru/LC_MESSAGES/django.po` - русский язык
- `locale/kk/LC_MESSAGES/django.po` - казахский язык

После редактирования скомпилируйте их командой:
```bash
python manage.py compilemessages
```

## Структура файла .po

```po
msgid "Русский текст"
msgstr "Казахский перевод"
```

После редактирования всегда выполняйте `compilemessages` для применения изменений.




