# DOCX анкеты: автоматическое заполнение

Скрипт `fill_docx.py` подставляет значения из JSON в DOCX-шаблон, сохраняя форматирование Word-документа.

## Что поддерживается

- Подстановка текста в поля через Jinja2-плейсхолдеры (например, `{{ full_name }}`).
- Варианты выбора/чекбоксы через функцию `checkbox()`.
- Условные блоки через Jinja2 `{% if ... %}`.
- (Опционально) вставка изображений подписи/печати через `images`.

## Установка

```bash
pip install -r requirements.txt
```

## Использование

```bash
python fill_docx.py template.docx filled_fields.json result.docx
```

## Формат JSON

```json
{
  "full_name": "Иванов Иван Иванович",
  "passport_series": "1234",
  "passport_number": "567890",
  "is_foreigner": false,
  "choices": {
    "basis": "charter",
    "representative_role": ["director"]
  },
  "images": {
    "signature": {
      "path": "assets/signature.png",
      "width_mm": 40
    }
  },
  "empty_placeholder": "—",
  "checkbox_symbols": {
    "checked": "☑",
    "unchecked": "☐"
  }
}
```

## Пример разметки шаблона

```text
ФИО: {{ value(full_name) }}
Паспорт: серия {{ value(passport_series) }} № {{ value(passport_number) }}

Основание полномочий:
{{ checkbox("basis", "charter") }} Устав
{{ checkbox("basis", "power_of_attorney") }} Доверенность

{% if is_foreigner %}
Гражданство: {{ value(citizenship) }}
{% endif %}

Подпись: {{ signature }}
```

> `value()` подставит `empty_placeholder`, если значение пустое/`null`.
