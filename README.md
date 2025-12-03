Скрипт `check_format.py` проверяет DOCX-файлы статей на соответствие правилам оформления (формат сборника академии).

## Запуск CLI

```bash
python check_format.py article.docx
python check_format.py article1.docx article2.docx
python check_format.py --json article.docx
python check_format.py --no-annotate article.docx
```

## Запуск GUI

Простейшее десктоп-приложение на Tkinter:

```bash
python format_checker_gui.py
```

В окне можно выбрать DOCX-файл, запустить проверку, увидеть список замечаний и подсветку проблемных абзацев. При наличии ошибок/предупреждений доступна кнопка для открытия аннотированной копии документа.

## Требования

* Python 3.10+
* Библиотека `python-docx` (`pip install python-docx`)
* Tkinter входит в стандартную поставку Python под Windows

## Отчёт

По каждому файлу выводится список правил со статусами `OK`, `WARN` или `ERROR` с указанием номера абзаца и краткого контекста. Для JSON-режима отчёт печатается в stdout как список объектов.

Дополнительно скрипт создаёт копию исходного документа с суффиксом `_annotated.docx`, где абзацы с ошибками и предупреждениями подсвечены красным. Если подсветка не нужна, используйте флаг `--no-annotate`.
