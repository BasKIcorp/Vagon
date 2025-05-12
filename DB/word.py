#!/usr/bin/env python3
# -*- coding: utf‑8 -*-
"""
word_marker_tools.py
~~~~~~~~~~~~~~~~~~~~
Замена маркеров вида [таблица.поле] в документах Word (.docx),
устойчивая к случаям, когда маркер разбит на несколько run'ов.

Функциональность
----------------
• create_test_document(path)      — генерирует тестовый шаблон с маркерами
• extract_placeholders(path)      — возвращает список всех маркеров в файле
• replace_placeholders(inp, out, mapping) — подменяет маркеры по словарю
• CLI‑режим: создаёт шаблон → печатает маркеры → делает замену

Требуется библиотека `python‑docx`:
    pip install python-docx
"""

import os
import re
from typing import Dict, Iterable, Set

from docx import Document


# ──────────────────────────── helpers ────────────────────────────────────── #
def _rewrite_paragraphs(
    paragraphs: Iterable,
    pattern: re.Pattern,
    mapping: Dict[str, str],
    document = None,  # Документ для создания новых параграфов
) -> None:
    """
    Склеивает текст всех run'ов параграфа, делает замену и,
    если что‑то изменилось, переписывает параграф одним run'ом.
    
    Обрабатывает специальное форматирование для значений с префиксом:
    - LIST: - создает маркированный список
    """
    for paragraph in paragraphs:
        original = "".join(run.text for run in paragraph.runs)
        
        # Проверка на наличие открывающей и закрывающей скобок в параграфе
        if '[' in original and ']' in original:
            # Ищем маркеры со значением LIST: в первую очередь
            list_marker_found = False
            for marker, value in mapping.items():
                marker_pattern = re.escape(f"[{marker}]")
                if re.search(marker_pattern, original) and value.startswith("LIST:") and document:
                    print(f"Найден LIST маркер: [{marker}], обрабатываем как список...")
                    match = re.search(marker_pattern, original)
                    
                    # Получаем текст до и после маркера
                    pre_text = original[:match.start()]
                    post_text = original[match.end():]
                    
                    # Получаем родительский элемент и индекс текущего параграфа
                    parent = paragraph._p.getparent()
                    idx = parent.index(paragraph._p)
                    
                    # Очищаем текущий параграф и добавляем в него текст перед маркером
                    for run in paragraph.runs[::-1]:
                        paragraph._p.remove(run._r)
                    
                    if pre_text:
                        paragraph.add_run(pre_text)
                        # Нужно создать новый параграф для первого элемента списка
                        next_p = document.add_paragraph()
                        next_p_element = next_p._p
                        # Вставляем новый параграф после текущего
                        parent.insert(idx + 1, next_p_element)
                        idx += 1  # Обновляем индекс для дальнейшей вставки
                    else:
                        # Используем текущий параграф для первого элемента списка
                        next_p = paragraph
                    
                    # Разбиваем список элементов
                    list_items = value[5:].split('|')  # Удаляем "LIST:" и разбиваем по разделителю
                    
                    # Для каждого элемента списка создаем параграф с маркером
                    for i, item in enumerate(list_items):
                        if i == 0:
                            # Первый элемент списка идет в параграф next_p
                            current_p = next_p
                        else:
                            # Создаем новый параграф для следующего элемента списка
                            current_p = document.add_paragraph()
                            current_p_element = current_p._p
                            # Вставляем его после предыдущего
                            parent.insert(idx + i, current_p_element)
                        
                        # Очищаем параграф (на всякий случай)
                        if current_p.text:
                            for run in current_p.runs[::-1]:
                                current_p._p.remove(run._r)
                        
                        # Добавляем маркер и текст элемента списка
                        current_p.add_run("• " + item.strip())
                        
                        # Устанавливаем отступ для элементов списка (в EMU - 1/100 точки)
                        current_p.paragraph_format.left_indent = 720000  # 0.5 дюйма = 720000 EMU
                    
                    # Если есть текст после маркера, добавляем его в новый параграф
                    if post_text:
                        last_p = document.add_paragraph()
                        last_p_element = last_p._p
                        parent.insert(idx + len(list_items), last_p_element)
                        last_p.add_run(post_text)
                    
                    list_marker_found = True
                    break  # Обрабатываем только первый найденный LIST: маркер
            
            if list_marker_found:
                continue  # Переходим к следующему параграфу
            
            # Если LIST: маркеров не найдено, выполняем обычную замену
            was_replaced = False
            for marker, value in mapping.items():
                marker_pattern = re.escape(f"[{marker}]")
                if re.search(marker_pattern, original):
                    print(f"Обычная замена маркера: [{marker}] -> {value}")
                    original = re.sub(marker_pattern, value, original)
                    was_replaced = True
            
            if was_replaced:
                # Удаляем все старые run'ы
                for run in paragraph.runs[::-1]:
                    paragraph._p.remove(run._r)
                # Добавляем один новый с замененным текстом
                paragraph.add_run(original)
        else:
            # Если нет маркеров в параграфе, пропускаем его
            continue


def _collect_markers(paragraph):
    """Собирает маркеры из параграфа, даже если они разбиты на несколько runs"""
    text = ""
    markers = set()
    
    # Обновляем регулярное выражение для поддержки нового формата
    pattern = re.compile(r"\[([^\[\]]+)\]")
    
    for run in paragraph.runs:
        text += run.text
        # Ищем маркеры в текущем тексте
        matches = pattern.finditer(text)
        for match in matches:
            markers.add(match.group(1))
    
    return markers


# ──────────────────────────── core API ───────────────────────────────────── #
def replace_placeholders(
    input_path: str,
    output_path: str,
    mapping: Dict[str, str],
) -> None:
    """
    Заменяет все маркеры, указанные в *mapping*, во всём документе *input_path*
    и сохраняет результат в *output_path*.

    mapping = { 
        "договоры.номер": "2024‑000001",
        "список_работ(договоры.номер)": "работа 1, работа 2",
        "сумма(договоры.номер)": "15000 руб."
    }
    
    Для маркированного списка используйте специальный префикс "LIST:" в значении.
    Например: {"список_работ(договоры.номер)": "LIST:работа 1|работа 2|работа 3"}
    Каждый элемент списка должен быть разделен символом '|'
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Входной файл не найден: {input_path}")

    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise ValueError("Входной и выходной пути не могут совпадать")

    out_dir = os.path.dirname(output_path) or "."
    if not os.access(out_dir, os.W_OK):
        raise PermissionError(f"Нет прав на запись в директорию: {out_dir}")

    # общий регекс вида \[(key1|key2|…)\]
    keys_re = "|".join(re.escape(k) for k in mapping)
    pattern = re.compile(r"\[(" + keys_re + r")\]")
    
    # Выводим все ключи для отладки
    print("Ключи для замены:")
    for k in mapping.keys():
        print(f"  - {k}")
    
    # Все маркеры, которые есть в документе
    doc_markers = extract_placeholders(input_path)
    print("Маркеры в документе:")
    for marker in doc_markers:
        print(f"  - {marker}")
        if marker in mapping:
            print(f"    Значение: {mapping[marker]}")
        else:
            print(f"    !!! НЕТ ЗНАЧЕНИЯ В СЛОВАРЕ !!!")

    doc = Document(input_path)

    # тело
    _rewrite_paragraphs(doc.paragraphs, pattern, mapping, doc)

    # таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _rewrite_paragraphs(cell.paragraphs, pattern, mapping, doc)

    # колонтитулы
    for section in doc.sections:
        _rewrite_paragraphs(section.header.paragraphs, pattern, mapping, doc)
        _rewrite_paragraphs(section.footer.paragraphs, pattern, mapping, doc)

    doc.save(output_path)
    print(f"✓ Файл сохранён: {output_path}")


def extract_placeholders(docx_path: str) -> list:
    """
    Возвращает отсортированный список всех уникальных маркеров в документе.
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)

    doc = Document(docx_path)
    markers: Set[str] = set()

    # тело
    for paragraph in doc.paragraphs:
        markers.update(_collect_markers(paragraph))

    # таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    markers.update(_collect_markers(paragraph))

    # колонтитулы
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            markers.update(_collect_markers(paragraph))
        for paragraph in section.footer.paragraphs:
            markers.update(_collect_markers(paragraph))

    return sorted(markers)


def create_test_document(output_path: str) -> None:
    """
    Создаёт небольшой шаблон для проверки работы замены.
    """
    doc = Document()

    # заголовок
    doc.add_heading("Тестовый документ", level=0)

    # основной текст
    doc.add_paragraph("Договор № [договоры.номер]")
    doc.add_paragraph("от [договоры.дата]")
    doc.add_paragraph("")
    doc.add_paragraph("Вагон № [вагоны.номер]")
    doc.add_paragraph("Подразделение: [вагоны.подразделение]")
    doc.add_paragraph("")
    doc.add_paragraph("Список работ: [список_работ(договоры.номер)]")
    doc.add_paragraph("Сумма: [сумма(договоры.номер)]")

    # таблица
    tbl = doc.add_table(rows=2, cols=2)
    tbl.style = "Table Grid"
    tbl.cell(0, 0).text = "Номер договора"
    tbl.cell(0, 1).text = "[договоры.номер]"
    tbl.cell(1, 0).text = "Дата договора"
    tbl.cell(1, 1).text = "[договоры.дата]"

    doc.save(output_path)
    print(f"Тестовый документ создан: {output_path}")


# ───────────────────────────── CLI demo ──────────────────────────────────── #
if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))

    # 1. Всегда создаём тестовый шаблон
    template = os.path.join(here, "word.docx")
    create_test_document(template)

    # 2. Ищем маркеры
    vars_found = extract_placeholders(template)
    print("Найдены маркеры:", vars_found)

    # 3. Подмена маркеров
    mapping_demo = {
        "договоры.номер":       "2024‑000001",
        "договоры.дата":        "23.12.2024",
        "вагоны.номер":         "12345",
        "вагоны.подразделение": "ПМС‑5, Свердловская ж/д",
        "список_работ(договоры.номер)": "LIST:Работа 1|Ремонт оси|Покраска",
        "сумма(договоры.номер)": "150000 руб."
    }
    result = os.path.join(here, "result.docx")
    replace_placeholders(template, result, mapping_demo)
