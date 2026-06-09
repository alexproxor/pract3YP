from __future__ import annotations

import csv
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
import re

import matplotlib.pyplot as plt
import networkx as nx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


ROOT = Path("/workspace")
DOCX_PATH = ROOT / "Прохоров.А.Ф_ПР3.docx"
GENERATED_DIR = ROOT / "generated"
GRAPH_PATH = GENERATED_DIR / "pr3_network_graph.png"
GANTT_PATH = GENERATED_DIR / "pr3_gantt_chart.png"
RESOURCE_HISTOGRAM_PATH = GENERATED_DIR / "pr3_resource_histogram.png"
TASKS_CSV_PATH = GENERATED_DIR / "pr3_projectlibre_tasks.csv"
RESOURCES_CSV_PATH = GENERATED_DIR / "pr3_projectlibre_resources.csv"
MARKDOWN_PATH = ROOT / "Прохоров.А.Ф_ПР3_дополнение.md"

APPENDIX_MARKER = "Установление зависимостей работ и типов связей между ними"
PROJECT_START = date(2026, 5, 6)
HOURS_PER_DAY = 8


TASKS = [
    {
        "code": "1",
        "name": "Инициация и планирование",
        "duration": "72",
        "assignee": "Руководитель проекта",
        "predecessor": "—",
        "link_type": "—",
    },
    {
        "code": "1.1",
        "name": "Формирование устава проекта",
        "duration": "16",
        "assignee": "Руководитель проекта",
        "predecessor": "—",
        "link_type": "—",
    },
    {
        "code": "1.2",
        "name": "Определение целей, границ и критериев успеха",
        "duration": "24",
        "assignee": "Руководитель проекта, Заказчик",
        "predecessor": "1.1",
        "link_type": "FS",
    },
    {
        "code": "1.3",
        "name": "План-график, бюджет, риски и коммуникации",
        "duration": "32",
        "assignee": "Руководитель проекта",
        "predecessor": "1.2",
        "link_type": "FS",
    },
    {
        "code": "2",
        "name": "Предпроектное обследование",
        "duration": "56",
        "assignee": "Бизнес-аналитик, Системный аналитик",
        "predecessor": "1",
        "link_type": "FS",
    },
    {
        "code": "2.1",
        "name": "Сбор и анализ требований (администрация, УК, жители)",
        "duration": "24",
        "assignee": "Бизнес-аналитик, Системный аналитик",
        "predecessor": "1.3",
        "link_type": "FS",
    },
    {
        "code": "2.2",
        "name": "Описание процессов «как есть» (AS-IS)",
        "duration": "16",
        "assignee": "Бизнес-аналитик",
        "predecessor": "2.1",
        "link_type": "FS",
    },
    {
        "code": "2.3",
        "name": "Формирование перечня функциональных и нефункциональных требований",
        "duration": "16",
        "assignee": "Системный аналитик",
        "predecessor": "2.2",
        "link_type": "FS",
    },
    {
        "code": "3",
        "name": "Проектирование ИС",
        "duration": "72",
        "assignee": "Системный архитектор, Системный аналитик",
        "predecessor": "2",
        "link_type": "FS",
    },
    {
        "code": "3.1",
        "name": "Архитектурное проектирование (модули, роли, интеграции)",
        "duration": "24",
        "assignee": "Системный архитектор",
        "predecessor": "2.3",
        "link_type": "FS",
    },
    {
        "code": "3.2",
        "name": "Проектирование БД и справочников",
        "duration": "16",
        "assignee": "Архитектор БД",
        "predecessor": "3.1",
        "link_type": "SS",
    },
    {
        "code": "3.3",
        "name": "Проектирование интерфейсов (web/моб. кабинет заявителя и оператора)",
        "duration": "16",
        "assignee": "UX/UI-дизайнер",
        "predecessor": "3.1",
        "link_type": "SS",
    },
    {
        "code": "3.4",
        "name": "Подготовка технического задания",
        "duration": "16",
        "assignee": "Системный аналитик, Руководитель проекта",
        "predecessor": "3.1, 3.2, 3.3",
        "link_type": "FS",
    },
    {
        "code": "4",
        "name": "Разработка",
        "duration": "80",
        "assignee": "Команда разработки",
        "predecessor": "3",
        "link_type": "FS",
    },
    {
        "code": "4.1",
        "name": "Модуль регистрации и маршрутизации заявок",
        "duration": "64",
        "assignee": "Backend-разработчик",
        "predecessor": "3.4",
        "link_type": "FS",
    },
    {
        "code": "4.2",
        "name": "Личный кабинет жителя (создание/отслеживание заявок)",
        "duration": "56",
        "assignee": "Frontend-разработчик",
        "predecessor": "4.1",
        "link_type": "SS",
    },
    {
        "code": "4.3",
        "name": "Кабинет диспетчера и исполнителя",
        "duration": "56",
        "assignee": "Backend-разработчик, Frontend-разработчик",
        "predecessor": "4.1",
        "link_type": "SS",
    },
    {
        "code": "4.4",
        "name": "Уведомления (email/SMS/push)",
        "duration": "24",
        "assignee": "Backend-разработчик",
        "predecessor": "4.1",
        "link_type": "FS",
    },
    {
        "code": "4.5",
        "name": "Отчеты и аналитика (сроки, SLA, проблемные зоны)",
        "duration": "40",
        "assignee": "BI-аналитик, Backend-разработчик",
        "predecessor": "4.2, 4.3",
        "link_type": "FF",
    },
    {
        "code": "5",
        "name": "Тестирование и контроль качества",
        "duration": "80",
        "assignee": "Команда QA",
        "predecessor": "4",
        "link_type": "SS",
    },
    {
        "code": "5.1",
        "name": "Подготовка тест-плана и сценариев",
        "duration": "16",
        "assignee": "Тестировщик",
        "predecessor": "4.1",
        "link_type": "SS",
    },
    {
        "code": "5.2",
        "name": "Функциональное и интеграционное тестирование",
        "duration": "40",
        "assignee": "Тестировщик",
        "predecessor": "4.2, 4.3, 4.4, 4.5",
        "link_type": "FS",
    },
    {
        "code": "5.3",
        "name": "Нагрузочное тестирование",
        "duration": "16",
        "assignee": "Инженер по нагрузочному тестированию",
        "predecessor": "5.2",
        "link_type": "SS",
    },
    {
        "code": "5.4",
        "name": "Исправление дефектов и регрессионная проверка",
        "duration": "24",
        "assignee": "Тестировщик, Разработчик",
        "predecessor": "5.2, 5.3",
        "link_type": "FS",
    },
    {
        "code": "6",
        "name": "Внедрение",
        "duration": "72",
        "assignee": "DevOps-инженер, Руководитель проекта",
        "predecessor": "5",
        "link_type": "SS",
    },
    {
        "code": "6.1",
        "name": "Подготовка инфраструктуры и развертывание",
        "duration": "16",
        "assignee": "DevOps-инженер",
        "predecessor": "5.2",
        "link_type": "SS",
    },
    {
        "code": "6.2",
        "name": "Миграция начальных данных и справочников",
        "duration": "16",
        "assignee": "Администратор БД, DevOps-инженер",
        "predecessor": "6.1",
        "link_type": "FS",
    },
    {
        "code": "6.3",
        "name": "Обучение пользователей (диспетчеры, исполнители, администраторы)",
        "duration": "16",
        "assignee": "Руководитель проекта, Бизнес-аналитик",
        "predecessor": "6.2",
        "link_type": "FS",
    },
    {
        "code": "6.4",
        "name": "Пилотная эксплуатация и корректировки",
        "duration": "24",
        "assignee": "Руководитель проекта, Команда проекта",
        "predecessor": "5.4, 6.2, 6.3",
        "link_type": "FS",
    },
    {
        "code": "7",
        "name": "Завершение проекта",
        "duration": "24",
        "assignee": "Руководитель проекта",
        "predecessor": "6",
        "link_type": "FS",
    },
    {
        "code": "7.1",
        "name": "Приемо-сдаточные испытания",
        "duration": "8",
        "assignee": "Заказчик, Руководитель проекта, Тестировщик",
        "predecessor": "6.4",
        "link_type": "FS",
    },
    {
        "code": "7.2",
        "name": "Ввод в промышленную эксплуатацию",
        "duration": "8",
        "assignee": "DevOps-инженер, Руководитель проекта",
        "predecessor": "7.1",
        "link_type": "FS",
    },
    {
        "code": "7.3",
        "name": "Закрытие проекта и передача в сопровождение",
        "duration": "8",
        "assignee": "Руководитель проекта, Служба сопровождения",
        "predecessor": "7.2",
        "link_type": "SF",
    },
]


LEAF_TASKS = [task for task in TASKS if "." in task["code"]]

DEPENDENCY_ROWS = [
    ("1.1", "—", "—"),
    ("1.2", "1.1", "—"),
    ("1.3", "1.2", "—"),
    ("2.1", "1.3", "—"),
    ("2.2", "2.1", "—"),
    ("2.3", "2.2", "—"),
    ("3.1", "2.3", "3.2, 3.3"),
    ("3.2", "2.3", "3.1, 3.3"),
    ("3.3", "2.3", "3.1, 3.2"),
    ("3.4", "3.1, 3.2, 3.3", "—"),
    ("4.1", "3.4", "4.2, 4.3, 5.1"),
    ("4.2", "3.4", "4.1, 4.3, 5.1"),
    ("4.3", "3.4", "4.1, 4.2, 5.1"),
    ("4.4", "4.1", "—"),
    ("4.5", "4.2, 4.3", "—"),
    ("5.1", "3.4", "4.1, 4.2, 4.3"),
    ("5.2", "4.2, 4.3, 4.4, 4.5", "5.3, 6.1"),
    ("5.3", "4.2, 4.3, 4.4, 4.5", "5.2, 6.1"),
    ("5.4", "5.2, 5.3", "—"),
    ("6.1", "4.2, 4.3, 4.4, 4.5", "5.2, 5.3"),
    ("6.2", "6.1", "—"),
    ("6.3", "6.2", "—"),
    ("6.4", "5.4, 6.2, 6.3", "—"),
    ("7.1", "6.4", "—"),
    ("7.2", "7.1", "—"),
    ("7.3", "7.2", "—"),
]

NON_SIMULTANEOUS_ROWS = [
    ("2.1", "2.2"),
    ("6.2", "6.4"),
    ("6.3", "6.4"),
]

CALENDAR_ROWS = [
    (
        "2.1",
        "1-я и 2-я недели проекта; интервью и сбор требований проводятся по заранее согласованному графику заказчика и управляющих компаний.",
    ),
    (
        "6.3",
        "За 2-3 рабочих дня до запуска пилотной эксплуатации, после подготовки данных и учетных записей.",
    ),
    (
        "7.1",
        "В течение 5 рабочих дней после завершения пилотной эксплуатации и устранения критических замечаний.",
    ),
]

EXPERT_PLACEHOLDERS = [
    "[Фамилия И.О. эксперта 1]",
    "[Фамилия И.О. эксперта 2]",
    "[Фамилия И.О. эксперта 3]",
]

RESOURCE_ASSIGNMENTS = {
    "1.1": {
        "labor": "- Руководитель проекта - 0,5 чел.\n- Бизнес-аналитик - 0,25 чел.",
        "material": "Канцелярские материалы, шаблон устава проекта.",
        "fixed": "8 тыс. руб.",
    },
    "1.2": {
        "labor": "- Руководитель проекта - 0,5 чел.\n- Представитель заказчика - 0,25 чел.",
        "material": "Материалы стратегической сессии, шаблоны целей и KPI.",
        "fixed": "10 тыс. руб.",
    },
    "1.3": {
        "labor": "- Руководитель проекта - 0,5 чел.\n- Финансовый аналитик - 0,25 чел.",
        "material": "Шаблоны календарного плана, реестр рисков.",
        "fixed": "12 тыс. руб.",
    },
    "2.1": {
        "labor": "- Бизнес-аналитик - 1 чел.\n- Системный аналитик - 0,5 чел.",
        "material": "Опросные листы, средства записи интервью, рабочие тетради обследования.",
        "fixed": "15 тыс. руб.",
    },
    "2.2": {
        "labor": "- Бизнес-аналитик - 1 чел.",
        "material": "Шаблоны BPMN/AS-IS диаграмм, канцелярские материалы.",
        "fixed": "8 тыс. руб.",
    },
    "2.3": {
        "labor": "- Системный аналитик - 1 чел.\n- Бизнес-аналитик - 0,5 чел.",
        "material": "Шаблон спецификации требований, реестр требований.",
        "fixed": "10 тыс. руб.",
    },
    "3.1": {
        "labor": "- Системный архитектор - 1 чел.\n- Системный аналитик - 0,5 чел.",
        "material": "Средства моделирования архитектуры, репозиторий проектных решений.",
        "fixed": "20 тыс. руб.",
    },
    "3.2": {
        "labor": "- Архитектор БД - 1 чел.",
        "material": "Шаблоны ER-диаграмм, модель данных, справочники предметной области.",
        "fixed": "12 тыс. руб.",
    },
    "3.3": {
        "labor": "- UX/UI-дизайнер - 1 чел.",
        "material": "Дизайн-система, библиотека компонентов, прототипы экранов.",
        "fixed": "18 тыс. руб.",
    },
    "3.4": {
        "labor": "- Системный аналитик - 1 чел.\n- Руководитель проекта - 0,25 чел.",
        "material": "Утвержденный шаблон ТЗ, комплект исходных требований.",
        "fixed": "8 тыс. руб.",
    },
    "4.1": {
        "labor": "- Backend-разработчик - 1 чел.\n- Системный аналитик - 0,25 чел.",
        "material": "Среда разработки, тестовый сервер, репозиторий исходного кода.",
        "fixed": "35 тыс. руб.",
    },
    "4.2": {
        "labor": "- Frontend-разработчик - 1 чел.",
        "material": "UI-kit, тестовое мобильное устройство, макеты интерфейсов.",
        "fixed": "28 тыс. руб.",
    },
    "4.3": {
        "labor": "- Backend-разработчик - 0,5 чел.\n- Frontend-разработчик - 0,5 чел.",
        "material": "Тестовые учетные записи, сервер приложений, API-спецификация.",
        "fixed": "30 тыс. руб.",
    },
    "4.4": {
        "labor": "- Backend-разработчик - 0,5 чел.\n- Инженер по интеграциям - 0,5 чел.",
        "material": "Тестовый SMS-шлюз, учетная запись почтового сервиса, шаблоны уведомлений.",
        "fixed": "25 тыс. руб.",
    },
    "4.5": {
        "labor": "- BI-аналитик - 0,5 чел.\n- Backend-разработчик - 0,5 чел.",
        "material": "Тестовый набор данных, шаблоны отчетов и дашбордов.",
        "fixed": "20 тыс. руб.",
    },
    "5.1": {
        "labor": "- Тестировщик - 1 чел.",
        "material": "Шаблоны тест-кейсов, чек-листы приемки.",
        "fixed": "8 тыс. руб.",
    },
    "5.2": {
        "labor": "- Тестировщик - 1 чел.\n- Системный аналитик - 0,25 чел.",
        "material": "Тестовый стенд, баг-трекер, комплект тестовых сценариев.",
        "fixed": "18 тыс. руб.",
    },
    "5.3": {
        "labor": "- Инженер по нагрузочному тестированию - 0,5 чел.\n- DevOps-инженер - 0,25 чел.",
        "material": "Инструмент нагрузочного тестирования, выделенный тестовый контур.",
        "fixed": "16 тыс. руб.",
    },
    "5.4": {
        "labor": "- Тестировщик - 0,5 чел.\n- Разработчик - 0,5 чел.",
        "material": "Журнал дефектов, отчеты о тестировании, регрессионный набор.",
        "fixed": "12 тыс. руб.",
    },
    "6.1": {
        "labor": "- DevOps-инженер - 1 чел.",
        "material": "Облачная виртуальная машина, скрипты развертывания, сертификаты доступа.",
        "fixed": "32 тыс. руб.",
    },
    "6.2": {
        "labor": "- Администратор БД - 0,5 чел.\n- DevOps-инженер - 0,5 чел.",
        "material": "Скрипты миграции, резервные копии, хранилище данных.",
        "fixed": "22 тыс. руб.",
    },
    "6.3": {
        "labor": "- Руководитель проекта - 0,25 чел.\n- Бизнес-аналитик - 0,5 чел.",
        "material": "Презентация обучения, инструкции пользователя, методические материалы.",
        "fixed": "14 тыс. руб.",
    },
    "6.4": {
        "labor": "- Руководитель проекта - 0,5 чел.\n- Специалист поддержки - 0,5 чел.\n- Разработчик - 0,25 чел.",
        "material": "Пилотный контур, журнал замечаний, канал оперативной поддержки.",
        "fixed": "24 тыс. руб.",
    },
    "7.1": {
        "labor": "- Представитель заказчика - 0,25 чел.\n- Руководитель проекта - 0,25 чел.\n- Тестировщик - 0,25 чел.",
        "material": "Протоколы испытаний, приемочные листы, комплект отчетных форм.",
        "fixed": "8 тыс. руб.",
    },
    "7.2": {
        "labor": "- DevOps-инженер - 0,5 чел.\n- Руководитель проекта - 0,25 чел.",
        "material": "Релизный пакет, чек-лист публикации, эксплуатационная инструкция.",
        "fixed": "10 тыс. руб.",
    },
    "7.3": {
        "labor": "- Руководитель проекта - 0,25 чел.\n- Специалист сопровождения - 0,25 чел.",
        "material": "Акт передачи, регламент сопровождения, архив проектной документации.",
        "fixed": "8 тыс. руб.",
    },
}

LINK_TYPE_EXPLANATION = {
    "FS": "окончание-начало",
    "SS": "начало-начало",
    "FF": "окончание-окончание",
    "SF": "начало-окончание",
}


def set_cell_text(cell, text: str) -> None:
    cell.text = text


def remove_generated_appendix(document: Document) -> None:
    body = document._body._element
    children = list(body)
    marker_index = None
    for idx, child in enumerate(children):
        if child.tag.endswith("}p"):
            text = "".join(node.text for node in child.iter() if node.text)
            if APPENDIX_MARKER in text:
                marker_index = idx
                break
    if marker_index is None:
        return
    for child in children[marker_index:]:
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)


def update_wbs_table(document: Document) -> None:
    task_by_code = {task["code"]: task for task in TASKS}
    table = document.tables[0]
    for row in table.rows[1:]:
        code = row.cells[0].text.strip()
        if code not in task_by_code:
            continue
        task = task_by_code[code]
        set_cell_text(row.cells[1], task["name"])
        set_cell_text(row.cells[2], task["duration"])
        set_cell_text(row.cells[3], task["assignee"])
        set_cell_text(row.cells[4], task["predecessor"])
        set_cell_text(row.cells[5], task["link_type"])


def add_table_title(document: Document, text: str) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(text).bold = True


def add_figure_title(document: Document, text: str) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(text).bold = True


def add_heading_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(13)


def add_subheading_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = True


def get_expert_estimates(duration: int) -> tuple[int, int, int, int]:
    if duration <= 8:
        return duration, duration, duration, duration

    delta = 2 if duration <= 24 else 4
    estimate_1 = max(8, duration - delta)
    estimate_2 = duration
    estimate_3 = duration + (duration - estimate_1)
    average = round((estimate_1 + estimate_2 + estimate_3) / 3)
    return estimate_1, estimate_2, estimate_3, average


def parse_resource_units(labor: str) -> list[tuple[str, float]]:
    resources = []
    for line in labor.splitlines():
        match = re.search(r"-\s*(.*?)\s*-\s*([\d,]+)\s*чел", line)
        if not match:
            continue
        name = match.group(1).strip()
        units = float(match.group(2).replace(",", "."))
        resources.append((name, units))
    return resources


def format_date(value: date) -> str:
    return value.strftime("%d.%m.%Y")


def is_workday(value: date) -> bool:
    return value.weekday() < 5


def next_workday(value: date) -> date:
    while not is_workday(value):
        value += timedelta(days=1)
    return value


def add_workdays(start: date, days: int) -> date:
    current = next_workday(start)
    remaining = days
    while remaining > 0:
        current += timedelta(days=1)
        if is_workday(current):
            remaining -= 1
    return current


def subtract_workdays(finish: date, days: int) -> date:
    current = next_workday(finish)
    remaining = days
    while remaining > 0:
        current -= timedelta(days=1)
        if is_workday(current):
            remaining -= 1
    return current


def workdays_between(start: date, finish: date) -> list[date]:
    current = next_workday(start)
    days = []
    while current <= finish:
        if is_workday(current):
            days.append(current)
        current += timedelta(days=1)
    return days


def get_leaf_task(task_code: str) -> dict[str, str]:
    for task in LEAF_TASKS:
        if task["code"] == task_code:
            return task
    raise KeyError(task_code)


def calculate_schedule() -> dict[str, dict[str, object]]:
    schedule: dict[str, dict[str, object]] = {}
    for task in LEAF_TASKS:
        duration_hours = int(task["duration"])
        duration_days = max(1, (duration_hours + HOURS_PER_DAY - 1) // HOURS_PER_DAY)
        predecessor_codes = [
            item.strip()
            for item in task["predecessor"].split(",")
            if item.strip() and item.strip() != "—" and "." in item
        ]

        start_candidates = [PROJECT_START]
        for predecessor_code in predecessor_codes:
            predecessor = schedule[predecessor_code]
            predecessor_task = get_leaf_task(predecessor_code)
            link_type = task["link_type"]
            if link_type == "SS":
                start_candidates.append(predecessor["start"])
            elif link_type == "FF":
                start_candidates.append(subtract_workdays(predecessor["finish"], duration_days - 1))
            elif link_type == "SF":
                start_candidates.append(next_workday(predecessor["start"]))
            else:
                start_candidates.append(add_workdays(predecessor["finish"], 1))

            # When several predecessors are present, ProjectLibre applies the same relation
            # shown in the WBS table to every listed predecessor for this учебная модель.
            if predecessor_task["link_type"] == "—":
                continue

        start = max(start_candidates)
        finish = add_workdays(start, duration_days - 1)
        schedule[task["code"]] = {
            "start": start,
            "finish": finish,
            "duration_days": duration_days,
            "duration_hours": duration_hours,
        }
    return schedule


def get_project_dates(schedule: dict[str, dict[str, object]]) -> tuple[date, date]:
    starts = [item["start"] for item in schedule.values()]
    finishes = [item["finish"] for item in schedule.values()]
    return min(starts), max(finishes)


def get_task_resources(task_code: str) -> list[tuple[str, float]]:
    return parse_resource_units(RESOURCE_ASSIGNMENTS[task_code]["labor"])


def build_daily_resource_usage(
    schedule: dict[str, dict[str, object]],
) -> tuple[dict[date, dict[str, float]], dict[str, dict[str, object]]]:
    daily_usage: dict[date, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    resource_summary: dict[str, dict[str, object]] = defaultdict(
        lambda: {"tasks": set(), "total_hours": 0.0, "max_units": 0.0}
    )

    for task in LEAF_TASKS:
        task_schedule = schedule[task["code"]]
        workdays = workdays_between(task_schedule["start"], task_schedule["finish"])
        for resource_name, units in get_task_resources(task["code"]):
            resource_summary[resource_name]["tasks"].add(task["code"])
            resource_summary[resource_name]["total_hours"] += int(task["duration"]) * units
            for current_day in workdays:
                daily_usage[current_day][resource_name] += units

    for day_usage in daily_usage.values():
        for resource_name, units in day_usage.items():
            resource_summary[resource_name]["max_units"] = max(
                resource_summary[resource_name]["max_units"], units
            )

    return daily_usage, resource_summary


def get_project_details(schedule: dict[str, dict[str, object]]) -> dict[str, str]:
    project_start, project_finish = get_project_dates(schedule)
    total_work_hours = sum(int(task["duration"]) for task in LEAF_TASKS)
    total_fixed_cost = sum(
        int(re.search(r"\d+", RESOURCE_ASSIGNMENTS[task["code"]]["fixed"]).group())
        for task in LEAF_TASKS
    )
    return {
        "Название проекта": "ИС «ЖКХ-Заявки»",
        "Дата начала": format_date(project_start),
        "Дата окончания": format_date(project_finish),
        "Календарь": "Стандартный: 5-дневная рабочая неделя, 8 часов в день",
        "Количество работ нижнего уровня": str(len(LEAF_TASKS)),
        "Плановая трудоемкость": f"{total_work_hours} ч",
        "Фиксированные затраты": f"{total_fixed_cost} тыс. руб.",
        "Критерий расчета": "ранее начало по зависимостям FS, SS, FF, SF",
    }


def get_who_does_what_rows() -> list[tuple[str, str, str, str]]:
    rows = []
    for task in LEAF_TASKS:
        resources = "; ".join(
            f"{name} ({str(units).replace('.', ',')} чел.)" for name, units in get_task_resources(task["code"])
        )
        rows.append((task["code"], task["name"], resources, RESOURCE_ASSIGNMENTS[task["code"]]["fixed"]))
    return rows


def write_projectlibre_csvs(
    schedule: dict[str, dict[str, object]], resource_summary: dict[str, dict[str, object]]
) -> None:
    GENERATED_DIR.mkdir(exist_ok=True)
    task_lines = [
        "ID;WBS;Name;DurationHours;Start;Finish;Predecessors;Resources;FixedCostThousandRub"
    ]
    for index, task in enumerate(LEAF_TASKS, start=1):
        task_schedule = schedule[task["code"]]
        resources = ", ".join(name for name, _ in get_task_resources(task["code"]))
        fixed_cost = re.search(r"\d+", RESOURCE_ASSIGNMENTS[task["code"]]["fixed"]).group()
        task_lines.append(
            ";".join(
                [
                    str(index),
                    task["code"],
                    task["name"],
                    task["duration"],
                    format_date(task_schedule["start"]),
                    format_date(task_schedule["finish"]),
                    task["predecessor"],
                    resources,
                    fixed_cost,
                ]
            )
        )
    TASKS_CSV_PATH.write_text("\n".join(task_lines) + "\n", encoding="utf-8")

    resource_lines = ["Name;Type;MaxUnits;TotalWorkHours;AssignedTasks"]
    for resource_name, summary in sorted(resource_summary.items()):
        assigned_tasks = ", ".join(sorted(summary["tasks"], key=lambda code: [int(part) for part in code.split(".")]))
        resource_lines.append(
            ";".join(
                [
                    resource_name,
                    "Work",
                    str(round(summary["max_units"], 2)).replace(".", ","),
                    str(round(summary["total_hours"], 1)).replace(".", ","),
                    assigned_tasks,
                ]
            )
        )
    RESOURCES_CSV_PATH.write_text("\n".join(resource_lines) + "\n", encoding="utf-8")


def create_gantt_chart(schedule: dict[str, dict[str, object]]) -> None:
    import matplotlib.dates as mdates

    GENERATED_DIR.mkdir(exist_ok=True)
    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(16, 11))

    colors = {
        "1": "#8dd3c7",
        "2": "#ffffb3",
        "3": "#bebada",
        "4": "#fb8072",
        "5": "#80b1d3",
        "6": "#fdb462",
        "7": "#b3de69",
    }
    ordered_tasks = list(reversed(LEAF_TASKS))
    y_positions = range(len(ordered_tasks))
    for y, task in zip(y_positions, ordered_tasks):
        task_schedule = schedule[task["code"]]
        start_num = mdates.date2num(task_schedule["start"])
        finish_num = mdates.date2num(task_schedule["finish"] + timedelta(days=1))
        width = finish_num - start_num
        phase = task["code"].split(".")[0]
        ax.barh(y, width, left=start_num, height=0.58, color=colors[phase], edgecolor="#333333")
        ax.text(start_num + 0.1, y, task["code"], va="center", ha="left", fontsize=8)

    ax.set_yticks(list(y_positions))
    ax.set_yticklabels([f"{task['code']} {task['name']}" for task in ordered_tasks], fontsize=8)
    ax.xaxis_date()
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m"))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO, interval=1))
    ax.grid(axis="x", linestyle=":", alpha=0.5)
    ax.set_title("Диаграмма Ганта проекта ИС «ЖКХ-Заявки»", fontsize=15)
    ax.set_xlabel("Календарные даты")
    plt.tight_layout()
    fig.savefig(GANTT_PATH, dpi=220, bbox_inches="tight")
    plt.close(fig)


def create_resource_histogram(daily_usage: dict[date, dict[str, float]]) -> None:
    import matplotlib.dates as mdates

    GENERATED_DIR.mkdir(exist_ok=True)
    key_resources = [
        "Руководитель проекта",
        "Бизнес-аналитик",
        "Системный аналитик",
        "Backend-разработчик",
        "Frontend-разработчик",
        "Тестировщик",
        "DevOps-инженер",
    ]
    days = sorted(daily_usage)
    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(16, 8))
    bottom = [0.0] * len(days)
    palette = ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc948", "#b07aa1"]

    for resource_name, color in zip(key_resources, palette):
        values = [daily_usage[day].get(resource_name, 0.0) for day in days]
        ax.bar(days, values, bottom=bottom, width=0.8, label=resource_name, color=color)
        bottom = [current + value for current, value in zip(bottom, values)]

    other_values = [
        sum(units for resource, units in daily_usage[day].items() if resource not in key_resources)
        for day in days
    ]
    ax.bar(days, other_values, bottom=bottom, width=0.8, label="Прочие ресурсы", color="#9c755f")

    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m"))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO, interval=1))
    ax.set_ylabel("Загрузка, чел.")
    ax.set_title("Гистограмма загрузки ресурсов по рабочим дням")
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(loc="upper left", ncol=2, fontsize=9)
    plt.tight_layout()
    fig.savefig(RESOURCE_HISTOGRAM_PATH, dpi=220, bbox_inches="tight")
    plt.close(fig)


def create_projectlibre_outputs() -> tuple[
    dict[str, dict[str, object]],
    dict[date, dict[str, float]],
    dict[str, dict[str, object]],
]:
    schedule = calculate_schedule()
    daily_usage, resource_summary = build_daily_resource_usage(schedule)
    create_gantt_chart(schedule)
    create_resource_histogram(daily_usage)
    write_projectlibre_csvs(schedule, resource_summary)
    return schedule, daily_usage, resource_summary


def add_projectlibre_sections(
    document: Document,
    schedule: dict[str, dict[str, object]],
    resource_summary: dict[str, dict[str, object]],
) -> None:
    add_heading_paragraph(document, "Расчет плановых характеристик проекта в ProjectLibre")
    document.add_paragraph(
        "Плановые характеристики рассчитаны для стандартного рабочего календаря: 5 рабочих дней в неделю, "
        "8 часов в день. В ProjectLibre заносятся работы нижнего уровня WBS, длительности, связи, исполнители "
        "и фиксированные затраты; подготовленные CSV-файлы можно использовать как основу для переноса данных."
    )

    document.add_picture(str(GANTT_PATH), width=Inches(7.0))
    add_figure_title(document, "Рисунок 2 - Диаграмма Ганта проекта")

    add_table_title(document, "Таблица 4.6")
    resource_table = document.add_table(rows=1, cols=4)
    resource_table.style = "Table Grid"
    headers = resource_table.rows[0].cells
    headers[0].text = "Ресурс"
    headers[1].text = "Тип"
    headers[2].text = "Макс. загрузка, чел."
    headers[3].text = "Плановая трудоемкость, ч"
    for resource_name, summary in sorted(resource_summary.items()):
        row = resource_table.add_row().cells
        row[0].text = resource_name
        row[1].text = "Трудовой"
        row[2].text = str(round(summary["max_units"], 2)).replace(".", ",")
        row[3].text = str(round(summary["total_hours"], 1)).replace(".", ",")

    add_table_title(document, "Таблица 4.7")
    details_table = document.add_table(rows=1, cols=2)
    details_table.style = "Table Grid"
    details_headers = details_table.rows[0].cells
    details_headers[0].text = "Параметр"
    details_headers[1].text = "Значение"
    for key, value in get_project_details(schedule).items():
        row = details_table.add_row().cells
        row[0].text = key
        row[1].text = value

    add_table_title(document, "Таблица 4.8")
    who_table = document.add_table(rows=1, cols=4)
    who_table.style = "Table Grid"
    who_headers = who_table.rows[0].cells
    who_headers[0].text = "Работа"
    who_headers[1].text = "Наименование"
    who_headers[2].text = "Назначенные ресурсы"
    who_headers[3].text = "Фиксированные затраты"
    for code, name, resources, fixed_cost in get_who_does_what_rows():
        row = who_table.add_row().cells
        row[0].text = code
        row[1].text = name
        row[2].text = resources
        row[3].text = fixed_cost

    document.add_paragraph()
    document.add_picture(str(RESOURCE_HISTOGRAM_PATH), width=Inches(7.0))
    add_figure_title(document, "Рисунок 3 - Гистограмма загрузки ресурсов")


def add_resource_sections(document: Document) -> None:
    add_heading_paragraph(document, "Назначение ресурсов на выполняемые работы")
    document.add_paragraph(
        "На данном этапе выполняется назначение временных, трудовых, материальных и финансовых ресурсов "
        "на работы проекта ИС «ЖКХ-Заявки»."
    )

    add_subheading_paragraph(document, "4.3.1. Временные ресурсы")
    document.add_paragraph(
        "Для оценки продолжительности работ использован метод экспертной оценки по трем точкам. "
        "В таблице 4.4 приведены оценки трех экспертов и средняя оценка по каждой работе."
    )
    document.add_paragraph("Эксперты для последующего заполнения отчета:")
    for expert in EXPERT_PLACEHOLDERS:
        document.add_paragraph(expert, style=None)

    add_table_title(document, "Таблица 4.4")
    duration_table = document.add_table(rows=1, cols=5)
    duration_table.style = "Table Grid"
    duration_headers = duration_table.rows[0].cells
    duration_headers[0].text = "Номер работы"
    duration_headers[1].text = "Оценка 1"
    duration_headers[2].text = "Оценка 2"
    duration_headers[3].text = "Оценка 3"
    duration_headers[4].text = "Средняя оценка"

    for task in LEAF_TASKS:
        estimate_1, estimate_2, estimate_3, average = get_expert_estimates(int(task["duration"]))
        row = duration_table.add_row().cells
        row[0].text = task["code"]
        row[1].text = str(estimate_1)
        row[2].text = str(estimate_2)
        row[3].text = str(estimate_3)
        row[4].text = str(average)

    add_subheading_paragraph(document, "4.3.2. Трудовые, материальные и финансовые ресурсы")
    document.add_paragraph(
        "Для выполнения работ проекта определены трудовые ресурсы команды, используемые материалы "
        "и фиксированные затраты. Итоговые назначения представлены в таблице 4.5."
    )

    add_table_title(document, "Таблица 4.5")
    resource_table = document.add_table(rows=1, cols=4)
    resource_table.style = "Table Grid"
    resource_headers = resource_table.rows[0].cells
    resource_headers[0].text = "Номер работы"
    resource_headers[1].text = "Трудовые ресурсы"
    resource_headers[2].text = "Материальные ресурсы"
    resource_headers[3].text = "Фиксированные затраты"

    for task in LEAF_TASKS:
        assignment = RESOURCE_ASSIGNMENTS[task["code"]]
        row = resource_table.add_row().cells
        row[0].text = task["code"]
        row[1].text = assignment["labor"]
        row[2].text = assignment["material"]
        row[3].text = assignment["fixed"]

    document.add_paragraph()
    add_subheading_paragraph(document, "4.3.3. Плановые характеристики в ProjectLibre")
    document.add_paragraph(
        "После назначения ресурсов модель проекта переносится в ProjectLibre. "
        "В отчете ниже приведены обязательные выгрузки: диаграмма Ганта, список ресурсов, "
        "Project Details, Who Does What и гистограмма загрузки ресурсов."
    )


def add_appendix(
    document: Document,
    schedule: dict[str, dict[str, object]],
    resource_summary: dict[str, dict[str, object]],
) -> None:
    document.add_page_break()
    add_heading_paragraph(document, APPENDIX_MARKER)

    document.add_paragraph(
        "Для проекта ИС «ЖКХ-Заявки» зависимости работ сформированы с учетом логики выполнения этапов, "
        "возможности параллельного старта отдельных работ и ограничений по использованию общих ресурсов."
    )
    document.add_paragraph(
        "Перечень обязательных работ-предшественников и одновременно начинаемых работ представлен в таблице 2."
    )

    add_table_title(document, "Таблица 2")
    dep_table = document.add_table(rows=1, cols=3)
    dep_table.style = "Table Grid"
    dep_headers = dep_table.rows[0].cells
    dep_headers[0].text = "Рассматриваемая работа"
    dep_headers[1].text = "Перечень обязательных работ-предшественников"
    dep_headers[2].text = "Перечень одновременно начинаемых работ"
    for code, predecessors, simultaneous in DEPENDENCY_ROWS:
        row = dep_table.add_row().cells
        row[0].text = code
        row[1].text = predecessors
        row[2].text = simultaneous

    document.add_paragraph(
        "Работы, которые не могут выполняться одновременно по организационным и ресурсным ограничениям, "
        "представлены в таблице 3."
    )
    add_table_title(document, "Таблица 3")
    restriction_table = document.add_table(rows=1, cols=2)
    restriction_table.style = "Table Grid"
    restriction_headers = restriction_table.rows[0].cells
    restriction_headers[0].text = "Рассматриваемая работа"
    restriction_headers[1].text = "Перечень работ, которые не могут выполняться одновременно"
    for code, restricted in NON_SIMULTANEOUS_ROWS:
        row = restriction_table.add_row().cells
        row[0].text = code
        row[1].text = restricted

    document.add_paragraph()
    add_heading_paragraph(document, "Построение сетевой модели")
    document.add_paragraph(
        "Сетевой график вида AoN позволяет связать все работы проекта в единую модель, "
        "проверить корректность зависимостей и использовать полученную структуру для последующего календарного планирования."
    )
    document.add_paragraph(
        "На рисунке 1 представлены работы нижнего уровня WBS. Типы связей показаны цветом и стилем линий: "
        "FS — сплошная черная, SS — синяя пунктирная, FF — оранжевая штрихпунктирная, SF — красная точечная."
    )
    document.add_paragraph(
        "Работы с календарной привязкой приведены в таблице 4. Они задают дополнительные ограничения при назначении дат в календарном плане."
    )

    add_table_title(document, "Таблица 4")
    calendar_table = document.add_table(rows=1, cols=2)
    calendar_table.style = "Table Grid"
    calendar_headers = calendar_table.rows[0].cells
    calendar_headers[0].text = "Рассматриваемая работа"
    calendar_headers[1].text = "Календарная привязка (период возможного выполнения, условие начала и др.)"
    for code, anchor in CALENDAR_ROWS:
        row = calendar_table.add_row().cells
        row[0].text = code
        row[1].text = anchor

    document.add_paragraph()
    document.add_picture(str(GRAPH_PATH), width=Inches(7.0))
    add_figure_title(document, "Рисунок 1 - Сетевая модель проекта ИС «ЖКХ-Заявки»")
    document.add_paragraph()
    add_resource_sections(document)
    document.add_paragraph()
    add_projectlibre_sections(document, schedule, resource_summary)


def create_graph() -> None:
    GENERATED_DIR.mkdir(exist_ok=True)

    graph = nx.DiGraph()
    graph.add_node("START")
    graph.add_node("FINISH")

    for task in LEAF_TASKS:
        graph.add_node(task["code"])

    edge_types: dict[tuple[str, str], str] = {}
    predecessors_map: dict[str, list[str]] = {}
    successor_count: defaultdict[str, int] = defaultdict(int)

    for task in LEAF_TASKS:
        predecessors = [item.strip() for item in task["predecessor"].split(",") if item.strip() and item.strip() != "—"]
        predecessors_map[task["code"]] = predecessors
        if not predecessors:
            graph.add_edge("START", task["code"])
            edge_types[("START", task["code"])] = "START"
        for predecessor in predecessors:
            if "." not in predecessor:
                continue
            graph.add_edge(predecessor, task["code"])
            edge_types[(predecessor, task["code"])] = task["link_type"]
            successor_count[predecessor] += 1

    for task in LEAF_TASKS:
        if successor_count[task["code"]] == 0:
            graph.add_edge(task["code"], "FINISH")
            edge_types[(task["code"], "FINISH")] = "END"

    layers = [
        ["START"],
        ["1.1"],
        ["1.2"],
        ["1.3"],
        ["2.1"],
        ["2.2"],
        ["2.3"],
        ["3.1", "3.2", "3.3"],
        ["3.4"],
        ["4.1", "4.2", "4.3", "5.1"],
        ["4.4", "4.5", "5.2", "5.3", "6.1"],
        ["5.4", "6.2"],
        ["6.3", "6.4"],
        ["7.1"],
        ["7.2", "7.3"],
        ["FINISH"],
    ]
    positions = {}
    for x, layer in enumerate(layers):
        count = len(layer)
        for idx, node in enumerate(layer):
            y = (count - 1) / 2 - idx
            positions[node] = (x * 2.3, y * 2.0)

    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(18, 10))
    ax.set_axis_off()

    node_colors = []
    for node in graph.nodes:
        if node in {"START", "FINISH"}:
            node_colors.append("#e8e8e8")
        else:
            node_colors.append("#ffffff")

    nx.draw_networkx_nodes(
        graph,
        positions,
        node_size=1700,
        node_color=node_colors,
        edgecolors="#333333",
        linewidths=1.5,
        ax=ax,
    )
    nx.draw_networkx_labels(graph, positions, font_size=10, ax=ax)

    styles = {
        "FS": {"edge_color": "#222222", "style": "solid", "width": 1.8},
        "SS": {"edge_color": "#1f77b4", "style": "dashed", "width": 2.0},
        "FF": {"edge_color": "#ff7f0e", "style": "dashdot", "width": 2.0},
        "SF": {"edge_color": "#d62728", "style": "dotted", "width": 2.4},
        "START": {"edge_color": "#777777", "style": "solid", "width": 1.5},
        "END": {"edge_color": "#777777", "style": "solid", "width": 1.5},
    }

    for edge_type, style in styles.items():
        edges = [edge for edge, current_type in edge_types.items() if current_type == edge_type]
        if not edges:
            continue
        nx.draw_networkx_edges(
            graph,
            positions,
            edgelist=edges,
            arrows=True,
            arrowstyle="-|>",
            arrowsize=18,
            min_source_margin=18,
            min_target_margin=18,
            connectionstyle="arc3,rad=0.05" if edge_type in {"SS", "FF", "SF"} else "arc3,rad=0.0",
            ax=ax,
            **style,
        )

    from matplotlib.lines import Line2D

    legend_items = [
        Line2D([0], [0], color="#222222", lw=1.8, linestyle="solid", label="FS"),
        Line2D([0], [0], color="#1f77b4", lw=2.0, linestyle="dashed", label="SS"),
        Line2D([0], [0], color="#ff7f0e", lw=2.0, linestyle="dashdot", label="FF"),
        Line2D([0], [0], color="#d62728", lw=2.4, linestyle="dotted", label="SF"),
    ]
    ax.legend(handles=legend_items, loc="lower center", ncol=4, frameon=False)
    ax.set_title("Сетевая модель проекта ИС «ЖКХ-Заявки»", fontsize=16, pad=18)
    plt.tight_layout()
    fig.savefig(GRAPH_PATH, dpi=220, bbox_inches="tight")
    plt.close(fig)


def write_markdown(
    schedule: dict[str, dict[str, object]],
    resource_summary: dict[str, dict[str, object]],
) -> None:
    task_name_by_code = {task["code"]: task["name"] for task in TASKS}
    lines = [
        "# Материалы по практической работе 3",
        "",
        "Документ подготовлен для проекта ИС «ЖКХ-Заявки» и содержит:",
        "- завершенную таблицу WBS с длительностями, исполнителями и зависимостями;",
        "- таблицу обязательных предшественников и одновременно начинаемых работ;",
        "- таблицу работ, которые не должны выполняться одновременно;",
        "- таблицу календарных привязок;",
        f"- рисунок сетевой модели: `{GRAPH_PATH.relative_to(ROOT)}`.",
        "",
        "## 1. Завершенная таблица WBS",
        "",
        "| Код WBS | Наименование работы | Длительность, ч | Исполнитель | Предшественник | Тип связи |",
        "|---|---|---:|---|---|---|",
    ]
    for task in TASKS:
        lines.append(
            f"| {task['code']} | {task['name']} | {task['duration']} | {task['assignee']} | {task['predecessor']} | {task['link_type']} |"
        )

    lines.extend(
        [
            "",
            "## 2. Обязательные предшественники и одновременно начинаемые работы",
            "",
            "| Рассматриваемая работа | Перечень обязательных работ-предшественников | Перечень одновременно начинаемых работ |",
            "|---|---|---|",
        ]
    )
    for code, predecessors, simultaneous in DEPENDENCY_ROWS:
        lines.append(f"| {code} | {predecessors} | {simultaneous} |")

    lines.extend(
        [
            "",
            "## 3. Работы, которые не должны выполняться одновременно",
            "",
            "| Рассматриваемая работа | Перечень работ, которые не могут выполняться одновременно |",
            "|---|---|",
        ]
    )
    for code, restricted in NON_SIMULTANEOUS_ROWS:
        lines.append(f"| {code} | {restricted} |")

    lines.extend(
        [
            "",
            "## 4. Работы с календарной привязкой",
            "",
            "| Рассматриваемая работа | Календарная привязка |",
            "|---|---|",
        ]
    )
    for code, calendar in CALENDAR_ROWS:
        lines.append(f"| {code} | {calendar} |")

    lines.extend(
        [
            "",
            "## 5. Пояснения к сетевой модели",
            "",
            "Для сетевого графика использованы работы нижнего уровня WBS.",
            "Обозначения типов связей:",
        ]
    )
    for code, explanation in LINK_TYPE_EXPLANATION.items():
        lines.append(f"- **{code}** — {explanation};")

    lines.extend(
        [
            "",
            "Дополнительные замечания:",
            "- для работы **6.4** уточнен набор предшественников: добавлены **6.2** и **6.3**, чтобы пилотная эксплуатация начиналась только после миграции данных и обучения пользователей;",
            "- связь **7.2 -> 7.3** оставлена в формате **SF**, чтобы сохранить в модели пример редкого типа зависимости «начало-окончание».",
            "",
            "## 6. Расшифровка кодов для рисунка",
            "",
            "| Код | Наименование |",
            "|---|---|",
        ]
    )
    for task in LEAF_TASKS:
        lines.append(f"| {task['code']} | {task_name_by_code[task['code']]} |")

    lines.extend(
        [
            "",
            "## 7. Назначение ресурсов на выполняемые работы",
            "",
            "### 7.1. Временные ресурсы",
            "",
            "Для оценки продолжительности работ использован метод экспертной оценки по трем точкам.",
            "Эксперты для заполнения отчета:",
        ]
    )
    for expert in EXPERT_PLACEHOLDERS:
        lines.append(f"- {expert}")

    lines.extend(
        [
            "",
            "| Номер работы | Оценка 1 | Оценка 2 | Оценка 3 | Средняя оценка |",
            "|---|---:|---:|---:|---:|",
        ]
    )
    for task in LEAF_TASKS:
        estimate_1, estimate_2, estimate_3, average = get_expert_estimates(int(task["duration"]))
        lines.append(f"| {task['code']} | {estimate_1} | {estimate_2} | {estimate_3} | {average} |")

    lines.extend(
        [
            "",
            "### 7.2. Трудовые, материальные и финансовые ресурсы",
            "",
            "| Номер работы | Трудовые ресурсы | Материальные ресурсы | Фиксированные затраты |",
            "|---|---|---|---|",
        ]
    )
    for task in LEAF_TASKS:
        assignment = RESOURCE_ASSIGNMENTS[task["code"]]
        labor = assignment["labor"].replace("\n", "<br>")
        lines.append(
            f"| {task['code']} | {labor} | {assignment['material']} | {assignment['fixed']} |"
        )

    lines.extend(
        [
            "",
            "## 8. Расчет плановых характеристик проекта в ProjectLibre",
            "",
            "Для расчета принят стандартный календарь: 5 рабочих дней в неделю, 8 часов в день.",
            f"Диаграмма Ганта сохранена в файле `{GANTT_PATH.relative_to(ROOT)}`.",
            f"Гистограмма загрузки ресурсов сохранена в файле `{RESOURCE_HISTOGRAM_PATH.relative_to(ROOT)}`.",
            f"CSV для переноса задач в ProjectLibre: `{TASKS_CSV_PATH.relative_to(ROOT)}`.",
            f"CSV со списком ресурсов: `{RESOURCES_CSV_PATH.relative_to(ROOT)}`.",
            "",
            "### 8.1. Project Details",
            "",
            "| Параметр | Значение |",
            "|---|---|",
        ]
    )
    for key, value in get_project_details(schedule).items():
        lines.append(f"| {key} | {value} |")

    lines.extend(
        [
            "",
            "### 8.2. Список ресурсов",
            "",
            "| Ресурс | Тип | Макс. загрузка, чел. | Плановая трудоемкость, ч |",
            "|---|---|---:|---:|",
        ]
    )
    for resource_name, summary in sorted(resource_summary.items()):
        max_units = str(round(summary["max_units"], 2)).replace(".", ",")
        total_hours = str(round(summary["total_hours"], 1)).replace(".", ",")
        lines.append(f"| {resource_name} | Трудовой | {max_units} | {total_hours} |")

    lines.extend(
        [
            "",
            "### 8.3. Who Does What",
            "",
            "| Работа | Наименование | Назначенные ресурсы | Фиксированные затраты |",
            "|---|---|---|---|",
        ]
    )
    for code, name, resources, fixed_cost in get_who_does_what_rows():
        lines.append(f"| {code} | {name} | {resources} | {fixed_cost} |")

    MARKDOWN_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    create_graph()
    schedule, _daily_usage, resource_summary = create_projectlibre_outputs()
    document = Document(str(DOCX_PATH))
    remove_generated_appendix(document)
    update_wbs_table(document)
    add_appendix(document, schedule, resource_summary)
    document.save(str(DOCX_PATH))
    write_markdown(schedule, resource_summary)


if __name__ == "__main__":
    main()
