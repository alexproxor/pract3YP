from __future__ import annotations

from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt
import networkx as nx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches


ROOT = Path("/workspace")
DOCX_PATH = ROOT / "Прохоров.А.Ф_ПР3.docx"
GENERATED_DIR = ROOT / "generated"
GRAPH_PATH = GENERATED_DIR / "pr3_network_graph.png"
MARKDOWN_PATH = ROOT / "Прохоров.А.Ф_ПР3_дополнение.md"

APPENDIX_MARKER = "Установление зависимостей работ и типов связей между ними"


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
        "duration": "240",
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
        "duration": "96",
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
    run.font.size = Inches(0.18)


def add_appendix(document: Document) -> None:
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


def write_markdown() -> None:
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

    MARKDOWN_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    create_graph()
    document = Document(str(DOCX_PATH))
    remove_generated_appendix(document)
    update_wbs_table(document)
    add_appendix(document)
    document.save(str(DOCX_PATH))
    write_markdown()


if __name__ == "__main__":
    main()
