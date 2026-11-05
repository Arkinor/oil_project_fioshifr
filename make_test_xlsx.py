"""
make_test_xlsx.py — генератор тестового Excel для проверки маскировки.
Запуск:  python make_test_xlsx.py
Результат: test_employees.xlsx  (30 сотрудников, все типы колонок)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from faker import Faker
from datetime import datetime, timedelta

fake = Faker('ru_RU')
Faker.seed(42)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Сотрудники"

# ---- Заголовки (все типы, которые умеет определять main.py) ----
headers = [
    "ФИО",
    "Телефон",
    "Паспорт",
    "Email",
    "Дата рождения",
    "Оклад",
    "Отдел",
    "Должность",
    "Комментарий",
    "Пустая колонка",
]

# ---- Стилизация заголовка ----
header_fill = PatternFill(start_color="4361EE", end_color="4361EE", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=11)
header_align = Alignment(horizontal="center", vertical="center")

for col_idx, h in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_idx, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_align

# ---- Данные ----
отделы = ["IT", "Бухгалтерия", "Отдел кадров", "Продажи", "Логистика", "Юридический"]
должности = [
    "Программист", "Ведущий разработчик", "Бухгалтер", "HR-менеджер",
    "Менеджер по продажам", "Логист", "Юрист", "Системный администратор"
]
комментарии = [
    "Работает удалённо с 2023 г.",
    "В отпуске до 15.06",
    "Переведён из другого отдела",
    "На испытательном сроке",
    "Требуется переобучение",
    "Ключевой сотрудник",
    "",
    "Работает по совместительству",
]

N = 30
for i in range(N):
    row = i + 2
    ws.cell(row=row, column=1, value=fake.name())
    ws.cell(row=row, column=2, value=fake.phone_number())
    ws.cell(row=row, column=3, value=f"{fake.random_int(1000, 9999)} {fake.random_int(100000, 999999)}")
    ws.cell(row=row, column=4, value=fake.email())
    ws.cell(row=row, column=5, value=fake.date_of_birth(minimum_age=22, maximum_age=60).strftime("%d.%m.%Y"))
    ws.cell(row=row, column=6, value=round(fake.random.uniform(40000, 250000), 2))
    ws.cell(row=row, column=7, value=fake.random_element(отделы))
    ws.cell(row=row, column=8, value=fake.random_element(должности))
    ws.cell(row=row, column=9, value=fake.random_element(комментарии))

# ---- Автоширина колонок ----
for col_idx in range(1, len(headers) + 1):
    letter = get_column_letter(col_idx)
    max_len = len(str(headers[col_idx - 1]))
    for row_idx in range(2, N + 2):
        val = ws.cell(row=row_idx, column=col_idx).value
        if val is not None:
            max_len = max(max_len, len(str(val)))
    ws.column_dimensions[letter].width = min(max_len + 3, 40)

# ---- Заморозка заголовка ----
ws.freeze_panes = "A2"

wb.save("test_employees.xlsx")
print(f"✅ Создан test_employees.xlsx: {N} строк, {len(headers)} колонок")
print("   Типы колонок, которые должны определиться:")
print("   ФИО → fio | Телефон → phone | Паспорт → passport | Email → email")
print("   Дата рождения → date | Оклад → number | Отдел/Должность/Комментарий → text")
print("   Пустая колонка → не маскируется")