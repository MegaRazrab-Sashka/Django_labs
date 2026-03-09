from django.shortcuts import render, redirect
from django.contrib import messages
from django.http import JsonResponse
import random
from datetime import datetime, timedelta
import os
from django.conf import settings
import json
import win32com.client  # Для работы с OLE (требуется pywin32)
from .models import DataStructure, ExchangeRate, DataExport
from .forms import DataStructureForm, ExchangeRateForm, DataGenerationForm, ExcelConnectionForm
import pythoncom


def index(request):
    """Главная страница"""
    structures = DataStructure.objects.all()
    last_export = DataExport.objects.last()
    return render(request, 'exchange/index.html', {
        'structures': structures,
        'view_last_export': last_export,  # Добавьте эту строку
    })


def structure_list(request):
    """Управление структурой данных"""
    if request.method == 'POST':
        form = DataStructureForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Поле структуры добавлено')
            return redirect('structure_list')
    else:
        form = DataStructureForm()

    structures = DataStructure.objects.all()
    return render(request, 'exchange/structure_list.html', {
        'form': form,
        'structures': structures
    })


def structure_delete(request, pk):
    """Удаление поля структуры"""
    structure = DataStructure.objects.get(id=pk)
    structure.delete()
    messages.success(request, 'Поле удалено')
    return redirect('structure_list')


def data_list(request):
    """Просмотр и редактирование данных"""
    data = ExchangeRate.objects.all()

    if request.method == 'POST':
        form = ExchangeRateForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Запись добавлена')
            return redirect('data_list')
    else:
        form = ExchangeRateForm()

    return render(request, 'exchange/data_list.html', {
        'form': form,
        'data': data
    })


def generate_data(request):
    """Генерация тестовых данных"""
    if request.method == 'POST':
        form = DataGenerationForm(request.POST)
        if form.is_valid():
            count = form.cleaned_data['record_count']
            start_date = form.cleaned_data['start_date']
            exchanges = form.cleaned_data['exchanges']

            # Очищаем старые данные
            ExchangeRate.objects.all().delete()

            # Генерируем новые
            sources = ['Reuters', 'Bloomberg', 'TASS', 'Интерфакс']
            base_rate = 75.0

            for i in range(count):
                date = start_date + timedelta(days=i)
                for exchange in exchanges:
                    # Случайные колебания курса
                    rate = base_rate + random.uniform(-5, 5)
                    base_rate = rate * 0.999 + 75 * 0.001  # Медленное возвращение к среднему

                    ExchangeRate.objects.create(
                        date=date,
                        exchange=exchange,
                        rate=round(rate, 4),
                        source=random.choice(sources)
                    )

            messages.success(request, f'Сгенерировано {count * len(exchanges)} записей')
            return redirect('data_list')
    else:
        form = DataGenerationForm()

    return render(request, 'exchange/generate_data.html', {'form': form})


def export_data(request):
    """Экспорт данных в файл (программа-источник)"""
    # Получаем структуру из БД или создаем стандартную
    if not DataStructure.objects.exists():
        create_default_structure()

    structure_fields = DataStructure.objects.filter(is_transferred=True).order_by('order')

    # Формируем структуру для экспорта
    structure = []
    for field in structure_fields:
        structure.append({
            'code': field.code,
            'name': field.name,
            'type': field.data_type,
            'precision': field.precision
        })

    # Получаем данные
    data = ExchangeRate.objects.all().order_by('date', 'exchange')

    # Формируем данные для экспорта
    export_data = []
    for item in data:
        row = {}
        for field in structure_fields:
            if field.code == 'date':
                row[field.code] = item.date.strftime('%Y-%m-%d')
            elif field.code == 'exchange':
                row[field.code] = item.get_exchange_display()
            elif field.code == 'rate':
                row[field.code] = item.rate
            elif field.code == 'source':
                row[field.code] = item.source
        export_data.append(row)

    # Создаем экспорт
    export = DataExport.objects.create(
        structure=structure,
        data=export_data
    )

    # Сохраняем в файл
    filename = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    file_path = export.export_to_file(filename)

    messages.success(request, f'Данные экспортированы в файл: {file_path}')
    return redirect('exchange_index')


def process_and_send_to_excel(request):
    """Программа-сервер: загрузка, обработка и отправка в Excel"""
    if request.method == 'POST':
        form = ExcelConnectionForm(request.POST)
        if form.is_valid():
            try:
                excel_file = form.cleaned_data['excel_file']
                start_cell = form.cleaned_data['start_cell']
                date_col = form.cleaned_data['date_column']
                rate_col = form.cleaned_data['rate_column']

                # 1. Загружаем последний экспортированный файл
                latest_export = DataExport.objects.last()
                if not latest_export:
                    messages.error(request, 'Нет экспортированных данных')
                    return redirect('exchange_index')

                # 2. Обрабатываем данные (усредняем по датам)
                processed_data = process_data(latest_export.data)

                # 3. Отправляем в Excel через OLE (убираем request)
                success = send_to_excel_ole(processed_data, excel_file, start_cell, date_col, rate_col)

                if success:
                    messages.success(request,
                                     f'Данные успешно отправлены в Excel. Обработано записей: {len(processed_data)}')
                else:
                    messages.error(request, 'Ошибка при отправке в Excel')

            except Exception as e:
                messages.error(request, f'Ошибка: {str(e)}')
    else:
        form = ExcelConnectionForm()

    return render(request, 'exchange/send_to_excel.html', {'form': form})


def create_default_structure():
    """Создание стандартной структуры данных"""
    default_fields = [
        {'code': 'date', 'order': 1, 'name': 'Дата', 'data_type': 'date', 'precision': 0},
        {'code': 'exchange', 'order': 2, 'name': 'Биржа', 'data_type': 'string', 'precision': 0},
        {'code': 'rate', 'order': 3, 'name': 'Курс USD', 'data_type': 'number', 'precision': 4},
        {'code': 'source', 'order': 4, 'name': 'Источник', 'data_type': 'string', 'precision': 0},
    ]

    for field in default_fields:
        DataStructure.objects.create(**field)


def process_data(raw_data):
    """Обработка данных: усреднение курса по датам"""
    from collections import defaultdict

    # Группируем по датам
    daily_rates = defaultdict(list)

    for item in raw_data:
        date = item.get('date', '')
        rate = item.get('rate', 0)
        if date and rate:
            daily_rates[date].append(float(rate))

    # Усредняем
    processed = []
    for date, rates in sorted(daily_rates.items()):
        avg_rate = sum(rates) / len(rates)
        processed.append({
            'date': date,
            'avg_rate': round(avg_rate, 4),
            'count': len(rates)
        })

    return processed


def send_to_excel_ole(data, excel_file, start_cell, date_col, rate_col):
    """Отправка данных в Excel через OLE интерфейс"""
    try:
        # Инициализируем COM
        pythoncom.CoInitialize()

        # Создаем объект Excel
        excel = win32com.client.Dispatch("Excel.Application")

        # Делаем Excel невидимым во время заполнения
        excel.Visible = False

        try:
            # Проверяем существование файла
            import os
            if not os.path.exists(excel_file):
                print(f"Файл не найден: {excel_file}")
                return False

            # Открываем книгу
            wb = excel.Workbooks.Open(excel_file)
            ws = wb.Worksheets(1)

            # Парсим начальную ячейку
            import re
            match = re.match(r"([A-Z]+)(\d+)", start_cell)
            if match:
                start_col_letter, start_row = match.groups()
                start_row = int(start_row)

                # Конвертируем букву колонки в номер
                start_col_num = 0
                for char in start_col_letter:
                    start_col_num = start_col_num * 26 + (ord(char.upper()) - ord('A') + 1)
            else:
                start_col_num, start_row = 1, 1

            # Очищаем старые данные
            ws.Range(ws.Cells(start_row, date_col),
                     ws.Cells(start_row + len(data), rate_col)).ClearContents()

            # Заполняем данными
            for i, item in enumerate(data):
                try:
                    # Преобразуем дату из строки в формат Excel
                    date_str = item['date']
                    # Разбираем строку даты (формат YYYY-MM-DD)
                    year, month, day = map(int, date_str.split('-'))
                    # Создаем объект даты
                    from datetime import date as date_class
                    excel_date = date_class(year, month, day)

                    # Записываем дату
                    ws.Cells(start_row + i, date_col).Value = excel_date
                    # Записываем курс
                    ws.Cells(start_row + i, rate_col).Value = float(item['avg_rate'])

                except Exception as e:
                    print(f"Ошибка при записи строки {i}: {str(e)}")
                    continue

            # Автоматически подбираем ширину колонок
            ws.Columns.AutoFit()

            # Обновляем диаграмму (если она есть)
            try:
                for chart in ws.ChartObjects():
                    chart.Chart.Refresh()
            except:
                pass  # Если нет диаграммы, игнорируем

            # Сохраняем и показываем
            wb.Save()
            excel.Visible = True

            return True

        except Exception as e:
            print(f"Ошибка при работе с Excel: {str(e)}")
            excel.Quit()
            return False

    except Exception as e:
        print(f"Ошибка OLE: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()