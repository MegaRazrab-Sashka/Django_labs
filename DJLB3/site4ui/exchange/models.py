from django.db import models
import random
from datetime import datetime, timedelta
import json


class DataStructure(models.Model):
    """Структура данных (метаданные)"""
    TYPE_CHOICES = [
        ('string', 'Строка'),
        ('date', 'Дата'),
        ('number', 'Число'),
    ]

    code = models.CharField(max_length=20, verbose_name="Код поля")
    order = models.IntegerField(verbose_name="Порядковый номер")
    name = models.CharField(max_length=100, verbose_name="Наименование")
    data_type = models.CharField(max_length=10, choices=TYPE_CHOICES, verbose_name="Тип данных")
    precision = models.IntegerField(default=0, verbose_name="Точность (для чисел)")
    is_transferred = models.BooleanField(default=True, verbose_name="Передавать")

    class Meta:
        ordering = ['order']

    def __str__(self):
        return f"{self.code} - {self.name}"


class ExchangeRate(models.Model):
    """Данные о курсах валют"""
    EXCHANGE_CHOICES = [
        ('MOEX', 'Московская биржа'),
        ('SPB', 'СПБ Биржа'),
        ('FOREX', 'Forex'),
        ('CBR', 'ЦБ РФ'),
    ]

    date = models.DateField(verbose_name="Дата")
    exchange = models.CharField(max_length=10, choices=EXCHANGE_CHOICES, verbose_name="Биржа")
    rate = models.FloatField(verbose_name="Курс USD")
    source = models.CharField(max_length=100, verbose_name="Источник информации")

    class Meta:
        ordering = ['-date', 'exchange']

    def __str__(self):
        return f"{self.date} - {self.exchange}: {self.rate}"


class DataExport(models.Model):
    """Экспортированные данные"""
    created_at = models.DateTimeField(auto_now_add=True)
    structure = models.JSONField(verbose_name="Структура данных")
    data = models.JSONField(verbose_name="Данные")
    file_path = models.CharField(max_length=500, blank=True, verbose_name="Путь к файлу")

    def export_to_file(self, filename):
        """Экспорт данных в текстовый файл"""
        import os
        from django.conf import settings

        file_path = os.path.join(settings.BASE_DIR, 'exports', filename)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        with open(file_path, 'w', encoding='utf-8') as f:
            # Записываем структуру
            f.write("#STRUCTURE\n")
            for field in self.structure:
                f.write(f"{field['code']}|{field['name']}|{field['type']}|{field['precision']}\n")

            # Записываем данные
            f.write("#DATA\n")
            for row in self.data:
                f.write("|".join(str(v) for v in row.values()) + "\n")

        self.file_path = file_path
        self.save()
        return file_path

    @staticmethod
    def import_from_file(file_path):
        """Импорт данных из файла"""
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        structure = []
        data = []
        is_data = False

        for line in lines:
            line = line.strip()
            if line == "#STRUCTURE":
                is_data = False
                continue
            elif line == "#DATA":
                is_data = True
                continue

            if not is_data and line:
                # Читаем структуру
                parts = line.split('|')
                structure.append({
                    'code': parts[0],
                    'name': parts[1],
                    'type': parts[2],
                    'precision': int(parts[3])
                })
            elif is_data and line:
                # Читаем данные
                parts = line.split('|')
                row = {}
                for i, field in enumerate(structure):
                    if i < len(parts):
                        if field['type'] == 'number':
                            row[field['code']] = float(parts[i])
                        elif field['type'] == 'date':
                            row[field['code']] = parts[i]
                        else:
                            row[field['code']] = parts[i]
                data.append(row)

        return DataExport.objects.create(
            structure=structure,
            data=data
        )