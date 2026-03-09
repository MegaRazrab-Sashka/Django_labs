from django import forms
from .models import DataStructure, ExchangeRate
from datetime import datetime


class DataStructureForm(forms.ModelForm):
    class Meta:
        model = DataStructure
        fields = ['code', 'order', 'name', 'data_type', 'precision', 'is_transferred']
        widgets = {
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'order': forms.NumberInput(attrs={'class': 'form-control'}),
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'data_type': forms.Select(attrs={'class': 'form-control'}),
            'precision': forms.NumberInput(attrs={'class': 'form-control'}),
            'is_transferred': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }


class ExchangeRateForm(forms.ModelForm):
    class Meta:
        model = ExchangeRate
        fields = ['date', 'exchange', 'rate', 'source']
        widgets = {
            'date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'exchange': forms.Select(attrs={'class': 'form-control'}),
            'rate': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.0001'}),
            'source': forms.TextInput(attrs={'class': 'form-control'}),
        }


class DataGenerationForm(forms.Form):
    record_count = forms.IntegerField(
        initial=10,
        min_value=1,
        max_value=100,
        label="Количество записей",
        widget=forms.NumberInput(attrs={'class': 'form-control'})
    )

    start_date = forms.DateField(
        initial=datetime.now().date,
        label="Начальная дата",
        widget=forms.DateInput(attrs={'class': 'form-control', 'type': 'date'})
    )

    exchanges = forms.MultipleChoiceField(
        choices=ExchangeRate.EXCHANGE_CHOICES,
        initial=['MOEX', 'SPB', 'FOREX'],
        label="Биржи",
        widget=forms.SelectMultiple(attrs={'class': 'form-control'})
    )


class ExcelConnectionForm(forms.Form):
    excel_file = forms.CharField(
        initial='C:\\Users\\Alex\\Desktop\\exchange_chart.xls',
        label="Путь к файлу Excel",
        widget=forms.TextInput(attrs={'class': 'form-control', 'size': 50})
    )

    start_cell = forms.CharField(
        initial='A2',
        label="Начальная ячейка",
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )

    date_column = forms.IntegerField(
        initial=1,
        label="Колонка с датой",
        widget=forms.NumberInput(attrs={'class': 'form-control'})
    )

    rate_column = forms.IntegerField(
        initial=2,
        label="Колонка с курсом",
        widget=forms.NumberInput(attrs={'class': 'form-control'})
    )