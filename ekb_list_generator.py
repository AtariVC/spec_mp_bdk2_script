import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def simplify_component_name(name):
    # Удаляем "ОСМ" в начале, если есть
    name = re.sub(r'^ОСМ\s+', '', name.strip())
    
    if name.startswith('Р'):
        match = re.match(r'Р(\d+-\d+)\s+(.*?)(\d+\.?\d*)\s*([кОмМк]?Ом?)\b', name)
        if match:
            base, _, value, unit = match.groups()
            unit = unit.replace('Ом', '').strip()
            if unit == 'к':
                value = f"{value}к"
            elif unit == 'М':
                value = f"{value}М"
            elif unit == 'кОм':
                value = f"{value}к"
            elif unit == 'МОм':
                value = f"{value}М"
            else:
                value = f"{value}"
            return f"Р{base} {value}"
    
    # Обработка конденсаторов (начинающихся на К)
    elif re.match(r'^К\d+-\d+', name):
        # Удаляем все технические пометки после ёмкости
        name = re.sub(r'([мкнп]Ф\s*).*', r'\1', name, flags=re.IGNORECASE)
        name = re.sub(r'\s*[±+].*', '', name)  # Удаляем допуски ±
        name = re.sub(r'\s*[МПН]\d*\b', '', name)  # Удаляем маркировки типа МП0, Н90
        
        # Паттерн для конденсаторов с напряжением
        match = re.match(
            r'(К\d+-\d+)\s+(\d+\s*В\s+)?(\d+[,.]?\d*)\s*([мкнп]?Ф?)\s*(\d+\s*В)?', 
            name, 
            flags=re.IGNORECASE
        )
        if match:
            base, volt_prefix, value, unit, volt_suffix = match.groups()
            value = value.replace(',', '.')
            unit = (unit or '').lower()
            
            # Определяем напряжение
            voltage = ''
            if volt_prefix:
                voltage = re.sub(r'\D', '', volt_prefix)
            elif volt_suffix:
                voltage = re.sub(r'\D', '', volt_suffix)
            
            # Форматируем выходную строку
            result = base
            if value:
                unit = unit.replace('ф', '')  # Удаляем "ф" если есть
                result += f" {value}{unit}Ф" if unit else f" {value}"
            if voltage:
                result += f" {voltage}В"
            
            # Удаляем возможные двойные пробелы
            return re.sub(r'\s+', ' ', result).strip()
        
        return name.split('(')[0].strip()
    
    # Для остальных типов оставляем как есть
    return name.split('(')[0].strip()

def convert_conclusions_to_passports(input_file, output_file):
    # Извлекаем номер паспорта из названия файла
    passport_prefix = re.search(r'(\d{2,3}[ПИ]\d{2,3})', input_file)
    if passport_prefix:
        passport_prefix = passport_prefix.group(1)
    else:
        passport_prefix = "XXПXX"
    
    # Чтение исходного файла
    df = pd.read_excel(input_file)
    
    # Создаем новый DataFrame для результата
    result_data = []
    
    # Обрабатываем каждую строку исходного файла
    for index, row in df.iterrows():
        # Получаем исходное название
        original_name = row['Тип изделия (номер партии)'].split('(')[0].strip()
        original_name = original_name.replace('ОСМ', '').strip()
        # Упрощаем название компонента
        simplified_name = simplify_component_name(original_name)
        
        # Формируем номер паспорта
        passport_number = f"ПДРФ.{passport_prefix}-{row['№']}"
        
        # Обрабатываем дату изготовления
        manufacture_date = str(row['Дата изготовления'])
        if 'нед.' in manufacture_date:
            year = re.search(r'(\d{4})', manufacture_date).group(1)[-2:]
            manufacture_date = f"{manufacture_date.split('нед.')[1].strip()}.{year}"
        elif 'пер.' in manufacture_date:
            manufacture_date = manufacture_date.split('пер.')[-1].strip()
        
        # Упрощаем дату до формата "М.ГГГГ" или "М.ГГ"
        if '.' in manufacture_date:
            parts = manufacture_date.split('.')
            if len(parts) >= 2:
                month = parts[0]
                year = parts[1][-2:] if len(parts[1]) > 2 else parts[1]
                manufacture_date = f"{month}.{year}"
        
        # Добавляем данные в результат
        result_data.append([simplified_name, passport_number, manufacture_date])
    
    # Создаем DataFrame из собранных данных
    result = pd.DataFrame(result_data, columns=['Тип изделия', 'Паспорт', 'Дата изготовления'])
    
    # Сохраняем результат в новый файл Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, header=False, sheet_name='Лист1')
        
        # Настраиваем ширину столбцов
        workbook = writer.book
        worksheet = writer.sheets['Лист1']
        
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 25
        worksheet.column_dimensions['C'].width = 15

    print(f"Файл успешно сохранен как {output_file}")

# Использование функции
input_filename = "/Users/vladk/dev/spec_mp_mk_bdk2/! Заключения 28П23.xlsx"
output_filename = "/Users/vladk/dev/spec_mp_mk_bdk2/список паспартов 28П23v2.xlsx"
convert_conclusions_to_passports(input_filename, output_filename)


