import pandas as pd
from openpyxl import Workbook

# Пути к файлам
spec_file = "/Users/vladk/Downloads/Telegram Desktop/ЮМП.250.212.045.03 Спецификация.xlsx"
ekb_file = "/Users/vladk/Downloads/Telegram Desktop/список паспартов ЭКБ.xlsx"
output_path = "/Users/vladk/Downloads/Telegram Desktop/output/merged_output.xlsx"

def load_data(spec_file, ekb_file):
    """Загрузка данных из всех листов файла спецификации."""
    all_sheets = pd.read_excel(spec_file, sheet_name=None)  # Загружаем все листы
    specification = pd.concat(all_sheets.values(), ignore_index=True)  # Объединяем в один DataFrame
    passports = pd.read_excel(ekb_file, sheet_name='Лист1')
    return specification, passports

def prepare_data(specification, passports):
    """Очистка и подготовка данных."""
    specification = specification.dropna(subset=['Наименование'])
    passports.columns = ['Наименование', 'Паспорт', 'Дата']
    return specification, passports

def filter_unwanted_sections(specification):
    """Фильтрация ненужных разделов."""
    unwanted_sections = ["Документация", "Сборочный чертеж", "Сборочные единицы", "Плата печатная", "Прочие изделия"]
    return specification[~specification['Наименование'].isin(unwanted_sections)]

def merge_data(specification, passports):
    """Объединение данных по наименованию."""
    merged_data = pd.merge(specification, passports, on='Наименование', how='left')
    return merged_data

def create_result_table(merged_data):
    """Создание результирующей таблицы с пустыми столбцами и нумерацией."""
    result = pd.DataFrame({
        'A': '',  
        'B': '',  
        'C': merged_data['Наименование'],  
        'D': '',  
        'E': merged_data.get('Кол.', ''),  
        'F': merged_data.get('Кол.', ''),  
        'G': merged_data['Паспорт'],
        'H': merged_data['Дата'],  
        'I': ''  
    })
    result.insert(0, '№', range(1, len(result) + 1))  
    return result

def add_section_names(result, specification):
    """Добавление названий разделов."""
    section_names = specification[specification['Наименование'].str.contains('Конденсаторы|Микросхемы|Диоды|Транзисторы', case=False, na=False)]['Наименование']
    final_data = []
    current_section = None
    for _, row in result.iterrows():
        if any(section in row['C'] for section in section_names):
            current_section = row['C']
            final_data.append(['', '', current_section, '', '', '', '', '', ''])  
        else:
            final_data.append([row['№'], row['A'], row['C'], row['D'], row['E'], row['F'], row['G'], row['H'], row['I']])
    return final_data

def save_to_excel(final_data, output_path):
    """Сохранение данных в Excel с разбиением на новые страницы, если больше 19 строк."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"  # Название первого листа
    
    row_count = 0
    sheet_number = 1
    
    for row in final_data:
        if row_count >= 19:  # Если 19 строк, создаем новый лист
            sheet_number += 1
            ws = wb.create_sheet(title=f"Лист{sheet_number}")
            row_count = 0  # Сбрасываем счетчик
        
        ws.append(row)
        row_count += 1

    wb.save(output_path)

def main():
    """Основная функция."""
    specification, passports = load_data(spec_file, ekb_file)
    specification, passports = prepare_data(specification, passports)
    specification = filter_unwanted_sections(specification)
    merged_data = merge_data(specification, passports)
    result = create_result_table(merged_data)
    final_data = add_section_names(result, specification)
    save_to_excel(final_data, output_path)

if __name__ == "__main__":
    main()
