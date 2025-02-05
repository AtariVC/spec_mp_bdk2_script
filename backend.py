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
    passports = pd.read_excel(ekb_file, sheet_name='Лист1', dtype={'Дата': str})  # Дата как текст
    return specification, passports

def prepare_data(specification, passports):
    """Очистка и подготовка данных."""
    specification = specification.dropna(subset=['Наименование'])
    passports.columns = ['Наименование', 'Паспорт', 'Дата']
    
    # Преобразуем дату в строку (если она не строка)
    passports['Дата'] = passports['Дата'].astype(str)
    
    return specification, passports

def extract_year_and_add_25(date_str):
    """Извлекает год из строки формата 'MM.YYYY' и прибавляет 25."""
    try:
        year = int(date_str.split('.')[-1])  # Берем только год
        return year + 25  # Прибавляем 25 и конвертируем в строку
    except (ValueError, IndexError):
        return ""  # Если ошибка, оставляем пустым

def merge_data(specification, passports):
    """Объединение данных по наименованию."""
    merged_data = pd.merge(specification, passports, on='Наименование', how='left')
    
    # Преобразуем столбец "Дата" в строку
    merged_data['Дата'] = merged_data['Дата'].astype(str)
    
    # Добавляем столбцы "I" и "J"
    merged_data['H'] = 25  # Весь столбец заполняем числом 25
    merged_data['I'] = merged_data['Дата'].apply(extract_year_and_add_25)  # Вычисляем "J"

    return merged_data

def create_result_table(merged_data):
    """Создание результирующей таблицы с пустыми столбцами и нумерацией."""
    result = pd.DataFrame({
        'A': merged_data['Поз.']-1,  
        'B': '',  
        'C': merged_data['Наименование'],  
        'D': merged_data.get('Кол.', ''),  
        'E': merged_data.get('Кол.', ''),  
        'F': merged_data['Паспорт'],
        'G': merged_data['Дата'],  
        'H': merged_data['H'],  
        'I': merged_data['I']
    })
    # result.insert(0, '№', range(1, len(result) + 1))  
    return result

def add_section_names(result, specification):
    """Добавление названий разделов."""
    section_names = specification[specification['Наименование'].str.contains('Конденсаторы|Микросхемы|Диоды|Транзисторы', case=False, na=False)]['Наименование']
    final_data = []
    for _, row in result.iterrows():
        if row['C'] in section_names.values:
            final_data.append(['', '', row['C'], '', '', '', '', '', ''])  
        else:
            final_data.append([row['A'], "", row['C'], row['D'], row['E'], row['F'], row['G'], row['H'], row['I']])
    return final_data

def save_to_excel(final_data, output_path):
    """Сохранение данных в Excel с разбиением на новые страницы."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"  

    row_count = 0
    sheet_number = 1

    for row in final_data:
        if row_count >= 19:  
            sheet_number += 1
            ws = wb.create_sheet(title=f"Лист{sheet_number}")
            row_count = 0  
        
        ws.append(row)
        row_count += 1

    # Устанавливаем формат столбцов "H", "I" и "J" как текст
    for sheet in wb.worksheets:
        for col in sheet.iter_cols(min_col=7, max_col=9):  # H, I, J - это 8, 9, 10 столбцы
            for cell in col:
                cell.number_format = '@'  # '@' означает текстовый формат

    wb.save(output_path)

def filter_unwanted_sections(specification):
    """Фильтрация ненужных разделов."""
    unwanted_sections = ["Документация", "Сборочный чертеж", "Сборочные единицы", "Плата печатная", "Прочие изделия"]
    
    # Убираем пробелы и приводим к нижнему регистру для точности сравнения
    specification['Наименование'] = specification['Наименование'].str.strip()
    
    # Применяем фильтрацию
    filtered_specification = specification[~specification['Наименование'].str.contains('|'.join(unwanted_sections), case=False, na=False)]
    
    return filtered_specification


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
