import os
import glob
import subprocess
from tqdm.notebook import trange
import pandas as pd
import re
from openpyxl import load_workbook

def convert_and_replace_xls_to_xlsx(root_folder):
    """
    Рекурсивно конвертирует все .xls-файлы в указанной папке (и вложенных папках) в .xlsx с помощью LibreOffice.
    Новые .xlsx-файлы сохраняются в тех же папках. Исходные .xls-файлы удаляются.
    Требует установленной LibreOffice (на Colab - !apt-get install -y libreoffice).
    """
    # Проверка наличия libreoffice
    try:
        subprocess.run(['libreoffice', '--version'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        raise RuntimeError('LibreOffice не установлена. Установите ее перед запуском этой функции!')

    # Рекурсивный поиск .xls
    xls_files = [y for x in os.walk(root_folder) for y in glob.glob(os.path.join(x[0], '*.xls'))]

    for i in trange(len(xls_files), desc="Конвертация xls → xlsx", unit="файл"):
        xls_path = xls_files[i]
        folder_path = os.path.dirname(xls_path)
        file_name = os.path.splitext(os.path.basename(xls_path))[0]
        xlsx_path = os.path.join(folder_path, file_name + '.xlsx')

        # Конвертация xls в xlsx
        try:
            subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'xlsx', '--outdir', folder_path, xls_path
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if os.path.exists(xlsx_path):
                os.remove(xls_path)
        except Exception as e:
            pass  # Можно добавить логгирование ошибок, если нужно

# Пример вызова:
# convert_and_replace_xls_to_xlsx('/content/gdrive/MyDrive/VLGR')



def parse_excel_to_df_1(file_path, level_names):
    """
    Парсит Excel-файл в потоковый DataFrame.

    Параметры:
    - file_path: str, путь к файлу Excel
    - level_names: dict, наименования уровней иерархии, например:
      {
          'account': 'Наименование счета',
          'sublevel': 'Подразделение',
          'detail': 'Статья затрат'
      }

    Возвращает:
    - DataFrame с потоковой структурой данных
    """
    def extract_month_year(text):
        match = re.search(r'([А-ЯЁа-яё]+)\s+(\d{4})', text)
        return f"{match.group(1)} {match.group(2)}" if match else text
    
    
    def get_cell_color(cell):
        return cell.fill.start_color.index if cell.fill.start_color.index else None
    
    wb = load_workbook(file_path, data_only=True)
    sheet = wb.active

    company_name = sheet['A1'].value.strip()
    date_info = extract_month_year(sheet['A2'].value.strip())

    start_row = 9
    max_row = sheet.max_row

    columns_mapping = {
        2: ('Сальдо на начало периода', 'Дебет'),
        3: ('Сальдо на начало периода', 'Кредит'),
        4: ('Обороты за период', 'Дебет'),
        5: ('Обороты за период', 'Кредит'),
        6: ('Сальдо на конец периода', 'Дебет'),
        7: ('Сальдо на конец периода', 'Кредит'),
    }

    current_account = None
    current_sublevel = None
    rows_data = []

    for row in range(start_row, max_row + 1):
        cell = sheet.cell(row=row, column=1)
        cell_value = cell.value
        cell_color = get_cell_color(cell)

        if cell_color == 'FFD6E5CB' and (cell_value is not None and str(cell_value).strip().lower() == 'итого'):
            for col_idx in range(2, 8):
                cell_data = sheet.cell(row=row, column=col_idx).value
                if cell_data not in (None, ''):
                    indicator, debit_credit = columns_mapping[col_idx]
                    rows_data.append({
                        'Компания': company_name,
                        'Период': date_info,
                        level_names['account']: 'Итого',
                        level_names['sublevel']: None,
                        level_names['detail']: None,
                        'Показатель': indicator,
                        'Дебет/Кредит': debit_credit,
                        'Значение': cell_data
                    })
            continue

        if cell_color == 'FFE4F0DD':
            current_account = cell_value
            for col_idx in range(2, 8):
                cell_data = sheet.cell(row=row, column=col_idx).value
                if cell_data not in (None, ''):
                    indicator, debit_credit = columns_mapping[col_idx]
                    rows_data.append({
                        'Компания': company_name,
                        'Период': date_info,
                        level_names['account']: current_account,
                        level_names['sublevel']: None,
                        level_names['detail']: None,
                        'Показатель': indicator,
                        'Дебет/Кредит': debit_credit,
                        'Значение': cell_data
                    })
            continue

        if cell_color == 'FFF0F6EF':
            current_sublevel = cell_value
            for col_idx in range(2, 8):
                cell_data = sheet.cell(row=row, column=col_idx).value
                if cell_data not in (None, ''):
                    indicator, debit_credit = columns_mapping[col_idx]
                    rows_data.append({
                        'Компания': company_name,
                        'Период': date_info,
                        level_names['account']: current_account,
                        level_names['sublevel']: current_sublevel,
                        level_names['detail']: None,
                        'Показатель': indicator,
                        'Дебет/Кредит': debit_credit,
                        'Значение': cell_data
                    })
            continue

        if cell_color == 'FFD6E5CB':
            continue

        for col_idx in range(2, 8):
            cell_data = sheet.cell(row=row, column=col_idx).value
            if cell_data not in (None, ''):
                indicator, debit_credit = columns_mapping[col_idx]
                rows_data.append({
                    'Компания': company_name,
                    'Период': date_info,
                    level_names['account']: current_account,
                    level_names['sublevel']: current_sublevel,
                    level_names['detail']: cell_value,
                    'Показатель': indicator,
                    'Дебет/Кредит': debit_credit,
                    'Значение': cell_data
                })

    return pd.DataFrame(rows_data)


# Пример использования:
# df = parse_excel_to_df_1('/content/Оборотно-сальдовая ведомость январь сч26.xlsx',
#                        {'account': 'Счёт', 'sublevel': 'Объект', 'detail': 'Статья'})
# print(df.head(20))
