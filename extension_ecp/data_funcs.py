import pandas as pd

from datetime import datetime, timedelta
from openpyxl import load_workbook



def extract_row_if_red(df: pd.DataFrame, path_to_file: str) -> list[dict]:
    extracted_rows: list[dict] = []

    wb = load_workbook(path_to_file)
    ws = wb['Лист1']

    registry_meta_info: dict = {
        'ИНН': 'A',
        'Наименование': 'B',
        'ДатаОкончанияЭЦП': 'E'
    }
    # Для фильтра
    today = datetime.today()
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)

    for idx, row in df.iterrows():
        is_add = False

        # Формируем похожий словарь, но с пустыми значениями. Этот словарь мы будем добавлять в список по заданными
        # требованиям ТЗ
        data_row: dict = {
            'ИНН': None,
            'Наименование': None,
            'ДатаОкончанияЭЦП': None
        }

        # ИНН
        current_cell_inn = f"{registry_meta_info['ИНН']}{idx + 2}"
        inn_value = ws[current_cell_inn].value

        current_cell_company_name = f"{registry_meta_info['Наименование']}{idx + 2}"
        company_name_value = ws[current_cell_company_name].value

        current_cell_ecp_date = f"{registry_meta_info['ДатаОкончанияЭЦП']}{idx + 2}"
        ecp_date_value = ws[current_cell_ecp_date].value


        if isinstance(ecp_date_value, datetime):
            if ecp_date_value <= today + timedelta(days=15):
                data_row['ИНН'] = inn_value
                data_row['Наименование'] = company_name_value

                formatted_ecp_date_value = ecp_date_value.strftime('%d.%m.%Y')

                data_row['ДатаОкончанияЭЦП'] = formatted_ecp_date_value

                extracted_rows.append(data_row)

    return extracted_rows


def app_data_extracter(path_file: str) -> list[dict]:
    df: pd.DataFrame = pd.read_excel(path_file)

    # Фильтруем df по критерию наличия значения в колонке 'Письмо'
    filtered_df = df[df['ИНН'].notna()]
    filtered_df = filtered_df[filtered_df['Дата окончания ЭЦП'].notna()]

    # Получаем извлеченные данные по критериям ТЗ
    extracted_data: list[dict] = extract_row_if_red(df=filtered_df, path_to_file=path_file)

    return extracted_data


if __name__ == '__main__':
    extracted_data: list[dict] = app_data_extracter(path_file='report_tokens/ТОКЕНЫ ОТЧЕТНОСТЬ.xlsx')