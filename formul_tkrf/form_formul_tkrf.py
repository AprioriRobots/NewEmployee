from openpyxl import load_workbook


def formul_tkrf_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    FORMUL_TKRF_INPUT_WORDS: dict = {
        '<Гражданство страны>': parsed_data_1c['personal_info']['Страна'],
        # '<Патент>': '2',
        # '<Серия патента>': '3',
        # '<Номер патента>': '4',
        # '<Дата выдачи патента>': '5',
        # '<Патент выдан>': '6',
        # '<Рабочая дата>': '7',
        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],
        '<Длительность рабочей недели>': parsed_data_1c['job_info']['ГрафикРаботы'],
        # '<Явка>': '10',
        '<Количество дней ежегодного отпуска>': '28',
        # '<Срочный трудовой договор до>': '12',
        '<Испытательный срок>': str(parsed_data_1c['job_info']['ИспытательныйСрок']),
        # '<Вид занятости>': '14',
        # '<Оплата по часовому тарифу>': '15',
        '<Должность руководителя>': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя']
        # '<Основание срочного договора>': '17'
    }

    try:
        wb = load_workbook(input_sample)
        ws = wb['Лист1']

        formul_tkrf_meta_info: dict = {
            'col_1': 'D',
            'col_2': 'E',
        }
        for i in range(1, 25):
            current_cell_col1 = f'{formul_tkrf_meta_info['col_1']}{i}'
            current_cell_col2 = f'{formul_tkrf_meta_info['col_2']}{i}'

            current_value_1 = ws[current_cell_col1].value
            current_value_2 = ws[current_cell_col2].value

            if '<Гражданство страны>' in current_value_2:
                ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
                    '<Гражданство страны>',
                    FORMUL_TKRF_INPUT_WORDS['<Гражданство страны>']
                )

            # if '<Патент>' in current_value_1 or current_value_2:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Патент>',
            #         FORMUL_TKRF_INPUT_WORDS['<Патент>']
            #     )
            #     ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
            #         '<Патент>',
            #         FORMUL_TKRF_INPUT_WORDS['<Патент>']
            #     )

            # if '<Серия патента>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Серия патента>',
            #         FORMUL_TKRF_INPUT_WORDS['<Серия патента>']
            #     )
            #
            # if '<Номер патента>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Номер патента>',
            #         FORMUL_TKRF_INPUT_WORDS['<Номер патента>']
            #     )
            #
            # if '<Дата выдачи патента>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Дата выдачи патента>',
            #         FORMUL_TKRF_INPUT_WORDS['<Дата выдачи патента>']
            #     )
            #
            # if '<Патент выдан>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Патент выдан>',
            #         FORMUL_TKRF_INPUT_WORDS['<Патент выдан>']
            #     )

            if '<Длительность рабочей недели>' in current_value_1:
                ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
                    '<Длительность рабочей недели>',
                    FORMUL_TKRF_INPUT_WORDS['<Длительность рабочей недели>']
                )

            # if '<Явка>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Явка>',
            #         FORMUL_TKRF_INPUT_WORDS['<Явка>']
            #     )

            if '<Количество дней ежегодного отпуска>' in current_value_1:
                ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
                    '<Количество дней ежегодного отпуска>',
                    FORMUL_TKRF_INPUT_WORDS['<Количество дней ежегодного отпуска>']
                )

            # if '<Срочный трудовой договор до>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Срочный трудовой договор до>',
            #         FORMUL_TKRF_INPUT_WORDS['<Срочный трудовой договор до>']
            #     )
            #
            # if '<Основание срочного договора>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Основание срочного договора>',
            #         FORMUL_TKRF_INPUT_WORDS['<Основание срочного договора>']
            #     )

            # if '<Оплата по часовому тарифу>' in current_value_1:
            #     ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
            #         '<Оплата по часовому тарифу>',
            #         FORMUL_TKRF_INPUT_WORDS['<Оплата по часовому тарифу>']
            #     )

            if '<Должность руководителя>' in current_value_1:
                ws[current_cell_col1].value = ws[current_cell_col1].value.replace(
                    '<Должность руководителя>',
                    FORMUL_TKRF_INPUT_WORDS['<Должность руководителя>']
                )

            # if '<Рабочая дата>' in current_value_2:
            #     ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
            #         '<Рабочая дата>',
            #         FORMUL_TKRF_INPUT_WORDS['<Рабочая дата>']
            #     )

            if '<Дата рождения>' in current_value_2:
                ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
                    '<Дата рождения>',
                    FORMUL_TKRF_INPUT_WORDS['<Дата рождения>']
                )

            if '<Испытательный срок>' in current_value_2:
                ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
                    '<Испытательный срок>',
                    FORMUL_TKRF_INPUT_WORDS['<Испытательный срок>']
                )

            # if '<Вид занятости>' in current_value_2:
            #     ws[current_cell_col2].value = ws[current_cell_col2].value.replace(
            #         '<Вид занятости>',
            #         FORMUL_TKRF_INPUT_WORDS['<Вид занятости>']
            #     )

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = formul_tkrf_sample_former(
        input_sample='input_sample/Формулировки для трудового договора.xlsx',
        output_result='output_result/output_Формулировки для трудового договора.xlsx'
    )
