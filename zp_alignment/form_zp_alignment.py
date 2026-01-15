from docx import Document


ZP_ALIGN_INPUT_WORDS: dict = {
    'Должность_руководителя': 'Директор',
    'Краткое_наименование': 'Априори',
    'Руководитель': 'Кабанов Кабан Кабаныч',
    'Должность_работника': 'Бухгалтер',
    'Полное_имя_работника': 'Иванов Иван Иванович',
    'Банк_работника': 'Т-Банк',
    'Корреспондентский_счёт_банка_работника': '301 01 810 0 00000000 608',
    'БИК_банка_работника': '044525974',
    'ИНН_банка_работника': '7710140679',
    'Номер_счёта_работника': '4200004000003001'
}


def zp_align_sample_former(input_sample: str, output_result: str) -> str:
    try:
        doc = Document(input_sample)

        paragraph_count: int = len(doc.paragraphs)
        table_count: int = len(doc.tables)

        for i in range(table_count):
            # Итерируемся по каждой строке
            for row in doc.tables[i].table.rows:
                current_row = row
                cell_count_of_current_row: int = len(current_row.cells)

                # Итерируемся по каждой ячейке
                for j in range(cell_count_of_current_row):
                    current_cell_txt = current_row.cells[j].text

                    if '<Должность руководителя>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Должность руководителя>',
                                                               ZP_ALIGN_INPUT_WORDS['Должность_руководителя'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Краткое наименование>',
                                                               ZP_ALIGN_INPUT_WORDS['Краткое_наименование'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Руководитель>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Руководитель>',
                                                               ZP_ALIGN_INPUT_WORDS['Руководитель'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

        for i in range(paragraph_count):
            print(doc.paragraphs[i].text)
            current_txt = doc.paragraphs[i].text

            if '<Должность работника>' in current_txt:
                updated_text = current_txt.replace('<Должность работника>',
                                                   ZP_ALIGN_INPUT_WORDS['Должность_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Полное имя работника>' in current_txt:
                updated_text = current_txt.replace('<Полное имя работника>',
                                                   ZP_ALIGN_INPUT_WORDS['Полное_имя_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Банк работника>' in current_txt:
                updated_text = current_txt.replace('<Банк работника>',
                                                   ZP_ALIGN_INPUT_WORDS['Банк_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Корреспондентский счёт банка работника>' in current_txt:
                updated_text = current_txt.replace('<Корреспондентский счёт банка работника>',
                                                   ZP_ALIGN_INPUT_WORDS['Корреспондентский_счёт_банка_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<БИК банка работника>' in current_txt:
                updated_text = current_txt.replace('<БИК банка работника>',
                                                   ZP_ALIGN_INPUT_WORDS['БИК_банка_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<ИНН банка работника>' in current_txt:
                updated_text = current_txt.replace('<ИНН банка работника>',
                                                   ZP_ALIGN_INPUT_WORDS['ИНН_банка_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Номер счёта работника>' in current_txt:
                updated_text = current_txt.replace('<Номер счёта работника>',
                                                   ZP_ALIGN_INPUT_WORDS['Номер_счёта_работника'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

        doc.save(output_result)
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = zp_align_sample_former(
        input_sample='input_sample/Заявление на перечисление заработной платы.docx',
        output_result='output_result/output_Заявление на перечисление заработной платы.docx'
    )

