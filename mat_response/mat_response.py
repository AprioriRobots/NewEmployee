from docx import Document

MAT_RESP_INPUT_WORDS: dict = {
    'Дата_приказа_о_приеме': '1',
    '<Полное наименование>': '2',
    '<Краткое наименование>': '3',
    '<Должность руководителя>': '4',
    '<Руководитель>': '5',
    '<Основание подписи>': '6',
    '<Гражданство страны>': '7',
    '<Полное имя работника>': '8',
    '<Должность работника>': '9',
    '<Юридический адрес>': '10',
    '<Фактический адрес>': '11',
    '<ИНН организации>': '12',
    '<КПП организации>': '13',
    '<ОГРН организации>': '14',
    '<Расчетный счет организации>': '15',
    '<БИК банка организации>': '16',
    '<Адрес по прописке>': '17',
    '<Дата рождения>': '18',
    '<Место рождения>': '19',
    '<Серия документа, удостоверяющего личность>': '20',
    '<Номер документа, удостоверяющего личность>': '21',
    '<Орган, выдавший документ, удостоверяющего личность>': '22',
    '<Дата выдачи документа, удостоверяющего личность>': '23',
    '<СНИЛС>': '24',
    '<Корреспондентский счет организации>': '25',
}


def mat_resp_sample_former(input_sample: str, output_result: str) -> str:
    try:
        doc = Document(input_sample)

        paragraph_count: int = len(doc.paragraphs)
        table_count: int = len(doc.tables)

        # Итерируемся по всем таблицам в документе
        for i in range(table_count):
            # Итерируемся по каждой строке
            for row in doc.tables[i].table.rows:
                current_row = row
                cell_count_of_current_row: int = len(current_row.cells)

                # Итерируемся по каждой ячейке
                for j in range(cell_count_of_current_row):
                    current_cell_txt = current_row.cells[j].text

                    if '<Дата приказа о приеме>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Дата приказа о приеме>',
                                                               MAT_RESP_INPUT_WORDS['Дата_приказа_о_приеме'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Краткое наименование>',
                                                               MAT_RESP_INPUT_WORDS['<Краткое наименование>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Юридический адрес>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Юридический адрес>',
                                                               MAT_RESP_INPUT_WORDS['<Юридический адрес>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Фактический адрес>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Фактический адрес>',
                                                               MAT_RESP_INPUT_WORDS['<Фактический адрес>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<ИНН организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<ИНН организации>',
                                                               MAT_RESP_INPUT_WORDS['<ИНН организации>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<КПП организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<КПП организации>',
                                                               MAT_RESP_INPUT_WORDS['<КПП организации>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<ОГРН организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<ОГРН организации>',
                                                               MAT_RESP_INPUT_WORDS['<ОГРН организации>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Расчетный счет организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Расчетный счет организации>',
                                                               MAT_RESP_INPUT_WORDS['<Расчетный счет организации>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<БИК банка организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<БИК банка организации>',
                                                               MAT_RESP_INPUT_WORDS['<БИК банка организации>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Адрес по прописке>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Адрес по прописке>',
                                                               MAT_RESP_INPUT_WORDS['<Адрес по прописке>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Дата рождения>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Дата рождения>',
                                                               MAT_RESP_INPUT_WORDS['<Дата рождения>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Место рождения>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Место рождения>',
                                                               MAT_RESP_INPUT_WORDS['<Место рождения>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Серия документа, удостоверяющего личность>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Серия документа, удостоверяющего личность>',
                            MAT_RESP_INPUT_WORDS['<Серия документа, удостоверяющего личность>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Номер документа, удостоверяющего личность>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Номер документа, удостоверяющего личность>',
                            MAT_RESP_INPUT_WORDS['<Номер документа, удостоверяющего личность>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Орган, выдавший документ, удостоверяющего личность>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Орган, выдавший документ, удостоверяющего личность>',
                            MAT_RESP_INPUT_WORDS['<Орган, выдавший документ, удостоверяющего личность>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Дата выдачи документа, удостоверяющего личность>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Дата выдачи документа, удостоверяющего личность>',
                            MAT_RESP_INPUT_WORDS['<Дата выдачи документа, удостоверяющего личность>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<СНИЛС>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<СНИЛС>',
                                                               MAT_RESP_INPUT_WORDS['<СНИЛС>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Корреспондентский счет организации>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Корреспондентский счет организации>',
                            MAT_RESP_INPUT_WORDS['<Корреспондентский счет организации>']
                        )
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Должность руководителя>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Должность руководителя>',
                                                               MAT_RESP_INPUT_WORDS['<Должность руководителя>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Руководитель>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Руководитель>',
                                                               MAT_RESP_INPUT_WORDS['<Руководитель>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

                    if '<Полное имя работника>' in current_cell_txt:
                        update_text = current_cell_txt.replace('<Полное имя работника>',
                                                               MAT_RESP_INPUT_WORDS['<Полное имя работника>'])
                        current_row.cells[j].text = update_text
                        current_cell_txt = update_text

        for i in range(paragraph_count):
            current_txt = doc.paragraphs[i].text
            print(current_txt)

            if '<Полное наименование>' in current_txt:
                updated_text = current_txt.replace('<Полное наименование>',
                                                   MAT_RESP_INPUT_WORDS['<Полное наименование>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Краткое наименование>' in current_txt:
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   MAT_RESP_INPUT_WORDS['<Краткое наименование>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Должность руководителя>' in current_txt:
                updated_text = current_txt.replace('<Должность руководителя>',
                                                   MAT_RESP_INPUT_WORDS['<Должность руководителя>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Руководитель>' in current_txt:
                updated_text = current_txt.replace('<Руководитель>',
                                                   MAT_RESP_INPUT_WORDS['<Руководитель>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Основание подписи>' in current_txt:
                updated_text = current_txt.replace('<Основание подписи>',
                                                   MAT_RESP_INPUT_WORDS['<Основание подписи>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Гражданство страны>' in current_txt:
                updated_text = current_txt.replace('<Гражданство страны>',
                                                   MAT_RESP_INPUT_WORDS['<Гражданство страны>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Гражданство страны>' in current_txt:
                updated_text = current_txt.replace('<Гражданство страны>',
                                                   MAT_RESP_INPUT_WORDS['<Гражданство страны>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Полное имя работника>' in current_txt:
                updated_text = current_txt.replace('<Полное имя работника>',
                                                   MAT_RESP_INPUT_WORDS['<Полное имя работника>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Должность работника>' in current_txt:
                updated_text = current_txt.replace('<Должность работника>',
                                                   MAT_RESP_INPUT_WORDS['<Должность работника>'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Дата приказа о приеме>' in current_txt:
                updated_text = current_txt.replace('<Дата приказа о приеме>',
                                                   MAT_RESP_INPUT_WORDS['Дата_приказа_о_приеме'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

        doc.save(output_result)
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = mat_resp_sample_former(
        input_sample='input_sample/Договор о полной индивидуальной материальной ответственности.docx',
        output_result='output_result/output_Договор о полной индивидуальной материальной ответственности.docx'
    )
