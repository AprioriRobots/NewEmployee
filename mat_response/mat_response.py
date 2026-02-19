from docx import Document
from python_docx_replace import docx_replace, docx_get_keys


# def mat_resp_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
#     MAT_RESP_INPUT_WORDS: dict = {
#         'Дата_приказа_о_приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
#         '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
#         '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         '<Должность руководителя>': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
#         '<Руководитель>': parsed_data_1c['job_info']['Организация']['Руководитель'],
#         '<Основание подписи>': '<Основание подписи>',
#         '<Гражданство страны>': parsed_data_1c['personal_info']['Страна'],
#         '<Полное имя работника>': parsed_data_1c['personal_info']['ФИО'],
#         '<Должность работника>': parsed_data_1c['job_info']['Должность'],
#         '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
#         '<Фактический адрес>': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
#         '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
#         '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
#         '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],
#
#         '<Расчетный счет организации>': parsed_data_1c['job_info']['Организация']['РасСчет'],
#         '<БИК банка организации>': parsed_data_1c['job_info']['Организация']['БикБанка'],
#
#         '<Адрес по прописке>': parsed_data_1c['personal_info']['Адрес_прописка'],
#         '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],
#         '<Место рождения>': parsed_data_1c['document_info']['МестоРождения'],
#         '<Серия документа, удостоверяющего личность>': parsed_data_1c['document_info']['Серия'],
#         '<Номер документа, удостоверяющего личность>': parsed_data_1c['document_info']['Номер'],
#         '<Орган, выдавший документ, удостоверяющего личность>': parsed_data_1c['document_info']['КемВыдан'],
#         '<Дата выдачи документа, удостоверяющего личность>': parsed_data_1c['document_info']['ДатаВыдачи'],
#         '<СНИЛС>': parsed_data_1c['personal_info']['СНИЛС'],
#
#         '<Корреспондентский счет организации>': parsed_data_1c['job_info']['Организация']['КорСчет'],
#     }
#
#     try:
#         doc = Document(input_sample)
#
#         paragraph_count: int = len(doc.paragraphs)
#         table_count: int = len(doc.tables)
#
#         # Итерируемся по всем таблицам в документе
#         for i in range(table_count):
#             # Итерируемся по каждой строке
#             for row in doc.tables[i].table.rows:
#                 current_row = row
#                 cell_count_of_current_row: int = len(current_row.cells)
#
#                 # Итерируемся по каждой ячейке
#                 for j in range(cell_count_of_current_row):
#                     current_cell_txt = current_row.cells[j].text
#
#                     if '<Дата приказа о приеме>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Дата приказа о приеме>',
#                                                                MAT_RESP_INPUT_WORDS['Дата_приказа_о_приеме'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Краткое наименование>',
#                                                                MAT_RESP_INPUT_WORDS['<Краткое наименование>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Юридический адрес>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Юридический адрес>',
#                                                                MAT_RESP_INPUT_WORDS['<Юридический адрес>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Фактический адрес>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Фактический адрес>',
#                                                                MAT_RESP_INPUT_WORDS['<Фактический адрес>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<ИНН организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<ИНН организации>',
#                                                                MAT_RESP_INPUT_WORDS['<ИНН организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<КПП организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<КПП организации>',
#                                                                MAT_RESP_INPUT_WORDS['<КПП организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<ОГРН организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<ОГРН организации>',
#                                                                MAT_RESP_INPUT_WORDS['<ОГРН организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Расчетный счет организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Расчетный счет организации>',
#                                                                MAT_RESP_INPUT_WORDS['<Расчетный счет организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<БИК банка организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<БИК банка организации>',
#                                                                MAT_RESP_INPUT_WORDS['<БИК банка организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Адрес по прописке>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Адрес по прописке>',
#                                                                MAT_RESP_INPUT_WORDS['<Адрес по прописке>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Дата рождения>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Дата рождения>',
#                                                                MAT_RESP_INPUT_WORDS['<Дата рождения>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Место рождения>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Место рождения>',
#                                                                MAT_RESP_INPUT_WORDS['<Место рождения>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Серия документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Серия документа, удостоверяющего личность>',
#                             MAT_RESP_INPUT_WORDS['<Серия документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Номер документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Номер документа, удостоверяющего личность>',
#                             MAT_RESP_INPUT_WORDS['<Номер документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Орган, выдавший документ, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Орган, выдавший документ, удостоверяющего личность>',
#                             MAT_RESP_INPUT_WORDS['<Орган, выдавший документ, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Дата выдачи документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Дата выдачи документа, удостоверяющего личность>',
#                             MAT_RESP_INPUT_WORDS['<Дата выдачи документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<СНИЛС>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<СНИЛС>',
#                                                                MAT_RESP_INPUT_WORDS['<СНИЛС>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Корреспондентский счет организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Корреспондентский счет организации>',
#                             MAT_RESP_INPUT_WORDS['<Корреспондентский счет организации>']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Должность руководителя>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Должность руководителя>',
#                                                                MAT_RESP_INPUT_WORDS['<Должность руководителя>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Руководитель>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Руководитель>',
#                                                                MAT_RESP_INPUT_WORDS['<Руководитель>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Полное имя работника>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Полное имя работника>',
#                                                                MAT_RESP_INPUT_WORDS['<Полное имя работника>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#         for i in range(paragraph_count):
#             current_txt = doc.paragraphs[i].text
#             print(current_txt)
#
#             if '<Полное наименование>' in current_txt:
#                 updated_text = current_txt.replace('<Полное наименование>',
#                                                    MAT_RESP_INPUT_WORDS['<Полное наименование>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Краткое наименование>' in current_txt:
#                 updated_text = current_txt.replace('<Краткое наименование>',
#                                                    MAT_RESP_INPUT_WORDS['<Краткое наименование>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Должность руководителя>' in current_txt:
#                 updated_text = current_txt.replace('<Должность руководителя>',
#                                                    MAT_RESP_INPUT_WORDS['<Должность руководителя>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Руководитель>' in current_txt:
#                 updated_text = current_txt.replace('<Руководитель>',
#                                                    MAT_RESP_INPUT_WORDS['<Руководитель>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Основание подписи>' in current_txt:
#                 updated_text = current_txt.replace('<Основание подписи>',
#                                                    MAT_RESP_INPUT_WORDS['<Основание подписи>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Гражданство страны>' in current_txt:
#                 updated_text = current_txt.replace('<Гражданство страны>',
#                                                    MAT_RESP_INPUT_WORDS['<Гражданство страны>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Гражданство страны>' in current_txt:
#                 updated_text = current_txt.replace('<Гражданство страны>',
#                                                    MAT_RESP_INPUT_WORDS['<Гражданство страны>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Полное имя работника>' in current_txt:
#                 updated_text = current_txt.replace('<Полное имя работника>',
#                                                    MAT_RESP_INPUT_WORDS['<Полное имя работника>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Должность работника>' in current_txt:
#                 updated_text = current_txt.replace('<Должность работника>',
#                                                    MAT_RESP_INPUT_WORDS['<Должность работника>'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Дата приказа о приеме>' in current_txt:
#                 updated_text = current_txt.replace('<Дата приказа о приеме>',
#                                                    MAT_RESP_INPUT_WORDS['Дата_приказа_о_приеме'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#         doc.save(output_result)
#         return output_result
#     except Exception as e:
#         return f'Ошибка во время формирования файла: {e}'


def mat_resp_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    mat_resp_input_words: dict = {
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Полное наименование': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Должность руководителя': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
        'Руководитель': parsed_data_1c['job_info']['Организация']['Руководитель'],
        'Основание подписи': 'Основание подписи',

        'Гражданство страны': parsed_data_1c['personal_info']['Страна'],
        'Полное имя работника': parsed_data_1c['personal_info']['ФИО'],
        'Должность работника': parsed_data_1c['job_info']['Должность'],
        'Юридический адрес': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        'Фактический адрес': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
        'ИНН организации': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        'КПП организации': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        'ОГРН организации': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],

        'Расчетный счет организации': parsed_data_1c['job_info']['Организация']['РасСчет'],
        'БИК банка организации': parsed_data_1c['job_info']['Организация']['БикБанка'],

        'Адрес по прописке': parsed_data_1c['personal_info']['Адрес_прописка'],
        'Дата рождения': parsed_data_1c['document_info']['ДатаРождения'],
        'Место рождения': parsed_data_1c['document_info']['МестоРождения'],
        'Серия документа, удостоверяющего личность': parsed_data_1c['document_info']['Серия'],
        'Номер документа, удостоверяющего личность': parsed_data_1c['document_info']['Номер'],
        'Орган, выдавший документ, удостоверяющего личность': parsed_data_1c['document_info']['КемВыдан'],
        'Дата выдачи документа, удостоверяющего личность': parsed_data_1c['document_info']['ДатаВыдачи'],
        'СНИЛС': parsed_data_1c['personal_info']['СНИЛС'],

        'Корреспондентский счет организации': parsed_data_1c['job_info']['Организация']['КорСчет']
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **mat_resp_input_words)

        doc.save(output_result)
        return output_result

    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'



if __name__ == '__main__':
    result: str = mat_resp_sample_former(
        input_sample='input_sample/Договор о полной индивидуальной материальной ответственности.docx',
        output_result='output_result/output_Договор о полной индивидуальной материальной ответственности.docx'
    )
