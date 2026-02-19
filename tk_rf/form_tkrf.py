from docx import Document
from python_docx_replace import docx_replace, docx_get_keys


# def tkrf_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
#     TKRF_INPUT_WORDS: dict = {
#         'Договор': parsed_data_1c['job_info']['ТкНомер'],
#         'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
#         'Полное наименование': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
#         'Должность руководителя': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
#         'Руководитель': parsed_data_1c['job_info']['Организация']['Руководитель'],
#         'Основания подписи': 'Основания_подписи',
#         'Гражданство страны': parsed_data_1c['personal_info']['Страна'],
#         'Полное имя работника': parsed_data_1c['personal_info']['ФИО'],
#         'Подразделение': parsed_data_1c['job_info']['Подразделение'],
#         'Должность работника': parsed_data_1c['job_info']['Должность'],
#         'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'Дата приема': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
#         'Испытательный срок': str(parsed_data_1c['job_info']['ИспытательныйСрок']),
#         'График работы': parsed_data_1c['job_info']['ГрафикРаботы'],
#
#         'Длительность рабочей недели': '',
#
#         'Количество дней ежегодного отпуска': parsed_data_1c['job_info']['Отпуск'],
#         'Оклад': str(parsed_data_1c['job_info']['Оклад']),
#         'Дата выплаты аванса': str(parsed_data_1c['job_info']['АвансПериод']),
#         'Дата выплаты зарплаты': str(parsed_data_1c['job_info']['ЗарплатаПериод']),
#         'Юридический адрес': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
#         'Фактический адрес': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
#         'ИНН организации': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
#         'КПП организации': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
#         'ОГРН организации': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],
#
#         'Расчетный счет организации': parsed_data_1c['job_info']['Организация']['РасСчет'],
#         'БИК банка организации': parsed_data_1c['job_info']['Организация']['БикБанка'],
#         'Корреспондентский счет организации': parsed_data_1c['job_info']['Организация']['КорСчет'],
#
#         'Дата рождения': parsed_data_1c['document_info']['ДатаРождения'],
#         'Место рождения': parsed_data_1c['document_info']['МестоРождения'],
#         'Серия документа, удостоверяющего личность': parsed_data_1c['document_info']['Серия'],
#         'Номер документа, удостоверяющего личность': parsed_data_1c['document_info']['Номер'],
#         'Орган, выдавший документ, удостоверяющего личность': parsed_data_1c['document_info']['КемВыдан'],
#         'Дата выдачи документа, удостоверяющего личность': parsed_data_1c['document_info']['ДатаВыдачи'],
#         'СНИЛС': parsed_data_1c['personal_info']['СНИЛС']
#     }
#
#     try:
#         doc = Document(input_sample)
#
#         paragraph_count: int = len(doc.paragraphs)
#         table_count: int = len(doc.tables)
#
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
#                                                                TKRF_INPUT_WORDS['Дата_приказа_о_приеме'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Краткое наименование>',
#                                                                TKRF_INPUT_WORDS['Краткое_наименование'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Юридический адрес>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Юридический адрес>',
#                                                                TKRF_INPUT_WORDS['<Юридический адрес>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Фактический адрес>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Фактический адрес>',
#                                                                TKRF_INPUT_WORDS['<Фактический адрес>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<ИНН организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<ИНН организации>',
#                                                                TKRF_INPUT_WORDS['<ИНН организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<КПП организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<КПП организации>',
#                                                                TKRF_INPUT_WORDS['<КПП организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<ОГРН организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<ОГРН организации>',
#                                                                TKRF_INPUT_WORDS['<ОГРН организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Расчетный счет организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Расчетный счет организации>',
#                                                                TKRF_INPUT_WORDS['<Расчетный счет организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<БИК банка организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<БИК банка организации>',
#                                                                TKRF_INPUT_WORDS['<БИК банка организации>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Дата рождения>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Дата рождения>',
#                                                                TKRF_INPUT_WORDS['<Дата рождения>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Место рождения>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Место рождения>',
#                                                                TKRF_INPUT_WORDS['<Место рождения>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Серия документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Серия документа, удостоверяющего личность>',
#                             TKRF_INPUT_WORDS['<Серия документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Номер документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Номер документа, удостоверяющего личность>',
#                             TKRF_INPUT_WORDS['<Номер документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Орган, выдавший документ, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Орган, выдавший документ, удостоверяющего личность>',
#                             TKRF_INPUT_WORDS['<Орган, выдавший документ, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Дата выдачи документа, удостоверяющего личность>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Дата выдачи документа, удостоверяющего личность>',
#                             TKRF_INPUT_WORDS['<Дата выдачи документа, удостоверяющего личность>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<СНИЛС>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<СНИЛС>',
#                                                                TKRF_INPUT_WORDS['<СНИЛС>'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Корреспондентский счет организации>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Корреспондентский счет организации>',
#                             TKRF_INPUT_WORDS['<Корреспондентский счет организации>']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Должность руководителя>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Должность руководителя>',
#                                                                TKRF_INPUT_WORDS['Должность_руководителя'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Руководитель>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Руководитель>',
#                                                                TKRF_INPUT_WORDS['Руководитель'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#                     if '<Полное имя работника>' in current_cell_txt:
#                         update_text = current_cell_txt.replace('<Полное имя работника>',
#                                                                TKRF_INPUT_WORDS['Полное_имя_работника'])
#                         current_row.cells[j].text = update_text
#                         current_cell_txt = update_text
#
#         for i in range(paragraph_count):
#             print(doc.paragraphs[i].text)
#             current_txt = doc.paragraphs[i].text
#
#             if '<Договор №>' in current_txt:
#                 updated_text = current_txt.replace('<Договор №>',
#                                                    TKRF_INPUT_WORDS['Договор'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#
#             if '<Договор №>' in current_txt:
#                 updated_text = current_txt.replace('<Договор №>',
#                                                    TKRF_INPUT_WORDS['Договор'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Полное наименование>' in current_txt:
#                 updated_text = current_txt.replace('<Полное наименование>',
#                                                    TKRF_INPUT_WORDS['Полное_наименование'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Должность руководителя>' in current_txt:
#                 updated_text = current_txt.replace('<Должность руководителя>',
#                                                    TKRF_INPUT_WORDS['Должность_руководителя'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Руководитель>' in current_txt:
#                 updated_text = current_txt.replace('<Руководитель>',
#                                                    TKRF_INPUT_WORDS['Руководитель'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Основания подписи>' in current_txt:
#                 updated_text = current_txt.replace('<Основания подписи>',
#                                                    TKRF_INPUT_WORDS['Основания_подписи'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Гражданство страны>' in current_txt:
#                 updated_text = current_txt.replace('<Гражданство страны>',
#                                                    TKRF_INPUT_WORDS['Гражданство_страны'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Полное имя работника>' in current_txt:
#                 updated_text = current_txt.replace('<Полное имя работника>',
#                                                    TKRF_INPUT_WORDS['Полное_имя_работника'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Подразделение>' in current_txt:
#                 updated_text = current_txt.replace('<Подразделение>',
#                                                    TKRF_INPUT_WORDS['Подразделение'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Должность работника>' in current_txt:
#                 updated_text = current_txt.replace('<Должность работника>',
#                                                    TKRF_INPUT_WORDS['Должность_работника'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Краткое наименование>' in current_txt:
#                 updated_text = current_txt.replace('<Краткое наименование>',
#                                                    TKRF_INPUT_WORDS['Краткое_наименование'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Дата приема>' in current_txt:
#                 updated_text = current_txt.replace('<Дата приема>',
#                                                    TKRF_INPUT_WORDS['Дата_приема'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Испытательный срок>' in current_txt:
#                 updated_text = current_txt.replace('<Испытательный срок>',
#                                                    TKRF_INPUT_WORDS['Испытательный_срок'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<График работы>' in current_txt:
#                 updated_text = current_txt.replace('<График работы>',
#                                                    TKRF_INPUT_WORDS['График_работы'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Длительность рабочей недели>' in current_txt:
#                 updated_text = current_txt.replace('<Длительность рабочей недели>',
#                                                    TKRF_INPUT_WORDS['Длительность_рабочей_недели'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Количество дней ежегодного отпуска>' in current_txt:
#                 updated_text = current_txt.replace('<Количество дней ежегодного отпуска>',
#                                                    TKRF_INPUT_WORDS['Количество_дней_ежегодного_отпуска'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Оклад>' in current_txt:
#                 updated_text = current_txt.replace('<Оклад>',
#                                                    TKRF_INPUT_WORDS['Оклад'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Дата выплаты аванса>' in current_txt:
#                 updated_text = current_txt.replace('<Дата выплаты аванса>',
#                                                    TKRF_INPUT_WORDS['Дата_выплаты_аванса'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Дата выплаты зарплаты>' in current_txt:
#                 updated_text = current_txt.replace('<Дата выплаты зарплаты>',
#                                                    TKRF_INPUT_WORDS['Дата_выплаты_зарплаты'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#         doc.save(output_result)
#         return f'Файл успешно сформирован. Сохранен в {output_result}'
#     except Exception as e:
#         return f'Ошибка во время формирования файла: {e}'


def tkrf_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    tkrf_input_words: dict = {
        'Договор №': parsed_data_1c['job_info']['ТкНомер'],
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Полное наименование': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        'Должность руководителя': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
        'Руководитель': parsed_data_1c['job_info']['Организация']['Руководитель'],
        'Основания подписи': 'Основания_подписи',
        'Гражданство страны': parsed_data_1c['personal_info']['Страна'],
        'Полное имя работника': parsed_data_1c['personal_info']['ФИО'],
        'Подразделение': parsed_data_1c['job_info']['Подразделение'],
        'Должность работника': parsed_data_1c['job_info']['Должность'],
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Дата приема': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Испытательный срок': str(parsed_data_1c['job_info']['ИспытательныйСрок']),
        'График работы': parsed_data_1c['job_info']['ГрафикРаботы'],
        'Длительность рабочей недели': parsed_data_1c['job_info']['ГрафикРаботы'],
        'Количество дней ежегодного отпуска': parsed_data_1c['job_info']['Отпуск'],
        'Оклад': str(parsed_data_1c['job_info']['Оклад']),
        'Дата выплаты аванса': str(parsed_data_1c['job_info']['АвансПериод']),
        'Дата выплаты зарплаты': str(parsed_data_1c['job_info']['ЗарплатаПериод']),
        'Юридический адрес': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        'Фактический адрес': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
        'ИНН организации': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        'КПП организации': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        'ОГРН организации': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],
        'Расчетный счет организации': parsed_data_1c['job_info']['Организация']['РасСчет'],
        'БИК банка организации': parsed_data_1c['job_info']['Организация']['БикБанка'],
        'Корреспондентский счет организации': parsed_data_1c['job_info']['Организация']['КорСчет'],
        'Дата рождения': parsed_data_1c['document_info']['ДатаРождения'],
        'Место рождения': parsed_data_1c['document_info']['МестоРождения'],
        'Серия документа, удостоверяющего личность': parsed_data_1c['document_info']['Серия'],
        'Номер документа, удостоверяющего личность': parsed_data_1c['document_info']['Номер'],
        'Орган, выдавший документ, удостоверяющего личность': parsed_data_1c['document_info']['КемВыдан'],
        'Дата выдачи документа, удостоверяющего личность': parsed_data_1c['document_info']['ДатаВыдачи'],
        'СНИЛС': parsed_data_1c['personal_info']['СНИЛС'],
        'Адрес по прописке': parsed_data_1c['personal_info']['Адрес_прописка']
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **tkrf_input_words)

        doc.save(output_result)
        return output_result

    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result = tkrf_sample_former(input_sample='input_sample/Трудовой договор.docx',
                                output_result='output_result/output_Трудовой договор.docx')

