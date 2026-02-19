from docx import Document
from python_docx_replace import docx_replace, docx_get_keys


# def lna_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
#     LNA_INPUT_WORDS: dict = {
#         'Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'Дата_приказа_о_приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
#         'трудового_распорядка_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'оплата_труда_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'расписание_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'Должность_работника': parsed_data_1c['job_info']['Должность'],
#         'инструкция_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'охрана_труда_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'пожарная_без_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'ком_тайна_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'персон_данн_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'мат_ответств_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'видеоконтроль_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
#         'первая_пом_Краткое_наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное']
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
#                     print(current_cell_txt)
#
#                     if 'трудового распорядка <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['трудового_распорядка_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'оплате труда <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['оплата_труда_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'расписание <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['расписание_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if '<Должность работника> <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Должность работника>',
#                             LNA_INPUT_WORDS['Должность_работника']
#                         ).replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['инструкция_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'охране труда <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['охрана_труда_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'безопасности <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['пожарная_без_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'коммерческой тайне <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['ком_тайна_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'данных работника и гарантии защиты их <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['персон_данн_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'материальной ответственности <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['мат_ответств_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if '(видеонаблюдения) на территории <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['видеоконтроль_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#                     if 'помощи при несчастном случае <Краткое наименование>' in current_cell_txt:
#                         update_text = current_cell_txt.replace(
#                             '<Краткое наименование>',
#                             LNA_INPUT_WORDS['первая_пом_Краткое_наименование']
#                         )
#                         current_row.cells[j].text = update_text
#                         current_txt = update_text
#
#         for i in range(paragraph_count):
#             current_txt = doc.paragraphs[i].text
#             print(current_txt)
#
#             # Тут строгое сравнение, так первое '<Краткое наименование>' идет отдельной строкой (без других слов).
#             if current_txt == '<Краткое наименование>':
#                 updated_text = current_txt.replace('<Краткое наименование>',
#                                                    LNA_INPUT_WORDS['Краткое_наименование'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#             if '<Дата приказа о приеме>' in current_txt:
#                 updated_text = current_txt.replace('<Дата приказа о приеме>',
#                                                    LNA_INPUT_WORDS['Дата_приказа_о_приеме'])
#                 doc.paragraphs[i].text = updated_text
#                 current_txt = updated_text
#
#         doc.save(output_result)
#     except Exception as e:
#         return f'Ошибка во время формирования файла: {e}'



def lna_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    lna_input_words: dict = {
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Должность работника': parsed_data_1c['job_info']['Должность'],
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **lna_input_words)

        doc.save(output_result)
        return output_result

    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = lna_sample_former(input_sample='input_sample/Лист ознакомления с ЛНА.docx',
                                    output_result='output_result/output_Лист ознакомления с ЛНА.docx')
    