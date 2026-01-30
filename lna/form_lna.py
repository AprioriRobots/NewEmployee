from docx import Document


LNA_INPUT_WORDS: dict = {
    'Краткое_наименование': '1',
    'Дата_приказа_о_приеме': '2',
    'трудового_распорядка_Краткое_наименование': '3',
    'оплата_труда_Краткое_наименование': '4',
    'расписание_Краткое_наименование': '5',
    'Должность_работника': '6',
    'инструкция_Краткое_наименование': '7',
    'охрана_труда_Краткое_наименование': '8',
    'пожарная_без_Краткое_наименование': '9',
    'ком_тайна_Краткое_наименование': '10',
    'персон_данн_Краткое_наименование': '11',
    'мат_ответств_Краткое_наименование': '12',
    'видеоконтроль_Краткое_наименование': '13',
    'первая_пом_Краткое_наименование': '14'
}


def lna_sample_former(input_sample: str, output_result: str) -> str:
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
                    print(current_cell_txt)

                    if 'трудового распорядка <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['трудового_распорядка_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'оплате труда <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['оплата_труда_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'расписание <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['расписание_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if '<Должность работника> <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Должность работника>',
                            LNA_INPUT_WORDS['Должность_работника']
                        ).replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['инструкция_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'охране труда <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['охрана_труда_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'безопасности <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['пожарная_без_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'коммерческой тайне <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['ком_тайна_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'данных работника и гарантии защиты их <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['персон_данн_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'материальной ответственности <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['мат_ответств_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if '(видеонаблюдения) на территории <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['видеоконтроль_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

                    if 'помощи при несчастном случае <Краткое наименование>' in current_cell_txt:
                        update_text = current_cell_txt.replace(
                            '<Краткое наименование>',
                            LNA_INPUT_WORDS['первая_пом_Краткое_наименование']
                        )
                        current_row.cells[j].text = update_text
                        current_txt = update_text

        for i in range(paragraph_count):
            current_txt = doc.paragraphs[i].text
            print(current_txt)

            # Тут строгое сравнение, так первое '<Краткое наименование>' идет отдельной строкой (без других слов).
            if current_txt == '<Краткое наименование>':
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   LNA_INPUT_WORDS['Краткое_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Дата приказа о приеме>' in current_txt:
                updated_text = current_txt.replace('<Дата приказа о приеме>',
                                                   LNA_INPUT_WORDS['Дата_приказа_о_приеме'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

        doc.save(output_result)
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = lna_sample_former(input_sample='input_sample/Лист ознакомления с ЛНА.docx',
                                    output_result='output_result/output_Лист ознакомления с ЛНА.docx')
    