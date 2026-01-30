from docx import Document


NDA_INPUT_WORDS: dict = {
    'Полное_наименование': '1',
    '(Краткое_наименование)': '2',
    'Краткое_наименование': '3',
    'работник_Краткое_наименование': 'работником 4',
    'Договор': '5',
    'Дата_приказа_о_приеме': '6',
    'работодатель_Краткое_наименование': '7',
    'конфиденциальности_Краткое_наименование': '8'
}


def lna_sample_former(input_sample: str, output_result: str) -> str:
    try:
        doc = Document(input_sample)

        paragraph_count: int = len(doc.paragraphs)
        table_count: int = len(doc.tables)

        for i in range(paragraph_count):
            current_txt = doc.paragraphs[i].text
            print(current_txt)

            if '<Полное наименование>' in current_txt:
                updated_text = current_txt.replace('<Полное наименование>',
                                                   NDA_INPUT_WORDS['Полное_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if current_txt == '(<Краткое наименование>)':
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   NDA_INPUT_WORDS['(Краткое_наименование)'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '\n<Краткое наименование>' in current_txt:
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   NDA_INPUT_WORDS['Краткое_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if 'работником <Краткое наименование>' in current_txt:
                updated_text = current_txt.replace('работником <Краткое наименование>',
                                                   NDA_INPUT_WORDS['работник_Краткое_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Договор №>' in current_txt:
                updated_text = current_txt.replace('<Договор №>',
                                                   NDA_INPUT_WORDS['Договор'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Дата приказа о приеме>' in current_txt:
                updated_text = current_txt.replace('<Дата приказа о приеме>',
                                                   NDA_INPUT_WORDS['Дата_приказа_о_приеме'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '<Краткое наименование> (далее – «Работодатель»)' in current_txt:
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   NDA_INPUT_WORDS['работодатель_Краткое_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

            if '(коммерческой тайне) <Краткое наименование>' in current_txt:
                updated_text = current_txt.replace('<Краткое наименование>',
                                                   NDA_INPUT_WORDS['конфиденциальности_Краткое_наименование'])
                doc.paragraphs[i].text = updated_text
                current_txt = updated_text

        doc.save(output_result)
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    result: str = lna_sample_former(
        input_sample='input_sample/Обязательство о неразглашении коммерческой тайны.docx',
        output_result='output_result/output_Обязательство о неразглашении коммерческой тайны.docx'
    )
