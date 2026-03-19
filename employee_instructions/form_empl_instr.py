import os

from docx import Document
from pymorphy3 import MorphAnalyzer
from python_docx_replace import docx_replace, docx_get_keys


parsed_data_1c: dict = {'document_info': {'ВидДокумента': 'Паспорт гражданина РФ', 'ДатаВыдачи': '26.06.2007', 'ДатаРождения': '12.06.1987', 'КемВыдан': 'Отделением по району Коньково ОУФМС России по гор. Москве в ЮЗАО ', 'КодПодразделения': '770-117', 'МестоРождения': 'Москва', 'Номер': '124984', 'Серия': '45 09', 'СрокДокумента': '01.01.0100'}, 'job_info': {'АвансПериод': 25, 'ГрафикРаботы': 'Пятидневная рабочая неделя', 'ДатаПриемаНаРаботу': '02.02.2026', 'Должность': 'Шеф-кондитер', 'ЗарплатаПериод': 10, 'ИспытательныйСрок': 3, 'Оклад': 41795, 'Организация': {'БикБанка': '044525593', 'ДолжностьРуководителя': 'Генеральный директор', 'ИНН_Орг': '9704045967', 'КПП_Орг': '770401001', 'КорСчет': '30101810200000000593', 'НаименованиеПолное': 'Общество с ограниченной ответственностью "МАРУЛА"', 'НаименованиеСокращенное': 'ООО "МАРУЛА"', 'ОГРН(ИП)': '1217700035805', 'РасСчет': '40702810402390007442 в АО «АЛЬФА-БАНК» г. Москва', 'Руководитель': 'Ильягуев Нувах Урилович', 'Телефон': '74951337214', 'ФактическийАдрес': '121059, Москва г, б-р Украинский, д. 15', 'ЮридическийАдрес': '119019, Город Москва, вн.тер.г. муниципальный округ Арбат, ул Арбат, д. 6/2, помещ. 1/1/4'}, 'ОснованиеПодписи': '', 'Отпуск': '28', 'Подразделение': 'Кухня', 'ТкНомер': '5-тр', 'ФИО': 'Молоканов Константин Владимирович'}, 'personal_info': {'Адрес_проживания': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'Адрес_прописка': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'АдресБанка': 'Ул. 2-я Хуторская, д.38А, стр.26', 'Банк': 'АО "ТБанк"', 'БикБанка': '044525974', 'ДатаРождения': '12.06.1987', 'ИНН': '772878576660', 'ИннБанка': '7710140679', 'КоррСчет': '', 'МестоРождения': '0,Москва,,,', 'НомерСчета': '40817810500040378646', 'РасчетныйСчет': '', 'СНИЛС': '135-203-129 07', 'Страна': 'РОССИЯ', 'Телефон': '79250051297', 'ФИО': 'Молоканов Константин Владимирович'}}

morph = MorphAnalyzer()


# Нужно доработать функцию
def morph_word(word: str, morph_type: str) -> str:
    param_words = word.split()      # если несколько слов, склоняем по отдельности
    backup = []

    param_gent = []
    for w in param_words:

        # Слова в кавычках не склоняем, еще morphy3 неправильно склоняет слово 'ответственностью' в данном контексте
        if "\"" in w or '\'' in w  or '«' in w or '»' in w or 'ответственностью' == w:
            param_gent.append(w)
            continue

        parsed = morph.parse(w)[0]
        inflected = parsed.inflect({morph_type})
        if inflected:
            param_gent.append(inflected.word)
        else:
            param_gent.append(w)

    result_param = ' '.join(param_gent)
    return result_param


def employee_instr_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    morphed_boss_position: str = morph_word(word=parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
                                            morph_type='gent')
    morphed_full_org_name: str = morph_word(word=parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
                                            morph_type='gent')

    employee_instr_input_words: dict = {
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Договор №': parsed_data_1c['job_info']['ТкНомер'],
        'Полное наименование': morphed_full_org_name,
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Должность руководителя': morphed_boss_position,
        'Фамилия': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        'Инициалы имени, отчества': f'{parsed_data_1c['personal_info']['ФИО'].split(' ')[1][0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[2][0]}',
        'Инициалы работника': f'{parsed_data_1c['personal_info']['ФИО'].split(' ')[0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[1][0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[2][0]}',
    }

    # output_results: list = []
    try:
        doc = Document(f'{input_sample}')
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **employee_instr_input_words)

        doc.save(output_result)
        return output_result
    except Exception as e:
        print(f'Ошибка во время формирования файла: {e}')
        return f'Ошибка во время формирования файла: {e}'


def employee_instr_sample_former2(input_sample: str, output_result: str, parsed_data_1c: dict) -> list[str]:
    morphed_boss_position: str = morph_word(word=parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
                                            morph_type='gent')
    morphed_full_org_name: str = morph_word(word=parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
                                            morph_type='gent')

    employee_instr_input_words: dict = {
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Договор №': parsed_data_1c['job_info']['ТкНомер'],
        'Полное наименование': morphed_full_org_name,
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Должность руководителя': morphed_boss_position,
        'Фамилия': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        'Инициалы имени, отчества': f'{parsed_data_1c['personal_info']['ФИО'].split(' ')[1][0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[2][0]}',
        'Инициалы работника': f'{parsed_data_1c['personal_info']['ФИО'].split(' ')[0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[1][0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[2][0]}',
    }
    output_results = []

    for file in os.listdir(input_sample):
        try:
            doc = Document(f'input_samples/{file}')
            keys = docx_get_keys(doc, pattern='<>')
            print(keys)
            docx_replace(doc, '<>', **employee_instr_input_words)

            doc.save(f'{output_result}/output_{file}')
            output_results.append(f'{output_result}/output_{file}')

        except Exception as e:
            print(f'Ошибка во время формирования файла: {e}')
            continue

    return output_results


if __name__ == '__main__':
    result: list[str] = employee_instr_sample_former2(input_sample='input_samples', output_result='output_result',
                                                      parsed_data_1c=parsed_data_1c)
