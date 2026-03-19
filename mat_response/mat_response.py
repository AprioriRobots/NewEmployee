from docx import Document
from pymorphy3 import MorphAnalyzer
from python_docx_replace import docx_replace, docx_get_keys


parsed_data_1c: dict = {'document_info': {'ВидДокумента': 'Паспорт гражданина РФ', 'ДатаВыдачи': '26.06.2007', 'ДатаРождения': '12.06.1987', 'КемВыдан': 'Отделением по району Коньково ОУФМС России по гор. Москве в ЮЗАО ', 'КодПодразделения': '770-117', 'МестоРождения': 'Москва', 'Номер': '124984', 'Серия': '45 09', 'СрокДокумента': '01.01.0100'}, 'job_info': {'АвансПериод': 25, 'ГрафикРаботы': 'Пятидневная рабочая неделя', 'ДатаПриемаНаРаботу': '02.02.2026', 'Должность': 'Шеф-кондитер', 'ЗарплатаПериод': 10, 'ИспытательныйСрок': 3, 'Оклад': 41795, 'Организация': {'БикБанка': '044525593', 'ДолжностьРуководителя': 'Генеральный директор', 'ИНН_Орг': '9704045967', 'КПП_Орг': '770401001', 'КорСчет': '30101810200000000593', 'НаименованиеПолное': 'Общество с ограниченной ответственностью "МАРУЛА"', 'НаименованиеСокращенное': 'ООО "МАРУЛА"', 'ОГРН(ИП)': '1217700035805', 'РасСчет': '40702810402390007442 в АО «АЛЬФА-БАНК» г. Москва', 'Руководитель': 'Ильягуев Нувах Урилович', 'Телефон': '74951337214', 'ФактическийАдрес': '121059, Москва г, б-р Украинский, д. 15', 'ЮридическийАдрес': '119019, Город Москва, вн.тер.г. муниципальный округ Арбат, ул Арбат, д. 6/2, помещ. 1/1/4'}, 'ОснованиеПодписи': '', 'Отпуск': '28', 'Подразделение': 'Кухня', 'ТкНомер': '5-тр', 'ФИО': 'Молоканов Константин Владимирович'}, 'personal_info': {'Адрес_проживания': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'Адрес_прописка': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'АдресБанка': 'Ул. 2-я Хуторская, д.38А, стр.26', 'Банк': 'АО "ТБанк"', 'БикБанка': '044525974', 'ДатаРождения': '12.06.1987', 'ИНН': '772878576660', 'ИннБанка': '7710140679', 'КоррСчет': '', 'МестоРождения': '0,Москва,,,', 'НомерСчета': '40817810500040378646', 'РасчетныйСчет': '', 'СНИЛС': '135-203-129 07', 'Страна': 'РОССИЯ', 'Телефон': '79250051297', 'ФИО': 'Молоканов Константин Владимирович'}}
CASES: list = ['nomn', 'gent', 'datv', 'accs', 'ablt', 'loct']

morph = MorphAnalyzer()

# Нужно доработать функцию
def morph_word(word: str, morph_type: str) -> str:

    def match_case(original: str, new_word: str) -> str:
        if original.isupper():
            return new_word.upper()
        elif original.istitle():
            return new_word.capitalize()
        elif original.islower():
            return new_word.lower()
        else:
            # Смешанный регистр оставим без изменений
            return new_word

    param_words = word.split()
    param_gent = []

    if 'шеф-' in word.lower():
        param_words = word.split('-')

    for w in param_words:
        if "\"" in w or '\'' in w or '«' in w or '»' in w or w.lower() == 'ответственностью':
            param_gent.append(w)
            continue

        if ('шеф' == w.lower() and len(param_words)) > 1:
            param_gent.append(w)
            continue

        parsed = morph.parse(w)[0]
        inflected = parsed.inflect({morph_type})
        if inflected:
            inflected_word = match_case(w, inflected.word)
            param_gent.append(inflected_word)
        else:
            param_gent.append(w)

    if 'шеф-' in word.lower():
        return '-'.join(param_gent)

    return ' '.join(param_gent)


def mat_resp_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    morphed_boss_position: str = morph_word(word=parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
                                            morph_type=CASES[1])
    morphed_full_org_name: str = morph_word(word=parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
                                            morph_type=CASES[2])

    morphed_boss_name: str = morph_word(word=parsed_data_1c['job_info']['Организация']['Руководитель'],
                                        morph_type=CASES[1])

    morphed_country: str = morph_word(word=parsed_data_1c['personal_info']['Страна'],
                                      morph_type=CASES[1])
    morphed_job: str = morph_word(word=parsed_data_1c['job_info']['Должность'], morph_type=CASES[1])

    mat_resp_input_words: dict = {
        'Дата приказа о приеме': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        'Полное наименование': morphed_full_org_name,
        'Краткое наименование': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        'Должность руководителя': morphed_boss_position,
        'Руководитель': parsed_data_1c['job_info']['Организация']['Руководитель'],
        'Основание подписи': 'Основание подписи',

        'Гражданство страны': morphed_country,
        'Полное имя работника': parsed_data_1c['personal_info']['ФИО'],
        'Должность работника': morphed_job,
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
        output_result='output_result/output_Договор о полной индивидуальной материальной ответственности.docx',
        parsed_data_1c=parsed_data_1c
    )
