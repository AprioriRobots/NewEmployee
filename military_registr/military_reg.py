from openpyxl import load_workbook


parsed_data_1c: dict = {'document_info': {'ВидДокумента': 'Паспорт гражданина РФ', 'ДатаВыдачи': '26.06.2007', 'ДатаРождения': '12.06.1987', 'КемВыдан': 'Отделением по району Коньково ОУФМС России по гор. Москве в ЮЗАО ', 'КодПодразделения': '770-117', 'МестоРождения': 'Москва', 'Номер': '124984', 'Серия': '45 09', 'СрокДокумента': '01.01.0100'}, 'job_info': {'АвансПериод': 25, 'ГрафикРаботы': 'Пятидневная рабочая неделя', 'ДатаПриемаНаРаботу': '02.02.2026', 'Должность': 'Шеф-кондитер', 'ЗарплатаПериод': 10, 'ИспытательныйСрок': 3, 'Оклад': 41795, 'Организация': {'БикБанка': '044525593', 'ДолжностьРуководителя': 'Генеральный директор', 'ИНН_Орг': '9704045967', 'КПП_Орг': '770401001', 'КорСчет': '30101810200000000593', 'НаименованиеПолное': 'Общество с ограниченной ответственностью "МАРУЛА"', 'НаименованиеСокращенное': 'ООО "МАРУЛА"', 'ОГРН(ИП)': '1217700035805', 'РасСчет': '40702810402390007442 в АО «АЛЬФА-БАНК» г. Москва', 'Руководитель': 'Ильягуев Нувах Урилович', 'Телефон': '74951337214', 'ФактическийАдрес': '121059, Москва г, б-р Украинский, д. 15', 'ЮридическийАдрес': '119019, Город Москва, вн.тер.г. муниципальный округ Арбат, ул Арбат, д. 6/2, помещ. 1/1/4'}, 'ОснованиеПодписи': '', 'Отпуск': '28', 'Подразделение': 'Кухня', 'ТкНомер': '5-тр', 'ФИО': 'Молоканов Константин Владимирович'}, 'personal_info': {'Адрес_проживания': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'Адрес_прописка': 'РОССИЯ, 117342, Город Москва, вн.тер.г. муниципальный округ Коньково, ул Островитянова, д. 43, к. 1, кв. 113', 'АдресБанка': 'Ул. 2-я Хуторская, д.38А, стр.26', 'Банк': 'АО "ТБанк"', 'БикБанка': '044525974', 'ДатаРождения': '12.06.1987', 'ИНН': '772878576660', 'ИннБанка': '7710140679', 'КоррСчет': '', 'МестоРождения': '0,Москва,,,', 'НомерСчета': '40817810500040378646', 'РасчетныйСчет': '', 'СНИЛС': '135-203-129 07', 'Страна': 'РОССИЯ', 'Телефон': '79250051297', 'ФИО': 'Молоканов Константин Владимирович'}}


def list_msg_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    list_msg_input_words: dict = {
        '<Фамилия>': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        '<Имя>': parsed_data_1c['personal_info']['ФИО'].split(' ')[1],
        '<Отчество>': parsed_data_1c['personal_info']['ФИО'].split(' ')[2],
        '<Подразделение>': parsed_data_1c['job_info']['Подразделение'],
       '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        '<Фактический адрес>': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],

        '<Воинское звание>': '<Воинское звание>',
        '<ВУС>': '<ВУС>',
        '<Вид образования>': '<Вид образования>',

        '<Должность работника>': parsed_data_1c['job_info']['Должность'],

        '<Фамилия родственника>': '<Фамилия родственника>',
        '<Имя родственника>': '<Имя родственника>',
        '<Отчество родственника>': '<Отчество родственника>',
        '<Дата рождения родственника>': '<Дата рождения родственника>',

        '<Адрес места проживания>': parsed_data_1c['personal_info']['Адрес_проживания'],

        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',

        '<Дата приказа о приеме>': parsed_data_1c['job_info']['ДатаПриемаНаРаботу']
    }

    try:
        wb = load_workbook(input_sample)
        ws = wb['Листок сообщения']

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if '<Фамилия>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия>', list_msg_input_words['<Фамилия>'])

                    if '<Имя>' in cell.value:
                        cell.value = cell.value.replace('<Имя>', list_msg_input_words['<Имя>'])

                    if '<Отчество>' in cell.value:
                        cell.value = cell.value.replace('<Отчество>', list_msg_input_words['<Отчество>'])

                    if '<Подразделение>' in cell.value:
                        cell.value = cell.value.replace('<Подразделение>',
                                                        list_msg_input_words['<Подразделение>'])

                    if '<Полное наименование>' in cell.value:
                        cell.value = cell.value.replace('<Полное наименование>',
                                                        list_msg_input_words['<Полное наименование>'])

                    if '<Краткое наименование>' in cell.value:
                        cell.value = cell.value.replace('<Краткое наименование>',
                                                        list_msg_input_words['<Краткое наименование>'])

                    if '<Юридический адрес>' in cell.value:
                        cell.value = cell.value.replace('<Юридический адрес>',
                                                        list_msg_input_words['<Юридический адрес>'])

                    if '<Фактический адрес>' in cell.value:
                        cell.value = cell.value.replace('<Фактический адрес>',
                                                        list_msg_input_words['<Фактический адрес>'])

                    if '<Дата рождения>' in cell.value:
                        cell.value = cell.value.replace('<Дата рождения>',
                                                        list_msg_input_words['<Дата рождения>'])

                    if '<Воинское звание>' in cell.value:
                        cell.value = cell.value.replace('<Воинское звание>',
                                                        list_msg_input_words['<Воинское звание>'])

                    if '<Воинское звание>' in cell.value:
                        cell.value = cell.value.replace('<Воинское звание>',
                                                        list_msg_input_words['<Воинское звание>'])

                    if '<ВУС>' in cell.value:
                        cell.value = cell.value.replace('<ВУС>',
                                                        list_msg_input_words['<ВУС>'])

                    if '<Вид образования>' in cell.value:
                        cell.value = cell.value.replace('<Вид образования>',
                                                        list_msg_input_words['<Вид образования>'])

                    if '<Должность работника>' in cell.value:
                        cell.value = cell.value.replace('<Должность работника>',
                                                        list_msg_input_words['<Должность работника>'])

                    if '<Фамилия родственника>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия родственника>',
                                                        list_msg_input_words['<Фамилия родственника>'])

                    if '<Имя родственника>' in cell.value:
                        cell.value = cell.value.replace('<Имя родственника>',
                                                        list_msg_input_words['<Имя родственника>'])

                    if '<Отчество родственника>' in cell.value:
                        cell.value = cell.value.replace('<Отчество родственника>',
                                                        list_msg_input_words['<Отчество родственника>'])

                    if '<Дата рождения родственника>' in cell.value:
                        cell.value = cell.value.replace('<Дата рождения родственника>',
                                                        list_msg_input_words['<Дата рождения родственника>'])

                    if '<Адрес места проживания>' in cell.value:
                        cell.value = cell.value.replace('<Адрес места проживания>',
                                                        list_msg_input_words['<Адрес места проживания>'])

                    if '<Инициалы имени, отчества ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace(
                            '<Инициалы имени, отчества ответственного за ВУ>',
                            list_msg_input_words['<Инициалы имени, отчества ответственного за ВУ>']
                        )

                    if '<Дата приказа о приеме>' in cell.value:
                        cell.value = cell.value.replace('<Дата приказа о приеме>',
                                                        list_msg_input_words['<Дата приказа о приеме>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def msg_13_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    list_msg_input_words: dict = {
        '<Полное имя работника>': parsed_data_1c['personal_info']['ФИО'],
        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],

        '<Военкомат>': '<Военкомат>',
        '<Адрес военкомата>': '<Адрес военкомата>',
        
        '<Дата приказа о приеме>': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        '<Номер приказа о приеме>': parsed_data_1c['job_info']['ТкНомер'],
        '<Серия документа, удостоверяющего личность>': parsed_data_1c['document_info']['Серия'],
        '<Номер документа, удостоверяющего личность>': parsed_data_1c['document_info']['Номер'],
        '<Орган, выдавший документ, удостоверяющего личность>': parsed_data_1c['document_info']['КемВыдан'],

        '<Код подразделения>': '<Код подразделения>',
        '<Документ воинского учёта>': '<Документ воинского учёта>',
        '<Серия документа воинского учёта>': '<Серия документа воинского учёта>',
        '<Номер документа воинского учёта>': '<Номер документа воинского учёта>',
        '<Отдел, выдавший документ воинского учёта>': '<Отдел, выдавший документ воинского учёта>',
        '<Дата выдачи документа воинского учёта>': '<Дата выдачи документа воинского учёта>',

        '<Должность руководителя>': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
        '<Инициалы имени, отчества руководителя>': f'{parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[1][0]} {parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[2][0]}',
        '<Фамилия руководителя>': parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[0]
    }

    try:
        wb = load_workbook(input_sample)
        ws = wb['Письмо 13']

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if '<Полное имя работника>' in cell.value:
                        cell.value = cell.value.replace('<Полное имя работника>', list_msg_input_words['<Полное имя работника>'])

                    if '<Полное наименование>' in cell.value:
                        cell.value = cell.value.replace('<Полное наименование>', list_msg_input_words['<Полное наименование>'])

                    if '<Краткое наименование>' in cell.value:
                        cell.value = cell.value.replace('<Краткое наименование>', list_msg_input_words['<Краткое наименование>'])

                    if '<Юридический адрес>' in cell.value:
                        cell.value = cell.value.replace('<Юридический адрес>',
                                                        list_msg_input_words['<Юридический адрес>'])

                    if '<ИНН организации>' in cell.value:
                        cell.value = cell.value.replace('<ИНН организации>',
                                                        list_msg_input_words['<ИНН организации>'])

                    if '<КПП организации>' in cell.value:
                        cell.value = cell.value.replace('<КПП организации>',
                                                        list_msg_input_words['<КПП организации>'])

                    if '<ОГРН организации>' in cell.value:
                        cell.value = cell.value.replace('<ОГРН организации>',
                                                        list_msg_input_words['<ОГРН организации>'])

                    if '<Фактический адрес>' in cell.value:
                        cell.value = cell.value.replace('<Фактический адрес>',
                                                        list_msg_input_words['<Фактический адрес>'])

                    if '<Военкомат>' in cell.value:
                        cell.value = cell.value.replace('<Военкомат>',
                                                        list_msg_input_words['<Военкомат>'])

                    if '<Адрес военкомата>' in cell.value:
                        cell.value = cell.value.replace('<Адрес военкомата>',
                                                        list_msg_input_words['<Адрес военкомата>'])

                    if '<Дата приказа о приеме>' in cell.value:
                        cell.value = cell.value.replace('<Дата приказа о приеме>',
                                                        list_msg_input_words['<Дата приказа о приеме>'])

                    if '<Номер приказа о приеме>' in cell.value:
                        cell.value = cell.value.replace('<Номер приказа о приеме>',
                                                        list_msg_input_words['<Номер приказа о приеме>'])

                    if '<Серия документа, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Серия документа, удостоверяющего личность>',
                                                        list_msg_input_words['<Серия документа, удостоверяющего личность>'])

                    if '<Номер документа, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Номер документа, удостоверяющего личность>',
                                                        list_msg_input_words['<Номер документа, удостоверяющего личность>'])

                    if '<Орган, выдавший документ, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Орган, выдавший документ, удостоверяющего личность>',
                                                        list_msg_input_words['<Орган, выдавший документ, удостоверяющего личность>'])

                    if '<Код подразделения>' in cell.value:
                        cell.value = cell.value.replace('<Код подразделения>',
                                                        list_msg_input_words['<Код подразделения>'])

                    if '<Документ воинского учёта>' in cell.value:
                        cell.value = cell.value.replace('<Документ воинского учёта>',
                                                        list_msg_input_words['<Документ воинского учёта>'])

                    if '<Серия документа воинского учёта>' in cell.value:
                        cell.value = cell.value.replace('<Серия документа воинского учёта>',
                                                        list_msg_input_words['<Серия документа воинского учёта>'])

                    if '<Номер документа воинского учёта>' in cell.value:
                        cell.value = cell.value.replace('<Номер документа воинского учёта>',
                                                        list_msg_input_words['<Номер документа воинского учёта>'])

                    if '<Отдел, выдавший документ воинского учёта>' in cell.value:
                        cell.value = cell.value.replace(
                            '<Отдел, выдавший документ воинского учёта>',
                            list_msg_input_words['<Отдел, выдавший документ воинского учёта>']
                        )

                    if '<Дата выдачи документа воинского учёта>' in cell.value:
                        cell.value = cell.value.replace('<Дата выдачи документа воинского учёта>',
                                                        list_msg_input_words['<Дата выдачи документа воинского учёта>'])

                    if '<Должность руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Должность руководителя>',
                                                        list_msg_input_words['<Должность руководителя>'])

                    if '<Инициалы имени, отчества руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества руководителя>',
                                                        list_msg_input_words['<Инициалы имени, отчества руководителя>'])

                    if '<Фамилия руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия руководителя>',
                                                        list_msg_input_words['<Фамилия руководителя>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def application_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict, sheet_name: str) -> str:
    application_input_words: dict = {
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],

        '<Военкомат>': '<Военкомат>',
        '<Адрес военкомата>': '<Адрес военкомата>',

        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],

        '<Должность ответственного за ВУ>': '<Должность ответственного за ВУ>',
        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',
        '<Фамилия ответственного за ВУ>': '<Фамилия ответственного за ВУ>',

        '<Фамилия>': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        '<Имя>': parsed_data_1c['personal_info']['ФИО'].split(' ')[1],
        '<Отчество>': parsed_data_1c['personal_info']['ФИО'].split(' ')[2],

        '<Воинское звание>': '<Воинское звание>',
        '<Категория запаса>': '<Категория запаса>',
        '<Состав (профиль)>': '<Состав (профиль)>',
        '<ВУС>': '<ВУС>',
        '<Годность к воинской службе>': '<Годность к воинской службе>',

        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],

        '<Место рождения>': '<Место рождения>',

        '<Вид образования>': '<Вид образования>',
        '<Учебное заведение>': '<Учебное заведение>',
        '<Квалификация>': '<Квалификация>',
        '<Специальность> ': '<Специальность>',
        '<Период обучения по>': '<Период обучения по>',

        '<Адрес места проживания>': parsed_data_1c['personal_info']['Адрес_проживания'],
        '<Адрес по прописке>':parsed_data_1c['personal_info']['Адрес_прописка'],
        '<Подразделение>': parsed_data_1c['job_info']['Подразделение'],
        '<Должность работника>': parsed_data_1c['job_info']['Должность'],
        '<Инициалы имени, отчества руководителя>': f'{parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[1][0]} {parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[2][0]}',
        '<Фамилия руководителя>': parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[0]
    }
    keys: list[str] = list(application_input_words.keys())

    try:
        wb = load_workbook(input_sample)
        ws = wb[sheet_name]

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in keys:
                        cell.value = cell.value.replace(cell.value, application_input_words[cell.value])

                    # Обрабатываем частные случаи (по типу '<ИНН организации>/<КПП организации>' или
                    # <Инициалы имени, отчества руководителя> <Фамилия руководителя>' если они в одной ячейке)

                    if '<ИНН организации>' in cell.value:
                        cell.value = cell.value.replace('<ИНН организации>',
                                                        application_input_words['<ИНН организации>'])

                    if '<КПП организации>' in cell.value:
                        cell.value = cell.value.replace('<КПП организации>',
                                                        application_input_words['<КПП организации>'])

                    if '<Должность ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Должность ответственного за ВУ>',
                                                        application_input_words['<Должность ответственного за ВУ>'])

                    if '<Инициалы имени, отчества ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества ответственного за ВУ>',
                                                        application_input_words['<Инициалы имени, отчества ответственного за ВУ>'])

                    if '<Квалификация> ' in cell.value:
                        cell.value = cell.value.replace('<Квалификация>',
                                                        application_input_words['<Квалификация>'])

                    if '<Специальность>' in cell.value:
                        cell.value = cell.value.replace('<Специальность>',
                                                        application_input_words['<Специальность>'])

                    if '<Инициалы имени, отчества руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества руководителя>',
                                                        application_input_words['<Инициалы имени, отчества руководителя>'])

                    if '<Фамилия руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия руководителя>',
                                                        application_input_words['<Фамилия руководителя>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def info_changes_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    info_changes_input_words: dict = {
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],

        '<Военкомат>': '<Военкомат>',
        '<Адрес военкомата>': '<Адрес военкомата>',

        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],

        '<Должность ответственного за ВУ>': '<Должность ответственного за ВУ>',
        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',
        '<Фамилия ответственного за ВУ>': '<Фамилия ответственного за ВУ>',

        '<Фамилия>': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        '<Имя>': parsed_data_1c['personal_info']['ФИО'].split(' ')[1],
        '<Отчество>': parsed_data_1c['personal_info']['ФИО'].split(' ')[2],

        '<Воинское звание>': '<Воинское звание>',

        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],
        '<Инициалы имени, отчества руководителя>': f'{parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[1][0]} {parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[2][0]}',
        '<Фамилия руководителя>': parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[0]
    }
    keys: list[str] = list(info_changes_input_words.keys())

    try:
        wb = load_workbook(input_sample)
        ws = wb['Сведения об изменениях']

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in keys:
                        cell.value = cell.value.replace(cell.value, info_changes_input_words[cell.value])

                    # Обрабатываем частные случаи (по типу '<ИНН организации>/<КПП организации>' или
                    # <Инициалы имени, отчества руководителя> <Фамилия руководителя>' если они в одной ячейке)

                    if '<ИНН организации>' in cell.value:
                        cell.value = cell.value.replace('<ИНН организации>',
                                                        info_changes_input_words['<ИНН организации>'])

                    if '<КПП организации>' in cell.value:
                        cell.value = cell.value.replace('<КПП организации>',
                                                        info_changes_input_words['<КПП организации>'])

                    if '<Должность ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Должность ответственного за ВУ>',
                                                        info_changes_input_words['<Должность ответственного за ВУ>'])

                    if '<Инициалы имени, отчества ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества ответственного за ВУ>',
                                                        info_changes_input_words['<Инициалы имени, отчества ответственного за ВУ>'])

                    if '<Инициалы имени, отчества руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества руководителя>',
                                                        info_changes_input_words['<Инициалы имени, отчества руководителя>'])

                    if '<Фамилия руководителя>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия руководителя>',
                                                        info_changes_input_words['<Фамилия руководителя>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def details_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    details_input_words: dict = {
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],

        '<Военкомат>': '<Военкомат>',
        '<Адрес военкомата>': '<Адрес военкомата>',

        '<Полное имя работника>': parsed_data_1c['personal_info']['ФИО'],

        '<Воинское звание>': '<Воинское звание>',

        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],
        '<Серия документа, удостоверяющего личность>': parsed_data_1c['document_info']['Серия'],
        '<Номер документа, удостоверяющего личность>': parsed_data_1c['document_info']['Номер'],
        '<СНИЛС>': parsed_data_1c['personal_info']['СНИЛС'],
        '<Адрес по прописке>': parsed_data_1c['personal_info']['Адрес_прописка'],
        '<Адрес места проживания>': parsed_data_1c['personal_info']['Адрес_проживания'],
        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        '<Фактический адрес>': parsed_data_1c['job_info']['Организация']['ФактическийАдрес'],
        '<Должность работника>': parsed_data_1c['job_info']['Должность'],
        '<Договор №>': parsed_data_1c['job_info']['ТкНомер'],
        '<Дата приказа о приеме>': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        '<Должность руководителя>': parsed_data_1c['job_info']['Организация']['ДолжностьРуководителя'],
        '<Инициалы имени, отчества руководителя>': f'{parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[1][0]} {parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[2][0]}',
        '<Фамилия руководителя>': parsed_data_1c['job_info']['Организация']['Руководитель'].split(' ')[0],

        '<Должность ответственного за ВУ>': '<Должность ответственного за ВУ>',
        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',
        '<Фамилия ответственного за ВУ>': '<Фамилия ответственного за ВУ>'
    }
    keys: list[str] = list(details_input_words.keys())

    try:
        wb = load_workbook(input_sample)
        ws = wb['Сведения']

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in keys:
                        cell.value = cell.value.replace(cell.value, details_input_words[cell.value])

                    # Обрабатываем частные случаи (по типу '<ИНН организации>/<КПП организации>' или
                    # <Инициалы имени, отчества руководителя> <Фамилия руководителя>' если они в одной ячейке)

                    if '<ИНН организации>' in cell.value:
                        cell.value = cell.value.replace('<ИНН организации>',
                                                        details_input_words['<ИНН организации>'])

                    if '<КПП организации>' in cell.value:
                        cell.value = cell.value.replace('<КПП организации>',
                                                        details_input_words['<КПП организации>'])

                    if '<Серия документа, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Серия документа, удостоверяющего личность>',
                                                        details_input_words['<Серия документа, удостоверяющего личность>'])

                    if '<Номер документа, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Номер документа, удостоверяющего личность>',
                                                        details_input_words['<Номер документа, удостоверяющего личность>'])

                    if '<Дата рождения>' in cell.value:
                        cell.value = cell.value.replace('<Дата рождения>',
                                                        details_input_words['<Дата рождения>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def notification_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict, sheet_name: str) -> str:
    notification_input_words: dict = {
        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Краткое наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеСокращенное'],
        '<Юридический адрес>': parsed_data_1c['job_info']['Организация']['ЮридическийАдрес'],
        '<ИНН организации>': parsed_data_1c['job_info']['Организация']['ИНН_Орг'],
        '<КПП организации>': parsed_data_1c['job_info']['Организация']['КПП_Орг'],
        '<ОГРН организации>': parsed_data_1c['job_info']['Организация']['ОГРН(ИП)'],
        '<Полное имя работника>': parsed_data_1c['personal_info']['ФИО'],

        '<Военкомат>': '<Военкомат>',
        '<Адрес военкомата>': '<Адрес военкомата>',
        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',
        '<Фамилия ответственного за ВУ>': '<Фамилия ответственного за ВУ>'
    }

    keys: list[str] = list(notification_input_words.keys())

    try:
        wb = load_workbook(input_sample)
        ws = wb[sheet_name]

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in keys:
                        cell.value = cell.value.replace(cell.value, notification_input_words[cell.value])

                    # Обрабатываем частные случаи (по типу '<ИНН организации>/<КПП организации>' или
                    # <Инициалы имени, отчества руководителя> <Фамилия руководителя>' если они в одной ячейке)

                    if '<ИНН организации>' in cell.value:
                        cell.value = cell.value.replace('<ИНН организации>',
                                                        notification_input_words['<ИНН организации>'])

                    if '<КПП организации>' in cell.value:
                        cell.value = cell.value.replace('<КПП организации>',
                                                        notification_input_words['<КПП организации>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def f10_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    f10_input_words: dict = {
        '<Полное наименование>': parsed_data_1c['job_info']['Организация']['НаименованиеПолное'],
        '<Инициалы имени, отчества ответственного за ВУ>': '<Инициалы имени, отчества ответственного за ВУ>',
        '<Фамилия ответственного за ВУ>': '<Фамилия ответственного за ВУ>',
        '<Фамилия>': parsed_data_1c['personal_info']['ФИО'].split(' ')[0],
        '<Имя>': parsed_data_1c['personal_info']['ФИО'].split(' ')[1],
        '<Отчество>': parsed_data_1c['personal_info']['ФИО'].split(' ')[2],
        '<Дата рождения>': parsed_data_1c['document_info']['ДатаРождения'],

        '<Место рождения>': '<Место рождения>',
        '<Вид образования>': '<Вид образования>',
        '<Учебное заведение>': '<Учебное заведение>',
        '<Квалификация>': '<Квалификация>',
        '<Специальность> ': '<Специальность>',
        '<Серия квалификационного документа>': '<Серия квалификационного документа>',
        '<Номер квалификационного документа>': '<Номер квалификационного документа>',
        '<Период обучения по>': '<Период обучения по>',

        '<Степень родства>': '<Степень родства>',
        '<Фамилия родственника>': '<Фамилия родственника>',
        '<Имя родственника>': '<Имя родственника>',
        '<Отчество родственника>': '<Отчество родственника>',
        '<Дата рождения родственника>': '<Дата рождения родственника>',

        '<Серия документа, удостоверяющего личность>': parsed_data_1c['document_info']['Серия'],
        '<Номер документа, удостоверяющего личность>': parsed_data_1c['document_info']['Номер'],
        '<Дата выдачи документа, удостоверяющего личность>': parsed_data_1c['document_info']['ДатаВыдачи'],
        '<Орган, выдавший документ, удостоверяющего личность>': parsed_data_1c['document_info']['КемВыдан'],

        '<Код подразделения>': '<Код подразделения>',

        '<Адрес по прописке>': parsed_data_1c['personal_info']['Адрес_прописка'],

        '<Дата регистрации>': '<Дата регистрации>',

        '<Адрес места проживания>': parsed_data_1c['personal_info']['Адрес_проживания'],
        '<Мобильный телефон>': parsed_data_1c['personal_info']['Телефон'],

        '<Категория запаса>': '<Категория запаса>',
        '<Военкомат>': '<Военкомат>',
        '<Воинское звание>': '<Воинское звание>',
        '<Состав (профиль)>': '<Состав (профиль)>',
        '<ВУС>': '<ВУС>',
        '<Годность к воинской службе>': '<Годность к воинской службе>',
        '<Документ воинского учёта>': '<Документ воинского учёта>',
        '<Серия документа воинского учёта>': '<Серия документа воинского учёта>',
        '<Номер документа воинского учёта>': '<Номер документа воинского учёта>',
        '<Отдел, выдавший документ воинского учёта>': '<Отдел, выдавший документ воинского учёта>',
        '<Дата выдачи документа воинского учёта>': '<Дата выдачи документа воинского учёта>',

        '<Номер приказа о приеме>': parsed_data_1c['job_info']['ТкНомер'],
        '<Дата приказа о приеме>': parsed_data_1c['job_info']['ДатаПриемаНаРаботу'],
        '<Должность работника>': parsed_data_1c['job_info']['Должность'],

        '<Должность ответственного за ВУ>': '<Должность ответственного за ВУ>',

        '<Инициалы имени, отчества>': f'{parsed_data_1c['personal_info']['ФИО'].split(' ')[1][0]} {parsed_data_1c['personal_info']['ФИО'].split(' ')[2][0]}'
    }

    keys: list[str] = list(f10_input_words.keys())

    try:
        wb = load_workbook(input_sample)
        ws = wb['Карточка']

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in keys:
                        cell.value = cell.value.replace(cell.value, f10_input_words[cell.value])

                    # Обрабатываем частные случаи (по типу '<ИНН организации>/<КПП организации>' или
                    # <Инициалы имени, отчества руководителя> <Фамилия руководителя>' если они в одной ячейке)

                    if '<Инициалы имени, отчества ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Инициалы имени, отчества ответственного за ВУ>',
                                                        f10_input_words['<Инициалы имени, отчества ответственного за ВУ>'])

                    if '<Фамилия ответственного за ВУ>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия ответственного за ВУ>',
                                                        f10_input_words['<Фамилия ответственного за ВУ>'])

                    if '<Фамилия родственника>' in cell.value:
                        cell.value = cell.value.replace('<Фамилия родственника>',
                                                        f10_input_words['<Фамилия родственника>'])

                    if '<Имя родственника>' in cell.value:
                        cell.value = cell.value.replace('<Имя родственника>',
                                                        f10_input_words['<Имя родственника>'])

                    if '<Отчество родственника>' in cell.value:
                        cell.value = cell.value.replace('<Отчество родственника>',
                                                        f10_input_words['<Отчество родственника>'])

                    if '<Дата рождения родственника>' in cell.value:
                        cell.value = cell.value.replace('<Дата рождения родственника>',
                                                        f10_input_words['<Дата рождения родственника>'])

                    if '<Дата рождения>' in cell.value:
                        cell.value = cell.value.replace('<Дата рождения>',
                                                        f10_input_words['<Дата рождения>'])

                    if '<Дата выдачи документа, удостоверяющего личность>' in cell.value:
                        cell.value = cell.value.replace('<Дата выдачи документа, удостоверяющего личность>',
                                                        f10_input_words['<Дата выдачи документа, удостоверяющего личность>'])

        wb.save(output_result)
        return output_result
    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'



if __name__ == '__main__':
    # list_msg: str = list_msg_sample_former(input_sample='input_samples/Листок-сообщения.xlsx',
    #                                        output_result='output_result/output_Листок-сообщения.xlsx',
    #                                        parsed_data_1c=parsed_data_1c)

    # msg_13: str = msg_13_sample_former(input_sample='input_samples/Письмо 13.xlsx',
    #                                    output_result='output_result/output_Письмо 13.xlsx',
    #                                    parsed_data_1c=parsed_data_1c)

    # appl_10: str = application_sample_former(input_sample='input_samples/Приложение 10.xlsx',
    #                                          output_result='output_result/output_Приложение 10.xlsx',
    #                                          parsed_data_1c=parsed_data_1c,
    #                                          sheet_name='Сведения прил. 10')
    #
    # appl_23: str = application_sample_former(input_sample='input_samples/Приложение 23.xlsx',
    #                                          output_result='output_result/output_Приложение 23.xlsx',
    #                                          parsed_data_1c=parsed_data_1c,
    #                                          sheet_name='Сверка')

    # info_change: str = info_changes_sample_former(input_sample='input_samples/Сведения об изменениях.xlsx',
    #                                               output_result='output_result/output_Сведения об изменениях.xlsx',
    #                                               parsed_data_1c=parsed_data_1c)

    # details: str = details_sample_former(input_sample='input_samples/Сведения.xlsx',
    #                                      output_result='output_result/output_Сведения.xlsx',
    #                                      parsed_data_1c=parsed_data_1c)

    # note_1: str = notification_sample_former(input_sample='input_samples/Уведомление 1.xlsx',
    #                                          output_result='output_result/output_Уведомление 1.xlsx',
    #                                          parsed_data_1c=parsed_data_1c, sheet_name='Уведомление 1')
    #
    # note_2: str = notification_sample_former(input_sample='input_samples/Уведомление 2.xlsx',
    #                                          output_result='output_result/output_Уведомление 2.xlsx',
    #                                          parsed_data_1c=parsed_data_1c, sheet_name='Уведомление 2')

    f10_doc: str = f10_sample_former(input_sample='input_samples/Ф10.xlsx',
                                     output_result='output_result/output_Ф10.xlsx',
                                     parsed_data_1c=parsed_data_1c)
    pass