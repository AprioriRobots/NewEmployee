import win32com.client
import datetime


def personal_info(connector: win32com.client, fio: str) -> dict:
    parsed_pers_info = {
        'ФИО': '',
        'ДатаРождения': '',
        'ИНН': '',
        'МестоРождения': '',
        'СНИЛС': '',
        'Адрес_прописка': '',
        'Адрес_проживания': '',
    }

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ
            Сотрудники.ФизическоеЛицо.ФИО КАК ФИО,
             Сотрудники.ФизическоеЛицо.ДатаРождения КАК ДатаРождения,
             Сотрудники.ФизическоеЛицо.ИНН КАК ИНН,
             Сотрудники.ФизическоеЛицо.МестоРождения КАК МестоРождения,
             Сотрудники.ФизическоеЛицо.СтраховойНомерПФР КАК СНИЛС
        ИЗ
            Справочник.Сотрудники КАК Сотрудники
        ГДЕ
            Сотрудники.ФизическоеЛицо.ФИО = &ФИО
        """

    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()

    while selection.Next():
        birth_date = selection.ДатаРождения
        formatted_date = birth_date.strftime('%d.%m.%Y')

        # print(
        #     f"ФИО: {selection.ФИО}, Дата Рождения: {selection.ДатаРождения},"
        #     f" ИНН: {selection.ИНН}, МестоРождения: {selection.МестоРождения},"
        #     f" СНИЛС: {selection.СНИЛС}"
        # )

        parsed_pers_info['ФИО'] = selection.ФИО
        parsed_pers_info['ДатаРождения'] = formatted_date
        parsed_pers_info['ИНН'] = selection.ИНН
        parsed_pers_info['МестоРождения'] = selection.МестоРождения
        parsed_pers_info['СНИЛС'] = selection.СНИЛС


    query_2 = connector.NewObject("Запрос")
    query_2.Text = """
                       ВЫБРАТЬ
                           ГражданствоФизическихЛиц.ФизическоеЛицо как ФизЛицо
                       ИЗ
                           РегистрСведений.ГражданствоФизическихЛиц КАК ГражданствоФизическихЛиц
                       ГДЕ
                           ГражданствоФизическихЛиц.ФизическоеЛицо.ФИО = &ФИО
                       """

    query_2.SetParameter("ФИО", fio)
    selection_2 = query_2.Execute().Choose()

    reg_address, fact_address = '', ''

    while selection_2.Next():
        print(f'Контактная информация: {selection_2.ФизЛицо.КонтактнаяИнформация}')

        if len(selection_2.ФизЛицо.КонтактнаяИнформация):
            row_counter: int = 0

            for row in selection_2.ФизЛицо.КонтактнаяИнформация:
                # reg_address = row[1][3]       # Индексация не работает у КОМ-Объектов :(
                # fact_address = row[2][3]      # Индексация не работает у КОМ-Объектов :(

                el_counter: int = 0

                for el in row:
                    if row_counter == 1 and el_counter == 3:
                        reg_address = el

                    if row_counter == 2 and el_counter == 3:
                        fact_address = el

                    el_counter += 1
                row_counter += 1

    parsed_pers_info['Адрес_прописка'] = reg_address
    parsed_pers_info['Адрес_проживания'] = fact_address

    return parsed_pers_info


def job_info(connector: win32com.client, fio: str) -> dict:
    parsed_job_info = {
        'ФИО': '',
        'ТкНомер': '',                      # Есть
        'ОснованиеПодписи': '',

        'Организация': {
            'Руководитель': '',             # Есть
            'ДолжностьРуководителя': '',    # Есть
            'НаименованиеПолное': '',       # Есть
            'НаименованиеСокращенное': '',  # Есть
            'ИНН_Орг': '',                  # Есть
            'ОГРН(ИП)': '',                 # Есть
            'КПП_Орг': '',                  # Есть
            'ЮридическийАдрес': '',         # Есть
            'ФактическийАдрес': '',         # Есть

        },

        'Должность': '',                    # Есть
        'Подразделение': '',                # Есть
        'Оклад': '',                        # Есть
        'ИспытательныйСрок': '',            # Есть
        'АвансПериод': '',                  # Нет
        'ГрафикРаботы': '',                 # Есть
        'ДатаПриемаНаРаботу': '',           # Есть
        'Отпуск': '28',                       # (Пока временно поставил 28)
    }

    employee = getattr(connector.Catalogs, "Сотрудники")
    needed_employee = employee.FindByDescription(fio)

    query = connector.NewObject("Запрос")
    query.Text = """
                       ВЫБРАТЬ
                           ПриемНаРаботу.Организация как Организация,
                           ПриемНаРаботу.Руководитель как Руководитель,
                           ПриемНаРаботу.ДолжностьРуководителя как ДолжностьРуководителя,
                           ПриемНаРаботу.ТрудовойДоговорНомер как ТрудовойДоговорНомер,
                           ПриемНаРаботу.ДатаПриема как ДатаПриема,
                           ПриемНаРаботу.Подразделение как Подразделение,
                           ПриемНаРаботу.Должность как Должность,
                           ПриемНаРаботу.ГрафикРаботы как График,
                           ПриемНаРаботу.ДлительностьИспытательногоСрока как ИспытательныйСрок,
                           ПриемНаРаботу.Аванс как Аванс,
                           ПриемНаРаботу.СовокупнаяТарифнаяСтавка как Оклад,
                           ПриемНаРаботу.ЕжегодныеОтпуска.КоличествоДнейВГод как Отпуск
                       ИЗ
                           Документ.ПриемНаРаботу КАК ПриемНаРаботу
                       ГДЕ
                           ПриемНаРаботу.Сотрудник = &ФИО
                       """

    # ПриемНаРаботу.ВидЗанятости как ВидЗанятости,

    query.SetParameter("ФИО", needed_employee)
    selection = query.Execute().Choose()

    while selection.Next():
        asigned_job_date = selection.ДатаПриема
        formatted_asigned_job_date = asigned_job_date.strftime('%d.%m.%Y')

        parsed_job_info['ФИО'] = fio
        parsed_job_info['ТкНомер'] = selection.ТрудовойДоговорНомер

        parsed_job_info['Организация']['НаименованиеПолное'] = selection.Организация.НаименованиеПолное
        parsed_job_info['Организация']['ДолжностьРуководителя'] = selection.Организация.НаименованиеСокращенное
        parsed_job_info['Организация']['Руководитель'] = selection.Руководитель.Наименование
        parsed_job_info['Организация']['ДолжностьРуководителя'] = selection.ДолжностьРуководителя.Наименование
        parsed_job_info['Организация']['ИНН_Орг'] = selection.Организация.ИНН
        parsed_job_info['Организация']['ОГРН(ИП)'] = selection.Организация.ОГРН
        parsed_job_info['Организация']['КПП_Орг'] = selection.Организация.РегистрацияВНалоговомОргане.КПП

        parsed_job_info['Подразделение'] = selection.Подразделение.Наименование
        parsed_job_info['Должность'] = selection.Должность.Наименование
        parsed_job_info['ГрафикРаботы'] = selection.График.Наименование
        parsed_job_info['ДатаПриемаНаРаботу'] = formatted_asigned_job_date

        parsed_job_info['ИспытательныйСрок'] = selection.ИспытательныйСрок
        parsed_job_info['Оклад'] = selection.Оклад

        reg_address, fact_address = '', ''

        if  len(selection.Организация.КонтактнаяИнформация):
                row_counter: int = 0

                for row in selection.Организация.КонтактнаяИнформация:
                    # reg_address = row[0][3]       # Индексация не работает у КОМ-Объектов :(
                    # fact_address = row[1][3]      # Индексация не работает у КОМ-Объектов :(

                    el_counter: int = 0

                    for el in row:
                        if row_counter == 0 and el_counter == 3:
                            reg_address = el

                        if row_counter == 1 and el_counter == 3:
                            fact_address = el

                        el_counter += 1
                    row_counter += 1

        # Пытался по аналогии проитерироваться по selection.Отпуск - возникает ошибка

        parsed_job_info['Организация']['ЮридическийАдрес'] = reg_address
        parsed_job_info['Организация']['ФактическийАдрес'] = fact_address

        # if len(selection.Отпуск):
        #     for row in selection.Отпуск:
        #         print(row)

        print(
            f"ДатаПриема: {selection.ДатаПриема}, Подразделение: {selection.Подразделение.Наименование},"
            f" ОрганизацияНаименованиеПолное: {selection.Организация.НаименованиеПолное},"
            f" ОрганизацияНаименованиеСокращенное: {selection.Организация.НаименованиеСокращенное},"
            
            f" ИНН Организации: {selection.Организация.ИНН}, ОГРН(ИП): {selection.Организация.ОГРН},"
            f" КПП Организации: {selection.Организация.РегистрацияВНалоговомОргане.КПП},"
            
            f" Факт адрес орг: {fact_address}\nЮр. адрес орг: {reg_address}"
            
            f" Должность: {selection.Должность.Наименование}, ГрафикРаботы: {selection.График.Наименование},"
            f" Руководитель: {selection.Руководитель.Наименование}, Должность руководителя: {selection.ДолжностьРуководителя.Наименование}"
            f" ТК Номер: {selection.ТрудовойДоговорНомер}, Аванс: {selection.Аванс},"
            f" ИспытательныйСрок: {selection.ИспытательныйСрок}, Оклад: {selection.Оклад},"
            f" Отпуск: {selection.Отпуск}"
        )
        # f" ОрганизацияНаименованиеСокращенное: {selection.Организация.ОрганизацияНаименованиеСокращенное},"
    return parsed_job_info



def document_info(connector: win32com.client, fio: str) -> dict:
    parsed_doc_info: dict = {
        'ВидДокумента': '',
        'Серия': '',
        'Номер': '',
        'ДатаВыдачи': '',
        'КемВыдан': '',
        'КодПодразделения': '',
        'СрокДокумента': '',
        'МестоРождения': '',
        'ДатаРождения': ''
    }

    query = connector.NewObject("Запрос")
    query.Text = """
                   ВЫБРАТЬ
                       ДокументыФизЛиц.ВидДокумента.Наименование как ВидДокумента,
                       ДокументыФизЛиц.Серия как Серия,
                       ДокументыФизЛиц.Номер как Номер,
                       ДокументыФизЛиц.ДатаВыдачи как ДатаВыдачи,
                       ДокументыФизЛиц.СрокДействия как СрокДействия,
                       ДокументыФизЛиц.КемВыдан как КемВыдан,
                       ДокументыФизЛиц.КодПодразделения как КодПодразделения
                   ИЗ
                       РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
                   ГДЕ
                       ДокументыФизЛиц.ФизЛицо.ФИО = &ФИО
                   """

    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()

    while selection.Next():
        doc_date = selection.ДатаВыдачи
        formatted_doc_date = doc_date.strftime('%d.%m.%Y')

        doc_expiration = selection.СрокДействия
        formatted_doc_expiration = doc_expiration.strftime('%d.%m.%Y')

        parsed_doc_info['ВидДокумента'] = selection.ВидДокумента
        parsed_doc_info['Серия'] = selection.Серия
        parsed_doc_info['Номер'] = selection.Номер
        parsed_doc_info['ДатаВыдачи'] = formatted_doc_date
        parsed_doc_info['СрокДокумента'] = formatted_doc_expiration
        parsed_doc_info['КемВыдан'] = selection.КемВыдан
        parsed_doc_info['КодПодразделения'] = selection.КодПодразделения


    query_2 = connector.NewObject("Запрос")
    query_2.Text = """
                           ВЫБРАТЬ
                               Сотрудники.ФизическоеЛицо как ФизЛицо
                           ИЗ
                               Справочник.Сотрудники КАК Сотрудники
                           ГДЕ
                               Сотрудники.ФизическоеЛицо.ФИО = &ФИО
                           """

    query_2.SetParameter("ФИО", fio)
    selection_2 = query_2.Execute().Choose()

    while selection_2.Next():
        print(f'Дата рождения: {selection_2.ФизЛицо.ДатаРождения}, Место рождения: {selection_2.ФизЛицо.МестоРождения}')

        birth_date = selection_2.ФизЛицо.ДатаРождения
        formatted_birth_date  = birth_date .strftime('%d.%m.%Y')

        birth_place = selection_2.ФизЛицо.МестоРождения
        formatted_birth_place = birth_place.replace(',', '').replace('0', '')

        parsed_doc_info['МестоРождения'] = formatted_birth_place
        parsed_doc_info['ДатаРождения'] = formatted_birth_date

    return parsed_doc_info

def app_1c(fio: str, base_1c: str, username: str, pswd: str) -> dict:
    V83_CONN_STRING = rf'Srvr="srv2024.apriori.loc:1541";Ref="{base_1c}";Usr="{username}";Pwd="{pswd}";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        print(f"Ошибка соединения к базе 1с: {e}")

    V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)

    personal: dict = personal_info(connector=V83, fio=fio)
    document: dict = document_info(connector=V83, fio=fio)
    job: dict = job_info(connector=V83, fio=fio)

    parsed_data_from_1c: dict = {
        'personal_info': personal,
        'document_info': document,
        'job_info': job
    }
    return parsed_data_from_1c


if __name__ == '__main__':
    app_1c(fio='Гордеева Дарья Васильевна', base_1c='zup_bazilevskii', username='Робот', pswd='YUSliot3Cb')

    # V83_CONN_STRING = r'Srvr="srv2024.apriori.loc:1541";Ref="zup_bazilevskii";Usr="Робот";Pwd="YUSliot3Cb";'
    #
    # try:
    #     V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    # except Exception as e:
    #     print(f"Ошибка соединения к базе 1с: {e}")
    #
    # V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    #
    # personal_info: dict = personal_info(connector=V83, fio='Гордеева Дарья Васильевна')
    # document_info: dict = document_info(connector=V83, fio='Гордеева Дарья Васильевна')
    # job_info: dict = job_info(connector=V83, fio='Гордеева Дарья Васильевна')

    pass
    # print(personal_info, document_info, sep='\n')
