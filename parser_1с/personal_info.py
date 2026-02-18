import win32com.client
import datetime


def personal_info(connector: win32com.client, fio: str) -> dict:
    parsed_pers_info = {
        'ФИО': '',
        'ДатаРождения': '',
        'ИНН': '',
        'Страна': '',
        'МестоРождения': '',
        'СНИЛС': '',
        'Адрес_прописка': '',
        'Адрес_проживания': '',
        'Телефон': '',
        'Банк': '',
        'АдресБанка': '',
        'ИннБанка': '',
        'НомерСчета': '',
        'БикБанка': '',
        'КоррСчет': '',
        'РасчетныйСчет': ''
    }

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ
            Сотрудники.ФизическоеЛицо.ФИО КАК ФИО,
             Сотрудники.ФизическоеЛицо.ДатаРождения КАК ДатаРождения,
             Сотрудники.ФизическоеЛицо.ИНН КАК ИНН,
             Сотрудники.ФизическоеЛицо.МестоРождения КАК МестоРождения,
             Сотрудники.ФизическоеЛицо.СтраховойНомерПФР КАК СНИЛС,
             Сотрудники.ФизическоеЛицо как ФизЛицо
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

        parsed_pers_info['ФИО'] = selection.ФИО
        parsed_pers_info['ДатаРождения'] = formatted_date
        parsed_pers_info['ИНН'] = selection.ИНН
        parsed_pers_info['МестоРождения'] = selection.МестоРождения
        parsed_pers_info['СНИЛС'] = selection.СНИЛС

        employee_ras_schet, employee_bik_bank = '', ''
        employee_core_schet = ''

        table_len_fiz_extra_req: int = len(selection.ФизЛицо.ДополнительныеРеквизиты)

        if len(selection.ФизЛицо.ДополнительныеРеквизиты):
            table_len: int = len(selection.ФизЛицо.ДополнительныеРеквизиты)

            for i in range(table_len):
                extra_req_value = selection.ФизЛицо.ДополнительныеРеквизиты[i].Свойство.Наименование

                if extra_req_value == 'Расчетный счет':
                    employee_ras_schet = selection.ФизЛицо.ДополнительныеРеквизиты[i].Значение

                if extra_req_value == 'БИК банка':
                    employee_bik_bank = selection.ФизЛицо.ДополнительныеРеквизиты[i].Значение

                if extra_req_value == 'Корр счет':
                    employee_core_schet = selection.ФизЛицо.ДополнительныеРеквизиты[i].Значение

        parsed_pers_info['КоррСчет'] = employee_core_schet
        parsed_pers_info['РасчетныйСчет'] = employee_ras_schet

    query_2 = connector.NewObject("Запрос")
    query_2.Text = """
                       ВЫБРАТЬ
                           ГражданствоФизическихЛиц.ФизическоеЛицо как ФизЛицо,
                           ГражданствоФизическихЛиц.Страна как Страна
                       ИЗ
                           РегистрСведений.ГражданствоФизическихЛиц КАК ГражданствоФизическихЛиц
                       ГДЕ
                           ГражданствоФизическихЛиц.ФизическоеЛицо.ФИО = &ФИО
                       """

    query_2.SetParameter("ФИО", fio)
    selection_2 = query_2.Execute().Choose()

    reg_address, fact_address = '', ''
    phone_number = ''

    while selection_2.Next():
        # print(f'Контактная информация: {selection_2.ФизЛицо.КонтактнаяИнформация}')

        entry_address: int = 0
        entry_telephone: int = 0

        if len(selection_2.ФизЛицо.КонтактнаяИнформация):
            table_len: int = len(selection_2.ФизЛицо.КонтактнаяИнформация)

            for i in range(table_len):

                if len(selection_2.ФизЛицо.КонтактнаяИнформация[i].НомерТелефона):

                    if entry_telephone == 0:

                        phone_number = selection_2.ФизЛицо.КонтактнаяИнформация[i].НомерТелефона
                        entry_telephone += 1

                if (len(selection_2.ФизЛицо.КонтактнаяИнформация[i].Страна)
                        or len(selection_2.ФизЛицо.КонтактнаяИнформация[i].Регион)):

                    if entry_address == 0:

                        reg_address = selection_2.ФизЛицо.КонтактнаяИнформация[i].Представление
                        fact_address = selection_2.ФизЛицо.КонтактнаяИнформация[i + 1].Представление

                        entry_address += 1

    parsed_pers_info['Адрес_прописка'] = reg_address
    parsed_pers_info['Адрес_проживания'] = fact_address
    parsed_pers_info['Телефон'] = phone_number

    try:
        parsed_pers_info['Страна'] = selection_2.Страна.Наименование    # один раз во время тестирование словил ошибку AttributeError
    except AttributeError:
        parsed_pers_info['Страна'] = ''

    query_3 = connector.NewObject("Запрос")
    query_3.Text = """
                           ВЫБРАТЬ
                               БанковскиеСчета.Банк.Наименование как Банк,
                               БанковскиеСчета.Банк.ИНН как ИНН,
                               БанковскиеСчета.Банк.Адрес как Адрес,
                               БанковскиеСчета.НомерСчета как НомерСчета,
                               БанковскиеСчета.Наименование как ИНФО,
                               БанковскиеСчета.ДополнительныеРеквизиты как ДополнительныеРеквизиты
                           ИЗ
                               Справочник.БанковскиеСчетаКонтрагентов КАК БанковскиеСчета
                           ГДЕ
                               БанковскиеСчета.Владелец.ФИО = &ФИО
                           """

    query_3.SetParameter("ФИО", fio)
    selection_3 = query_3.Execute().Choose()

    while selection_3.Next():
        parsed_pers_info['Банк'] = selection_3.Банк
        parsed_pers_info['АдресБанка'] = selection_3.Адрес
        parsed_pers_info['ИннБанка'] = selection_3.ИНН
        parsed_pers_info['НомерСчета'] = selection_3.НомерСчета

        bank_info: str = selection_3.ИНФО
        bank_bik: str = bank_info.split(',')[0]

        parsed_pers_info['БикБанка'] = bank_bik

    return parsed_pers_info


def job_info(connector: win32com.client, fio: str) -> dict:
    parsed_job_info = {
        'ФИО': '',                          # Есть
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
            'Телефон': '',                  # Есть

            'БикБанка': '',
            'РасСчет': '',
            'КорСчет': ''

        },
        'Должность': '',                    # Есть
        'Подразделение': '',                # Есть
        'Оклад': '',                        # Есть
        'ИспытательныйСрок': '',            # Есть
        'АвансПериод': '',                  # Есть
        'ЗарплатаПериод': '',               # Есть
        'ГрафикРаботы': '',                 # Есть
        'ДатаПриемаНаРаботу': '',           # Есть
        'Отпуск': '28',                     # (Пока временно поставил 28)
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
                           ПриемНаРаботу.ЕжегодныеОтпуска как ЕжегодныеОтпуска
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
        parsed_job_info['Организация']['НаименованиеСокращенное'] = selection.Организация.НаименованиеСокращенное
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
        phone_number = ''

        entry_address: int = 0
        entry_telephone: int = 0

        if  len(selection.Организация.КонтактнаяИнформация):
                table_len: int = len(selection.Организация.КонтактнаяИнформация)

                for i in range(table_len):

                    if len(selection.Организация.КонтактнаяИнформация[i].НомерТелефона):
                        if entry_telephone == 0:

                            phone_number = selection.Организация.КонтактнаяИнформация[i].НомерТелефона
                            entry_telephone += 1

                    if (len(selection.Организация.КонтактнаяИнформация[i].Страна)
                            or len(selection.Организация.КонтактнаяИнформация[i].Регион)):

                        if entry_address == 0:

                            reg_address = selection.Организация.КонтактнаяИнформация[i].Представление
                            fact_address = selection.Организация.КонтактнаяИнформация[i + 1].Представление

                            entry_address += 1

        org_ras_schet, org_bik_bank = '', ''
        org_core_schet = ''

        if len(selection.Организация.ДополнительныеРеквизиты):
            table_len: int = len(selection.Организация.ДополнительныеРеквизиты)

            for i in range(table_len):
                extra_req_value = selection.Организация.ДополнительныеРеквизиты[i].Свойство.Наименование

                if extra_req_value == 'Расчетный счет':
                    org_ras_schet = selection.Организация.ДополнительныеРеквизиты[i].Значение

                if extra_req_value == 'БИК банка':
                    org_bik_bank = selection.Организация.ДополнительныеРеквизиты[i].Значение

                if extra_req_value == 'Корр счет':
                    org_core_schet = selection.Организация.ДополнительныеРеквизиты[i].Значение

        # Пытался по аналогии проитерироваться по selection.Отпуск - возникает ошибка

        parsed_job_info['Организация']['ЮридическийАдрес'] = reg_address
        parsed_job_info['Организация']['ФактическийАдрес'] = fact_address
        parsed_job_info['Организация']['Телефон'] = phone_number

        parsed_job_info['Организация']['БикБанка'] = org_bik_bank
        parsed_job_info['Организация']['РасСчет'] = org_ras_schet
        parsed_job_info['Организация']['КорСчет'] = org_core_schet


        # vac_days: int | float = selection.ЕжегодныеОтпуска.КоличествоДнейВГод

        # if len(selection.Отпуск):
        #     table_len: int = len(selection.Отпуск)
        #
        #     for i in range(table_len):
        #         vac_days: int | float = selection.Отпуск[i].КоличествоДнейВГод
        #         pass

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
            f" ИспытательныйСрок: {selection.ИспытательныйСрок}, Оклад: {selection.Оклад}"
        )
        # f" ОрганизацияНаименованиеСокращенное: {selection.Организация.ОрганизацияНаименованиеСокращенное},"


    query_2 = connector.NewObject("Запрос")
    query_2.Text = """
                                   ВЫБРАТЬ
                                       ЗарплатаКадры.ДатаВыплатыАвансаНеПозжеЧем как АвансПериод,
                                       ЗарплатаКадры.ДатаВыплатыЗарплатыНеПозжеЧем как ЗарплатаПериод
                                   ИЗ
                                       РегистрСведений.НастройкиЗарплатаКадрыРасширенная КАК ЗарплатаКадры
                                   ГДЕ
                                       ЗарплатаКадры.Организация = &Организация
                                   """

    query_2.SetParameter("Организация", selection.Организация)
    selection_2 = query_2.Execute().Choose()

    while selection_2.Next():
        parsed_job_info['АвансПериод'] = selection_2.АвансПериод
        parsed_job_info['ЗарплатаПериод'] = selection_2.ЗарплатаПериод

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
    V83_CONN_STRING = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'

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
    parsed_data: dict = app_1c(fio='Молоканов Константин Владимирович', base_1c='Srvr="srv2024.apriori.loc:1541";Ref="zup_poezd";',
                               username='Робот', pswd='YUSliot3Cb')
    pass
