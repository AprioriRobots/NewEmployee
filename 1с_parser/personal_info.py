import win32com.client
import datetime


def personal_info(connector: win32com.client, fio: str) -> dict:
    parsed_pers_info = {
        'ФИО': '',
        'ДатаРождения': '',
        'ИНН': '',
        'МестоРождения': '',
        'СНИЛС': ''
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

    return parsed_pers_info


def document_info(connector: win32com.client, fio: str) -> dict:
    parsed_doc_info: dict = {
        'ВидДокумента': '',
        'Серия': '',
        'Номер': '',
        'ДатаВыдачи': '',
        'КемВыдан': '',
        'КодПодразделения': ''
    }

    query = connector.NewObject("Запрос")
    query.Text = """
                   ВЫБРАТЬ
                       ДокументыФизЛиц.ВидДокумента.Наименование как ВидДокумента,
                       ДокументыФизЛиц.Серия как Серия,
                       ДокументыФизЛиц.Номер как Номер,
                       ДокументыФизЛиц.ДатаВыдачи как ДатаВыдачи,
                       ДокументыФизЛиц.КемВыдан как КемВыдан,
                       ДокументыФизЛиц.КодПодразделения как КодПодразделения
                   ИЗ
                       РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
                   ГДЕ
                       ДокументыФизЛиц.ФизЛицо.ФИО= &ФИО
                   """

    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()

    while selection.Next():
        doc_date = selection.ДатаВыдачи
        formatted_doc_date = doc_date.strftime('%d.%m.%Y')

        # print(
        #     f"Вид документа: {selection.ВидДокумента}, Серия: {selection.Серия}, Номер: {selection.Номер},"
        #     f" Дата выдачи: {selection.ДатаВыдачи}, Кем выдан: {selection.КемВыдан},"
        #     f" Код подразделения: {selection.КодПодразделения}"
        # )

        parsed_doc_info['ВидДокумента'] = selection.ВидДокумента
        parsed_doc_info['Серия'] = selection.Серия
        parsed_doc_info['Номер'] = selection.Номер
        parsed_doc_info['ДатаВыдачи'] = formatted_doc_date
        parsed_doc_info['КемВыдан'] = selection.КемВыдан
        parsed_doc_info['КодПодразделения'] = selection.КодПодразделения

    return parsed_doc_info


def contact_info(connector: win32com.client, fio: str) -> dict:
    pass


if __name__ == '__main__':
    V83_CONN_STRING = r'Srvr="srv2024.apriori.loc:1541";Ref="zup_bazilevskii";Usr="Робот";Pwd="YUSliot3Cb";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        print(f"Ошибка соединения к базе 1с: {e}")

    V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)

    personal_info: dict = personal_info(connector=V83, fio='Абдимамат кызы Айжаркын')
    document_info: dict = document_info(connector=V83, fio='Абдимамат кызы Айжаркын')

    print(personal_info, document_info, sep='\n')


    # query = V83.NewObject("Запрос")
    # query.Text = """
    #            ВЫБРАТЬ
    #                Справочник.ФизическиеЛица.Наименование КАК ФИО,
    #                Справочник.ФизическиеЛица.ИНН КАК ИНН,
    #                СтраховойНомерПФР
    #            ИЗ
    #                Справочник.ФизическиеЛица КАК ФизическиеЛица
    #            ГДЕ
    #                ФизическиеЛица.Наименование = &ФИО
    #            """
    #
    # query.SetParameter("ФИО", 'Абдимамат кызы Айжаркын')
    # selection = query.Execute().Choose()
    #
    # while selection.Next():
    #     print(
    #         f"Вид документа: {selection.ВидДокумента}, Серия: {selection.Серия}, Номер: {selection.Номер},"
    #         f" Дата выдачи: {selection.ДатаВыдачи}, Кем выдан: {selection.КемВыдан},"
    #         f" Код подразделения: {selection.КодПодразделения}"
    #     )
    # Пример СНИЛС'А 145-434-285 57