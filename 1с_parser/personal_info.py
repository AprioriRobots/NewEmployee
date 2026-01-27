import win32com.client
import datetime


def personal_info(connector: win32com.client, fio: str) -> dict:
    query = connector.NewObject("Запрос")

    parsed_pers_info = {
        'ФИО': '',
        'ДатаРождения': '',
        'ИНН': '',
        'МестоРождения': '',
        'СНИЛС': ''
    }
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
    query = connector.NewObject("Запрос")

    parsed_doc_info: dict = {
        'Документ': '',
    }
    query.Text = """
           ВЫБРАТЬ
                ДокументыФизЛиц.Представление как Документ
           ИЗ
               РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
           ГДЕ
               ДокументыФизЛиц.ФизЛицо.ФИО= &ФИО
           """

    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()

    while selection.Next():
        # print(
        #     f"Паспортные данные: {selection.Документ}"
        # )
        parsed_doc_info['Документ'] = selection.Документ

    return parsed_doc_info


if __name__ == '__main__':
    V83_CONN_STRING = r'File="C:\Users\serge\OneDrive\Документы\1C\DemoHRMBase6";Usr="Администратор";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        print(f"Ошибка соединения к базе 1с: {e}")

    V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)

    # personal_info: dict = personal_info(connector=V83, fio='Федотов Сергей Сергеевич')
    # document_info: dict = document_info(connector=V83, fio='Федотов Сергей Сергеевич')
    #
    # print(personal_info, document_info, sep='\n')

