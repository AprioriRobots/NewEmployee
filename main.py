import win32com.client
import datetime


if __name__ == '__main__':
    V83_CONN_STRING = r'File="C:\Users\serge\OneDrive\Документы\1C\DemoHRMBase6";Usr="Администратор";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        print(f"Ошибка соединения к базе 1с: {e}")

    V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)

    query = V83.NewObject("Запрос")

    # Сотрудники.ФизическоеЛицо.ФИО КАК ФИО,
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

    query.SetParameter("ФИО", "Федотов Сергей Сергеевич")
    selection = query.Execute().Choose()
    while selection.Next():
        print(
            f"ФИО: {selection.ФИО}, Дата Рождения: {selection.ДатаРождения},"
            f" ИНН: {selection.ИНН}, МестоРождения: {selection.МестоРождения},"
            f" СНИЛС: {selection.СНИЛС}"
        )

    query_2 = V83.NewObject("Запрос")
    query_2.Text = """
        ВЫБРАТЬ
             ДокументыФизЛиц.Представление как Документ
        ИЗ
            РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
        ГДЕ
            ДокументыФизЛиц.ФизЛицо.ФИО= &ФИО
        """

    query_2.SetParameter("ФИО", "Федотов Сергей Сергеевич")
    selection = query_2.Execute().Choose()

    while selection.Next():
        print(
            f"Паспортные данные: {selection.Документ}"
        )
