import win32com.client
from datetime import timedelta, datetime


# Банковские данные исполнителя
CONTRACTORS_REQ_DATA: dict = {
    'БИК_банка_исполнителя': '044525225',
    'ИНН_исполнителя': '507501320569',
    'Корр_счет_исполнителя': '30101810400000000225',
    'Рас_счёт_исполнителя': '40802810938000059729'
}
# Инф-а орг-ии заказчика
CUSTOMER_ORG_INFO: dict = {
    'org_info': {
        'БикБанка': '044525225',
        'ДолжностьРуководителя': 'Генеральный директор',
        'ИНН_Орг': '9704047266',
        'КПП_Орг': '770401001',
        'КорСчет': '30101810400000000225',
        'НаименованиеПолное': 'Общество с ограниченной ответственностью "РЕСТОБУРГ"',
        'НаименованиеСокращенное': 'ООО "РЕСТОБУРГ"',
        'ОГРН(ИП)': '1217700050512',
        'РабочаяДата': '18.03.2026',
        'РасСчет': '40702810338000175400',
        'Руководитель': 'Тузова Любовь Андреевна',
        'Телефон': '74951337214',
        'ФактическийАдрес': '121099, Город Москва, вн.тер.г. муниципальный округ Арбат, ул Новый Арбат, д. 27, пом/ком/оф 7Б/П',
        'ЮридическийАдрес': '121099, Город Москва, вн.тер.г. муниципальный округ Арбат, ул Новый Арбат, д. 27',
        'НомерДокумента': 2
    }
}


def create_ecp_bill(connector: win32com.client, customer_org_info: dict, contractors_req_data: dict, nds_sum='175,00') -> str | None:
    orgs = getattr(connection.Catalogs, "Организации")

    # Ссылка орг-ии заказчика
    customer_org = orgs.FindByAttribute("ИНН", contractors_req_data['org_info']['ИНН_Орг'])

    # Ссылка орг-ии исполнителя
    contractors_org = orgs.FindByAttribute("ИНН", contractors_req_data['org_info']['ИНН_исполнителя'])

    bills = getattr(connector.Documents, "СчетНаОплатуПокупателю")
    bill = bills.CreateDocument()
    # bill.Fill(True)

    # Заполнение реквизиты покупателя
    setattr(bill, "Организация", customer_org)
    setattr(bill, "ИННПлательщика", customer_org.ИНН)
    setattr(bill, "КПППлательщика", customer_org.КПП)

    # Заполняем реквизиты получателя
    setattr(bill, "ОрганизацияПолучатель", contractors_org)


    # Вот тут кусок кода (пример получения банковсокго счета исполнителя, взято из проекта Bills)
    # if connection.Справочники.БанковскиеСчета.FindByAttribute("НомерСчета",
    #                                                           acc).Владелец.Наименование != needed_org.Наименование:
    #
    #     needed_contr_bank_acc = connection.Справочники.БанковскиеСчета.FindByAttribute("НомерСчета", acc)
    #
    #

    setattr(bill, "СчетОрганизации", needed_org_bank_acc)
    setattr(bill, "Контрагент", needed_contr)
    setattr(bill, "ИННПолучателя", needed_contr.ИНН)
    setattr(bill, "КПППолучателя", needed_contr.КПП)

    # setattr(bill, "СтатьяДвиженияДенежныхСредств",
    #         connector.Справочники.СтатьиДвиженияДенежныхСредств.FindByDescription("Оплата поставщикам (подрядчикам)"))

    setattr(bill, "ВидОперации", connector.Перечисления.ВидыОперацийСписаниеДенежныхСредств.ОплатаПоставщику)
    setattr(bill, "СуммаДокумента", str(sum_doc))

    setattr(bill, "СтавкаНДС", 5)
    setattr(bill, "СуммаНДС", nds_sum)

    current_date = str(datetime.datetime.now().date())
    formatted_date = current_date[8:] + '.' + current_date[5:7] + '.' + current_date[:4]

    setattr(bill, "Дата", datetime.datetime.strptime(formatted_date, '%d.%m.%Y'))
    etattr(bill, "Номер", customer_org_info['org_info']['НомерДокумента'])
    # setattr(bill, "ОчередностьПлатежа", 5)

    setattr(bill, "НазначениеПлатежа", nazn)
    if set_bank:
        print('SET BANK')
        setattr(bill, "СчетКонтрагента", needed_contr_bank_acc)

    print(bill.СчетКонтрагента.Наименование)
    bill.DeletionMark = True
    bill.Write()

    print \
        (f"СчетНаОплату создано: Номер: {bill.Номер}, Дата: {bill.Дата}, Сумма: {bill.СуммаДокумента}, Организация: {bill.Организация.Наименование}, Контрагент: {bill.Контрагент.Наименование}")
    return bill.Ссылка


def create_ecp_act(connector: win32com.client, customer_org_info: dict, contractors_req_data: dict) -> str | None:
    acts = getattr(connector.Documents, "РеализацияТоваровУслуг")
    act = acts.CreateDocument()
    # act.Fill(True)
    # Заполнение реквизитов
    setattr(act, "Организация", needed_org)
    setattr(act, "ИННПлательщика", needed_org.ИНН)
    setattr(act, "КПППлательщика", needed_org.КПП)
    setattr(act, "СчетОрганизации", needed_org_bank_acc)
    setattr(act, "Контрагент", needed_contr)
    setattr(act, "ИННПолучателя", needed_contr.ИНН)
    setattr(act, "КПППолучателя", needed_contr.КПП)

    setattr(act, "СтатьяДвиженияДенежныхСредств",
            connector.Справочники.СтатьиДвиженияДенежныхСредств.FindByDescription("Оплата поставщикам (подрядчикам)"))
    setattr(act, "ВидОперации", connector.Перечисления.ВидыОперацийСписаниеДенежныхСредств.ОплатаПоставщику)
    setattr(act, "СуммаДокумента", str(sum_doc))
    if NDS_rate:
        setattr(act, "СтавкаНДС", NDS_rate)
    setattr(act, "СуммаНДС", NDS_sum)

    current_date = str(datetime.datetime.now().date())
    formatted_date = current_date[8:] + '.' + current_date[5:7] + '.' + current_date[:4]

    setattr(act, "Дата", datetime.datetime.strptime(formatted_date, '%d.%m.%Y'))
    setattr(act, "ОчередностьПлатежа", 5)

    setattr(act, "НазначениеПлатежа", nazn)
    if set_bank:
        print('SET BANK')
        setattr(act, "СчетКонтрагента", needed_contr_bank_acc)

    print(act.СчетКонтрагента.Наименование)
    act.DeletionMark = True
    act.Write()

    print \
        (f"СчетНаОплату создано: Номер: {act.Номер}, Дата: {act.Дата}, Сумма: {act.СуммаДокумента}, Организация: {act.Организация.Наименование}, Контрагент: {act.Контрагент.Наименование}")
    return act.Ссылка


if __name__ == '__main__':

    pass