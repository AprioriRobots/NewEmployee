import os
import win32com.client
from datetime import timedelta, datetime
from typing import Any



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
# Для счета
BILL_DATA: dict[str, Any] = {
    # Дефолты по примеру "ООО РЕСТОБУРГ_Пример счёта.pdf"
    "organization_inn": "507501320569",
    "organization_bank_account": "40802810938000059729",
    "counterparty_inn": "9704047266",
    "counterparty_kpp": "770401001",
    "document_number": "2029",
    "document_date": "2026-03-12",
    "bill_amount": 3500.00,
    "nds_rate": "НДС5",
    "nds_sum": 350.00,
    "payment_purpose": (
        "Продление сертификата ЭЦП Отчетность в ЛК н/п ФНС/Работа IT специалиста по установке и активации ЭЦП"
    ),
    "item_name": "Вводный курс",
    "item_article": "",
}


def _env_str(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip()


def _env_float(name: str, default: float) -> float:
    value = _env_str(name)
    if not value:
        return default
    return float(value.replace(",", "."))


def _env_bool(name: str, default: bool = True) -> bool:
    value = _env_str(name).lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "y", "on", "да"}


def _is_empty_ref(obj: Any) -> bool:
    if obj is None:
        return True
    try:
        return bool(obj.Empty())
    except Exception:
        return False


def _try_getattr(obj: Any, attr_name: str) -> Any:
    try:
        return getattr(obj, attr_name)
    except Exception:
        return None


def _set_if_exists(target: Any, attr_name: str, value: Any) -> bool:
    try:
        getattr(target, attr_name)
    except Exception:
        return False

    try:
        setattr(target, attr_name, value)
        return True
    except Exception:
        return False


def _parse_1c_date(value: str) -> datetime:
    if "." in value:
        return datetime.strptime(value, "%d.%m.%Y")
    return datetime.strptime(value, "%Y-%m-%d")


def connect_1c() -> Any:
    base_1c = _env_str("ONEC_BASE")
    username = _env_str("ONEC_USER")
    password = _env_str("ONEC_PASSWORD")

    if not base_1c:
        raise RuntimeError("Переменная ONEC_BASE обязательна (например File=\"...\"; или Srvr=\"...\";Ref=\"...\";).")
    if not username:
        raise RuntimeError("Переменная ONEC_USER обязательна.")

    conn_string = f'{base_1c}Usr="{username}";'
    if password:
        conn_string += f'Pwd="{password}";'

    return win32com.client.Dispatch("V83.COMConnector").Connect(conn_string)


def resolve_bill_document_manager(connector: Any) -> Any:
    doc_name = "СчетНаОплатуПокупателю"
    manager = _try_getattr(connector.Documents, doc_name)
    if manager is None:
        raise RuntimeError(
            f'Документ "{doc_name}" не найден в текущей конфигурации 1С. '
            "Проверьте, что подключение выполнено к базе 'Бухгалтерия предприятия учебная'."
        )
    return manager


def _find_org_by_inn(connector: Any, inn: str) -> Any:
    orgs = _try_getattr(connector.Catalogs, "Организации")
    if orgs is None:
        return None
    org = orgs.FindByAttribute("ИНН", inn)
    return None if _is_empty_ref(org) else org


def _find_counterparty_by_inn(connector: Any, inn: str) -> Any:
    counterparties = _try_getattr(connector.Catalogs, "Контрагенты")
    if counterparties is None:
        return None
    counterparty = counterparties.FindByAttribute("ИНН", inn)
    return None if _is_empty_ref(counterparty) else counterparty


def _find_org_bank_account(connector: Any, account_number: str) -> Any:
    bank_accounts = _try_getattr(connector.Catalogs, "БанковскиеСчета")
    if bank_accounts is None:
        return None
    account_ref = bank_accounts.FindByAttribute("НомерСчета", account_number)
    return None if _is_empty_ref(account_ref) else account_ref


def _find_first_contract_by_counterparty(connector: Any, counterparty_ref: Any) -> Any:
    if counterparty_ref is None:
        return None

    counterparty_name = counterparty_ref.Наименование

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Договоры.Ссылка КАК Ссылка
        ИЗ
            Справочник.ДоговорыКонтрагентов КАК Договоры
        ГДЕ
            Договоры.Владелец = &Контрагент
    """
    query.SetParameter("Контрагент", counterparty_ref)
    selection = query.Execute().Choose()
    if selection.Next():
        return selection.Ссылка
    return None


def resolve_references(connector: Any, payload: dict[str, Any]) -> dict[str, Any]:
    payload["organization_inn"] = "7799292647"
    payload["counterparty_inn"] = "7799292647"
    payload["organization_bank_account"] = "40101810000000010004"

    org_ref = _find_org_by_inn(connector, payload["organization_inn"])
    counterparty_ref = _find_counterparty_by_inn(connector, payload["counterparty_inn"])
    org_bank_ref = _find_org_bank_account(connector, payload["organization_bank_account"])
    contract_ref = _find_first_contract_by_counterparty(connector, counterparty_ref)

    return {
        "organization": org_ref,
        "counterparty": counterparty_ref,
        "organization_bank_account": org_bank_ref,
        "contract": contract_ref,
    }


def build_bill_payload_from_defaults() -> dict[str, Any]:
    payload = dict(DEFAULT_BILL_DATA)

    payload["organization_inn"] = _env_str("ONEC_BILL_ORG_INN", payload["organization_inn"])
    payload["organization_bank_account"] = _env_str(
        "ONEC_BILL_ORG_BANK_ACCOUNT",
        payload["organization_bank_account"],
    )
    payload["counterparty_inn"] = _env_str("ONEC_BILL_COUNTERPARTY_INN", payload["counterparty_inn"])
    payload["counterparty_kpp"] = _env_str("ONEC_BILL_COUNTERPARTY_KPP", payload["counterparty_kpp"])
    payload["document_number"] = _env_str("ONEC_BILL_NUMBER", payload["document_number"])
    payload["document_date"] = _env_str("ONEC_BILL_DATE", payload["document_date"])
    payload["bill_amount"] = _env_float("ONEC_BILL_AMOUNT", payload["bill_amount"])
    payload["nds_sum"] = _env_float("ONEC_BILL_NDS_SUM", payload["nds_sum"])
    payload["nds_rate"] = _env_str("ONEC_BILL_NDS_RATE", payload["nds_rate"])
    payload["payment_purpose"] = _env_str("ONEC_BILL_PURPOSE", payload["payment_purpose"])
    payload["item_name"] = _env_str("ONEC_BILL_ITEM_NAME", payload["item_name"])
    payload["item_article"] = _env_str("ONEC_BILL_ITEM_ARTICLE", payload["item_article"])

    return payload


def _set_nds_rate_if_possible(connector: Any, bill: Any, nds_rate_name: str) -> None:
    if not nds_rate_name:
        return
    try:
        enum_value = getattr(connector.Enums.СтавкиНДС, nds_rate_name)
        _set_if_exists(bill, "СтавкаНДС", enum_value)
    except Exception:
        # В разных версиях/конфигурациях ставка НДС может называться иначе.
        pass


def resolve_nomenclature(connector: Any, article: str, name: str) -> Any:
    catalog = _try_getattr(connector.Catalogs, "Номенклатура")
    if catalog is None:
        return None

    if article:
        by_article = catalog.FindByAttribute("Артикул", article)
        if not _is_empty_ref(by_article):
            return by_article

    if name:
        by_name = catalog.FindByDescription(name)
        if not _is_empty_ref(by_name):
            return by_name

    return None


def resolve_bill_tabular_section(bill: Any) -> Any:
    candidates = ("Товары", "Услуги", "ТоварыУслуги")
    for section_name in candidates:
        section = _try_getattr(bill, section_name)
        if section is not None:
            return section

    raise RuntimeError(
        'Не найдена табличная часть для строк счета. Ожидались: "Товары", "Услуги", "ТоварыУслуги".'
    )


def add_bill_line(connector: Any, bill: Any, payload: dict[str, Any]) -> None:
    item_name = str(payload.get("item_name", "")).strip()
    item_article = str(payload.get("item_article", "")).strip()
    line_sum = float(payload["bill_amount"])
    line_nds_sum = float(payload["nds_sum"])

    nomenclature_ref = resolve_nomenclature(connector=connector, article=item_article, name=item_name)
    if nomenclature_ref is None:
        item_hint = item_article or item_name or "без указания названия"
        raise RuntimeError(f'Не найдена номенклатура для добавления в счет: "{item_hint}".')

    section = resolve_bill_tabular_section(bill)
    row = _try_getattr(section, "Add")
    if row is None:
        add_method = _try_getattr(section, "Добавить")
        if add_method is None:
            raise RuntimeError("Не удалось добавить строку в табличную часть счета: метод Add/Добавить не найден.")
        row = add_method()
    else:
        row = row()

    assigned_fields = 0
    if _set_if_exists(row, "Номенклатура", nomenclature_ref):
        assigned_fields += 1

    if _set_if_exists(row, "Количество", 1):
        assigned_fields += 1
    if _set_if_exists(row, "Цена", line_sum):
        assigned_fields += 1
    if _set_if_exists(row, "Сумма", line_sum):
        assigned_fields += 1

    if _set_if_exists(row, "Содержание", item_name):
        assigned_fields += 1
    if _set_if_exists(row, "Наименование", item_name):
        assigned_fields += 1

    nds_rate_name = str(payload.get("nds_rate", ""))
    if nds_rate_name:
        try:
            enum_value = getattr(connector.Enums.СтавкиНДС, nds_rate_name)
            if _set_if_exists(row, "СтавкаНДС", enum_value):
                assigned_fields += 1
        except Exception:
            pass

    if _set_if_exists(row, "СуммаНДС", line_nds_sum):
        assigned_fields += 1

    if assigned_fields == 0:
        raise RuntimeError("Не удалось добавить строку счета: у строки отсутствуют ожидаемые реквизиты.")


def create_and_post_bill(connector: Any, payload: dict[str, Any], doc_number: str) -> Any:
    bill_manager = resolve_bill_document_manager(connector)
    refs = resolve_references(connector, payload)

    missing: list[str] = []
    if refs["organization"] is None:
        missing.append(f'Организация по ИНН "{payload["organization_inn"]}"')
    if refs["counterparty"] is None:
        missing.append(f'Контрагент по ИНН "{payload["counterparty_inn"]}"')
        missing.append(f'Контрагент по ИНН "{payload["counterparty_inn"]}"')
    if refs["organization_bank_account"] is None:
        missing.append(f'Банковский счет организации "{payload["organization_bank_account"]}"')
    if missing:
        raise RuntimeError("Не найдены обязательные ссылки: " + "; ".join(missing))

    bill = bill_manager.CreateDocument()

    _set_if_exists(bill, "Организация", refs["organization"])
    _set_if_exists(bill, "ОрганизацияПолучатель", refs["organization"])

    _set_if_exists(bill, "Контрагент", refs["counterparty"])
    _set_if_exists(bill, "СтруктурнаяЕдиница", refs["organization_bank_account"])

    contract_name = refs["contract"].Наименование

    _set_if_exists(bill, "ДоговорКонтрагент", refs["contract"])     # Кодекс выдал атрибут Договор, вместо ДоговорКонтрагент

    if refs["organization"] is not None:
        _set_if_exists(bill, "ИННПлательщика", refs["organization"].ИНН)
        _set_if_exists(bill, "КПППлательщика", refs["organization"].КПП)

    if refs["counterparty"] is not None:
        _set_if_exists(bill, "ИННПолучателя", refs["counterparty"].ИНН)
        _set_if_exists(bill, "КПППолучателя", refs["counterparty"].КПП)

    _set_if_exists(bill, "СуммаДокумента", payload["bill_amount"])
    _set_if_exists(bill, "СуммаНДС", payload["nds_sum"])
    _set_nds_rate_if_possible(connector, bill, str(payload["nds_rate"]))
    _set_if_exists(bill, "НазначениеПлатежа", payload["payment_purpose"])
    _set_if_exists(bill, "Дата", _parse_1c_date(payload["document_date"]))

    payload["document_number"] = doc_number

    _set_if_exists(bill, "Номер", payload["document_number"])
    _set_if_exists(bill, "КПППолучателя", payload["counterparty_kpp"])

    add_bill_line(connector=connector, bill=bill, payload=payload)

    bill.Write()

    # if post_document:
    #     bill.Post()

    return bill.Ссылка


def _print_result(bill: Any, posted_expected: bool) -> None:
    posted_value = _try_getattr(bill, "Проведен")
    posted_text = posted_value if posted_value is not None else f"н/д (ожидалось: {posted_expected})"

    organization_name = ""
    counterparty_name = ""
    try:
        organization_name = bill.Организация.Наименование
    except Exception:
        pass
    try:
        counterparty_name = bill.Контрагент.Наименование
    except Exception:
        pass

    print("Счет на оплату успешно создан.")
    print(f"Номер: {getattr(bill, 'Номер', '')}")
    print(f"Дата: {getattr(bill, 'Дата', '')}")
    print(f"Сумма: {getattr(bill, 'СуммаДокумента', '')}")
    print(f"Организация: {organization_name}")
    print(f"Контрагент: {counterparty_name}")
    print(f"Проведен: {posted_text}")
    print(f"Ссылка: {getattr(bill, 'Ссылка', '')}")


def create_bill(connector: win32com.client, number: str) -> win32com.client:
    payload = build_bill_payload_from_defaults()
    post_document = _env_bool("ONEC_BILL_POST", default=True)

    bill = create_and_post_bill(connector=connector, payload=payload, doc_number=number)
    # _print_result(bill=bill, posted_expected=post_document)
    return bill.Ссылка


if __name__ == "__main__":
    V83_CONN_STRING = 'Srvr="srv2024.apriori.loc:1541";Ref="buh_apriori";Usr="Робот";Pwd="YUSliot3Cb";'
    V83 = None

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        print(f"Ошибка соединение к базе 1с: {e}")

    if V83 is not None:
        try:
            create_bill(connector=connector, number='')
        except Exception as error:
            print(f"Ошибка во время создания счета в 1С: {error}")
