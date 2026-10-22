from datetime import datetime

import win32com.client


def _normalize_requisite_name(value) -> str:
    return str(value or "").replace(" ", "").replace("_", "").replace("ё", "е").lower()


def _fill_bank_fields(target: dict, additional_requisites, bik_key: str, corr_key: str, account_key: str) -> None:
    try:
        table_len = len(additional_requisites)
    except Exception:
        return

    for i in range(table_len):
        try:
            title = additional_requisites[i].Свойство.Заголовок
            value = additional_requisites[i].Значение
        except Exception:
            continue

        if not value:
            continue

        normalized = _normalize_requisite_name(title)
        if "бик" in normalized and not target.get(bik_key):
            target[bik_key] = _coerce_1c_scalar_text(value)
        elif ("коррсчет" in normalized or "корсчет" in normalized or "корреспондент" in normalized) and not target.get(corr_key):
            target[corr_key] = _coerce_1c_scalar_text(value)
        elif ("расчетныйсчет" in normalized or "рассчетныйсчет" in normalized or normalized == "расчетныйсчет") and not target.get(account_key):
            target[account_key] = _coerce_1c_scalar_text(value)


def _get_connector(base_1c: str, username: str, pswd: str):
    conn_string = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'
    return win32com.client.Dispatch("V83.COMConnector").Connect(conn_string)


def _safe_value(value):
    try:
        return value if value is not None else ""
    except Exception:
        return ""


def _is_unknown_com_text(value: str) -> bool:
    normalized = str(value or "").strip().lower()
    return "comobject" in normalized and "unknown" in normalized


def _coerce_1c_scalar_text(value) -> str:
    if value in (None, ""):
        return ""

    if isinstance(value, str):
        text = value.strip()
        return "" if _is_unknown_com_text(text) else text

    if isinstance(value, (int, float, bool)):
        return str(value).strip()

    try:
        if hasattr(value, "strftime"):
            return value.strftime("%d.%m.%Y")
    except Exception:
        pass

    for attr_name in (
        "Код",
        "Наименование",
        "Представление",
        "Description",
        "БИК",
        "БИКРКЦ",
        "КоррСчет",
        "КорСчет",
        "НомерСчета",
        "ИНН",
        "Адрес",
    ):
        attr_value = _safe_getattr(value, attr_name)
        if attr_value in (None, "", value):
            continue
        coerced_value = _coerce_1c_scalar_text(attr_value)
        if coerced_value:
            return coerced_value

    try:
        text = str(value).strip()
    except Exception:
        return ""

    return "" if _is_unknown_com_text(text) else text


def _normalize_account_number(value) -> str:
    raw_value = _coerce_1c_scalar_text(value)
    if not raw_value:
        return ""
    digits_only = "".join(ch for ch in raw_value if ch.isdigit())
    return digits_only or raw_value


def _has_bank_account_number(bank_data: dict) -> bool:
    return bool(bank_data.get("НомерСчета") or bank_data.get("РасчетныйСчет"))


def _iter_1c_collection_items(value):
    if value is None:
        return

    try:
        for item in value:
            yield item
        return
    except Exception:
        pass

    for length_attr in ("Count", "Количество"):
        try:
            count = getattr(value, length_attr)
            if callable(count):
                count = count()
            count = int(count)
            for index in range(count):
                try:
                    yield value[index]
                except Exception:
                    try:
                        yield value.Get(index)
                    except Exception:
                        continue
            return
        except Exception:
            continue


def _parse_date(value: str):
    if not value:
        return None
    try:
        return datetime.strptime(str(value), "%d.%m.%Y")
    except Exception:
        return None


def _coerce_datetime(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    try:
        year = getattr(value, "year", None)
        month = getattr(value, "month", None)
        day = getattr(value, "day", None)
        if year and month and day:
            return datetime(year, month, day)
    except Exception:
        pass
    try:
        return _parse_date(str(value))
    except Exception:
        return None


def _describe_owner(owner) -> str:
    if owner is None:
        return "<None>"

    for attr_name in ("Наименование", "НаименованиеПолное", "ФИО", "ИНН"):
        try:
            attr_value = getattr(owner, attr_name)
            if attr_value:
                return f"{owner} [{attr_name}={attr_value}]"
        except Exception:
            continue

    return str(owner)


def _safe_getattr(obj, attr_name: str):
    try:
        return getattr(obj, attr_name)
    except Exception:
        return None


def _log_bank_source(source: str, bank_data: dict, **context) -> None:
    # Detailed employee-bank diagnostics are intentionally disabled in normal runs.
    return None


def _extract_bank_data_from_account_ref(account_ref) -> dict:
    if account_ref is None:
        return {}

    if isinstance(account_ref, str):
        stripped_value = account_ref.strip()
        if stripped_value:
            return {
                "Банк": "",
                "ИннБанка": "",
                "АдресБанка": "",
                "НомерСчета": stripped_value,
                "РасчетныйСчет": stripped_value,
                "БикБанка": "",
                "КоррСчет": "",
            }
        return {}

    number = _safe_getattr(account_ref, "НомерСчета")
    bank = _safe_getattr(account_ref, "Банк")
    requisites = _safe_getattr(account_ref, "ДополнительныеРеквизиты")
    presentation = _safe_getattr(account_ref, "Наименование")

    result = {
        "Банк": _coerce_1c_scalar_text(_safe_getattr(bank, "Наименование")),
        "ИннБанка": _coerce_1c_scalar_text(_safe_getattr(bank, "ИНН")),
        "АдресБанка": _coerce_1c_scalar_text(_safe_getattr(bank, "Адрес")),
        "НомерСчета": _coerce_1c_scalar_text(number),
        "РасчетныйСчет": _coerce_1c_scalar_text(number),
        "БикБанка": "",
        "КоррСчет": "",
    }

    _fill_bank_fields(result, requisites, "БикБанка", "КоррСчет", "РасчетныйСчет")

    if not result["БикБанка"]:
        presentation_str = _coerce_1c_scalar_text(presentation)
        if "," in presentation_str:
            result["БикБанка"] = presentation_str.split(",")[0].strip()

    if result.get("НомерСчета") or result.get("РасчетныйСчет") or result.get("Банк"):
        return result
    return {}


def _extract_bank_data_from_requisites_table(additional_requisites) -> dict:
    result = {
        "Банк": "",
        "ИннБанка": "",
        "АдресБанка": "",
        "НомерСчета": "",
        "РасчетныйСчет": "",
        "БикБанка": "",
        "КоррСчет": "",
    }

    _fill_bank_fields(result, additional_requisites, "БикБанка", "КоррСчет", "РасчетныйСчет")
    if result["РасчетныйСчет"] and not result["НомерСчета"]:
        result["НомерСчета"] = result["РасчетныйСчет"]

    if any(result.values()):
        return result
    return {}


def _extract_bank_data_from_scalar_attrs(owner) -> tuple[dict, list[dict]]:
    result = {
        "Банк": "",
        "ИннБанка": "",
        "АдресБанка": "",
        "НомерСчета": "",
        "РасчетныйСчет": "",
        "БикБанка": "",
        "КоррСчет": "",
    }
    attempts = []

    mapping = {
        "Банк": "Банк",
        "НаименованиеБанка": "Банк",
        "ИннБанка": "ИннБанка",
        "ИННБанка": "ИннБанка",
        "АдресБанка": "АдресБанка",
        "НомерСчета": "НомерСчета",
        "РасчетныйСчет": "РасчетныйСчет",
        "РасчетныйСчёт": "РасчетныйСчет",
        "БикБанка": "БикБанка",
        "БИК": "БикБанка",
        "КоррСчет": "КоррСчет",
        "КорСчет": "КоррСчет",
        "КорреспондентскийСчет": "КоррСчет",
    }

    for attr_name, target_key in mapping.items():
        attr_value = _safe_getattr(owner, attr_name)
        attempts.append({
            "attribute": attr_name,
            "exists": attr_value is not None,
            "value": _coerce_1c_scalar_text(attr_value) if attr_value is not None else "",
        })
        if attr_value and not result[target_key]:
            result[target_key] = _coerce_1c_scalar_text(attr_value)

    if result["РасчетныйСчет"] and not result["НомерСчета"]:
        result["НомерСчета"] = result["РасчетныйСчет"]
    if result["НомерСчета"] and not result["РасчетныйСчет"]:
        result["РасчетныйСчет"] = result["НомерСчета"]

    if any(result.values()):
        return result, attempts
    return {}, attempts


def _is_payroll_bank_related_name(name: str) -> bool:
    normalized = _normalize_requisite_name(name)
    keywords = (
        "банк",
        "счет",
        "счёт",
        "зарп",
        "выплат",
        "лицев",
        "проект",
        "карт",
        "перечисл",
    )
    return any(keyword in normalized for keyword in keywords)


def _extract_bank_data_by_metadata_scan(owner) -> tuple[dict, list[dict]]:
    attempts = []

    metadata = _safe_getattr(owner, "Метаданные")
    if callable(metadata):
        try:
            metadata = metadata()
        except Exception:
            metadata = None

    if metadata is None:
        return {}, attempts

    requisites = _safe_getattr(metadata, "Реквизиты") or []
    for requisite in requisites:
        requisite_name = _coerce_1c_scalar_text(_safe_getattr(requisite, "Имя"))
        if not requisite_name or not _is_payroll_bank_related_name(requisite_name):
            continue
        attr_value = _safe_getattr(owner, requisite_name)
        attempt = {
            "metadata_source": "Реквизиты",
            "attribute": requisite_name,
            "exists": attr_value is not None,
            "value": _describe_owner(attr_value) if attr_value is not None else "",
            "data": _extract_bank_data_from_value(attr_value) if attr_value is not None else {},
        }
        attempts.append(attempt)
        if attempt["data"]:
            return attempt["data"], attempts

    table_sections = _safe_getattr(metadata, "ТабличныеЧасти") or []
    for table_section in table_sections:
        section_name = _coerce_1c_scalar_text(_safe_getattr(table_section, "Имя"))
        if not section_name or not _is_payroll_bank_related_name(section_name):
            continue
        section_value = _safe_getattr(owner, section_name)
        section_data = _extract_bank_data_from_value(section_value) if section_value is not None else {}
        attempt = {
            "metadata_source": "ТабличныеЧасти",
            "attribute": section_name,
            "exists": section_value is not None,
            "value": _describe_owner(section_value) if section_value is not None else "",
            "data": section_data,
        }
        attempts.append(attempt)
        if section_data:
            return section_data, attempts

    return {}, attempts


def _extract_bank_data_from_value(value, depth: int = 0, seen: set | None = None) -> dict:
    if value is None or depth > 3:
        return {}

    if seen is None:
        seen = set()

    object_key = str(value)
    if object_key in seen:
        return {}
    seen.add(object_key)

    collection_data = _extract_bank_data_from_collection(value, depth, seen)
    if collection_data:
        return collection_data

    direct_data = _extract_bank_data_from_account_ref(value)
    if direct_data:
        return direct_data

    requisites = _safe_getattr(value, "ДополнительныеРеквизиты")
    if requisites is not None:
        requisites_data = _extract_bank_data_from_requisites_table(requisites)
        if requisites_data:
            return requisites_data

    scalar_data, _ = _extract_bank_data_from_scalar_attrs(value)
    if scalar_data:
        return scalar_data

    nested_attrs = (
        "Счет",
        "Счёт",
        "БанковскийСчет",
        "БанковскийСчёт",
        "БанковскийСчетКонтрагента",
        "БанковскийСчетФизическогоЛица",
        "БанковскийСчетДляПеречисленияЗарплаты",
        "СчетДляПеречисленияЗарплаты",
        "ЗарплатныйСчет",
        "ЛицевойСчет",
        "ЛицевойСчетСотрудника",
        "Банк",
        "ЗарплатныйПроект",
        "Проект",
        "СпособВыплатыЗарплаты",
        "ВыплатаЗарплаты",
        "ВыплатыИУчетЗатрат",
    )

    for attr_name in nested_attrs:
        nested_value = _safe_getattr(value, attr_name)
        if nested_value is None:
            continue
        nested_data = _extract_bank_data_from_value(nested_value, depth + 1, seen)
        if nested_data:
            return nested_data

    return {}


def _extract_bank_data_from_collection(value, depth: int = 0, seen: set | None = None) -> dict:
    candidates = []

    for item in _iter_1c_collection_items(value):
        if item is None:
            continue

        item_data = _extract_bank_data_from_value(item, depth + 1, seen)
        if not item_data:
            continue

        account_number = _normalize_account_number(
            _safe_getattr(item, "НомерСчета")
            or _safe_getattr(item, "НомерЛицевогоСчета")
            or _safe_getattr(item, "ЛицевойСчет")
            or item_data.get("НомерСчета")
        )
        if account_number:
            item_data["НомерСчета"] = account_number
            item_data["РасчетныйСчет"] = account_number

        project_name = _coerce_1c_scalar_text(
            _safe_getattr(item, "ЗарплатныйПроект")
            or _safe_getattr(item, "Проект")
            or _safe_getattr(item, "СпособВыплатыЗарплаты")
        )
        open_date = _coerce_datetime(
            _safe_getattr(item, "Открыт")
            or _safe_getattr(item, "ДатаОткрытия")
            or _safe_getattr(item, "ДатаОткрытияСчета")
        )

        candidates.append({
            "bank_data": item_data,
            "account_number": account_number,
            "project_name": project_name,
            "open_date": open_date,
            "period_date": None,
        })

    return _choose_best_payroll_candidate(candidates)


def _describe_requisites_table(additional_requisites) -> list[dict]:
    try:
        table_len = len(additional_requisites)
    except Exception:
        return []

    result = []
    for i in range(table_len):
        try:
            title = additional_requisites[i].Свойство.Заголовок
            value = additional_requisites[i].Значение
        except Exception:
            continue
        result.append({
            "title": _coerce_1c_scalar_text(title),
            "value": _coerce_1c_scalar_text(value),
        })
    return result


def _read_bank_details_from_owner_attributes(owner) -> tuple[dict, list[dict]]:
    attempts = []
    candidate_attrs = (
        "ОсновнойБанковскийСчет",
        "БанковскийСчет",
        "СчетВБанке",
        "ОсновнойСчет",
        "ЗарплатныйСчет",
        "СчетДляПеречисленияЗарплаты",
        "БанковскийСчетДляПеречисленияЗарплаты",
        "ЛицевойСчет",
        "ЛицевойСчетСотрудника",
        "НомерЛицевогоСчета",
        "СпособВыплатыЗарплаты",
        "ВыплатаЗарплаты",
        "ВыплатыИУчетЗатрат",
        "БанковскиеРеквизиты",
        "ДополнительныеРеквизиты",
    )

    for attr_name in candidate_attrs:
        attr_value = _safe_getattr(owner, attr_name)
        attempt = {
            "attribute": attr_name,
            "exists": attr_value is not None,
            "value": _describe_owner(attr_value) if attr_value is not None else "",
            "data": {},
        }
        if attr_value is not None:
            if attr_name == "ДополнительныеРеквизиты":
                attempt["requisites"] = _describe_requisites_table(attr_value)
                attempt["data"] = _extract_bank_data_from_requisites_table(attr_value)
            else:
                attempt["data"] = _extract_bank_data_from_value(attr_value)
        attempts.append(attempt)

        if attempt["data"]:
            return attempt["data"], attempts

    scalar_data, scalar_attempts = _extract_bank_data_from_scalar_attrs(owner)
    attempts.append({
        "attribute_group": "scalar_attrs",
        "found": bool(scalar_data),
        "attempts": scalar_attempts,
        "data": scalar_data,
    })
    if scalar_data:
        return scalar_data, attempts

    metadata_data, metadata_attempts = _extract_bank_data_by_metadata_scan(owner)
    attempts.append({
        "attribute_group": "metadata_scan",
        "found": bool(metadata_data),
        "attempts": metadata_attempts,
        "data": metadata_data,
    })
    if metadata_data:
        return metadata_data, attempts

    return {}, attempts


def _find_physical_person(connector, fio: str):
    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Сотрудники.ФизическоеЛицо КАК ФизЛицо
        ИЗ
            Справочник.Сотрудники КАК Сотрудники
        ГДЕ
            Сотрудники.ФизическоеЛицо.ФИО = &ФИО
    """
    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()
    if selection.Next():
        return selection.ФизЛицо
    return None


def _find_employee(connector, fio: str, inn_org: str = "", hire_date=None):
    request_object_name = "\u0417\u0430\u043f\u0440\u043e\u0441"
    if inn_org:
        query = connector.NewObject(request_object_name)
        hire_date_filter = ""
        if hire_date is not None:
            hire_date_filter = """
                \u0418 \u041d\u0410\u0427\u0410\u041b\u041e\u041f\u0415\u0420\u0418\u041e\u0414\u0410(\u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0414\u0430\u0442\u0430, \u0414\u0415\u041d\u042c) = &\u0414\u0430\u0442\u0430\u041f\u0440\u0438\u0435\u043c\u0430"""
        query.Text = f"""
            \u0412\u042b\u0411\u0420\u0410\u0422\u042c \u041f\u0415\u0420\u0412\u042b\u0415 1
                \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a \u041a\u0410\u041a \u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a
            \u0418\u0417
                \u0414\u043e\u043a\u0443\u043c\u0435\u043d\u0442.\u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443 \u041a\u0410\u041a \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443
            \u0413\u0414\u0415
                \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0424\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u043e\u0435\u041b\u0438\u0446\u043e.\u0424\u0418\u041e = &\u0424\u0418\u041e
                \u0418 \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u041e\u0440\u0433\u0430\u043d\u0438\u0437\u0430\u0446\u0438\u044f.\u0418\u041d\u041d = &\u0418\u041d\u041d\u041e\u0440\u0433{hire_date_filter}
            \u0423\u041f\u041e\u0420\u042f\u0414\u041e\u0427\u0418\u0422\u042c \u041f\u041e
                \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0414\u0430\u0442\u0430 \u0423\u0411\u042b\u0412,
                \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0421\u0441\u044b\u043b\u043a\u0430 \u0423\u0411\u042b\u0412
        """
        query.SetParameter("\u0424\u0418\u041e", fio)
        query.SetParameter("\u0418\u041d\u041d\u041e\u0440\u0433", inn_org)
        if hire_date is not None:
            query.SetParameter("\u0414\u0430\u0442\u0430\u041f\u0440\u0438\u0435\u043c\u0430", hire_date)
        selection = query.Execute().Choose()
        if selection.Next():
            return _safe_getattr(selection, "\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a")

    query = connector.NewObject(request_object_name)
    query.Text = """
        \u0412\u042b\u0411\u0420\u0410\u0422\u042c \u041f\u0415\u0420\u0412\u042b\u0415 1
            \u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u0438.\u0421\u0441\u044b\u043b\u043a\u0430 \u041a\u0410\u041a \u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a
        \u0418\u0417
            \u0421\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a.\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u0438 \u041a\u0410\u041a \u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u0438
        \u0413\u0414\u0415
            \u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u0438.\u0424\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u043e\u0435\u041b\u0438\u0446\u043e.\u0424\u0418\u041e = &\u0424\u0418\u041e
    """
    query.SetParameter("\u0424\u0418\u041e", fio)
    selection = query.Execute().Choose()
    if selection.Next():
        return _safe_getattr(selection, "\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a")
    return None

def _find_latest_hire_doc(connector, fio: str, inn_org: str = "", hire_date=None):
    query = connector.NewObject("\u0417\u0430\u043f\u0440\u043e\u0441")
    company_filter = ""
    if inn_org:
        company_filter = """
            \u0418 \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u041e\u0440\u0433\u0430\u043d\u0438\u0437\u0430\u0446\u0438\u044f.\u0418\u041d\u041d = &\u0418\u041d\u041d\u041e\u0440\u0433"""
    hire_date_filter = ""
    if hire_date is not None:
        hire_date_filter = """
            \u0418 \u041d\u0410\u0427\u0410\u041b\u041e\u041f\u0415\u0420\u0418\u041e\u0414\u0410(\u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0414\u0430\u0442\u0430, \u0414\u0415\u041d\u042c) = &\u0414\u0430\u0442\u0430\u041f\u0440\u0438\u0435\u043c\u0430"""

    query.Text = f"""
        \u0412\u042b\u0411\u0420\u0410\u0422\u042c \u041f\u0415\u0420\u0412\u042b\u0415 1
            \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0421\u0441\u044b\u043b\u043a\u0430 \u041a\u0410\u041a \u0421\u0441\u044b\u043b\u043a\u0430
        \u0418\u0417
            \u0414\u043e\u043a\u0443\u043c\u0435\u043d\u0442.\u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443 \u041a\u0410\u041a \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443
        \u0413\u0414\u0415
            \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0424\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u043e\u0435\u041b\u0438\u0446\u043e.\u0424\u0418\u041e = &\u0424\u0418\u041e{company_filter}{hire_date_filter}
        \u0423\u041f\u041e\u0420\u042f\u0414\u041e\u0427\u0418\u0422\u042c \u041f\u041e
            \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0414\u0430\u0442\u0430 \u0423\u0411\u042b\u0412,
            \u041f\u0440\u0438\u0435\u043c\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443.\u0421\u0441\u044b\u043b\u043a\u0430 \u0423\u0411\u042b\u0412
    """
    query.SetParameter("\u0424\u0418\u041e", fio)
    if inn_org:
        query.SetParameter("\u0418\u041d\u041d\u041e\u0440\u0433", inn_org)
    if hire_date is not None:
        query.SetParameter("\u0414\u0430\u0442\u0430\u041f\u0440\u0438\u0435\u043c\u0430", hire_date)

    try:
        selection = query.Execute().Choose()
    except Exception:
        return None

    if selection.Next():
        return _safe_getattr(selection, "\u0421\u0441\u044b\u043b\u043a\u0430")
    return None

def _read_bank_details_from_hire_doc(connector, fio: str, inn_org: str = "", hire_date=None) -> dict:
    hire_doc_ref = _find_latest_hire_doc(connector, fio, inn_org=inn_org, hire_date=hire_date)
    if hire_doc_ref is None:
        return {}

    get_object_method = _safe_getattr(hire_doc_ref, "\u041f\u043e\u043b\u0443\u0447\u0438\u0442\u044c\u041e\u0431\u044a\u0435\u043a\u0442")
    if get_object_method is None:
        return {}

    try:
        hire_doc_obj = get_object_method()
    except Exception:
        return {}

    direct_bank_data, _ = _read_bank_details_from_owner_attributes(hire_doc_obj)
    if direct_bank_data:
        _log_bank_source(
            "hire_doc_owner_attributes",
            direct_bank_data,
            fio=fio,
            inn_org=inn_org,
            hire_date=hire_date,
        )
        return direct_bank_data

    metadata_bank_data, metadata_attempts = _extract_bank_data_by_metadata_scan(hire_doc_obj)
    if metadata_attempts:
        _log_bank_source(
            "hire_doc_metadata_scan_attempts",
            {"attempts_count": len(metadata_attempts)},
            fio=fio,
            inn_org=inn_org,
            hire_date=hire_date,
            attempts=metadata_attempts,
        )
    if metadata_bank_data:
        _log_bank_source(
            "hire_doc_metadata_scan",
            metadata_bank_data,
            fio=fio,
            inn_org=inn_org,
            hire_date=hire_date,
        )
        return metadata_bank_data

    candidate_attrs = (
        "\u0421\u043f\u043e\u0441\u043e\u0431\u0412\u044b\u043f\u043b\u0430\u0442\u044b\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u0412\u044b\u043f\u043b\u0430\u0442\u0430\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u0412\u044b\u043f\u043b\u0430\u0442\u044b\u0418\u0423\u0447\u0435\u0442\u0417\u0430\u0442\u0440\u0430\u0442",
        "\u041c\u0435\u0441\u0442\u0430\u0412\u044b\u043f\u043b\u0430\u0442\u044b",
        "\u0414\u0435\u0442\u0430\u043b\u0438\u0437\u0430\u0446\u0438\u044f\u041c\u0435\u0441\u0442\u0412\u044b\u043f\u043b\u0430\u0442",
        "\u0414\u0435\u0442\u0430\u043b\u0438\u0437\u0430\u0446\u0438\u044f\u041c\u0435\u0441\u0442\u0412\u044b\u043f\u043b\u0430\u0442\u044b\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0430\u0412\u044b\u043f\u043b\u0430\u0442\u044b\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u043d\u044b\u0439\u041f\u0440\u043e\u0435\u043a\u0442",
        "\u0411\u0430\u043d\u043a\u043e\u0432\u0441\u043a\u0438\u0439\u0421\u0447\u0435\u0442",
        "\u0411\u0430\u043d\u043a\u043e\u0432\u0441\u043a\u0438\u0439\u0421\u0447\u0451\u0442",
        "\u0411\u0430\u043d\u043a\u043e\u0432\u0441\u043a\u0438\u0439\u0421\u0447\u0435\u0442\u0414\u043b\u044f\u041f\u0435\u0440\u0435\u0447\u0438\u0441\u043b\u0435\u043d\u0438\u044f\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u0421\u0447\u0435\u0442\u0414\u043b\u044f\u041f\u0435\u0440\u0435\u0447\u0438\u0441\u043b\u0435\u043d\u0438\u044f\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u044b",
        "\u0417\u0430\u0440\u043f\u043b\u0430\u0442\u043d\u044b\u0439\u0421\u0447\u0435\u0442",
        "\u041b\u0438\u0446\u0435\u0432\u043e\u0439\u0421\u0447\u0435\u0442",
        "\u041b\u0438\u0446\u0435\u0432\u043e\u0439\u0421\u0447\u0435\u0442\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u0430",
    )

    for attr_name in candidate_attrs:
        attr_value = _safe_getattr(hire_doc_obj, attr_name)
        if attr_value is None:
            continue
        bank_data = _extract_bank_data_from_value(attr_value)
        if bank_data.get("\u041d\u043e\u043c\u0435\u0440\u0421\u0447\u0435\u0442\u0430") or bank_data.get("\u0420\u0430\u0441\u0447\u0435\u0442\u043d\u044b\u0439\u0421\u0447\u0435\u0442") or bank_data.get("\u0411\u0430\u043d\u043a"):
            _log_bank_source(
                "hire_doc_attr",
                bank_data,
                fio=fio,
                inn_org=inn_org,
                hire_date=hire_date,
                attr_name=attr_name,
            )
            return bank_data

    return {}

def _find_counterparties_by_fio(connector, fio: str) -> list:
    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 5
            Контрагенты.Ссылка КАК Контрагент
        ИЗ
            Справочник.Контрагенты КАК Контрагенты
        ГДЕ
            Контрагенты.Наименование = &ФИО
            ИЛИ Контрагенты.Наименование ПОДОБНО &ФИОПодобно
    """
    query.SetParameter("ФИО", fio)
    query.SetParameter("ФИОПодобно", f"%{fio}%")
    selection = query.Execute().Choose()

    result = []
    while selection.Next():
        result.append(selection.Контрагент)
    return result


def _find_organization(connector, inn_org: str):
    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Организации.Ссылка КАК Организация
        ИЗ
            Справочник.Организации КАК Организации
        ГДЕ
            Организации.ИНН = &ИНН
    """
    query.SetParameter("ИНН", inn_org)
    selection = query.Execute().Choose()
    if selection.Next():
        return selection.Организация
    return None


def _find_organization_from_hire_doc(
    connector,
    fio: str,
    employee_ref,
    phys_person,
    inn_org: str,
    hire_date=None,
) -> tuple[object | None, str]:
    """Return the organization reference stored in the selected hire document."""
    filters = []
    parameters = {}
    if employee_ref is not None:
        filters.append("ПриемНаРаботу.Сотрудник = &Сотрудник")
        parameters["Сотрудник"] = employee_ref
    elif phys_person is not None:
        filters.append("ПриемНаРаботу.ФизическоеЛицо = &ФизическоеЛицо")
        parameters["ФизическоеЛицо"] = phys_person
    else:
        filters.append(
            "(ПриемНаРаботу.ФизическоеЛицо.ФИО = &ФИО "
            "ИЛИ ПриемНаРаботу.Сотрудник.Наименование = &ФИО)"
        )
        parameters["ФИО"] = fio

    if inn_org:
        filters.append("ПриемНаРаботу.Организация.ИНН = &ИННОрг")
        parameters["ИННОрг"] = inn_org

    def find(use_hire_date: bool):
        date_filter = " И ПриемНаРаботу.ДатаПриема = &ДатаПриема" if use_hire_date else ""
        query = connector.NewObject("Запрос")
        query.Text = f"""
            ВЫБРАТЬ ПЕРВЫЕ 1
                ПриемНаРаботу.Организация КАК Организация
            ИЗ
                Документ.ПриемНаРаботу КАК ПриемНаРаботу
            ГДЕ
                {' И '.join(filters)}{date_filter}
            УПОРЯДОЧИТЬ ПО
                ПриемНаРаботу.Дата УБЫВ,
                ПриемНаРаботу.Ссылка УБЫВ
        """
        for name, value in parameters.items():
            query.SetParameter(name, value)
        if use_hire_date:
            query.SetParameter("ДатаПриема", hire_date)
        try:
            selection = query.Execute().Choose()
        except Exception:
            return None
        return selection.Организация if selection.Next() else None

    if hire_date is not None:
        organization = find(use_hire_date=True)
        if organization is not None:
            return organization, "hire_document_by_hire_date"

    organization = find(use_hire_date=False)
    if organization is not None:
        return organization, "latest_hire_document"

    return None, "catalog_by_inn"


def _read_bank_account_by_owner(connector, owner, catalog_name: str):
    if owner is None:
        return {}

    query = connector.NewObject("Запрос")
    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            БанковскиеСчета.Ссылка КАК СчетСсылка,
            БанковскиеСчета.Банк КАК БанкСсылка,
            БанковскиеСчета.Банк.Наименование КАК Банк,
            БанковскиеСчета.Банк.ИНН КАК ИннБанка,
            БанковскиеСчета.Банк.Адрес КАК АдресБанка,
            БанковскиеСчета.Банк.Код КАК БикБанка,
            БанковскиеСчета.Банк.КоррСчет КАК КоррСчет,
            БанковскиеСчета.НомерСчета КАК НомерСчета,
            БанковскиеСчета.Наименование КАК ПредставлениеСчета,
            БанковскиеСчета.ДополнительныеРеквизиты КАК ДополнительныеРеквизиты
        ИЗ
            Справочник.{catalog_name} КАК БанковскиеСчета
        ГДЕ
            БанковскиеСчета.Владелец = &Владелец
        УПОРЯДОЧИТЬ ПО
            БанковскиеСчета.Ссылка УБЫВ
    """
    query.SetParameter("Владелец", owner)

    try:
        selection = query.Execute().Choose()
    except Exception:
        return {}

    if not selection.Next():
        return {}

    result = {
        "Банк": _coerce_1c_scalar_text(selection.Банк),
        "ИннБанка": _coerce_1c_scalar_text(selection.ИннБанка),
        "АдресБанка": _coerce_1c_scalar_text(selection.АдресБанка),
        "НомерСчета": _coerce_1c_scalar_text(selection.НомерСчета),
        "РасчетныйСчет": _coerce_1c_scalar_text(selection.НомерСчета),
        "БикБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "БикБанка")),
        "КоррСчет": _coerce_1c_scalar_text(_safe_getattr(selection, "КоррСчет")),
    }

    _fill_bank_fields(result, selection.ДополнительныеРеквизиты, "БикБанка", "КоррСчет", "РасчетныйСчет")

    # В обычном сценарии корр. счет часто хранится напрямую в счете или в объекте банка,
    # а не в дополнительных реквизитах. Добираем недостающие поля оттуда.
    if not result["БикБанка"] or not result["КоррСчет"]:
        for source_value in (_safe_getattr(selection, "СчетСсылка"), _safe_getattr(selection, "БанкСсылка")):
            source_data = _extract_bank_data_from_value(source_value)
            if source_data.get("БикБанка") and not result["БикБанка"]:
                result["БикБанка"] = source_data["БикБанка"]
            if source_data.get("КоррСчет") and not result["КоррСчет"]:
                result["КоррСчет"] = source_data["КоррСчет"]
            if source_data.get("Банк") and not result["Банк"]:
                result["Банк"] = source_data["Банк"]
            if source_data.get("ИннБанка") and not result["ИннБанка"]:
                result["ИннБанка"] = source_data["ИннБанка"]
            if source_data.get("АдресБанка") and not result["АдресБанка"]:
                result["АдресБанка"] = source_data["АдресБанка"]

    if not result["БикБанка"]:
        account_presentation = _coerce_1c_scalar_text(selection.ПредставлениеСчета)
        result["БикБанка"] = account_presentation.split(",")[0].strip()

    return result


def _bank_account_selection_to_dict(selection) -> dict:
    result = {
        "Банк": _coerce_1c_scalar_text(_safe_getattr(selection, "Банк")),
        "ИннБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "ИннБанка")),
        "АдресБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "АдресБанка")),
        "НомерСчета": _coerce_1c_scalar_text(_safe_getattr(selection, "НомерСчета")),
        "РасчетныйСчет": _coerce_1c_scalar_text(_safe_getattr(selection, "НомерСчета")),
        "БикБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "БикБанка")),
        "КоррСчет": _coerce_1c_scalar_text(_safe_getattr(selection, "КоррСчет")),
    }

    additional_requisites = _safe_getattr(selection, "ДополнительныеРеквизиты")
    _fill_bank_fields(result, additional_requisites, "БикБанка", "КоррСчет", "РасчетныйСчет")

    if not result["БикБанка"] or not result["КоррСчет"]:
        for source_value in (_safe_getattr(selection, "СчетСсылка"), _safe_getattr(selection, "БанкСсылка")):
            source_data = _extract_bank_data_from_value(source_value)
            for key in ("БикБанка", "КоррСчет", "Банк", "ИннБанка", "АдресБанка"):
                if source_data.get(key) and not result[key]:
                    result[key] = source_data[key]

    if result["РасчетныйСчет"] and not result["НомерСчета"]:
        result["НомерСчета"] = result["РасчетныйСчет"]
    if result["НомерСчета"] and not result["РасчетныйСчет"]:
        result["РасчетныйСчет"] = result["НомерСчета"]

    if not result["БикБанка"]:
        account_presentation = _coerce_1c_scalar_text(_safe_getattr(selection, "ПредставлениеСчета"))
        result["БикБанка"] = account_presentation.split(",")[0].strip()

    return result if any(result.values()) else {}


def _read_bank_account_by_owner_presentation(connector, fio: str, catalog_name: str):
    if not fio:
        return {}

    query = connector.NewObject("Запрос")
    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            БанковскиеСчета.Ссылка КАК СчетСсылка,
            БанковскиеСчета.Банк КАК БанкСсылка,
            БанковскиеСчета.Банк.Наименование КАК Банк,
            БанковскиеСчета.Банк.ИНН КАК ИннБанка,
            БанковскиеСчета.Банк.Адрес КАК АдресБанка,
            БанковскиеСчета.Банк.Код КАК БикБанка,
            БанковскиеСчета.Банк.КоррСчет КАК КоррСчет,
            БанковскиеСчета.НомерСчета КАК НомерСчета,
            БанковскиеСчета.Наименование КАК ПредставлениеСчета,
            БанковскиеСчета.ДополнительныеРеквизиты КАК ДополнительныеРеквизиты
        ИЗ
            Справочник.{catalog_name} КАК БанковскиеСчета
        ГДЕ
            ПРЕДСТАВЛЕНИЕ(БанковскиеСчета.Владелец) = &ФИО
            ИЛИ ПРЕДСТАВЛЕНИЕ(БанковскиеСчета.Владелец) ПОДОБНО &ФИОПодобно
        УПОРЯДОЧИТЬ ПО
            БанковскиеСчета.Ссылка УБЫВ
    """
    query.SetParameter("ФИО", fio)
    query.SetParameter("ФИОПодобно", f"%{fio}%")

    try:
        selection = query.Execute().Choose()
    except Exception:
        return {}

    if not selection.Next():
        return {}

    return _bank_account_selection_to_dict(selection)


def _normalize_fio_match_text(value: str) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").lower().split())


def _read_bank_account_by_owner_name(connector, fio: str, catalog_name: str):
    if not fio:
        return {}

    query = connector.NewObject("Запрос")
    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            БанковскиеСчета.Ссылка КАК СчетСсылка,
            БанковскиеСчета.Банк КАК БанкСсылка,
            БанковскиеСчета.Банк.Наименование КАК Банк,
            БанковскиеСчета.Банк.ИНН КАК ИннБанка,
            БанковскиеСчета.Банк.Адрес КАК АдресБанка,
            БанковскиеСчета.Банк.Код КАК БикБанка,
            БанковскиеСчета.Банк.КоррСчет КАК КоррСчет,
            БанковскиеСчета.НомерСчета КАК НомерСчета,
            БанковскиеСчета.Наименование КАК ПредставлениеСчета,
            БанковскиеСчета.ДополнительныеРеквизиты КАК ДополнительныеРеквизиты
        ИЗ
            Справочник.{catalog_name} КАК БанковскиеСчета
        ГДЕ
            БанковскиеСчета.Владелец.Наименование = &ФИО
            ИЛИ БанковскиеСчета.Владелец.Наименование ПОДОБНО &ФИОПодобно
        УПОРЯДОЧИТЬ ПО
            БанковскиеСчета.Ссылка УБЫВ
    """
    query.SetParameter("ФИО", fio)
    query.SetParameter("ФИОПодобно", f"%{fio}%")

    try:
        selection = query.Execute().Choose()
    except Exception:
        return {}

    if not selection.Next():
        return {}

    return _bank_account_selection_to_dict(selection)


def _scan_bank_accounts_by_owner_fio(connector, fio: str, catalog_name: str):
    if not fio:
        return {}

    query = connector.NewObject("Запрос")
    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1000
            БанковскиеСчета.Ссылка КАК СчетСсылка,
            БанковскиеСчета.Владелец КАК Владелец,
            БанковскиеСчета.Банк КАК БанкСсылка,
            БанковскиеСчета.Банк.Наименование КАК Банк,
            БанковскиеСчета.Банк.ИНН КАК ИннБанка,
            БанковскиеСчета.Банк.Адрес КАК АдресБанка,
            БанковскиеСчета.Банк.Код КАК БикБанка,
            БанковскиеСчета.Банк.КоррСчет КАК КоррСчет,
            БанковскиеСчета.НомерСчета КАК НомерСчета,
            БанковскиеСчета.Наименование КАК ПредставлениеСчета,
            БанковскиеСчета.ДополнительныеРеквизиты КАК ДополнительныеРеквизиты
        ИЗ
            Справочник.{catalog_name} КАК БанковскиеСчета
        УПОРЯДОЧИТЬ ПО
            БанковскиеСчета.Ссылка УБЫВ
    """

    try:
        selection = query.Execute().Choose()
    except Exception:
        return {}

    fio_match = _normalize_fio_match_text(fio)
    scanned_count = 0
    while selection.Next():
        scanned_count += 1
        owner_text = _normalize_fio_match_text(_describe_owner(_safe_getattr(selection, "Владелец")))
        if fio_match not in owner_text:
            continue

        bank_data = _bank_account_selection_to_dict(selection)
        if _has_bank_account_number(bank_data):
            print(
                "bank_details_old_catalog_scan_match: "
                f"fio={fio!r}; catalog_name={catalog_name!r}; owner={owner_text!r}; scanned={scanned_count}"
            )
            return bank_data

    print(
        "bank_details_old_catalog_scan_not_found: "
        f"fio={fio!r}; catalog_name={catalog_name!r}; scanned={scanned_count}"
    )
    return {}


def _get_employee_counterparty_bank_details(connector, fio: str, hire_date=None, inn_org: str = "") -> dict:
    counterparties = _find_counterparties_by_fio(connector, fio)
    for counterparty in counterparties:
        bank_data = _read_bank_account_by_owner(connector, counterparty, "БанковскиеСчетаКонтрагентов")
        if _has_bank_account_number(bank_data):
            _log_bank_source(
                "employee_bank_final_counterparty_catalog",
                bank_data,
                fio=fio,
                inn_org=inn_org,
                hire_date=hire_date,
                owner=_describe_owner(counterparty),
                catalog_name="БанковскиеСчетаКонтрагентов",
            )
            return bank_data
    return {}


def _get_employee_bank_details_by_old_logic(connector, fio: str, phys_person, employee_ref, hire_date=None, inn_org: str = "") -> dict:
    owner_candidates = []
    for owner in (employee_ref, phys_person):
        if owner is None:
            continue
        owner_candidates.append(owner)
        for attr_name in ("Контрагент", "ФизическоеЛицо", "ФизЛицо"):
            linked_owner = _safe_getattr(owner, attr_name)
            if linked_owner is not None:
                owner_candidates.append(linked_owner)
    owner_candidates.extend(_find_counterparties_by_fio(connector, fio))

    seen = set()
    for owner in owner_candidates:
        owner_key = str(owner)
        if owner_key in seen:
            continue
        seen.add(owner_key)

        direct_bank_data, _ = _read_bank_details_from_owner_attributes(owner)
        if direct_bank_data:
            _log_bank_source("employee_bank_final_old_owner_attrs", direct_bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date, owner=_describe_owner(owner))
            return direct_bank_data

        for catalog_name in ("БанковскиеСчетаФизическихЛиц", "БанковскиеСчетаКонтрагентов"):
            bank_data = _read_bank_account_by_owner(connector, owner, catalog_name)
            if _has_bank_account_number(bank_data):
                _log_bank_source("employee_bank_final_old_catalog", bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date, owner=_describe_owner(owner), catalog_name=catalog_name)
                return bank_data

    for catalog_name in ("БанковскиеСчетаКонтрагентов", "БанковскиеСчетаФизическихЛиц"):
        bank_data = _read_bank_account_by_owner_presentation(connector, fio, catalog_name)
        if _has_bank_account_number(bank_data):
            _log_bank_source(
                "employee_bank_final_old_catalog_by_fio",
                bank_data,
                fio=fio,
                inn_org=inn_org,
                hire_date=hire_date,
                catalog_name=catalog_name,
            )
            return bank_data

        bank_data = _read_bank_account_by_owner_name(connector, fio, catalog_name)
        if _has_bank_account_number(bank_data):
            _log_bank_source(
                "employee_bank_final_old_catalog_by_owner_name",
                bank_data,
                fio=fio,
                inn_org=inn_org,
                hire_date=hire_date,
                catalog_name=catalog_name,
            )
            return bank_data

        bank_data = _scan_bank_accounts_by_owner_fio(connector, fio, catalog_name)
        if _has_bank_account_number(bank_data):
            _log_bank_source(
                "employee_bank_final_old_catalog_scan",
                bank_data,
                fio=fio,
                inn_org=inn_org,
                hire_date=hire_date,
                catalog_name=catalog_name,
            )
            return bank_data

    return {}


def _bank_data_from_payroll_selection(selection) -> dict:
    account_number = _normalize_account_number(_safe_getattr(selection, "НомерСчета"))
    bank_name = _coerce_1c_scalar_text(_safe_getattr(selection, "Банк"))
    bik = _coerce_1c_scalar_text(_safe_getattr(selection, "БикБанка"))
    corr_account = _coerce_1c_scalar_text(_safe_getattr(selection, "КоррСчет"))
    bank_inn = _coerce_1c_scalar_text(_safe_getattr(selection, "ИннБанка"))
    bank_address = _coerce_1c_scalar_text(_safe_getattr(selection, "АдресБанка"))

    result = {
        "Банк": bank_name,
        "ИннБанка": bank_inn,
        "АдресБанка": bank_address,
        "НомерСчета": account_number,
        "РасчетныйСчет": account_number,
        "БикБанка": bik,
        "КоррСчет": corr_account,
    }

    return result if any(result.values()) else {}


def _build_payroll_candidate(selection, bank_data: dict, source: str) -> dict:
    account_number = _normalize_account_number(_safe_getattr(selection, "НомерСчета"))
    raw_open_date = (
        _safe_getattr(selection, "ДатаОткрытия")
        or _safe_getattr(selection, "ДатаОткрытияСчета")
        or _safe_getattr(selection, "ДатаОткрытияЛицевогоСчета")
        or _safe_getattr(selection, "Открыт")
    )
    open_date = _coerce_datetime(raw_open_date)
    period_date = _coerce_datetime(_safe_getattr(selection, "Период"))
    project_name = _coerce_1c_scalar_text(_safe_getattr(selection, "ЗарплатныйПроект"))

    candidate_data = dict(bank_data or {})
    if account_number:
        candidate_data["НомерСчета"] = account_number
        candidate_data["РасчетныйСчет"] = account_number

    return {
        "source": source,
        "account_number": account_number,
        "project_name": project_name,
        "open_date": open_date,
        "period_date": period_date,
        "bank_data": candidate_data,
    }


def _choose_best_payroll_candidate(candidates: list[dict], hire_date=None) -> dict:
    if not candidates:
        return {}

    deduped_candidates = []
    seen = set()
    for candidate in candidates:
        account_number = candidate.get("account_number") or ""
        project_name = candidate.get("project_name") or ""
        bank_name = (candidate.get("bank_data") or {}).get("Банк", "")
        key = (account_number, project_name, bank_name)
        if key in seen:
            continue
        seen.add(key)
        deduped_candidates.append(candidate)

    def _score(candidate: dict):
        account_number = candidate.get("account_number") or ""
        bank_data = candidate.get("bank_data") or {}
        open_date = candidate.get("open_date")
        period_date = candidate.get("period_date")
        reference_date = open_date or period_date or datetime.min
        has_bank_payload = any(bank_data.values())
        has_open_date = open_date is not None

        if hire_date is not None and reference_date != datetime.min:
            distance_days = abs((hire_date - reference_date).days)
            is_not_future = reference_date <= hire_date
            return (
                1 if account_number else 0,
                1 if has_bank_payload else 0,
                1 if has_open_date else 0,
                1 if is_not_future else 0,
                -distance_days,
                reference_date,
            )

        return (
            1 if account_number else 0,
            1 if has_bank_payload else 0,
            1 if has_open_date else 0,
            reference_date,
        )

    best_candidate = max(deduped_candidates, key=_score)
    return best_candidate.get("bank_data") or {}


def _query_current_payroll_project_account(connector, phys_person, organization=None) -> dict:
    if phys_person is None:
        return {}

    where_clauses = ["ЛицевыеСчета.ФизическоеЛицо = &ФизЛицо"]
    if organization is not None:
        where_clauses.append("ЛицевыеСчета.Организация = &Организация")

    query = connector.NewObject("Запрос")
    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            ЛицевыеСчета.НомерЛицевогоСчета КАК НомерСчета,
            ЛицевыеСчета.ДатаОткрытияЛицевогоСчета КАК ДатаОткрытияЛицевогоСчета,
            ЛицевыеСчета.ЗарплатныйПроект.Наименование КАК ЗарплатныйПроект,
            ЛицевыеСчета.СпособЗачисления КАК СпособЗачисления,
            ЛицевыеСчета.БанковскийСчет КАК БанковскийСчет,
            ЛицевыеСчета.БанковскаяКарта КАК БанковскаяКарта,
            ЛицевыеСчета.Телефон КАК Телефон,
            ЛицевыеСчета.ДокументОснование КАК ДокументОснование,
            ЛицевыеСчета.ЗарплатныйПроект.Банк КАК БанкСсылка,
            ЛицевыеСчета.ЗарплатныйПроект.Банк.Наименование КАК Банк,
            ЛицевыеСчета.ЗарплатныйПроект.Банк.ИНН КАК ИннБанка,
            ЛицевыеСчета.ЗарплатныйПроект.Банк.Адрес КАК АдресБанка,
            ЛицевыеСчета.ЗарплатныйПроект.Банк.Код КАК БикБанка,
            ЛицевыеСчета.ЗарплатныйПроект.Банк.КоррСчет КАК КоррСчет
        ИЗ
            РегистрСведений.ЛицевыеСчетаСотрудниковПоЗарплатнымПроектам КАК ЛицевыеСчета
        ГДЕ
            {" И ".join(where_clauses)}
        УПОРЯДОЧИТЬ ПО
            ЛицевыеСчета.ДатаОткрытияЛицевогоСчета УБЫВ
    """
    query.SetParameter("ФизЛицо", phys_person)
    if organization is not None:
        query.SetParameter("Организация", organization)

    try:
        selection = query.Execute().Choose()
    except Exception as e:
        print(f"Error occurred while executing query: {e}")
        return {}

    if not selection.Next():
        return {}

    account_number = _normalize_account_number(_safe_getattr(selection, "НомерСчета"))
    if not account_number:
        return {}

    bank_data = {
        "Банк": _coerce_1c_scalar_text(_safe_getattr(selection, "Банк")),
        "ИннБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "ИннБанка")),
        "АдресБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "АдресБанка")),
        "НомерСчета": account_number,
        "РасчетныйСчет": account_number,
        "БикБанка": _coerce_1c_scalar_text(_safe_getattr(selection, "БикБанка")),
        "КоррСчет": _coerce_1c_scalar_text(_safe_getattr(selection, "КоррСчет")),
    }

    for attr_name in ("БанковскийСчет", "БанковскаяКарта", "ДокументОснование", "БанкСсылка"):
        fallback_data = _extract_bank_data_from_value(_safe_getattr(selection, attr_name))
        if fallback_data:
            bank_data.update({key: value for key, value in fallback_data.items() if value and not bank_data.get(key)})

    return bank_data


def _read_current_payroll_project_account(connector, phys_person, organization=None, hire_date=None) -> dict:
    bank_data = _query_current_payroll_project_account(
        connector,
        phys_person=phys_person,
        organization=organization,
    )
    if bank_data:
        return bank_data

    if organization is not None:
        return _query_current_payroll_project_account(
            connector,
            phys_person=phys_person,
            organization=None,
        )

    return {}


def _read_employee_payroll_project_bank_details(connector, employee_ref, phys_person, organization=None, hire_date=None) -> dict:
    return _read_current_payroll_project_account(
        connector,
        phys_person=phys_person,
        organization=organization,
        hire_date=hire_date,
    )


def _read_employee_payroll_project_bank_details_actual(connector, employee_ref, phys_person, organization=None, hire_date=None) -> dict:
    return _read_employee_payroll_project_bank_details(
        connector,
        employee_ref,
        phys_person,
        organization=organization,
        hire_date=hire_date,
    )


def _get_employee_bank_details(
    connector,
    fio: str,
    phys_person,
    employee_ref,
    hire_date=None,
    inn_org: str = "",
    organization=None,
) -> dict:
    if phys_person is None and employee_ref is not None:
        phys_person = _safe_getattr(employee_ref, "ФизическоеЛицо")

    if organization is None and inn_org:
        organization = _find_organization(connector, inn_org)
    payroll_bank_data = _read_employee_payroll_project_bank_details_actual(
        connector,
        employee_ref,
        phys_person,
        organization=organization,
        hire_date=hire_date,
    )
    if _has_bank_account_number(payroll_bank_data):
        _log_bank_source("employee_bank_final_payroll", payroll_bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date)
        return payroll_bank_data

    old_logic_bank_data = _get_employee_bank_details_by_old_logic(
        connector,
        fio,
        phys_person,
        employee_ref,
        hire_date=hire_date,
        inn_org=inn_org,
    )
    if old_logic_bank_data:
        return old_logic_bank_data

    counterparty_bank_data = _get_employee_counterparty_bank_details(
        connector,
        fio,
        hire_date=hire_date,
        inn_org=inn_org,
    )
    if counterparty_bank_data:
        return counterparty_bank_data

    hire_doc_bank_data = _read_bank_details_from_hire_doc(connector, fio, inn_org=inn_org, hire_date=hire_date)
    if hire_doc_bank_data:
        _log_bank_source("employee_bank_final_hire_doc", hire_doc_bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date)
        return hire_doc_bank_data

    owner_candidates = []
    if employee_ref is not None:
        owner_candidates.append(employee_ref)
    if phys_person is not None:
        owner_candidates.append(phys_person)
    owner_candidates.extend(_find_counterparties_by_fio(connector, fio))

    seen = set()
    deduped_owner_candidates = []
    for owner in owner_candidates:
        owner_key = str(owner)
        if owner_key in seen:
            continue
        seen.add(owner_key)
        deduped_owner_candidates.append(owner)

    for owner in deduped_owner_candidates:
        direct_bank_data, _ = _read_bank_details_from_owner_attributes(owner)
        if direct_bank_data:
            _log_bank_source("employee_bank_final_owner_attrs", direct_bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date, owner=_describe_owner(owner))
            return direct_bank_data
        for catalog_name in (
            "\u0411\u0430\u043d\u043a\u043e\u0432\u0441\u043a\u0438\u0435\u0421\u0447\u0435\u0442\u0430\u0424\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u0438\u0445\u041b\u0438\u0446",
            "\u0411\u0430\u043d\u043a\u043e\u0432\u0441\u043a\u0438\u0435\u0421\u0447\u0435\u0442\u0430\u041a\u043e\u043d\u0442\u0440\u0430\u0433\u0435\u043d\u0442\u043e\u0432",
        ):
            bank_data = _read_bank_account_by_owner(connector, owner, catalog_name)
            if bank_data.get("\u041d\u043e\u043c\u0435\u0440\u0421\u0447\u0435\u0442\u0430") or bank_data.get("\u0420\u0430\u0441\u0447\u0435\u0442\u043d\u044b\u0439\u0421\u0447\u0435\u0442") or bank_data.get("\u0411\u0430\u043d\u043a"):
                _log_bank_source("employee_bank_final_catalog", bank_data, fio=fio, inn_org=inn_org, hire_date=hire_date, owner=_describe_owner(owner), catalog_name=catalog_name)
                return bank_data
    return {}

def _get_org_bank_details(connector, organization, inn_org: str, debug_attempts: list | None = None) -> dict:
    owner_candidates = []
    if organization is not None:
        owner_candidates.append(organization)

    if inn_org:
        owner_candidates.extend(_find_counterparties_by_fio(connector, inn_org))

        query = connector.NewObject("Запрос")
        query.Text = """
            ВЫБРАТЬ ПЕРВЫЕ 5
                Контрагенты.Ссылка КАК Контрагент
            ИЗ
                Справочник.Контрагенты КАК Контрагенты
            ГДЕ
                Контрагенты.ИНН = &ИНН
        """
        query.SetParameter("ИНН", inn_org)
        try:
            selection = query.Execute().Choose()
            while selection.Next():
                owner_candidates.append(selection.Контрагент)
        except Exception:
            pass

    seen = set()
    deduped_owners = []
    for owner in owner_candidates:
        owner_key = str(owner)
        if owner_key in seen:
            continue
        seen.add(owner_key)
        deduped_owners.append(owner)

    for owner in deduped_owners:
        direct_bank_data, direct_attempts = _read_bank_details_from_owner_attributes(owner)
        if debug_attempts is not None:
            debug_attempts.append({
                "owner": _describe_owner(owner),
                "source": "attributes",
                "data": direct_bank_data,
                "requisites": _describe_requisites_table(_safe_getattr(owner, "ДополнительныеРеквизиты")),
            })
        if direct_bank_data:
            return direct_bank_data
        for catalog_name in ("БанковскиеСчетаОрганизаций", "БанковскиеСчетаКонтрагентов"):
            bank_data = _read_bank_account_by_owner(connector, owner, catalog_name)
            if debug_attempts is not None:
                debug_attempts.append({
                    "owner": _describe_owner(owner),
                    "source": catalog_name,
                    "data": bank_data,
                })
            if bank_data.get("НомерСчета") or bank_data.get("РасчетныйСчет") or bank_data.get("Банк"):
                return bank_data
    return {}


def _print_organization_bank_details_debug(
    inn_org: str,
    organization,
    organization_source: str,
    source_values: dict,
    attempts: list[dict],
    final_values: dict,
) -> None:
    """Print one focused diagnostic block for organization banking requisites."""
    try:
        print("=== ORGANIZATION BANK DETAILS DEBUG START ===")
        print(f"inn_org: {inn_org!r}")
        print(f"organization: {_describe_owner(organization)!r}")
        print(f"organization_source: {organization_source!r}")
        print(f"values_before_enrichment: {source_values!r}")
        for index, attempt in enumerate(attempts, start=1):
            print(
                "organization_bank_attempt: "
                f"index={index}; owner={attempt['owner']!r}; "
                f"source={attempt['source']!r}; data={attempt['data']!r}"
            )
            if attempt.get("requisites"):
                print(f"organization_additional_requisites: {attempt['requisites']!r}")
        print(f"values_after_enrichment: {final_values!r}")
        print("=== ORGANIZATION BANK DETAILS DEBUG END ===")
    except Exception as error:
        print(f"organization_bank_debug_print_error: {error}")


def _collect_bank_details_debug(connector, fio: str, inn_org: str = "", hire_date=None) -> dict:
    debug = {
        "fio": fio,
        "inn_org": inn_org,
        "hire_date": hire_date,
        "employee": "",
        "phys_person": "",
        "employee_bank_final": {},
        "employee_counterparties": [],
        "hire_doc_attempt": {},
        "payroll_actual_attempt": {},
        "organization": "",
        "organization_counterparties_by_inn": [],
        "employee_attempts": [],
        "organization_attempts": [],
    }

    employee_ref = _find_employee(connector, fio, inn_org=inn_org, hire_date=hire_date)
    debug["employee"] = _describe_owner(employee_ref)

    phys_person = _find_physical_person(connector, fio)
    debug["phys_person"] = _describe_owner(phys_person)
    organization = _find_organization(connector, inn_org) if inn_org else None
    debug["organization"] = _describe_owner(organization)

    hire_doc_bank_data = _read_bank_details_from_hire_doc(connector, fio, inn_org=inn_org, hire_date=hire_date)
    debug["hire_doc_attempt"] = {
        "source": "hire_doc_current",
        "found": bool(hire_doc_bank_data),
        "data": hire_doc_bank_data,
    }

    payroll_bank_data = _read_employee_payroll_project_bank_details(
        connector,
        employee_ref,
        phys_person,
        organization=organization,
        hire_date=hire_date,
    )
    debug["employee_attempts"].append({
        "owner": _describe_owner(employee_ref or phys_person),
        "source": "payroll_project_register",
        "found": bool(payroll_bank_data),
        "data": payroll_bank_data,
    })
    payroll_actual_bank_data = _read_employee_payroll_project_bank_details_actual(
        connector,
        employee_ref,
        phys_person,
        organization=organization,
        hire_date=hire_date,
    )
    debug["payroll_actual_attempt"] = {
        "owner": _describe_owner(employee_ref or phys_person),
        "source": "payroll_project_register_actual",
        "found": bool(payroll_actual_bank_data),
        "data": payroll_actual_bank_data,
    }
    debug["employee_bank_final"] = _get_employee_bank_details(
        connector,
        fio,
        phys_person,
        employee_ref,
        hire_date=hire_date,
        inn_org=inn_org,
    )

    employee_owners = []
    if employee_ref is not None:
        employee_owners.append(employee_ref)
    if phys_person is not None:
        employee_owners.append(phys_person)

    employee_counterparties = _find_counterparties_by_fio(connector, fio)
    debug["employee_counterparties"] = [_describe_owner(item) for item in employee_counterparties]
    employee_owners.extend(employee_counterparties)

    for owner in employee_owners:
        direct_bank_data, direct_attempts = _read_bank_details_from_owner_attributes(owner)
        debug["employee_attempts"].append({
            "owner": _describe_owner(owner),
            "source": "owner_attributes",
            "found": bool(direct_bank_data),
            "attempts": direct_attempts,
            "data": direct_bank_data,
        })
        for catalog_name in ("БанковскиеСчетаФизическихЛиц", "БанковскиеСчетаКонтрагентов"):
            bank_data = _read_bank_account_by_owner(connector, owner, catalog_name)
            debug["employee_attempts"].append({
                "owner": _describe_owner(owner),
                "source": "catalog",
                "catalog": catalog_name,
                "found": bool(bank_data),
                "data": bank_data,
            })

    org_owners = []
    if organization is not None:
        org_owners.append(organization)

    if inn_org:
        query = connector.NewObject("Запрос")
        query.Text = """
            ВЫБРАТЬ ПЕРВЫЕ 10
                Контрагенты.Ссылка КАК Контрагент
            ИЗ
                Справочник.Контрагенты КАК Контрагенты
            ГДЕ
                Контрагенты.ИНН = &ИНН
        """
        query.SetParameter("ИНН", inn_org)
        try:
            selection = query.Execute().Choose()
            while selection.Next():
                org_owners.append(selection.Контрагент)
                debug["organization_counterparties_by_inn"].append(_describe_owner(selection.Контрагент))
        except Exception as error:
            debug["organization_counterparties_by_inn"].append(f"<query error: {error}>")

    seen = set()
    deduped_org_owners = []
    for owner in org_owners:
        owner_key = str(owner)
        if owner_key in seen:
            continue
        seen.add(owner_key)
        deduped_org_owners.append(owner)

    for owner in deduped_org_owners:
        direct_bank_data, direct_attempts = _read_bank_details_from_owner_attributes(owner)
        debug["organization_attempts"].append({
            "owner": _describe_owner(owner),
            "source": "owner_attributes",
            "found": bool(direct_bank_data),
            "attempts": direct_attempts,
            "data": direct_bank_data,
        })
        for catalog_name in ("БанковскиеСчетаОрганизаций", "БанковскиеСчетаКонтрагентов"):
            bank_data = _read_bank_account_by_owner(connector, owner, catalog_name)
            debug["organization_attempts"].append({
                "owner": _describe_owner(owner),
                "source": "catalog",
                "catalog": catalog_name,
                "found": bool(bank_data),
                "data": bank_data,
            })

    return debug


def get_bank_details_debug_from_1c(
    fio: str,
    base_1c: str,
    username: str,
    pswd: str,
    inn_org: str = "",
    hire_date_raw: str = "",
) -> dict:
    connector = _get_connector(base_1c=base_1c, username=username, pswd=pswd)
    hire_date = _parse_date(hire_date_raw)
    return _collect_bank_details_debug(connector, fio, inn_org=inn_org, hire_date=hire_date)


def enrich_bank_details_from_1c(parsed_data: dict, fio: str, base_1c: str, username: str, pswd: str) -> dict:
    connector = _get_connector(base_1c=base_1c, username=username, pswd=pswd)
    hire_date_raw = parsed_data.get("job_info", {}).get("ДатаПриемаНаРаботу", "")
    hire_date = _parse_date(hire_date_raw)

    inn_org = parsed_data.get("job_info", {}).get("Организация", {}).get("ИНН_Орг", "")
    employee_ref = _find_employee(connector, fio, inn_org=inn_org, hire_date=hire_date)
    phys_person = _find_physical_person(connector, fio)
    organization, organization_source = _find_organization_from_hire_doc(
        connector,
        fio,
        employee_ref,
        phys_person,
        inn_org,
        hire_date=hire_date,
    )
    if organization is None and inn_org:
        organization = _find_organization(connector, inn_org)
    employee_bank = _get_employee_bank_details(
        connector,
        fio,
        phys_person,
        employee_ref,
        hire_date=hire_date,
        inn_org=inn_org,
        organization=organization,
    )
    if employee_bank:
        personal_info = parsed_data.setdefault("personal_info", {})
        for key, value in employee_bank.items():
            if not value:
                continue
            personal_info[key] = value

        if personal_info.get("РасчетныйСчет") and not personal_info.get("НомерСчета"):
            personal_info["НомерСчета"] = personal_info["РасчетныйСчет"]
        if personal_info.get("НомерСчета") and not personal_info.get("РасчетныйСчет"):
            personal_info["РасчетныйСчет"] = personal_info["НомерСчета"]

    org_target = parsed_data.setdefault("job_info", {}).setdefault("Организация", {})
    org_values_before_enrichment = {
        "РасСчет": org_target.get("РасСчет", ""),
        "БикБанка": org_target.get("БикБанка", ""),
        "КорСчет": org_target.get("КорСчет", ""),
    }
    org_bank_attempts = []
    org_bank = _get_org_bank_details(
        connector,
        organization,
        inn_org,
        debug_attempts=org_bank_attempts,
    )
    if org_bank:
        if org_bank.get("БикБанка"):
            org_target["БикБанка"] = org_bank["БикБанка"]
        if org_bank.get("РасчетныйСчет"):
            org_target["РасСчет"] = org_bank["РасчетныйСчет"]
        if org_bank.get("КоррСчет"):
            org_target["КорСчет"] = org_bank["КоррСчет"]
        if org_bank.get("Банк"):
            org_target["Банк"] = org_bank["Банк"]

    _print_organization_bank_details_debug(
        inn_org=inn_org,
        organization=organization,
        organization_source=organization_source,
        source_values=org_values_before_enrichment,
        attempts=org_bank_attempts,
        final_values={
            "РасСчет": org_target.get("РасСчет", ""),
            "БикБанка": org_target.get("БикБанка", ""),
            "КорСчет": org_target.get("КорСчет", ""),
            "Банк": org_target.get("Банк", ""),
        },
    )

    return parsed_data
