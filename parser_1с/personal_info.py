import win32com.client
import datetime
from pathlib import Path

from name_utils import get_fio_part, get_name_patronymic_initials


_PERSONAL_INFO_DEBUG_PATH = Path("personal_info_debug.log")


def _write_personal_info_debug(message: str) -> None:
    # Temporary 1C diagnostics are disabled in normal runs.
    return None


def _normalize_requisite_name(value) -> str:
    return str(value or "").replace(" ", "").replace("_", "").replace("ё", "е").lower()


def _safe_getattr(obj, attr_name, default=""):
    try:
        value = getattr(obj, attr_name)
    except Exception:
        return default
    return default if value is None else value


def _format_1c_date(value) -> str:
    if not value:
        return ""

    try:
        return value.strftime('%d.%m.%Y')
    except Exception:
        try:
            return str(value).strip()
        except Exception:
            return ""


def _normalize_1c_empty_date(value):
    if value in (None, ""):
        return None

    year = getattr(value, "year", None)
    if isinstance(year, int) and year <= 1901:
        return None

    raw_variants = []
    try:
        raw_variants.append(str(value).strip())
    except Exception:
        pass

    formatted_value = _format_1c_date(value)
    if formatted_value:
        raw_variants.append(formatted_value)

    for raw_value in raw_variants:
        normalized = raw_value.replace("T", " ").strip()
        if normalized.startswith(("01.01.0100", "01.01.0001", "0100-01-01", "0001-01-01")):
            return None

    return value


def _extract_first_present(source, attr_names: tuple[str, ...], formatter=None):
    for attr_name in attr_names:
        value = _safe_getattr(source, attr_name, None)
        if value is None or value == "":
            continue
        if formatter is not None:
            try:
                return formatter(value)
            except Exception:
                continue
        return value
    return ""


def _query_single_value(connector, query_text: str, parameter_name: str, parameter_value, alias: str):
    try:
        query = connector.NewObject("Запрос")
        query.Text = query_text
        query.SetParameter(parameter_name, parameter_value)
        selection = query.Execute().Choose()
        if selection.Next():
            return _safe_getattr(selection, alias, None)
    except Exception:
        return None
    return None


def _query_single_value_with_params(connector, query_text: str, parameters: dict, alias: str):
    try:
        query = connector.NewObject("Запрос")
        query.Text = query_text
        for parameter_name, parameter_value in parameters.items():
            query.SetParameter(parameter_name, parameter_value)
        selection = query.Execute().Choose()
        if selection.Next():
            return _safe_getattr(selection, alias, None)
    except Exception:
        return None
    return None


def _query_first_phys_person_field(connector, fio: str, field_name: str):
    return _query_single_value(
        connector,
        f"""
            ВЫБРАТЬ ПЕРВЫЕ 1
                ФизЛица.{field_name} КАК Значение
            ИЗ
                Справочник.ФизическиеЛица КАК ФизЛица
            ГДЕ
                ФизЛица.ФИО = &ФИО
        """,
        "ФИО",
        fio,
        "Значение"
    )


def _query_first_phys_person_field_presentation(connector, fio: str, field_name: str) -> str:
    value = _query_single_value(
        connector,
        f"""
            ВЫБРАТЬ ПЕРВЫЕ 1
                ПРЕДСТАВЛЕНИЕ(ФизЛица.{field_name}) КАК Значение
            ИЗ
                Справочник.ФизическиеЛица КАК ФизЛица
            ГДЕ
                ФизЛица.ФИО = &ФИО
        """,
        "ФИО",
        fio,
        "Значение"
    )
    return _sanitize_1c_text(value)


def _query_family_status_from_marriage_register(connector, fio: str) -> str:
    value = _query_single_value(
        connector,
        """
            ВЫБРАТЬ ПЕРВЫЕ 1
                ПРЕДСТАВЛЕНИЕ(СостоянияВБраке.СостояниеВБраке) КАК Значение
            ИЗ
                РегистрСведений.СостоянияВБракеФизическихЛиц КАК СостоянияВБраке
            ГДЕ
                СостоянияВБраке.ФизическоеЛицо.ФИО = &ФИО
        """,
        "ФИО",
        fio,
        "Значение"
    )
    return _sanitize_1c_text(value)


def _get_driver_license_info(connector, fio: str) -> dict:
    driver_license_info = {
        'Серия': '',
        'Номер': '',
        'ДатаВыдачи': '',
        'КатегорииТС': '',
    }

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ
            ДокументыФизЛиц.ВидДокумента.Наименование КАК ВидДокумента,
            ДокументыФизЛиц.Серия КАК Серия,
            ДокументыФизЛиц.Номер КАК Номер,
            ДокументыФизЛиц.ДатаВыдачи КАК ДатаВыдачи
        ИЗ
            РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
        ГДЕ
            ДокументыФизЛиц.ФизЛицо.ФИО = &ФИО
    """
    query.SetParameter("ФИО", fio)

    try:
        selection = query.Execute().Choose()
    except Exception:
        selection = None

    if selection is not None:
        while selection.Next():
            doc_type = _sanitize_1c_text(_safe_getattr(selection, "ВидДокумента", ""))
            normalized_doc_type = doc_type.lower()
            if "водитель" not in normalized_doc_type and "водит" not in normalized_doc_type:
                continue

            driver_license_info['Серия'] = str(_safe_getattr(selection, "Серия", "") or "").strip()
            driver_license_info['Номер'] = str(_safe_getattr(selection, "Номер", "") or "").strip()
            driver_license_info['ДатаВыдачи'] = _format_1c_date(_safe_getattr(selection, "ДатаВыдачи", None))
            break

    category_field_candidates = (
        "КатегорииТС",
        "КатегорияТС",
        "КатегорииТранспортныхСредств",
        "КатегорииТранспортныхСредствВУ",
        "Категории",
    )
    for field_name in category_field_candidates:
        category_value = _query_single_value_with_params(
            connector,
            f"""
                ВЫБРАТЬ ПЕРВЫЕ 1
                    ДокументыФизЛиц.{field_name} КАК Значение
                ИЗ
                    РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
                ГДЕ
                    ДокументыФизЛиц.ФизЛицо.ФИО = &ФИО
                    И ДокументыФизЛиц.ВидДокумента.Наименование ПОДОБНО &ВидДокумента
            """,
            {
                "ФИО": fio,
                "ВидДокумента": "%водитель%",
            },
            "Значение"
        )
        formatted_category = _extract_1c_presentation(connector, category_value)
        if formatted_category:
            driver_license_info['КатегорииТС'] = formatted_category
            break

    return driver_license_info


def _get_military_document_info(connector, fio: str) -> dict:
    military_document_info = {
        'ДокументВоинскогоУчета': '',
        'СерияДокументаВоинскогоУчета': '',
        'НомерДокументаВоинскогоУчета': '',
        'ОтделВыдавшийДокументВоинскогоУчета': '',
        'ДатаВыдачиДокументаВоинскогоУчета': '',
    }

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ
            ДокументыФизЛиц.ВидДокумента.Наименование КАК ВидДокумента,
            ДокументыФизЛиц.Серия КАК Серия,
            ДокументыФизЛиц.Номер КАК Номер,
            ДокументыФизЛиц.ДатаВыдачи КАК ДатаВыдачи,
            ДокументыФизЛиц.КемВыдан КАК КемВыдан
        ИЗ
            РегистрСведений.ДокументыФизическихЛиц КАК ДокументыФизЛиц
        ГДЕ
            ДокументыФизЛиц.ФизЛицо.ФИО = &ФИО
    """
    query.SetParameter("ФИО", fio)

    try:
        selection = query.Execute().Choose()
    except Exception:
        selection = None

    best_match = None
    best_priority = -1

    if selection is not None:
        while selection.Next():
            doc_type = _sanitize_1c_text(_safe_getattr(selection, "ВидДокумента", ""))
            normalized_doc_type = doc_type.lower()

            priority = -1
            if "военн" in normalized_doc_type and "билет" in normalized_doc_type:
                priority = 3
            elif "припис" in normalized_doc_type:
                priority = 2
            elif "воин" in normalized_doc_type or "военн" in normalized_doc_type:
                priority = 1

            if priority > best_priority:
                best_priority = priority
                best_match = {
                    'ДокументВоинскогоУчета': doc_type,
                    'СерияДокументаВоинскогоУчета': str(_safe_getattr(selection, "Серия", "") or "").strip(),
                    'НомерДокументаВоинскогоУчета': str(_safe_getattr(selection, "Номер", "") or "").strip(),
                    'ОтделВыдавшийДокументВоинскогоУчета': str(_safe_getattr(selection, "КемВыдан", "") or "").strip(),
                    'ДатаВыдачиДокументаВоинскогоУчета': _format_1c_date(_safe_getattr(selection, "ДатаВыдачи", None)),
                }

    if best_match:
        military_document_info.update(best_match)

    return military_document_info


def _extract_country_name(country_ref) -> str:
    if not country_ref:
        return ""

    for attr_name in ("НаименованиеПолное", "Наименование", "Представление"):
        value = _safe_getattr(country_ref, attr_name, "")
        if value:
            return str(value).strip()

    try:
        return str(country_ref).strip()
    except Exception:
        return ""


def _extract_1c_presentation(connector, value) -> str:
    if not value:
        return ""

    for attr_name in ("Наименование", "Представление", "Description"):
        attr_value = _safe_getattr(value, attr_name, "")
        if attr_value:
            sanitized = _sanitize_1c_text(attr_value)
            if sanitized:
                return sanitized

    try:
        presented = connector.String(value)
        if presented:
            sanitized = _sanitize_1c_text(presented)
            if sanitized:
                return sanitized
    except Exception:
        pass

    try:
        return _sanitize_1c_text(value)
    except Exception:
        return ""


def _sanitize_1c_text(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""

    normalized = text.lower()
    if "comobject" in normalized and "unknown" in normalized:
        return ""

    return text


def _collect_requisite_names(obj, markers: tuple[str, ...]) -> list[str]:
    collected: list[str] = []
    lowered_markers = tuple(marker.lower() for marker in markers)

    try:
        for name in dir(obj):
            lowered_name = str(name).lower()
            if any(marker in lowered_name for marker in lowered_markers) and name not in collected:
                collected.append(name)
    except Exception:
        pass

    try:
        metadata_method = getattr(obj, "Метаданные", None)
        metadata = metadata_method() if callable(metadata_method) else None
        if metadata is not None:
            for requisite in metadata.Реквизиты:
                req_name = _safe_getattr(requisite, "Имя", "")
                lowered_name = str(req_name).lower()
                if any(marker in lowered_name for marker in lowered_markers) and req_name not in collected:
                    collected.append(req_name)
    except Exception:
        pass

    return collected


def _extract_first_meaningful_attr_text(
    connector,
    source_objects: list,
    attr_names: list[str],
    *,
    skip_boolean_like: bool = False,
) -> str:
    for source_obj in source_objects:
        if source_obj is None:
            continue
        for attr_name in attr_names:
            raw_value = _safe_getattr(source_obj, attr_name, "")
            if raw_value in (None, ""):
                continue
            value = _sanitize_1c_text(_extract_1c_presentation(connector, raw_value))
            if skip_boolean_like and _is_boolean_like_text(value):
                continue
            if value:
                return value
    return ""


def _is_truthy_1c(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0

    normalized = str(value or "").strip().lower()
    return normalized in ("true", "1", "yes", "y", "да", "истина")


def _is_boolean_like_text(value: str) -> bool:
    normalized = str(value or "").strip().lower()
    return normalized in ("true", "false", "1", "0", "yes", "no", "y", "n", "да", "нет", "истина", "ложь")


def _get_employee_position_by_ref(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    query = connector.NewObject("Запрос")
    filters = []
    if employee_ref is not None:
        filters.append("ПриемНаРаботу.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("ПриемНаРаботу.ФизическоеЛицо = &ФизическоеЛицо")

    if not filters:
        return ""

    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            ПриемНаРаботу.Должность КАК Должность
        ИЗ
            Документ.ПриемНаРаботу КАК ПриемНаРаботу
        ГДЕ
            {' ИЛИ '.join(filters)}
        УПОРЯДОЧИТЬ ПО
            ПриемНаРаботу.Дата УБЫВ
    """

    try:
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизическоеЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return _extract_1c_presentation(connector, _safe_getattr(selection, "Должность", ""))
    return ""


def _get_employee_position_by_fio(connector, fio: str) -> str:
    if not fio:
        return ""

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            ПриемНаРаботу.Должность КАК Должность
        ИЗ
            Документ.ПриемНаРаботу КАК ПриемНаРаботу
        ГДЕ
            ПриемНаРаботу.ФизическоеЛицо.ФИО = &ФИО
            ИЛИ ПриемНаРаботу.Сотрудник.Наименование = &ФИО
        УПОРЯДОЧИТЬ ПО
            ПриемНаРаботу.Дата УБЫВ
    """

    try:
        query.SetParameter("ФИО", fio)
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return _extract_1c_presentation(connector, _safe_getattr(selection, "Должность", ""))
    return ""


def _get_employee_ref_by_fio(connector, fio: str):
    if not fio:
        return None

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Сотрудники.Ссылка КАК Сотрудник
        ИЗ
            Справочник.Сотрудники КАК Сотрудники
        ГДЕ
            Сотрудники.ФизическоеЛицо.ФИО = &ФИО
            ИЛИ Сотрудники.Наименование = &ФИО
    """

    try:
        query.SetParameter("ФИО", fio)
        selection = query.Execute().Choose()
    except Exception:
        return None

    if selection.Next():
        return _safe_getattr(selection, "Сотрудник", None)
    return None


def _get_phys_person_ref_by_fio(connector, fio: str):
    if not fio:
        return None

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            ФизЛица.Ссылка КАК ФизЛицо
        ИЗ
            Справочник.ФизическиеЛица КАК ФизЛица
        ГДЕ
            ФизЛица.ФИО = &ФИО
            ИЛИ ФизЛица.Наименование = &ФИО
    """

    try:
        query.SetParameter("ФИО", fio)
        selection = query.Execute().Choose()
    except Exception:
        return None

    if selection.Next():
        return _safe_getattr(selection, "ФизЛицо", None)
    return None


def _get_employee_ref_by_phys_person(connector, phys_person_ref):
    if phys_person_ref is None:
        return None

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Сотрудники.Ссылка КАК Сотрудник
        ИЗ
            Справочник.Сотрудники КАК Сотрудники
        ГДЕ
            Сотрудники.ФизическоеЛицо = &ФизЛицо
    """

    try:
        query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return None

    if selection.Next():
        return _safe_getattr(selection, "Сотрудник", None)
    return None


def _get_employee_position_name_by_ref(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    query = connector.NewObject("Запрос")
    filters = []
    if employee_ref is not None:
        filters.append("ПриемНаРаботу.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("ПриемНаРаботу.ФизическоеЛицо = &ФизическоеЛицо")

    if not filters:
        return ""

    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            ПриемНаРаботу.Должность.Наименование КАК ДолжностьНаименование
        ИЗ
            Документ.ПриемНаРаботу КАК ПриемНаРаботу
        ГДЕ
            {' ИЛИ '.join(filters)}
        УПОРЯДОЧИТЬ ПО
            ПриемНаРаботу.Дата УБЫВ
    """

    try:
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизическоеЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return str(_safe_getattr(selection, "ДолжностьНаименование", "") or "").strip()
    return ""


def _get_current_position_from_hr_register(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    query = connector.NewObject("Запрос")
    filters = []
    if employee_ref is not None:
        filters.append("КадровыеДанные.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("КадровыеДанные.ФизическоеЛицо = &ФизЛицо")

    if not filters:
        return ""

    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            КадровыеДанные.ТекущаяДолжность КАК ТекущаяДолжность
        ИЗ
            РегистрСведений.ТекущиеКадровыеДанныеСотрудников КАК КадровыеДанные
        ГДЕ
            {' ИЛИ '.join(filters)}
    """

    try:
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return _extract_1c_presentation(connector, _safe_getattr(selection, "ТекущаяДолжность", ""))
    return ""


def _get_current_position_name_from_hr_register(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    query = connector.NewObject("Запрос")
    filters = []
    if employee_ref is not None:
        filters.append("КадровыеДанные.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("КадровыеДанные.ФизическоеЛицо = &ФизЛицо")

    if not filters:
        return ""

    query.Text = f"""
        ВЫБРАТЬ ПЕРВЫЕ 1
            КадровыеДанные.ТекущаяДолжность.Наименование КАК ТекущаяДолжностьНаименование
        ИЗ
            РегистрСведений.ТекущиеКадровыеДанныеСотрудников КАК КадровыеДанные
        ГДЕ
            {' ИЛИ '.join(filters)}
    """

    try:
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return str(_safe_getattr(selection, "ТекущаяДолжностьНаименование", "") or "").strip()
    return ""


def _get_current_position_name_from_hr_register_slice(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    dimension_filters = []
    if employee_ref is not None:
        dimension_filters.append("РЎРѕС‚СЂСѓРґРЅРёРє = &РЎРѕС‚СЂСѓРґРЅРёРє")
    if phys_person_ref is not None:
        dimension_filters.append("Р¤РёР·РёС‡РµСЃРєРѕРµР›РёС†Рѕ = &Р¤РёР·Р›РёС†Рѕ")

    if not dimension_filters:
        return ""

    query = connector.NewObject("Р—Р°РїСЂРѕСЃ")
    query.Text = f"""
        Р’Р«Р‘Р РђРўР¬ РџР•Р Р’Р«Р• 1
            РљР°РґСЂРѕРІС‹РµР”Р°РЅРЅС‹Рµ.РўРµРєСѓС‰Р°СЏР”РѕР»Р¶РЅРѕСЃС‚СЊ.РќР°РёРјРµРЅРѕРІР°РЅРёРµ РљРђРљ РўРµРєСѓС‰Р°СЏР”РѕР»Р¶РЅРѕСЃС‚СЊРќР°РёРјРµРЅРѕРІР°РЅРёРµ
        РР—
            Р РµРіРёСЃС‚СЂРЎРІРµРґРµРЅРёР№.РўРµРєСѓС‰РёРµРљР°РґСЂРѕРІС‹РµР”Р°РЅРЅС‹РµРЎРѕС‚СЂСѓРґРЅРёРєРѕРІ.РЎСЂРµР·РџРѕСЃР»РµРґРЅРёС…(&РќР°Р”Р°С‚Сѓ, {' РР›Р '.join(dimension_filters)}) РљРђРљ РљР°РґСЂРѕРІС‹РµР”Р°РЅРЅС‹Рµ
    """

    try:
        if employee_ref is not None:
            query.SetParameter("РЎРѕС‚СЂСѓРґРЅРёРє", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("Р¤РёР·Р›РёС†Рѕ", phys_person_ref)
        query.SetParameter("РќР°Р”Р°С‚Сѓ", datetime.datetime.now())
        selection = query.Execute().Choose()
    except Exception:
        return ""

    if selection.Next():
        return str(_safe_getattr(selection, "РўРµРєСѓС‰Р°СЏР”РѕР»Р¶РЅРѕСЃС‚СЊРќР°РёРјРµРЅРѕРІР°РЅРёРµ", "") or "").strip()
    return _get_current_position_name_from_hr_register(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )


def _get_current_position_name_from_hr_register_slice_v2(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    dimension_filters = []
    if employee_ref is not None:
        dimension_filters.append("Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        dimension_filters.append("ФизическоеЛицо = &ФизЛицо")

    if not dimension_filters:
        return ""

    try:
        query = connector.NewObject("Запрос")
        query.Text = f"""
            ВЫБРАТЬ ПЕРВЫЕ 1
                КадровыеДанные.ТекущаяДолжность.Наименование КАК ТекущаяДолжностьНаименование
            ИЗ
                РегистрСведений.ТекущиеКадровыеДанныеСотрудников.СрезПоследних(&НаДату, {' ИЛИ '.join(dimension_filters)}) КАК КадровыеДанные
        """
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизЛицо", phys_person_ref)
        query.SetParameter("НаДату", datetime.datetime.now())
        selection = query.Execute().Choose()
        if selection.Next():
            position = str(_safe_getattr(selection, "ТекущаяДолжностьНаименование", "") or "").strip()
            if position:
                return position
    except Exception:
        pass

    return _get_current_position_name_from_hr_register(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )


def _get_current_position_name_from_hr_register_ordered(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    filters = []
    if employee_ref is not None:
        filters.append("КадровыеДанные.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("КадровыеДанные.ФизическоеЛицо = &ФизЛицо")

    if not filters:
        return ""

    try:
        query = connector.NewObject("Запрос")
        query.Text = f"""
            ВЫБРАТЬ
                КадровыеДанные.ТекущаяДолжность.Наименование КАК ТекущаяДолжностьНаименование
            ИЗ
                РегистрСведений.ТекущиеКадровыеДанныеСотрудников КАК КадровыеДанные
            ГДЕ
                {' ИЛИ '.join(filters)}
            УПОРЯДОЧИТЬ ПО
                КадровыеДанные.Период УБЫВ
        """
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
        while selection.Next():
            position = str(_safe_getattr(selection, "ТекущаяДолжностьНаименование", "") or "").strip()
            if position:
                return position
    except Exception:
        return ""

    return ""


def _get_position_name_from_employee_history(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    filters = []
    if employee_ref is not None:
        filters.append("КадроваяИстория.Сотрудник = &Сотрудник")
    if phys_person_ref is not None:
        filters.append("КадроваяИстория.ФизическоеЛицо = &ФизЛицо")

    if not filters:
        return ""

    try:
        query = connector.NewObject("Запрос")
        query.Text = f"""
            ВЫБРАТЬ
                КадроваяИстория.Должность.Наименование КАК ДолжностьНаименование
            ИЗ
                РегистрСведений.КадроваяИсторияСотрудников КАК КадроваяИстория
            ГДЕ
                {' ИЛИ '.join(filters)}
            УПОРЯДОЧИТЬ ПО
                КадроваяИстория.Период УБЫВ
        """
        if employee_ref is not None:
            query.SetParameter("Сотрудник", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
        while selection.Next():
            position = str(_safe_getattr(selection, "ДолжностьНаименование", "") or "").strip()
            if position:
                return position
    except Exception:
        pass

    return _get_current_position_name_from_hr_register_ordered(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )


def _get_position_name_from_employee_history_v2(connector, employee_ref=None, phys_person_ref=None) -> str:
    if employee_ref is None and phys_person_ref is None:
        return ""

    filters = []
    if employee_ref is not None:
        filters.append("\u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f.\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a = &\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a")
    if phys_person_ref is not None:
        filters.append("\u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f.\u0424\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u043e\u0435\u041b\u0438\u0446\u043e = &\u0424\u0438\u0437\u041b\u0438\u0446\u043e")

    if not filters:
        return ""

    try:
        or_delimiter = "\u0020\u0418\u041b\u0418\u0020"
        query = connector.NewObject("\u0417\u0430\u043f\u0440\u043e\u0441")
        query.Text = f"""
            \u0412\u042b\u0411\u0420\u0410\u0422\u042c
                \u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f.\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c.\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u041a\u0410\u041a \u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435
            \u0418\u0417
                \u0420\u0435\u0433\u0438\u0441\u0442\u0440\u0421\u0432\u0435\u0434\u0435\u043d\u0438\u0439.\u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a\u043e\u0432 \u041a\u0410\u041a \u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f
            \u0413\u0414\u0415
                {or_delimiter.join(filters)}
            \u0423\u041f\u041e\u0420\u042f\u0414\u041e\u0427\u0418\u0422\u042c \u041f\u041e
                \u041a\u0430\u0434\u0440\u043e\u0432\u0430\u044f\u0418\u0441\u0442\u043e\u0440\u0438\u044f.\u041f\u0435\u0440\u0438\u043e\u0434 \u0423\u0411\u042b\u0412
        """
        if employee_ref is not None:
            query.SetParameter("\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a", employee_ref)
        if phys_person_ref is not None:
            query.SetParameter("\u0424\u0438\u0437\u041b\u0438\u0446\u043e", phys_person_ref)
        selection = query.Execute().Choose()
        while selection.Next():
            position = str(_safe_getattr(selection, "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435", "") or "").strip()
            if position:
                return position
    except Exception:
        pass

    return _get_current_position_name_from_hr_register_ordered(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )


def _extract_position_from_source(connector, source) -> str:
    if not source:
        return ""

    for attr_name in (
        "Должность",
        "ТекущаяДолжность",
        "ДолжностьПоШтатномуРасписанию",
        "РабочееМесто",
        "МестоРаботы",
    ):
        candidate = _safe_getattr(source, attr_name, "")
        position = _extract_1c_presentation(connector, candidate)
        if position:
            return position
        for nested_attr_name in ("РќР°РёРјРµРЅРѕРІР°РЅРёРµ", "РќР°РёРјРµРЅРѕРІР°РЅРёРµРџРѕР»РЅРѕРµ", "РџСЂРµРґСЃС‚Р°РІР»РµРЅРёРµ", "Description", "Name"):
            nested_value = _safe_getattr(candidate, nested_attr_name, "")
            nested_position = _sanitize_1c_text(nested_value)
            if nested_position:
                return nested_position

    return ""


def _extract_position_from_object(connector, source) -> str:
    if not source:
        return ""

    try:
        obj = source.GetObject()
    except Exception:
        return ""

    return _extract_position_from_source(connector, obj)


def _resolve_responsible_position(connector, value, fio: str) -> str:
    position = _extract_position_from_source(connector, value)
    if position:
        return position

    position = _extract_position_from_object(connector, value)
    if position:
        return position

    employee_ref = (
        _safe_getattr(value, "Сотрудник", None)
        or _safe_getattr(value, "Employee", None)
    )
    phys_person_ref = (
        _safe_getattr(value, "ФизическоеЛицо", None)
        or _safe_getattr(value, "ФизЛицо", None)
        or _safe_getattr(value, "Person", None)
    )

    if phys_person_ref is None and _safe_getattr(value, "ДатаРождения", None) is not None:
        phys_person_ref = value

    if employee_ref is None and phys_person_ref is not None and _safe_getattr(value, "Наименование", None):
        employee_ref = value

    if employee_ref is not None:
        position = _get_position_name_from_employee_history_v2(connector, employee_ref=employee_ref)
        if position:
            return position
        position = _get_current_position_name_from_hr_register_slice_v2(connector, employee_ref=employee_ref)
        if position:
            return position
        position = _extract_position_from_source(connector, employee_ref)
        if position:
            return position
        position = _extract_position_from_object(connector, employee_ref)
        if position:
            return position

    if phys_person_ref is not None:
        position = _get_position_name_from_employee_history_v2(connector, phys_person_ref=phys_person_ref)
        if position:
            return position
        position = _get_current_position_name_from_hr_register_slice_v2(connector, phys_person_ref=phys_person_ref)
        if position:
            return position
        position = _extract_position_from_source(connector, phys_person_ref)
        if position:
            return position
        position = _extract_position_from_object(connector, phys_person_ref)
        if position:
            return position

    position = _get_employee_position_name_by_ref(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )
    if position:
        return position

    position = _get_employee_position_by_ref(
        connector,
        employee_ref=employee_ref,
        phys_person_ref=phys_person_ref,
    )
    if position:
        return position

    if fio:
        phys_person_ref_by_fio = _get_phys_person_ref_by_fio(connector, fio)
        if phys_person_ref_by_fio is not None:
            position = _get_position_name_from_employee_history_v2(connector, phys_person_ref=phys_person_ref_by_fio)
            if position:
                return position
            position = _get_current_position_name_from_hr_register_slice_v2(connector, phys_person_ref=phys_person_ref_by_fio)
            if position:
                return position
            position = _extract_position_from_object(connector, phys_person_ref_by_fio)
            if position:
                return position

            position = _get_employee_position_name_by_ref(connector, phys_person_ref=phys_person_ref_by_fio)
            if position:
                return position

            position = _get_employee_position_by_ref(connector, phys_person_ref=phys_person_ref_by_fio)
            if position:
                return position

            employee_ref_by_phys_person = _get_employee_ref_by_phys_person(connector, phys_person_ref_by_fio)
            if employee_ref_by_phys_person is not None:
                position = _get_position_name_from_employee_history_v2(connector, employee_ref=employee_ref_by_phys_person)
                if position:
                    return position
                position = _get_current_position_name_from_hr_register_slice_v2(connector, employee_ref=employee_ref_by_phys_person)
                if position:
                    return position
                position = _extract_position_from_object(connector, employee_ref_by_phys_person)
                if position:
                    return position

                position = _get_employee_position_name_by_ref(connector, employee_ref=employee_ref_by_phys_person)
                if position:
                    return position

                position = _get_employee_position_by_ref(connector, employee_ref=employee_ref_by_phys_person)
                if position:
                    return position

        employee_ref_by_fio = _get_employee_ref_by_fio(connector, fio)
        if employee_ref_by_fio is not None:
            position = _get_position_name_from_employee_history_v2(connector, employee_ref=employee_ref_by_fio)
            if position:
                return position
            position = _get_current_position_name_from_hr_register_slice_v2(connector, employee_ref=employee_ref_by_fio)
            if position:
                return position
            position = _extract_position_from_object(connector, employee_ref_by_fio)
            if position:
                return position

            position = _get_employee_position_name_by_ref(connector, employee_ref=employee_ref_by_fio)
            if position:
                return position

            position = _get_employee_position_by_ref(connector, employee_ref=employee_ref_by_fio)
            if position:
                return position

    if fio:
        return _get_employee_position_by_fio(connector, fio)

    return ""


def _extract_commission_member_data(connector, value) -> dict:
    result = {
        "ФИО": "",
        "Фамилия": "",
        "ИнициалыИмениОтчества": "",
        "Должность": "",
    }
    if not value:
        return result

    employee_ref = None
    phys_person_ref = None

    phys_person_ref = _safe_getattr(value, "ФизическоеЛицо", None)
    employee_ref = value if phys_person_ref is not None else None

    fio = ""
    if phys_person_ref is not None:
        fio = (
            _safe_getattr(phys_person_ref, "Наименование", "")
            or _safe_getattr(phys_person_ref, "Description", "")
            or _safe_getattr(phys_person_ref, "Представление", "")
        )

    if not fio:
        fio = (
            _safe_getattr(value, "Наименование", "")
            or _safe_getattr(value, "Description", "")
            or _safe_getattr(value, "Представление", "")
        )

    position = ""
    for attr_name in (
        "Должность",
        "ТекущаяДолжность",
        "ДолжностьПоШтатномуРасписанию",
        "РабочееМесто",
    ):
        candidate = _safe_getattr(value, attr_name, "")
        position = _extract_1c_presentation(connector, candidate)
        if position:
            break

    if not position:
        position = _get_employee_position_by_ref(connector, employee_ref=employee_ref, phys_person_ref=phys_person_ref)

    fio = str(fio or "").strip()
    result["ФИО"] = fio
    result["Фамилия"] = get_fio_part(fio, 0)
    result["ИнициалыИмениОтчества"] = get_name_patronymic_initials(fio)
    result["Должность"] = str(position or "").strip()
    return result


def _extract_commission_member_data_with_fallback(connector, value) -> dict:
    result = _extract_commission_member_data(connector, value)
    if result.get("Должность") or result.get("Р”РѕР»Р¶РЅРѕСЃС‚СЊ"):
        return result

    fio = str(result.get("ФИО", "") or result.get("Р¤РРћ", "")).strip()
    fallback_position = _resolve_responsible_position(connector, value, fio)
    if fallback_position:
        if "Должность" in result:
            result["Должность"] = str(fallback_position).strip()
        if "Р”РѕР»Р¶РЅРѕСЃС‚СЊ" in result:
            result["Р”РѕР»Р¶РЅРѕСЃС‚СЊ"] = str(fallback_position).strip()

    return result


def _try_get_signing_basis(connector, organization_ref, responsible_person_ref=None) -> tuple[str, list[str]]:
    debug_lines: list[str] = []
    if organization_ref is None:
        debug_lines.append("Организация для поиска основания подписи не определена.")
        return "", debug_lines

    today = datetime.datetime.now()

    query_variants = [
        """
        ВЫБРАТЬ ПЕРВЫЕ 1
            Основания.ОснованиеПодписи КАК ОснованиеПодписи,
            Основания.НомерПоПорядку КАК НомерПоПорядку,
            Основания.ДатаНачала КАК ДатаНачала,
            Основания.ДатаОкончания КАК ДатаОкончания
        ИЗ
            РегистрСведений.ОснованияПолномочийОтветственныхЛиц КАК Основания
        ГДЕ
            Основания.Организация = &Организация
            И (Основания.ДатаНачала ЕСТЬ NULL ИЛИ Основания.ДатаНачала <= &ТекущаяДата)
            И (Основания.ДатаОкончания ЕСТЬ NULL ИЛИ Основания.ДатаОкончания = ДАТАВРЕМЯ(1, 1, 1, 0, 0, 0) ИЛИ Основания.ДатаОкончания >= &ТекущаяДата)
        УПОРЯДОЧИТЬ ПО
            Основания.НомерПоПорядку УБЫВ,
            Основания.ДатаНачала УБЫВ
        """,
    ]

    if responsible_person_ref is not None:
        debug_lines.append("Поиск основания подписи: пробуем варианты по организации и физлицу.")
        query_variants = [
            """
            ВЫБРАТЬ ПЕРВЫЕ 1
                Основания.ОснованиеПодписи КАК ОснованиеПодписи,
                Основания.НомерПоПорядку КАК НомерПоПорядку,
                Основания.ДатаНачала КАК ДатаНачала,
                Основания.ДатаОкончания КАК ДатаОкончания
            ИЗ
                РегистрСведений.ОснованияПолномочийОтветственныхЛиц КАК Основания
            ГДЕ
                Основания.Организация = &Организация
                И Основания.ФизическоеЛицо = &ФизическоеЛицо
                И (Основания.ДатаНачала ЕСТЬ NULL ИЛИ Основания.ДатаНачала <= &ТекущаяДата)
                И (Основания.ДатаОкончания ЕСТЬ NULL ИЛИ Основания.ДатаОкончания = ДАТАВРЕМЯ(1, 1, 1, 0, 0, 0) ИЛИ Основания.ДатаОкончания >= &ТекущаяДата)
            УПОРЯДОЧИТЬ ПО
                Основания.НомерПоПорядку УБЫВ,
                Основания.ДатаНачала УБЫВ
            """,
        ] + query_variants
    else:
        debug_lines.append("Поиск основания подписи: физлицо не определено, ищем только по организации.")

    for variant_index, query_text in enumerate(query_variants, start=1):
        try:
            query = connector.NewObject("Запрос")
            query.Text = query_text
            query.SetParameter("Организация", organization_ref)
            query.SetParameter("ТекущаяДата", today)
            if responsible_person_ref is not None:
                try:
                    query.SetParameter("ФизическоеЛицо", responsible_person_ref)
                except Exception:
                    debug_lines.append(f"Вариант {variant_index}: не удалось установить параметр ФизическоеЛицо.")

            selection = query.Execute().Choose()
            found_rows = False
            while selection.Next():
                found_rows = True
                value = _safe_getattr(selection, "ОснованиеПодписи", "")
                if value:
                    value_str = str(value).strip()
                    order_no = _safe_getattr(selection, "НомерПоПорядку", "")
                    date_start = _safe_getattr(selection, "ДатаНачала", "")
                    date_end = _safe_getattr(selection, "ДатаОкончания", "")
                    debug_lines.append(
                        f"Вариант {variant_index}: найдено значение '{value_str}'. "
                        f"НомерПоПорядку={order_no}, ДатаНачала={date_start}, ДатаОкончания={date_end}."
                    )
                    return value_str, debug_lines

            if found_rows:
                debug_lines.append(f"Вариант {variant_index}: строки есть, но поле ОснованиеПодписи пустое.")
            else:
                debug_lines.append(f"Вариант {variant_index}: строки не найдены.")
        except Exception as error:
            debug_lines.append(f"Вариант {variant_index}: ошибка запроса: {error}")
            continue

    return "", debug_lines


def _fill_migrant_data(connector, phys_person_ref, target: dict) -> None:
    if phys_person_ref is None:
        return

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            ДанныеМигранта.Виза КАК Виза,
            ДанныеМигранта.Патент КАК ПатентФлаг,
            ДанныеМигранта.ПатентДатаВыдачи КАК ПатентДатаВыдачи,
            ДанныеМигранта.ПатентНомер КАК ПатентНомер,
            ДанныеМигранта.ПатентСерия КАК ПатентСерия,
            ДанныеМигранта.Разрешение КАК РазрешениеФлаг,
            ДанныеМигранта.РазрешениеНачало КАК РазрешениеНачало,
            ДанныеМигранта.РазрешениеНомер КАК РазрешениеНомер,
            ДанныеМигранта.РазрешениеСерия КАК РазрешениеСерия,
            ДанныеМигранта.КемВыдан КАК КемВыдан,
            ДанныеМигранта.ДатаВыдачи КАК ДатаВыдачи,
            ДанныеМигранта.СрокДействия КАК СрокДействия,
            ДанныеМигранта.ВидДокумента КАК ВидДокумента
        ИЗ
            РегистрСведений.ДанныеМигранта КАК ДанныеМигранта
        ГДЕ
            ДанныеМигранта.ФизЛицо = &ФизЛицо
    """

    try:
        query.SetParameter("ФизЛицо", phys_person_ref)
        selection = query.Execute().Choose()
    except Exception:
        return

    while selection.Next():
        patent_flag = _safe_getattr(selection, "ПатентФлаг", "")
        permit_flag = _safe_getattr(selection, "РазрешениеФлаг", "")
        migrant_doc_type = _extract_1c_presentation(connector, _safe_getattr(selection, "ВидДокумента", ""))
        patent_checked = str(_extract_1c_presentation(connector, patent_flag) or "").strip()
        permit_checked = str(_extract_1c_presentation(connector, permit_flag) or "").strip()

        target["ЧекбоксПатент"] = patent_checked
        target["ЧекбоксРазрешение"] = permit_checked
        target["ВидДокументаМигранта"] = migrant_doc_type

        issuer = str(_safe_getattr(selection, "КемВыдан", "") or "").strip()
        target["КемВыданПатент"] = issuer
        target["КемВыданРазрешения"] = issuer

        patent_date = _safe_getattr(selection, "ПатентДатаВыдачи", None)
        if patent_date:
            try:
                target["ДатаВыдачиПатента"] = patent_date.strftime('%d.%m.%Y')
            except Exception:
                target["ДатаВыдачиПатента"] = str(patent_date).strip()

        permit_date = _safe_getattr(selection, "РазрешениеНачало", None)
        if permit_date:
            try:
                target["ДатаВыдачиРазрешения"] = permit_date.strftime('%d.%m.%Y')
            except Exception:
                target["ДатаВыдачиРазрешения"] = str(permit_date).strip()

        patent_number = str(_safe_getattr(selection, "ПатентНомер", "") or "").strip()
        patent_series = str(_safe_getattr(selection, "ПатентСерия", "") or "").strip()
        permit_number = str(_safe_getattr(selection, "РазрешениеНомер", "") or "").strip()
        permit_series = str(_safe_getattr(selection, "РазрешениеСерия", "") or "").strip()

        if patent_checked:
            target["Патент"] = migrant_doc_type or "Патент"
        target["СерияПатента"] = patent_series
        target["НомерПатента"] = patent_number

        if permit_checked:
            target["Разрешение"] = migrant_doc_type or "Разрешение"
        target["СерияРазрешения"] = permit_series
        target["НомерРазрешения"] = permit_number

        if patent_checked and not target["ДатаВыдачиПатента"]:
            issue_date = _safe_getattr(selection, "ДатаВыдачи", None)
            if issue_date:
                try:
                    target["ДатаВыдачиПатента"] = issue_date.strftime('%d.%m.%Y')
                except Exception:
                    target["ДатаВыдачиПатента"] = str(issue_date).strip()

        if permit_checked and not target["ДатаВыдачиРазрешения"]:
            issue_date = _safe_getattr(selection, "ДатаВыдачи", None)
            if issue_date:
                try:
                    target["ДатаВыдачиРазрешения"] = issue_date.strftime('%d.%m.%Y')
                except Exception:
                    target["ДатаВыдачиРазрешения"] = str(issue_date).strip()

        return


def _describe_info_register_metadata(connector, register_name: str) -> list[str]:
    lines: list[str] = [f"РегистрСведений: {register_name}"]

    try:
        metadata = connector.Metadata
    except Exception as error:
        return lines + [f"Не удалось получить Metadata: {error}"]

    register_meta = None

    try:
        info_registers = metadata.InformationRegisters
        try:
            register_meta = getattr(info_registers, register_name)
        except Exception:
            for item in info_registers:
                if _safe_getattr(item, "Name", "") == register_name:
                    register_meta = item
                    break
    except Exception as error:
        return lines + [f"Не удалось получить InformationRegisters: {error}"]

    if register_meta is None:
        return lines + ["Регистр не найден в метаданных."]

    def append_collection(title: str, collection_name: str) -> None:
        try:
            collection = getattr(register_meta, collection_name)
        except Exception as error:
            lines.append(f"{title}: ошибка чтения коллекции: {error}")
            return

        names: list[str] = []
        try:
            for item in collection:
                names.append(_safe_getattr(item, "Name", "") or _safe_getattr(item, "Имя", "") or str(item))
        except Exception as error:
            lines.append(f"{title}: ошибка перебора: {error}")
            return

        lines.append(f"{title}: {', '.join(names) if names else '<пусто>'}")

    append_collection("Измерения", "Dimensions")
    append_collection("Ресурсы", "Resources")
    append_collection("Реквизиты", "Attributes")
    return lines


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

        normalized_title = _normalize_requisite_name(title)

        if "бик" in normalized_title and not target.get(bik_key):
            target[bik_key] = value
        elif ("коррсчет" in normalized_title or "корсчет" in normalized_title or "корреспондент" in normalized_title) and not target.get(corr_key):
            target[corr_key] = value
        elif ("расчетныйсчет" in normalized_title or "расчетный" in normalized_title or "рассчетный" in normalized_title) and not target.get(account_key):
            target[account_key] = value


def personal_info(connector: win32com.client, fio: str) -> dict:
    parsed_pers_info = {
        'ФИО': '',
        'ПОЛ': '',
        'ДатаРождения': '',
        'ИНН': '',
        'Страна': '',
        'МестоРождения': '',
        'СНИЛС': '',
        'Адрес_прописка': '',
        'Адрес_проживания': '',
        'Телефон': '',
        'Банк': '',
        'СемейноеПоложение': '',
        'ДатаРегистрации': '',

        'Родственники': {
            'Фамилия': '',
            'Имя': '',
            'Отчество': '',
            'ДатаРождения': '',
            'СтепеньРодства': ''
        },
        'РодственникиСписок': [],

        'Образование': {
            'УчебноеЗаведение': '',
            'ВидОбразования': '',
            'Квалификация': '',
            'Специальность': '',
            'ПериодОбучения': '',
            'Серия': '',
            'Номер': '',
        },

        'АдресБанка': '',
        'ИннБанка': '',
        'НомерСчета': '',
        'БикБанка': '',
        'КоррСчет': '',
        'РасчетныйСчет': '',
        'ВодительскоеУдостоверение': {
            'Серия': '',
            'Номер': '',
            'ДатаВыдачи': '',
            'КатегорииТС': '',
        }
    }

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ
             Сотрудники.ФизическоеЛицо.ФИО КАК ФИО,
             Сотрудники.ФизическоеЛицо.Пол КАК Пол,
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
    phys_person_ref = None

    while selection.Next():
        gender = selection.Пол
        gender_str = connector.String(gender)

        birth_date = selection.ДатаРождения
        formatted_date = birth_date.strftime('%d.%m.%Y')

        parsed_pers_info['ФИО'] = selection.ФИО
        parsed_pers_info['ДатаРождения'] = formatted_date
        parsed_pers_info['ПОЛ'] = gender_str
        parsed_pers_info['ИНН'] = selection.ИНН
        parsed_pers_info['МестоРождения'] = selection.МестоРождения
        parsed_pers_info['СНИЛС'] = selection.СНИЛС
        phys_person_ref = selection.ФизЛицо


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
    country_name = ''

    while selection_2.Next():
        if not country_name:
            country_name = _extract_country_name(_safe_getattr(selection_2, "Страна", None))

        # print(f'Контактная информация: {selection_2.ФизЛицо.КонтактнаяИнформация}')

        entry_address: int = 0
        entry_telephone: int = 0

        if len(phys_person_ref.КонтактнаяИнформация):

            table_len: int = len(phys_person_ref.КонтактнаяИнформация)

            for i in range(table_len):

                if len(phys_person_ref.КонтактнаяИнформация[i].НомерТелефона):
                    if entry_telephone == 0:
                        phone_number = phys_person_ref.КонтактнаяИнформация[i].НомерТелефона
                        entry_telephone += 1

                if (len(phys_person_ref.КонтактнаяИнформация[i].Страна)
                        or len(phys_person_ref.КонтактнаяИнформация[i].Регион)):

                    if entry_address == 0:

                        reg_address = phys_person_ref.КонтактнаяИнформация[i].Представление

                        fact_address = phys_person_ref.КонтактнаяИнформация[i + 1].Представление

                        entry_address += 1

    parsed_pers_info['Адрес_прописка'] = reg_address
    parsed_pers_info['Адрес_проживания'] = fact_address
    parsed_pers_info['Телефон'] = phone_number
    parsed_pers_info['Страна'] = country_name

    # Личные банковские реквизиты здесь намеренно не читаем из общего справочника
    # счетов контрагентов. Для повторно оформляемых сотрудников он часто содержит
    # исторические счета из прошлых приемов. Актуальные реквизиты добираются позже
    # через enrich_bank_details_from_1c, где сначала проверяется текущий зарплатный
    # проект и данные, привязанные непосредственно к сотруднику/физлицу.


    query_4 = connector.NewObject("Запрос")
    query_4.Text = """
           ВЫБРАТЬ
               Родственники.Фамилия КАК Фамилия,
               Родственники.Имя КАК Имя,
               Родственники.Отчество КАК Отчество,
               Родственники.ДатаРождения КАК ДатаРождения,
               Родственники.СтепеньРодства.Наименование КАК СтепеньРодства 
           ИЗ
               Справочник.РодственникиФизическихЛиц КАК Родственники
           ГДЕ
               Родственники.Владелец.ФИО = &ФИО
    """

    query_4.SetParameter("ФИО", fio)
    selection_4 = query_4.Execute().Choose()

    relatives_list: list[dict] = []

    while selection_4.Next():
        relative_birth_date = selection_4.ДатаРождения
        formatted_date = ''
        if relative_birth_date:
            formatted_date = relative_birth_date.strftime('%d.%m.%Y')

        relative_data = {
            'Фамилия': selection_4.Фамилия or '',
            'Имя': selection_4.Имя or '',
            'Отчество': selection_4.Отчество or '',
            'ДатаРождения': formatted_date,
            'СтепеньРодства': selection_4.СтепеньРодства or '',
        }
        relatives_list.append(relative_data)

    parsed_pers_info['РодственникиСписок'] = relatives_list

    if relatives_list:
        parsed_pers_info['Родственники'] = relatives_list[0].copy()

    family_status = ''
    registration_date = ''

    family_status = _query_family_status_from_marriage_register(connector, fio)

    if not family_status and phys_person_ref is not None:
        family_status = _extract_first_present(
            phys_person_ref,
            (
                'СемейноеПоложение',
                'СемейноеПоложениеФизЛица',
                'СемейноеПоложениеФизическогоЛица',
                'БрачноеСостояние',
                'СоставСемьи',
            ),
            lambda value: _extract_1c_presentation(connector, value)
        )
        registration_date = _extract_first_present(
            phys_person_ref,
            (
                'ДатаРегистрации',
                'ДатаРегистрацииПоМестуЖительства',
                'ДатаРегистрацииПоМестуПроживания',
                'ДатаПостояннойРегистрации',
            ),
            _format_1c_date
        )

    if not family_status:
        for field_name in (
            'СемейноеПоложение',
            'СемейноеПоложениеФизЛица',
            'СемейноеПоложениеФизическогоЛица',
            'БрачноеСостояние',
            'СоставСемьи',
        ):
            family_status = _query_first_phys_person_field_presentation(connector, fio, field_name)
            if not family_status:
                family_status = _extract_1c_presentation(connector, _query_first_phys_person_field(connector, fio, field_name))
            if family_status:
                break

    if not registration_date:
        for field_name in (
            'ДатаРегистрации',
            'ДатаРегистрацииПоМестуЖительства',
            'ДатаРегистрацииПоМестуПроживания',
            'ДатаПостояннойРегистрации',
        ):
            registration_date = _format_1c_date(_query_first_phys_person_field(connector, fio, field_name))
            if registration_date:
                break

    parsed_pers_info['СемейноеПоложение'] = family_status
    parsed_pers_info['ДатаРегистрации'] = registration_date
    parsed_pers_info['ВодительскоеУдостоверение'] = _get_driver_license_info(connector, fio)

    query_5 = connector.NewObject("Запрос")
    query_5.Text = """
               ВЫБРАТЬ
                   Образование.УчебноеЗаведение.Наименование КАК УчебноеЗаведение,
                   Образование.ВидОбразования.Наименование КАК ВидОбразования,
                   Образование.Квалификация КАК Квалификация,
                   Образование.Специальность.Наименование КАК Специальность,
                   Образование.Серия КАК Серия,
                   Образование.Номер КАК Номер,
                   Образование.Начало КАК Начало,
                   Образование.Окончание КАК Окончание
               ИЗ
                   Справочник.ОбразованиеФизическихЛиц КАК Образование
               ГДЕ
                   Образование.Владелец.ФИО = &ФИО
        """

    query_5.SetParameter("ФИО", fio)
    selection_5 = query_5.Execute().Choose()

    while selection_5.Next():
        parsed_pers_info['Образование']['УчебноеЗаведение'] = selection_5.УчебноеЗаведение
        parsed_pers_info['Образование']['ВидОбразования'] = selection_5.ВидОбразования
        parsed_pers_info['Образование']['Квалификация'] = selection_5.Квалификация
        parsed_pers_info['Образование']['Специальность'] = selection_5.Специальность
        parsed_pers_info['Образование']['Серия'] = selection_5.Серия
        parsed_pers_info['Образование']['Номер'] = selection_5.Номер

        end_date = _normalize_1c_empty_date(selection_5.Окончание)
        parsed_pers_info['Образование']['ПериодОбучения'] = _format_1c_date(end_date)

    military_document_info = _get_military_document_info(connector, fio)
    return parsed_pers_info


def job_info(connector: win32com.client, fio: str, inn_org: str = "") -> dict:
    parsed_job_info = {
        'ФИО': '',                          # Есть
        'ТкНомер': '',                      # Есть
        'ТабельныйНомерРаботника': '',
        'НомерПриказаОПриеме': '',
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
            'КорСчет': '',

            'ФИОПредседателяКомиссии': '',
            'ФамилияПредседателяКомиссии': '',
            'ИнициалыИмениОтчестваПредседателяКомиссии': '',
            'ДолжностьПредседателяКомиссии': '',
            'ФИОЧленаКомиссии1': '',
            'ФамилияЧленаКомиссии1': '',
            'ИнициалыИмениОтчестваЧленаКомиссии1': '',
            'ДолжностьЧленаКомиссии1': '',
            'ФИОЧленаКомиссии2': '',
            'ФамилияЧленаКомиссии2': '',
            'ИнициалыИмениОтчестваЧленаКомиссии2': '',
            'ДолжностьЧленаКомиссии2': '',
            'ФИООтветственногоЗаОТ': '',
            'ФамилияОтветственногоЗаОТ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаОТ': '',
            'ДолжностьОтветственногоЗаОТ': '',
            'ФИООтветственногоЗаПБ': '',
            'ФамилияОтветственногоЗаПБ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаПБ': '',
            'ДолжностьОтветственногоЗаПБ': '',
            'ФИООтветственногоЗаПроведениеИнструктажейПоПБ': '',
            'ФамилияОтветственногоЗаПроведениеИнструктажейПоПБ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаПроведениеИнструктажейПоПБ': '',
            'ДолжностьОтветственногоЗаПроведениеИнструктажейПоПБ': '',
            'ФИООтветственногоЗаПроведениеИнструктажейПоОТ': '',
            'ФамилияОтветственногоЗаПроведениеИнструктажейПоОТ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаПроведениеИнструктажейПоОТ': '',
            'ДолжностьОтветственногоЗаПроведениеИнструктажейПоОТ': '',
            'ФИООтветственногоЗаВыдачуСИЗ': '',
            'ФамилияОтветственногоЗаВыдачуСИЗ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаВыдачуСИЗ': '',
            'ДолжностьОтветственногоЗаВыдачуСИЗ': '',
            'ФИООтветственногоЗаГОиЧС': '',
            'ФамилияОтветственногоЗаГОиЧС': '',
            'ИнициалыИмениОтчестваОтветственногоЗаГОиЧС': '',
            'ДолжностьОтветственногоЗаГОиЧС': '',
            'ФИООтветственногоЗаЭБ': '',
            'ФамилияОтветственногоЗаЭБ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаЭБ': '',
            'ДолжностьОтветственногоЗаЭБ': '',
            'ФИООтветственногоЗаАТЗ': '',
            'ФамилияОтветственногоЗаАТЗ': '',
            'ИнициалыИмениОтчестваОтветственногоЗаАТЗ': '',
            'ДолжностьОтветственногоЗаАТЗ': '',

        },
        'Должность': '',                    # Есть
        'Подразделение': '',                # Есть
        'Оклад': '',                        # Есть
        'ИспытательныйСрок': '',            # Есть
        'АвансПериод': '',                  # Есть
        'ЗарплатаПериод': '',               # Есть
        'ГрафикРаботы': '',                 # Есть
        'ДлительностьРабочейНедели': '',
        'ВидЗанятости': '',
        'Начисление': '',
        'ОплатаПоЧасовомуТарифу': '',
        'ВидТарифнойСтавки': '',
        'СрочныйДоговор': '',
        'СрочныйТрудовойДоговорДо': '',
        'ОснованиеСрочногоДоговора': '',
        'Патент': '',
        'ЧекбоксПатент': '',
        'СерияПатента': '',
        'НомерПатента': '',
        'ДатаВыдачиПатента': '',
        'КемВыданПатент': '',
        'Разрешение': '',
        'ЧекбоксРазрешение': '',
        'СерияРазрешения': '',
        'НомерРазрешения': '',
        'ДатаВыдачиРазрешения': '',
        'КемВыданРазрешения': '',
        'ВидДокументаМигранта': '',
        'ДатаПриемаНаРаботу': '',           # Есть
        'Отпуск': '28',                     # (Пока временно поставил 28)
    }

    employee = getattr(connector.Catalogs, "Сотрудники")
    needed_employee = employee.FindByDescription(fio)
    phys_person_ref = _safe_getattr(needed_employee, "ФизическоеЛицо", None)

    query = connector.NewObject("Запрос")
    query_filters = []
    if phys_person_ref is not None:
        query_filters.append("ПриемНаРаботу.ФизическоеЛицо = &ФизическоеЛицо")
    if needed_employee:
        query_filters.append("ПриемНаРаботу.Сотрудник = &ФИО")
    if not query_filters:
        return parsed_job_info

    company_filter = ""
    if str(inn_org or "").strip():
        company_filter = "\n                            И ПриемНаРаботу.Организация.ИНН = &ИННОрг"

    query.Text = f"""
                       ВЫБРАТЬ ПЕРВЫЕ 1
                           ПриемНаРаботу.Ссылка как Ссылка,
                           ПриемНаРаботу.Организация как ОрганизацияСсылка,
                           ПриемНаРаботу.Руководитель как Руководитель,
                           ПриемНаРаботу.ДолжностьРуководителя как ДолжностьРуководителя,
                           ПриемНаРаботу.ТрудовойДоговорНомер как ТрудовойДоговорНомер,
                           ПриемНаРаботу.Номер как НомерПриказаОПриеме,
                           ПриемНаРаботу.ДатаПриема как ДатаПриема,
                           ПриемНаРаботу.Подразделение как Подразделение,
                           ПриемНаРаботу.Должность как Должность,
                           ПриемНаРаботу.ГрафикРаботы как График,
                           ПриемНаРаботу.ВидЗанятости как ВидЗанятости,
                           ПриемНаРаботу.ДлительностьИспытательногоСрока как ИспытательныйСрок,
                           ПриемНаРаботу.Аванс как Аванс,
                           ПриемНаРаботу.СовокупнаяТарифнаяСтавка как Оклад,
                           ПриемНаРаботу.ЕжегодныеОтпуска как ЕжегодныеОтпуска
                       ИЗ
                           Документ.ПриемНаРаботу КАК ПриемНаРаботу
                       ГДЕ
                            ({" ИЛИ ".join(query_filters)}){company_filter}
                       УПОРЯДОЧИТЬ ПО
                           ПриемНаРаботу.Дата УБЫВ,
                           ПриемНаРаботу.Ссылка УБЫВ
                        """

    if needed_employee:
        query.SetParameter("ФИО", needed_employee)
    if phys_person_ref is not None:
        query.SetParameter("ФизическоеЛицо", phys_person_ref)
    if str(inn_org or "").strip():
        query.SetParameter("ИННОрг", str(inn_org).strip())
    selection = query.Execute().Choose()
    organization_ref = None
    responsible_person_ref = None
    signing_basis_debug: list[str] = []

    while selection.Next():
        asigned_job_date = selection.ДатаПриема
        formatted_asigned_job_date = asigned_job_date.strftime('%d.%m.%Y')

        parsed_job_info['ФИО'] = fio
        parsed_job_info['ТкНомер'] = selection.ТрудовойДоговорНомер
        parsed_job_info['ТабельныйНомерРаботника'] = (
            _safe_getattr(needed_employee, "ТабельныйНомер", "")
            or _safe_getattr(needed_employee, "Код", "")
        )
        parsed_job_info['НомерПриказаОПриеме'] = selection.НомерПриказаОПриеме

        organization_ref = selection.ОрганизацияСсылка
        responsible_person_ref = _safe_getattr(selection.Руководитель, "ФизическоеЛицо", None)
        if responsible_person_ref is None:
            responsible_person_ref = _safe_getattr(selection, "Руководитель", None)

        parsed_job_info['Организация']['НаименованиеПолное'] = organization_ref.НаименованиеПолное
        parsed_job_info['Организация']['НаименованиеСокращенное'] = organization_ref.НаименованиеСокращенное
        parsed_job_info['Организация']['Руководитель'] = selection.Руководитель.Наименование
        parsed_job_info['Организация']['ДолжностьРуководителя'] = selection.ДолжностьРуководителя.Наименование
        parsed_job_info['Организация']['ИНН_Орг'] = organization_ref.ИНН
        parsed_job_info['Организация']['ОГРН(ИП)'] = organization_ref.ОГРН
        parsed_job_info['Организация']['КПП_Орг'] = organization_ref.РегистрацияВНалоговомОргане.КПП

        parsed_job_info['Подразделение'] = selection.Подразделение.Наименование
        parsed_job_info['Должность'] = selection.Должность.Наименование
        parsed_job_info['ГрафикРаботы'] = selection.График.Наименование
        parsed_job_info['ДлительностьРабочейНедели'] = _safe_getattr(selection.График, "ДлительностьРабочейНедели", "")
        parsed_job_info['ВидЗанятости'] = _extract_1c_presentation(connector, selection.ВидЗанятости)
        parsed_job_info['ДатаПриемаНаРаботу'] = formatted_asigned_job_date
        parsed_job_info['ИспытательныйСрок'] = selection.ИспытательныйСрок
        parsed_job_info['Оклад'] = selection.Оклад

        hire_doc_ref = _safe_getattr(selection, "Ссылка", None)
        hire_doc_obj = None
        if hire_doc_ref is not None:
            try:
                hire_doc_obj = hire_doc_ref.ПолучитьОбъект()
            except Exception:
                hire_doc_obj = hire_doc_ref

        payroll_source_objects = []
        if hire_doc_obj is not None:
            payroll_source_objects.append(hire_doc_obj)
        payroll_source_objects.append(selection)

        payroll_type_candidates = (
            "Начисление",
            "ОсновноеНачисление",
            "НачислениеОплатыТруда",
            "ВидНачисления",
        )
        hourly_rate_candidates = (
            "ОплатаПоЧасовомуТарифу",
            "ЧасоваяТарифнаяСтавка",
            "ТарифнаяСтавка",
            "СтоимостьЧаса",
            "ОплатаЗаЧас",
        )
        tariff_kind_candidates = (
            "ВидТарифнойСтавки",
            "ТипТарифнойСтавки",
        )
        payroll_type_discovered = _collect_requisite_names(
            hire_doc_obj,
            ("начисл", "оплат", "тариф"),
        )
        hourly_rate_discovered = _collect_requisite_names(
            hire_doc_obj,
            ("час", "тариф", "ставк", "оплат"),
        )
        tariff_kind_discovered = _collect_requisite_names(
            hire_doc_obj,
            ("видтариф", "типтариф", "час", "тариф"),
        )
        payroll_type_attr_names = list(dict.fromkeys(payroll_type_candidates + tuple(payroll_type_discovered)))
        hourly_rate_attr_names = list(dict.fromkeys(hourly_rate_candidates + tuple(hourly_rate_discovered)))
        tariff_kind_attr_names = list(dict.fromkeys(tariff_kind_candidates + tuple(tariff_kind_discovered)))

        parsed_job_info["Начисление"] = _extract_first_meaningful_attr_text(
            connector,
            payroll_source_objects,
            payroll_type_attr_names,
            skip_boolean_like=True,
        )
        parsed_job_info["ОплатаПоЧасовомуТарифу"] = _extract_first_meaningful_attr_text(
            connector,
            payroll_source_objects,
            hourly_rate_attr_names,
        )
        parsed_job_info["ВидТарифнойСтавки"] = _extract_first_meaningful_attr_text(
            connector,
            payroll_source_objects,
            tariff_kind_attr_names,
            skip_boolean_like=True,
        )

        if (
            "оплата по часовому тарифу" in _sanitize_1c_text(parsed_job_info["Начисление"]).lower()
            and not parsed_job_info["ОплатаПоЧасовомуТарифу"]
            and parsed_job_info["Оклад"] not in (None, "")
        ):
            parsed_job_info["ОплатаПоЧасовомуТарифу"] = str(parsed_job_info["Оклад"]).strip()

        _write_personal_info_debug(
            "payroll_debug: "
            f"fio={fio!r}; "
            f"payroll_type={parsed_job_info['Начисление']!r}; "
            f"hourly_rate={parsed_job_info['ОплатаПоЧасовомуТарифу']!r}; "
            f"tariff_kind={parsed_job_info['ВидТарифнойСтавки']!r}; "
            f"salary_value={parsed_job_info['Оклад']!r}; "
            f"payroll_type_attrs={payroll_type_attr_names!r}; "
            f"hourly_rate_attrs={hourly_rate_attr_names!r}; "
            f"tariff_kind_attrs={tariff_kind_attr_names!r}"
        )

        urgent_contract_till = None
        urgent_contract_basis = ""
        urgent_date_candidates = (
            "ДатаЗавершенияТрудовогоДоговора",
            "СрочныйТрудовойДоговорДо",
            "СрокТрудовогоДоговораДо",
            "ДатаОкончанияСрочногоТрудовогоДоговора",
            "ДатаОкончанияТрудовогоДоговора",
            "ДатаОкончанияДоговора",
            "СрокДоговораДо",
            "ДатаОкончанияСрочногоДоговора",
            "СрочныйДоговорДо",
            "ДатаПо",
        )
        urgent_basis_candidates = (
            "ОснованиеСрочногоДоговора",
            "ОснованиеЗаключенияСрочногоТрудовогоДоговора",
            "ОснованиеСрочногоТрудовогоДоговора",
        )
        urgent_contract_flag = _safe_getattr(hire_doc_obj, "СрочныйДоговор", None) if hire_doc_obj is not None else None
        urgent_source_objects = []
        if hire_doc_obj is not None:
            urgent_source_objects.append(hire_doc_obj)
        urgent_source_objects.append(selection)

        for source_obj in urgent_source_objects:
            for attr_name in urgent_date_candidates:
                candidate_date = _normalize_1c_empty_date(_safe_getattr(source_obj, attr_name, None))
                if candidate_date not in (None, ""):
                    urgent_contract_till = candidate_date
                    break
            if urgent_contract_till not in (None, ""):
                break

        for source_obj in urgent_source_objects:
            for attr_name in urgent_basis_candidates:
                raw_basis = _safe_getattr(source_obj, attr_name, "")
                urgent_contract_basis = _sanitize_1c_text(_extract_1c_presentation(connector, raw_basis))
                if urgent_contract_basis:
                    break
            if urgent_contract_basis:
                break

        urgent_contract_enabled = _is_truthy_1c(urgent_contract_flag) or urgent_contract_till not in (None, "")
        if not urgent_contract_enabled:
            urgent_contract_basis = ""

        _write_personal_info_debug("=== URGENT CONTRACT DEBUG START ===")
        _write_personal_info_debug(f"fio: {fio}")
        _write_personal_info_debug(f"hire_doc_ref: {hire_doc_ref}")
        _write_personal_info_debug(
            f"hire_doc_obj_type: {type(hire_doc_obj).__name__ if hire_doc_obj is not None else ''}"
        )
        _write_personal_info_debug(f"urgent_contract_flag: {urgent_contract_flag!r}")
        if hire_doc_obj is not None:
            try:
                available_attrs = [name for name in dir(hire_doc_obj) if any(marker in name for marker in ("Срок", "Сроч", "Оконч", "Договор", "Дата"))]
                _write_personal_info_debug(f"available_attrs_filtered: {available_attrs}")
            except Exception as debug_error:
                _write_personal_info_debug(f"available_attrs_filtered_error: {debug_error}")
            try:
                metadata_method = getattr(hire_doc_obj, "Метаданные", None)
                metadata = metadata_method() if callable(metadata_method) else None
                metadata_requisites = []
                if metadata is not None:
                    for requisite in metadata.Реквизиты:
                        req_name = _safe_getattr(requisite, "Имя", "")
                        if req_name and any(marker in str(req_name) for marker in ("Срок", "Сроч", "Оконч", "Договор", "Дата")):
                            metadata_requisites.append(str(req_name))
                _write_personal_info_debug(f"metadata_requisites_filtered: {metadata_requisites}")
            except Exception as metadata_error:
                _write_personal_info_debug(f"metadata_requisites_filtered_error: {metadata_error}")
        for attr_name in urgent_date_candidates:
            obj_value = _safe_getattr(hire_doc_obj, attr_name, None) if hire_doc_obj is not None else None
            sel_value = _safe_getattr(selection, attr_name, None)
            _write_personal_info_debug(f"date_attr {attr_name}: obj={obj_value!r}; selection={sel_value!r}")
        for attr_name in urgent_basis_candidates:
            obj_value = _safe_getattr(hire_doc_obj, attr_name, None) if hire_doc_obj is not None else None
            sel_value = _safe_getattr(selection, attr_name, None)
            _write_personal_info_debug(f"basis_attr {attr_name}: obj={obj_value!r}; selection={sel_value!r}")
        _write_personal_info_debug(f"urgent_contract_till_result: {urgent_contract_till!r}")
        _write_personal_info_debug(f"urgent_contract_basis_result: {urgent_contract_basis!r}")
        _write_personal_info_debug(f"urgent_contract_enabled: {urgent_contract_enabled!r}")
        _write_personal_info_debug("=== URGENT CONTRACT DEBUG END ===")

        parsed_job_info['СрочныйДоговор'] = 'Да' if urgent_contract_enabled else 'Нет'

        if urgent_contract_enabled and urgent_contract_till not in (None, ""):
            try:
                parsed_job_info['СрочныйТрудовойДоговорДо'] = urgent_contract_till.strftime('%d.%m.%Y')
            except Exception:
                parsed_job_info['СрочныйТрудовойДоговорДо'] = str(urgent_contract_till).strip()

        parsed_job_info['ОснованиеСрочногоДоговора'] = urgent_contract_basis if urgent_contract_enabled else ''

        reg_address, fact_address = '', ''
        phone_number = ''

        entry_address: int = 0
        entry_telephone: int = 0

        if  len(organization_ref.КонтактнаяИнформация):
                table_len: int = len(organization_ref.КонтактнаяИнформация)

                for i in range(table_len):

                    if len(organization_ref.КонтактнаяИнформация[i].НомерТелефона):
                        if entry_telephone == 0:

                            phone_number = organization_ref.КонтактнаяИнформация[i].НомерТелефона
                            entry_telephone += 1

                    if (len(organization_ref.КонтактнаяИнформация[i].Страна)
                            or len(organization_ref.КонтактнаяИнформация[i].Регион)):

                        if entry_address == 0:
                            reg_address = organization_ref.КонтактнаяИнформация[i].Представление
                            fact_address = organization_ref.КонтактнаяИнформация[i + 1].Представление

                            entry_address += 1

        # Пытался по аналогии проитерироваться по selection.Отпуск - возникает ошибка

        parsed_job_info['Организация']['ЮридическийАдрес'] = reg_address
        parsed_job_info['Организация']['ФактическийАдрес'] = fact_address
        parsed_job_info['Организация']['Телефон'] = phone_number

        additional_requisites = _safe_getattr(organization_ref, "ДополнительныеРеквизиты", None)
        try:
            additional_requisites_len = len(additional_requisites)
        except Exception:
            additional_requisites_len = 0

        commission_requisite_map = {
            "председателькомиссии": (
                "ФИОПредседателяКомиссии",
                "ФамилияПредседателяКомиссии",
                "ИнициалыИмениОтчестваПредседателяКомиссии",
                "ДолжностьПредседателяКомиссии",
            ),
            "членкомиссии1": (
                "ФИОЧленаКомиссии1",
                "ФамилияЧленаКомиссии1",
                "ИнициалыИмениОтчестваЧленаКомиссии1",
                "ДолжностьЧленаКомиссии1",
            ),
            "членкомиссии2": (
                "ФИОЧленаКомиссии2",
                "ФамилияЧленаКомиссии2",
                "ИнициалыИмениОтчестваЧленаКомиссии2",
                "ДолжностьЧленаКомиссии2",
            ),
        }

        responsible_requisite_map = {
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u043f\u043e\u043e\u0445\u0440\u0430\u043d\u0435\u0442\u0440\u0443\u0434\u0430": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041e\u0422",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041e\u0422",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041e\u0422",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041e\u0422",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u043f\u043e\u0436\u0430\u0440\u043d\u0443\u044e\u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u043e\u0441\u0442\u044c": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0411",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0411",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0411",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0411",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u0433\u043e\u0438\u0447\u0441": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0413\u041e\u0438\u0427\u0421",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0413\u041e\u0438\u0427\u0421",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0413\u041e\u0438\u0427\u0421",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0413\u041e\u0438\u0427\u0421",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u043f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u043f\u043e\u043f\u0431": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041f\u0411",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041f\u0411",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041f\u0411",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041f\u0411",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u043f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u043f\u043e\u043e\u0442": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041e\u0422",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041e\u0422",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041e\u0422",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u041f\u0440\u043e\u0432\u0435\u0434\u0435\u043d\u0438\u0435\u0418\u043d\u0441\u0442\u0440\u0443\u043a\u0442\u0430\u0436\u0435\u0439\u041f\u043e\u041e\u0422",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u0432\u044b\u0434\u0430\u0447\u0443\u0441\u0438\u0437": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0412\u044b\u0434\u0430\u0447\u0443\u0421\u0418\u0417",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0412\u044b\u0434\u0430\u0447\u0443\u0421\u0418\u0417",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0412\u044b\u0434\u0430\u0447\u0443\u0421\u0418\u0417",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0412\u044b\u0434\u0430\u0447\u0443\u0421\u0418\u0417",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u044d\u043b\u0435\u043a\u0442\u0440\u043e\u0431\u0435\u0437\u043e\u043f\u0430\u0441\u043d\u043e\u0441\u0442\u044c": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u042d\u0411",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u042d\u0411",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u042d\u0411",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u042d\u0411",
            ),
            "\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439\u0437\u0430\u0430\u0442\u0437": (
                "\u0424\u0418\u041e\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0410\u0422\u0417",
                "\u0424\u0430\u043c\u0438\u043b\u0438\u044f\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0410\u0422\u0417",
                "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b\u0418\u043c\u0435\u043d\u0438\u041e\u0442\u0447\u0435\u0441\u0442\u0432\u0430\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0410\u0422\u0417",
                "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0433\u043e\u0417\u0430\u0410\u0422\u0417",
            ),
        }

        for i in range(additional_requisites_len):
            try:
                title = additional_requisites[i].Свойство.Заголовок
                value = additional_requisites[i].Значение
            except Exception:
                continue

            normalized_title = _normalize_requisite_name(title)
            target_fields = commission_requisite_map.get(normalized_title)
            if not target_fields or not value:
                continue

            commission_member = _extract_commission_member_data_with_fallback(connector, value)
            org_target = parsed_job_info['Организация']
            org_target[target_fields[0]] = commission_member["ФИО"]
            org_target[target_fields[1]] = commission_member["Фамилия"]
            org_target[target_fields[2]] = commission_member["ИнициалыИмениОтчества"]
            org_target[target_fields[3]] = commission_member["Должность"]

        for i in range(additional_requisites_len):
            try:
                title = _safe_getattr(_safe_getattr(additional_requisites[i], "Свойство", None), "Заголовок", "")
                value = _safe_getattr(additional_requisites[i], "Значение", "")
            except Exception:
                continue

            normalized_title = _normalize_requisite_name(title)
            responsible_fields = responsible_requisite_map.get(normalized_title)
            if not responsible_fields or not value:
                continue

            responsible_member = _extract_commission_member_data_with_fallback(connector, value)
            org_target = parsed_job_info['Организация']
            org_target[responsible_fields[0]] = responsible_member["ФИО"]
            org_target[responsible_fields[1]] = responsible_member["Фамилия"]
            org_target[responsible_fields[2]] = responsible_member["ИнициалыИмениОтчества"]
            org_target[responsible_fields[3]] = responsible_member["Должность"]
            if not responsible_member["Должность"]:
                phys_person_ref_by_fio = _get_phys_person_ref_by_fio(connector, responsible_member["ФИО"])
                employee_ref_by_phys_person = _get_employee_ref_by_phys_person(connector, phys_person_ref_by_fio)
                employee_ref_by_fio = _get_employee_ref_by_fio(connector, responsible_member["ФИО"])
                position_by_fio = _get_employee_position_by_fio(connector, responsible_member["ФИО"])
                current_position_by_phys_person = _get_current_position_name_from_hr_register_slice_v2(connector, phys_person_ref=phys_person_ref_by_fio)
                current_position_by_employee = _get_current_position_name_from_hr_register_slice_v2(connector, employee_ref=employee_ref_by_fio)
                current_position_by_phys_person_ordered = _get_current_position_name_from_hr_register_ordered(connector, phys_person_ref=phys_person_ref_by_fio)
                current_position_by_employee_ordered = _get_current_position_name_from_hr_register_ordered(connector, employee_ref=employee_ref_by_fio)
                _write_personal_info_debug(
                    "responsible_position_missing: "
                    f"title={title!r}; "
                    f"fio={responsible_member['ФИО']!r}; "
                    f"value_type={type(value)}; "
                    f"raw_value={value!r}; "
                    f"phys_person_ref_by_fio={phys_person_ref_by_fio!r}; "
                    f"employee_ref_by_phys_person={employee_ref_by_phys_person!r}; "
                    f"employee_ref_by_fio={employee_ref_by_fio!r}; "
                    f"position_by_fio={position_by_fio!r}; "
                    f"current_position_by_phys_person={current_position_by_phys_person!r}; "
                    f"current_position_by_employee={current_position_by_employee!r}; "
                    f"current_position_by_phys_person_ordered={current_position_by_phys_person_ordered!r}; "
                    f"current_position_by_employee_ordered={current_position_by_employee_ordered!r}"
                )

        # vac_days: int | float = selection.ЕжегодныеОтпуска.КоличествоДнейВГод

        # if len(selection.Отпуск):
        #     table_len: int = len(selection.Отпуск)
        #
        #     for i in range(table_len):
        #         vac_days: int | float = selection.Отпуск[i].КоличествоДнейВГод
        #         pass

    if organization_ref is None:
        return parsed_job_info

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

    query_2.SetParameter("Организация", organization_ref)
    selection_2 = query_2.Execute().Choose()

    while selection_2.Next():
        parsed_job_info['АвансПериод'] = selection_2.АвансПериод
        parsed_job_info['ЗарплатаПериод'] = selection_2.ЗарплатаПериод

    parsed_job_info['ОснованиеПодписи'], signing_basis_debug = _try_get_signing_basis(
        connector=connector,
        organization_ref=organization_ref,
        responsible_person_ref=responsible_person_ref,
    )
    parsed_job_info['_debug_ОснованиеПодписи'] = signing_basis_debug
    _fill_migrant_data(connector, phys_person_ref, parsed_job_info)

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
        birth_date = selection_2.ФизЛицо.ДатаРождения
        formatted_birth_date  = birth_date .strftime('%d.%m.%Y')

        birth_place = selection_2.ФизЛицо.МестоРождения
        formatted_birth_place = birth_place.replace(',', '').replace('0', '')

        parsed_doc_info['МестоРождения'] = formatted_birth_place
        parsed_doc_info['ДатаРождения'] = formatted_birth_date

    return parsed_doc_info


def military_info(connector: win32com.client, fio: str, boss_fio: str, boss_position: str, org_name: str) -> dict:
    parsed_military_info: dict = {
        'ВУС': '',
        'Военкомат': '',
        'АдресВоенкомата': '',
        'КатегорияЗапаса': '',
        'ВоинскоеЗвание': '',
        'СоставПрофиль': '',
        'Годность': '',
        'ДокументВоинскогоУчета': '',
        'СерияДокументаВоинскогоУчета': '',
        'НомерДокументаВоинскогоУчета': '',
        'ОтделВыдавшийДокументВоинскогоУчета': '',
        'ДатаВыдачиДокументаВоинскогоУчета': '',

        'ИнициалыИмениОтчестваОтветственногоЗаВУ': '',
        'ДолжностьОтветственногоЗаВУ': '',
        'ФамилияОтветсвтенногоЗаВУ': ''
    }

    query = connector.NewObject("Запрос")
    query.Text = """
                      ВЫБРАТЬ
                          ВоинскийУчет.ВУС КАК ВУС,
                          ВоинскийУчет.Военкомат.Наименование КАК Военкомат,
                          ВоинскийУчет.Военкомат.Адрес КАК АдресВоенкомата,
                          ВоинскийУчет.КатегорияЗапаса.Ссылка КАК КатегорияЗапаса,
                          ВоинскийУчет.Звание.Наименование КАК ВоинскоеЗвание,
                          ВоинскийУчет.Состав.Наименование КАК СоставПрофиль,
                          ВоинскийУчет.Годность КАК Годность                      
    
                      ИЗ
                          РегистрСведений.ВоинскийУчет КАК ВоинскийУчет
                      ГДЕ
                          ВоинскийУчет.ФизическоеЛицо.ФИО = &ФИО
                      """

    query.SetParameter("ФИО", fio)
    selection = query.Execute().Choose()

    while selection.Next():
        parsed_military_info['ВУС'] = selection.ВУС
        parsed_military_info['Военкомат'] = selection.Военкомат
        parsed_military_info['АдресВоенкомата'] = selection.АдресВоенкомата
        parsed_military_info['КатегорияЗапаса'] = connector.String(selection.КатегорияЗапаса)
        parsed_military_info['ВоинскоеЗвание'] = selection.ВоинскоеЗвание
        parsed_military_info['СоставПрофиль'] = selection.СоставПрофиль
        parsed_military_info['Годность'] = connector.String(selection.Годность)
        parsed_military_info['ДокументВоинскогоУчета'] = _extract_first_present(
            selection,
            (
                'ДокументВоинскогоУчета',
                'ВидДокументаВоинскогоУчета',
                'ВидДокумента',
                'ВоинскийДокумент',
                'ВоенныйБилет',
                'ПриписноеСвидетельство',
            ),
            lambda value: _extract_1c_presentation(connector, value)
        )
        parsed_military_info['СерияДокументаВоинскогоУчета'] = _extract_first_present(
            selection,
            (
                'СерияДокументаВоинскогоУчета',
                'СерияДокумента',
                'Серия',
            ),
            lambda value: str(value).strip()
        )
        parsed_military_info['НомерДокументаВоинскогоУчета'] = _extract_first_present(
            selection,
            (
                'НомерДокументаВоинскогоУчета',
                'НомерДокумента',
                'Номер',
            ),
            lambda value: str(value).strip()
        )
        parsed_military_info['ОтделВыдавшийДокументВоинскогоУчета'] = _extract_first_present(
            selection,
            (
                'ОтделВыдавшийДокументВоинскогоУчета',
                'ОрганВыдачиДокументаВоинскогоУчета',
                'КемВыданДокументВоинскогоУчета',
                'ОрганВыдачи',
                'КемВыдан',
            ),
            lambda value: _extract_1c_presentation(connector, value)
        )
        parsed_military_info['ДатаВыдачиДокументаВоинскогоУчета'] = _extract_first_present(
            selection,
            (
                'ДатаВыдачиДокументаВоинскогоУчета',
                'ДатаВыдачиДокумента',
                'ДатаВыдачи',
            ),
            _format_1c_date
        )

        parsed_military_info['ИнициалыИмениОтчестваОтветственногоЗаВУ'] = get_name_patronymic_initials(boss_fio)
        parsed_military_info['ДолжностьОтветственногоЗаВУ'] = boss_position
        parsed_military_info['ФамилияОтветсвтенногоЗаВУ'] = get_fio_part(boss_fio, 0)

    extra_military_queries = {
        'ДокументВоинскогоУчета': (
            (
                'ДокументВоинскогоУчета',
                'ВидДокументаВоинскогоУчета',
                'ВидДокумента',
                'ВоинскийДокумент',
                'ВоенныйБилет',
                'ПриписноеСвидетельство',
            ),
            lambda value: _extract_1c_presentation(connector, value)
        ),
        'СерияДокументаВоинскогоУчета': (
            (
                'СерияДокументаВоинскогоУчета',
                'СерияДокумента',
                'Серия',
            ),
            lambda value: str(value).strip()
        ),
        'НомерДокументаВоинскогоУчета': (
            (
                'НомерДокументаВоинскогоУчета',
                'НомерДокумента',
                'Номер',
            ),
            lambda value: str(value).strip()
        ),
        'ОтделВыдавшийДокументВоинскогоУчета': (
            (
                'ОтделВыдавшийДокументВоинскогоУчета',
                'ОрганВыдачиДокументаВоинскогоУчета',
                'КемВыданДокументВоинскогоУчета',
                'ОрганВыдачи',
                'КемВыдан',
            ),
            lambda value: _extract_1c_presentation(connector, value)
        ),
        'ДатаВыдачиДокументаВоинскогоУчета': (
            (
                'ДатаВыдачиДокументаВоинскогоУчета',
                'ДатаВыдачиДокумента',
                'ДатаВыдачи',
            ),
            _format_1c_date
        ),
    }

    for target_key, (candidate_fields, formatter) in extra_military_queries.items():
        if parsed_military_info[target_key]:
            continue
        for field_name in candidate_fields:
            value = _query_single_value(
                connector,
                f"""
                      ВЫБРАТЬ ПЕРВЫЕ 1
                          ВоинскийУчет.{field_name} КАК Значение
                      ИЗ
                          РегистрСведений.ВоинскийУчет КАК ВоинскийУчет
                      ГДЕ
                          ВоинскийУчет.ФизическоеЛицо.ФИО = &ФИО
                """,
                "ФИО",
                fio,
                "Значение"
            )
            if value is None or value == "":
                continue
            try:
                parsed_military_info[target_key] = formatter(value)
            except Exception:
                parsed_military_info[target_key] = ""
            if parsed_military_info[target_key]:
                break

    military_document_info = _get_military_document_info(connector, fio)
    for key, value in military_document_info.items():
        if value and not parsed_military_info.get(key):
            parsed_military_info[key] = value

    print(
        "military_info_debug: "
        f"fio={fio!r}; "
        f"military_info={parsed_military_info!r}"
    )

    # query_2 = connector.NewObject("Запрос")
    # query_2.Text = """
    #                       ВЫБРАТЬ
    #                           ОснованияПолномочийОтветственныхЛиц.Должность.Наименование КАК Должность,
    #                           ОснованияПолномочийОтветственныхЛиц.ФизическоеЛицо.Наименование КАК ФИО
    #                       ИЗ
    #                           РегистрСведений.ОснованияПолномочийОтветственныхЛиц КАК ОснованияПолномочийОтветственныхЛиц
    #                       ГДЕ
    #                           ОснованияПолномочийОтветственныхЛиц.Организация.НаименованиеСокращенное = &НаименованиеКомпании
    # """
    #
    # query_2.SetParameter("НаименованиеКомпании", org_name)
    # selection_2 = query_2.Execute().Choose()
    #
    # while selection_2.Next():
    #     boss_fio_2 = selection_2.ФИО
    #     boss_position_2 = selection_2.Должность


    return parsed_military_info


def organisation_info(connector: win32com.client, inn_org: str) -> dict:
    org_info: dict = {
            'РабочаяДата': '',              # Есть
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
            'БикБанка': '',                 # Есть
            'РасСчет': '',                  # Есть
            'КорСчет': ''                   # Есть
    }

    query = connector.NewObject("Запрос")
    query.Text = """
           ВЫБРАТЬ
               ПриемНаРаботу.Организация как Организация,
               ПриемНаРаботу.Руководитель как Руководитель,
               ПриемНаРаботу.ДолжностьРуководителя как ДолжностьРуководителя,
               ПриемНаРаботу.ДатаПриема как ДатаПриема
           ИЗ
               Документ.ПриемНаРаботу КАК ПриемНаРаботу
           ГДЕ
               ПриемНаРаботу.Организация.ИНН = &ИНН
    """

    query.SetParameter("ИНН", inn_org)
    selection = query.Execute().Choose()

    while selection.Next():
        org_info['НаименованиеПолное'] = selection.Организация.НаименованиеПолное
        org_info['НаименованиеСокращенное'] = selection.Организация.НаименованиеСокращенное
        org_info['Руководитель'] = selection.Руководитель.Наименование
        org_info['ДолжностьРуководителя'] = selection.ДолжностьРуководителя.Наименование
        org_info['ИНН_Орг'] = selection.Организация.ИНН
        org_info['ОГРН(ИП)'] = selection.Организация.ОГРН
        org_info['КПП_Орг'] = selection.Организация.РегистрацияВНалоговомОргане.КПП

        reg_address, fact_address = '', ''
        phone_number = ''

        entry_address: int = 0
        entry_telephone: int = 0

        if len(selection.Организация.КонтактнаяИнформация):
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

        org_info['ЮридическийАдрес'] = reg_address
        org_info['ФактическийАдрес'] = fact_address
        org_info['Телефон'] = phone_number

        if len(selection.Организация.ДополнительныеРеквизиты):
            table_len: int = len(selection.Организация.ДополнительныеРеквизиты)

            for i in range(table_len):
                value = selection.Организация.ДополнительныеРеквизиты[i].Значение
                col_name = selection.Организация.ДополнительныеРеквизиты[i].Свойство.Заголовок

                if col_name == 'Расчетный счет':
                    org_info['РасСчет'] = value

                if col_name == 'БИК банка':
                    org_info['БикБанка'] = value

                if col_name == 'Корр счет':
                    org_info['КорСчет'] = value

        _fill_bank_fields(
            org_info,
            selection.Организация.ДополнительныеРеквизиты,
            'БикБанка',
            'КорСчет',
            'РасСчет'
        )


        today = datetime.datetime.today()

        today_str = today.strftime("%d.%m.%Y")
        org_info['РабочаяДата'] = today_str
        break

    return org_info


def app_1c(fio: str, base_1c: str, username: str, pswd: str, inn_org: str = "") -> dict | str:
    V83_CONN_STRING = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        return f"Ошибка соединения к базе 1С: {e}"

    personal: dict = personal_info(connector=V83, fio=fio)
    document: dict = document_info(connector=V83, fio=fio)
    job: dict = job_info(connector=V83, fio=fio, inn_org=inn_org)
    military: dict = military_info(connector=V83, fio=fio,
                                   boss_fio=job['Организация']['Руководитель'],
                                   boss_position=job['Организация']['ДолжностьРуководителя'],
                                   org_name=job['Организация']['НаименованиеСокращенное'])

    parsed_data_from_1c: dict = {
        'personal_info': personal,
        'document_info': document,
        'job_info': job,
        'military_info': military,
        '_debug_migrant_register_metadata': _describe_info_register_metadata(V83, "ДанныеМигранта"),
    }
    return parsed_data_from_1c


def only_org_info(inn_org: str, base_1c: str, username: str, pswd: str) -> dict | str:
    V83_CONN_STRING = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'

    try:
        V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except Exception as e:
        return f"Ошибка соединения к базе 1С: {e}"

    org_info: dict = organisation_info(connector=V83, inn_org=inn_org)

    parsed_data_org_1c: dict = {
        'org_info': org_info
    }
    return parsed_data_org_1c


if __name__ == '__main__':
    parsed_data: dict = app_1c(fio='Минин Павел Сергеевич', base_1c='Srvr="srv2024.apriori.loc:1541";Ref="zup_poezd";',
                               username='Робот', pswd='YUSliot3Cb')
    pass
