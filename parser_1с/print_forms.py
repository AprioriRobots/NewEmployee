import os
import shutil

import win32com.client


def _split_env_names(env_name: str, default_names: list[str]) -> list[str]:
    raw_value = os.getenv(env_name, "")
    if not raw_value.strip():
        return default_names

    names = [item.strip() for item in raw_value.split(",") if item.strip()]
    return names or default_names


ORDER_PRINT_FORM_NAMES = _split_env_names(
    "ONEC_HR_ORDER_PRINT_FORM_NAMES",
    [
        os.getenv("ONEC_HR_ORDER_PRINT_FORM_NAME", "Приказ о приеме"),
        "Приказ о приеме",
        "ПриказОПриеме",
        "ПриемНаРаботуТ1",
        "Прием на работу Т-1",
        "Т-1",
        "Т1",
        "ФормаТ1",
    ],
)

SEDO_PRINT_FORM_NAMES = _split_env_names(
    "ONEC_HR_SEDO_PRINT_FORM_NAMES",
    [
        os.getenv("ONEC_HR_SEDO_PRINT_FORM_NAME", "СЭДО"),
        "СЭДО",
        "СЕДО",
        "Сведения для СЭДО",
        "Трудовая деятельность СЭДО",
        "Мероприятие трудовой деятельности",
        "ЕФС-1",
    ],
)


def connect_1c(base_1c: str, username: str, pswd: str):
    conn_string = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'
    return win32com.client.Dispatch("V83.COMConnector").Connect(conn_string)


def get_hire_doc_ref(connector, fio: str):
    employees = getattr(connector.Catalogs, "Сотрудники")
    needed_employee = employees.FindByDescription(fio)

    query = connector.NewObject("Запрос")
    query.Text = """
        ВЫБРАТЬ ПЕРВЫЕ 1
            ПриемНаРаботу.Ссылка КАК Ссылка
        ИЗ
            Документ.ПриемНаРаботу КАК ПриемНаРаботу
        ГДЕ
            ПриемНаРаботу.Сотрудник = &ФИО
        УПОРЯДОЧИТЬ ПО
            ПриемНаРаботу.Дата УБЫВ,
            ПриемНаРаботу.Номер УБЫВ
    """
    query.SetParameter("ФИО", needed_employee)
    selection = query.Execute().Choose()

    if selection.Next():
        return selection.Ссылка
    return None


def _make_array(connector, *items):
    array_1c = connector.NewObject("Массив")
    for item in items:
        array_1c.Добавить(item)
    return array_1c


def _save_print_result(result, output_path: str) -> bool:
    save_methods = ["Записать", "Write", "Сохранить", "Save"]
    for method_name in save_methods:
        try:
            method = getattr(result, method_name)
        except Exception:
            continue

        try:
            method(output_path)
            return os.path.isfile(output_path)
        except Exception:
            continue

    if isinstance(result, str) and os.path.isfile(result):
        shutil.copyfile(result, output_path)
        return True

    return False


def _try_call(target, method_name: str, *args):
    try:
        method = getattr(target, method_name)
    except Exception:
        return None

    try:
        return method(*args)
    except Exception:
        return None


def _build_print_form_collection(connector, form_name: str):
    collection = connector.NewObject("ТаблицаЗначений")
    for column_name in (
        "ИмяВРЕГ",
        "ИмяМакета",
        "СинонимМакета",
        "ТабличныйДокумент",
        "ПолныйПутьКМакету",
    ):
        try:
            collection.Колонки.Добавить(column_name)
        except Exception:
            pass

    try:
        row = collection.Добавить()
        for attr_name in ("ИмяВРЕГ", "ИмяМакета", "СинонимМакета"):
            try:
                setattr(row, attr_name, form_name)
            except Exception:
                continue
    except Exception:
        pass

    return collection


def _save_first_tabdoc_from_collection(collection, output_path: str) -> bool:
    try:
        count = collection.Количество()
    except Exception:
        try:
            count = len(collection)
        except Exception:
            return False

    for index in range(count):
        try:
            row = collection.Get(index)
        except Exception:
            try:
                row = collection[index]
            except Exception:
                continue

        for attr_name in ("ТабличныйДокумент",):
            try:
                tab_doc = getattr(row, attr_name)
            except Exception:
                continue
            if tab_doc and _save_print_result(tab_doc, output_path):
                return True

    return False


def _safe_dir(target) -> list[str]:
    try:
        return sorted({name for name in dir(target) if not name.startswith("_")})
    except Exception:
        return []


def _extract_collection_names(collection) -> list[str]:
    names: list[str] = []
    if collection is None:
        return names

    try:
        count = collection.Count()
    except Exception:
        try:
            count = len(collection)
        except Exception:
            return names

    for index in range(count):
        try:
            item = collection.Get(index)
        except Exception:
            try:
                item = collection[index]
            except Exception:
                continue

        for attr_name in ("Имя", "Наименование", "Представление", "Name"):
            try:
                value = getattr(item, attr_name)
            except Exception:
                continue
            if value:
                names.append(str(value))
                break

    return names


def _discover_print_commands(manager, doc_object) -> list[str]:
    discovered: list[str] = []
    for target in (manager, doc_object):
        for attr_name in ("КомандыПечати", "PrintCommands", "ПечатныеФормы"):
            try:
                collection = getattr(target, attr_name)
            except Exception:
                continue
            discovered.extend(_extract_collection_names(collection))

    return sorted({item for item in discovered if item})


def _try_procedure_print(connector, target, docs_array, doc_ref, form_name: str, output_path: str) -> tuple[bool, list[str]]:
    bsp_variants = [
        ("МассивДокументов", "КоллекцияПечатныхФорм", "ОбъектыПечати", "ПараметрыВывода"),
        ("МассивДокументов", "ПараметрыПечати", "КоллекцияПечатныхФорм", "ОбъектыПечати", "ПараметрыВывода"),
    ]

    procedure_variants = [
        ("ТабличныйДокумент", "МассивДокументов", "ИмяФормы"),
        ("МассивДокументов", "ТабличныйДокумент", "ИмяФормы"),
        ("ТабличныйДокумент", "ИмяФормы", "МассивДокументов"),
        ("ИмяФормы", "ТабличныйДокумент", "МассивДокументов"),
        ("МассивДокументов", "ИмяФормы", "ТабличныйДокумент"),
        ("ИмяФормы", "МассивДокументов", "ТабличныйДокумент"),
        ("ТабличныйДокумент", "Документ", "ИмяФормы"),
        ("Документ", "ТабличныйДокумент", "ИмяФормы"),
        ("ТабличныйДокумент", "ИмяФормы", "Документ"),
        ("ИмяФормы", "ТабличныйДокумент", "Документ"),
        ("Документ", "ИмяФормы", "ТабличныйДокумент"),
        ("ИмяФормы", "Документ", "ТабличныйДокумент"),
        ("ТабличныйДокумент", "МассивДокументов"),
        ("МассивДокументов", "ТабличныйДокумент"),
        ("ТабличныйДокумент", "Документ"),
        ("Документ", "ТабличныйДокумент"),
    ]

    value_factories = {
        "ТабличныйДокумент": lambda: connector.NewObject("ТабличныйДокумент"),
        "МассивДокументов": lambda: docs_array,
        "Документ": lambda: doc_ref,
        "ИмяФормы": lambda: form_name,
    }
    errors: list[str] = []

    for signature in bsp_variants:
        collection = _build_print_form_collection(connector, form_name)
        objects_print = connector.NewObject("СписокЗначений")
        params_output = connector.NewObject("Структура")
        params_print = connector.NewObject("Структура")

        bsp_value_map = {
            "МассивДокументов": docs_array,
            "КоллекцияПечатныхФорм": collection,
            "ОбъектыПечати": objects_print,
            "ПараметрыВывода": params_output,
            "ПараметрыПечати": params_print,
        }

        try:
            args = [bsp_value_map[token] for token in signature]
            getattr(target, "Печать")(*args)
        except Exception as error:
            errors.append(f"Печать({', '.join(signature)}): {error}")
            continue

        if _save_first_tabdoc_from_collection(collection, output_path):
            return True, errors

        errors.append(f"Печать({', '.join(signature)}): выполнено без ошибки, но печатная форма не найдена в коллекции")

    for signature in procedure_variants:
        tab_doc = connector.NewObject("ТабличныйДокумент")
        args = []
        for token in signature:
            if token == "ТабличныйДокумент":
                args.append(tab_doc)
            else:
                args.append(value_factories[token]())

        try:
            getattr(target, "Печать")(*args)
        except Exception as error:
            errors.append(f"Печать({', '.join(signature)}): {error}")
            continue

        if _save_print_result(tab_doc, output_path):
            return True, errors

        errors.append(f"Печать({', '.join(signature)}): выполнено без ошибки, но файл не сохранился")

    return False, errors


def _render_print_form(connector, doc_ref, print_form_names: list[str], output_path: str) -> str:
    manager = getattr(connector.Documents, "ПриемНаРаботу")
    doc_object = doc_ref.GetObject()
    docs_array = _make_array(connector, doc_ref)

    manager_methods = _safe_dir(manager)
    doc_methods = _safe_dir(doc_object)
    discovered_commands = _discover_print_commands(manager, doc_object)
    procedure_errors: list[str] = []

    candidates = []
    for form_name in print_form_names:
        candidates.extend(
            [
                (manager, "ПолучитьПечатнуюФорму", (form_name, docs_array)),
                (manager, "ПолучитьПечатнуюФорму", (form_name, doc_ref)),
                (manager, "СформироватьПечатнуюФорму", (form_name, docs_array)),
                (manager, "СформироватьПечатнуюФорму", (form_name, doc_ref)),
                (manager, "ПолучитьМакет", (form_name,)),
                (doc_object, "ПолучитьМакет", (form_name,)),
                (doc_object, "ПолучитьПечатнуюФорму", (form_name,)),
                (doc_object, "СформироватьПечатнуюФорму", (form_name,)),
            ]
        )

    for command_name in discovered_commands:
        candidates.extend(
            [
                (manager, "ПолучитьПечатнуюФорму", (command_name, docs_array)),
                (manager, "СформироватьПечатнуюФорму", (command_name, docs_array)),
            ]
        )

    for target, method_name, args in candidates:
        result = _try_call(target, method_name, *args)
        if result and _save_print_result(result, output_path):
            return output_path

    for form_name in print_form_names:
        success, errors = _try_procedure_print(
            connector=connector,
            target=manager,
            docs_array=docs_array,
            doc_ref=doc_ref,
            form_name=form_name,
            output_path=output_path,
        )
        procedure_errors.extend([f"{form_name}: {error}" for error in errors])
        if success:
            return output_path

    for command_name in discovered_commands:
        success, errors = _try_procedure_print(
            connector=connector,
            target=manager,
            docs_array=docs_array,
            doc_ref=doc_ref,
            form_name=command_name,
            output_path=output_path,
        )
        procedure_errors.extend([f"{command_name}: {error}" for error in errors])
        if success:
            return output_path

    return (
        "Ошибка во время формирования файла в 1С: "
        f"не удалось получить печатную форму. Проверить имя формы: {print_form_names[0]}. "
        f"Доступные команды печати: {', '.join(discovered_commands[:20]) or 'не найдены'}. "
        f"Методы менеджера: {', '.join(manager_methods[:20]) or 'не найдены'}. "
        f"Методы объекта: {', '.join(doc_methods[:20]) or 'не найдены'}. "
        f"Ошибки процедуры печати: {' | '.join(procedure_errors[:20]) or 'не зафиксированы'}."
    )


def form_order_and_sedo_docs_from_1c(fio: str, base_1c: str, username: str, pswd: str) -> dict[str, str]:
    try:
        connector = connect_1c(base_1c=base_1c, username=username, pswd=pswd)
    except Exception as error:
        error_msg = f"Ошибка соединения к базе 1С при формировании печатных форм: {error}"
        return {"order": error_msg, "sedo": error_msg}

    hire_doc_ref = get_hire_doc_ref(connector=connector, fio=fio)
    if not hire_doc_ref:
        error_msg = f"Ошибка во время формирования файла в 1С: не найден документ ПриемНаРаботу для {fio}"
        return {"order": error_msg, "sedo": error_msg}

    output_dir = os.path.join(os.path.dirname(__file__), "output_result")
    os.makedirs(output_dir, exist_ok=True)

    safe_fio = "_".join(fio.split())
    order_path = os.path.join(output_dir, f"output_Приказ_о_приеме_{safe_fio}.mxl")
    sedo_path = os.path.join(output_dir, f"output_СЭДО_{safe_fio}.mxl")

    return {
        "order": _render_print_form(
            connector=connector,
            doc_ref=hire_doc_ref,
            print_form_names=ORDER_PRINT_FORM_NAMES,
            output_path=order_path,
        ),
        "sedo": _render_print_form(
            connector=connector,
            doc_ref=hire_doc_ref,
            print_form_names=SEDO_PRINT_FORM_NAMES,
            output_path=sedo_path,
        ),
    }
