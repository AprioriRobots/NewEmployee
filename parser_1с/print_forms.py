import os
import shutil
import json
import time

import win32com.client
from pypdf import PdfReader, PdfWriter


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

DEFAULT_EXTERNAL_PROCESSOR_PATH = os.getenv(
    "ONEC_HR_EXTERNAL_PROCESSOR_PATH",
    os.path.join(os.path.dirname(__file__), "export_hr_docs_1c.epf"),
)
EXTERNAL_PROCESSOR_METHODS = _split_env_names(
    "ONEC_HR_EXTERNAL_PROCESSOR_METHODS",
    [
        os.getenv("ONEC_HR_EXTERNAL_PROCESSOR_METHOD", "СформироватьКадровыеДокументыПоФИОJSON"),
        "СформироватьКадровыеДокументыПоФИОJSON",
        "СформироватьКадровыеДокументыПоФИО",
    ],
)


def connect_1c(base_1c: str, username: str, pswd: str):
    conn_string = rf'{base_1c}Usr="{username}";Pwd="{pswd}";'
    return win32com.client.Dispatch("V83.COMConnector").Connect(conn_string)


def _create_external_processor(connector, processor_path: str):
    errors: list[str] = []
    normalized_path = os.path.abspath(processor_path)

    if not os.path.isfile(normalized_path):
        return None, [f"Файл внешней обработки не найден: {normalized_path}"]

    manager_names = ("ВнешниеОбработки", "ExternalDataProcessors")
    create_methods = ("Создать", "Create")
    argument_variants = ((normalized_path, False), (normalized_path,))

    for manager_name in manager_names:
        try:
            manager = getattr(connector, manager_name)
        except Exception as error:
            errors.append(f"{manager_name}: {error}")
            continue

        for create_method in create_methods:
            try:
                creator = getattr(manager, create_method)
            except Exception as error:
                errors.append(f"{manager_name}.{create_method}: {error}")
                continue

            for args in argument_variants:
                try:
                    return creator(*args), []
                except Exception as error:
                    errors.append(f"{manager_name}.{create_method}{args}: {error}")

    return None, errors


def _extract_result_value(result, key: str):
    if result is None:
        return None

    if isinstance(result, dict):
        return result.get(key)

    try:
        return getattr(result, key)
    except Exception:
        pass

    for method_name in ("Get", "Получить"):
        try:
            method = getattr(result, method_name)
            return method(key)
        except Exception:
            continue

    return None


def _normalize_external_result(raw_result):
    if isinstance(raw_result, str):
        stripped = raw_result.strip()
        if stripped.startswith("{") and stripped.endswith("}"):
            try:
                parsed = json.loads(stripped)
                if isinstance(parsed, dict) and "#value" in parsed and isinstance(parsed["#value"], list):
                    normalized = {}
                    for item in parsed["#value"]:
                        if not isinstance(item, dict):
                            continue
                        key_meta = item.get("name", {})
                        value_meta = item.get("Value", {})
                        key = key_meta.get("#value") if isinstance(key_meta, dict) else None
                        value = value_meta.get("#value") if isinstance(value_meta, dict) else None
                        if key is not None:
                            normalized[key] = value
                    return normalized
                return parsed
            except Exception:
                return {"success": False, "log": raw_result, "order_error": raw_result, "sedo_error": raw_result}
        return {"success": False, "log": raw_result, "order_error": raw_result, "sedo_error": raw_result}

    return {
        "success": bool(_extract_result_value(raw_result, "success")),
        "order_path": _extract_result_value(raw_result, "order_path") or "",
        "sedo_path": _extract_result_value(raw_result, "sedo_path") or "",
        "order_error": _extract_result_value(raw_result, "order_error") or "",
        "sedo_error": _extract_result_value(raw_result, "sedo_error") or "",
        "log": _extract_result_value(raw_result, "log") or "",
    }


def _try_external_processor_hr_docs(connector, fio: str, output_dir: str) -> dict[str, str] | None:
    processor, create_errors = _create_external_processor(connector, DEFAULT_EXTERNAL_PROCESSOR_PATH)
    if processor is None:
        return {
            "order": "Ошибка вызова внешней обработки 1С: " + " | ".join(create_errors[:10]),
            "sedo": "Ошибка вызова внешней обработки 1С: " + " | ".join(create_errors[:10]),
        }

    call_errors: list[str] = []
    for method_name in EXTERNAL_PROCESSOR_METHODS:
        try:
            method = getattr(processor, method_name)
        except Exception as error:
            call_errors.append(f"{method_name}: {error}")
            continue

        for args in ((fio, output_dir), (output_dir, fio)):
            try:
                raw_result = method(*args)
                parsed_result = _normalize_external_result(raw_result)
                external_log = str(parsed_result.get("log") or "").strip()
                order_value = parsed_result.get("order_path") or parsed_result.get("order_error") or ""
                sedo_value = parsed_result.get("sedo_path") or parsed_result.get("sedo_error") or ""

                if external_log:
                    print("=== 1C HR DOCS LOG START ===")
                    print(external_log)
                    print("=== 1C HR DOCS LOG END ===")
                    if order_value and not _is_existing_file_path(order_value):
                        order_value = f"{order_value}\nЛог 1С:\n{external_log}"
                    if sedo_value and not _is_existing_file_path(sedo_value):
                        sedo_value = f"{sedo_value}\nЛог 1С:\n{external_log}"

                return {
                    "order": order_value,
                    "sedo": sedo_value,
                }
            except Exception as error:
                call_errors.append(f"{method_name}{args}: {error}")

    return {
        "order": "Ошибка вызова внешней обработки 1С: " + " | ".join(call_errors[:10]),
        "sedo": "Ошибка вызова внешней обработки 1С: " + " | ".join(call_errors[:10]),
    }


def _is_existing_file_path(value: str) -> bool:
    return bool(value) and os.path.isfile(value)


def _keep_only_first_pdf_page(file_path: str) -> str:
    if not _is_existing_file_path(file_path) or not str(file_path).lower().endswith(".pdf"):
        return file_path

    try:
        reader = PdfReader(file_path)
        if len(reader.pages) <= 1:
            return file_path

        writer = PdfWriter()
        writer.add_page(reader.pages[0])

        temp_path = f"{file_path}.tmp"
        with open(temp_path, "wb") as target_file:
            writer.write(target_file)

        os.replace(temp_path, file_path)
        return file_path
    except Exception as error:
        print(f"Не удалось обрезать PDF до первой страницы: {file_path}: {error}")
        return file_path


def _should_return_external_result(external_result: dict[str, str] | None) -> bool:
    if not external_result:
        return False

    if _is_existing_file_path(external_result.get("order", "")) or _is_existing_file_path(external_result.get("sedo", "")):
        return True

    external_errors = " ".join(
        [
            str(external_result.get("order", "")),
            str(external_result.get("sedo", "")),
        ]
    ).lower()

    if "внешней обработки 1с" in external_errors:
        return True

    if DEFAULT_EXTERNAL_PROCESSOR_PATH and not os.path.isfile(os.path.abspath(DEFAULT_EXTERNAL_PROCESSOR_PATH)):
        return True

    return False


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
            time.sleep(5)
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
    order_path = os.path.join(output_dir, f"output_Приказ_о_приеме_{safe_fio}.pdf")
    sedo_path = os.path.join(output_dir, f"output_СЭДО_{safe_fio}.pdf")

    if DEFAULT_EXTERNAL_PROCESSOR_PATH:
        external_result = _try_external_processor_hr_docs(connector=connector, fio=fio, output_dir=output_dir)
        if _should_return_external_result(external_result):
            order_ready = external_result.get("order", "") if external_result else ""
            sedo_ready = external_result.get("sedo", "") if external_result else ""
            order_result = order_ready if _is_existing_file_path(order_ready) else ""
            sedo_result = sedo_ready if _is_existing_file_path(sedo_ready) else ""

            if not order_result:
                order_result = _render_print_form(
                    connector=connector,
                    doc_ref=hire_doc_ref,
                    print_form_names=ORDER_PRINT_FORM_NAMES,
                    output_path=order_path,
                )
            if _is_existing_file_path(order_result):
                order_result = _keep_only_first_pdf_page(order_result)

            if not sedo_result:
                sedo_result = _render_print_form(
                    connector=connector,
                    doc_ref=hire_doc_ref,
                    print_form_names=SEDO_PRINT_FORM_NAMES,
                    output_path=sedo_path,
                )

            return {
                "order": order_result or order_ready or "",
                "sedo": sedo_result or sedo_ready or "",
            }

    order_result = _render_print_form(
        connector=connector,
        doc_ref=hire_doc_ref,
        print_form_names=ORDER_PRINT_FORM_NAMES,
        output_path=order_path,
    )
    if _is_existing_file_path(order_result):
        order_result = _keep_only_first_pdf_page(order_result)

    return {
        "order": order_result,
        "sedo": _render_print_form(
            connector=connector,
            doc_ref=hire_doc_ref,
            print_form_names=SEDO_PRINT_FORM_NAMES,
            output_path=sedo_path,
        ),
    }
