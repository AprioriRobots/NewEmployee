import os

from dotenv import load_dotenv

from bitrix.bitrix import add_checklist_item, add_task, get_deals_bitrix, upload_and_attach_file_to_task
from employee_instructions.form_empl_instr import employee_instr_sample_former
from formul_tkrf.form_formul_tkrf import formul_tkrf_sample_former
from lna.form_lna import lna_sample_former
from mat_response.mat_response import mat_resp_sample_former
from military_registr.military_reg import (
    application_sample_former,
    details_sample_former,
    f10_sample_former,
    info_changes_sample_former,
    list_msg_sample_former,
    msg_13_sample_former,
    notification_sample_former,
)
from nda.form_nda import nda_sample_former
from occup_safety.form_occup_safety import occup_safety_former
from parser_1с.print_forms import form_order_and_sedo_docs_from_1c
from parser_1с.personal_info import app_1c
from personal_secret.form_personal_secret import personal_secret_sample_former
from tk_rf.form_tkrf import tkrf_sample_former
from zp_alignment.form_zp_alignment import zp_align_sample_former


load_dotenv()

BITRIX_TOKEN = os.getenv("BITRIX_TOKEN")
ROBOT_LOGIN = os.getenv("ROBOT_LOGIN", "Робот")
ROBOT_PSWD = os.getenv("ROBOT_PSWD", "")
ROBOT_PSWD_DEFAULT = os.getenv("ROBOT_PSWD_DEFAULT", "YUSliot3Cb")
ROBOT_PSWD_PERELMAN = os.getenv("ROBOT_PSWD_PERELMAN", "Q1234567890Q")
DEFAULT_UPLOAD_FOLDER_ID = int(os.getenv("BITRIX_TASK_UPLOAD_FOLDER_ID", "193"))

HR_FLAG_FIELD = "ufCrm3_1770106591"
OCCUP_SAFETY_FLAG_FIELD = "ufCrm3_1770106632"
MILITARY_FLAG_FIELD = "ufCrm3_1770106622"
BASE_1C_FIELD = "ufCrm3_1770808935"

EMPLOYEE_INSTRUCTIONS = [
    "Бармен",
    "Бартендер",
    "Кондитер",
    "Котлочист",
    "Менеджер",
    "Мойщик посуды",
    "Оператор (call-центра)",
    "Официант",
    "Пекарь",
    "Повар",
    "Повар ГЦ",
    "Повар ЗЦ",
    "Повар ХЦ",
    "Су-шеф",
    "Уборщик производственных и служебных помещений",
    "Управляющий",
    "Хостес",
    "Шеф-повар",
]

TASK_CONFIG = {
    "hr": {
        "flag_field": HR_FLAG_FIELD,
        "responsible_env": "BITRIX_HR_TASK_RESPONSIBLE_ID",
        "title": "Отдел кадров: Подготовить и подписать документы с работником",
        "checklist": [
            "Приказ о приёме",
            "Трудовой договор",
            "Должностная инструкция",
            "Договор о материальной ответственности",
            "Обязательно о неразглашении коммерческой тайны",
            "Лист ознакомления с ЛНА",
            "Согласие на обработку персональных данных",
            "Согласие на распространение персональных данных",
            "Заявление на перечисление заработной платы",
            "СЭДО",
        ],
    },
    "occup_safety": {
        "flag_field": OCCUP_SAFETY_FLAG_FIELD,
        "responsible_env": "BITRIX_OCCUP_SAFETY_TASK_RESPONSIBLE_ID",
        "title": "Служба охраны труда: Подготовить и подписать документы с работником",
        "checklist": [
            "Провести вводный и первичный инструктаж по ОТ",
            "Получить подпись в журнале",
            "Провести вводный и первичный инструктаж по ПБ",
            "Получить подпись в журнале",
            "Провести вводный и первичный инструктаж по ЭБ - 1 группа",
            "Получить подпись в журнале",
            "Провести вводный и первичный инструктаж по ГО ЧС",
            "Получить подпись в журнале",
            "Провести вводный и первичный инструктаж по АТ",
            "Получить подпись в журнале",
            "Знакомство с ЛНА и инструкциями по ОТ",
            "Получить подписи об ознакомлении",
        ],
    },
    "military": {
        "flag_field": MILITARY_FLAG_FIELD,
        "responsible_env": "BITRIX_MILITARY_TASK_RESPONSIBLE_ID",
        "title": "Отдел воинского учета: Подготовить и подписать документы с работником",
        "checklist": [
            "Подготовить Форму 10",
            "Подготовить листок сообщения",
            "Подготовить направление в военкомат",
            "Подготовить сведения о приёме на работу",
            "Выдать сотруднику подготовленные документы",
            "Подписать подготовленные документы с работником",
            "Получить в журнале подпись работника в получении документов",
            "Провести разъяснительную работу",
        ],
    },
}


def is_enabled(deal: dict, field_name: str) -> bool:
    return str(deal.get(field_name, "")).upper() == "Y"


def get_fio(deal: dict) -> str:
    fio = (
        f"{deal.get('ufCrm3_1768399087', '').strip()} "
        f"{deal.get('ufCrm3_1768400217', '').strip()} "
        f"{deal.get('ufCrm3_1768399180', '').strip()}"
    )
    return " ".join(fio.split())


def get_1c_password(base_1c: str) -> str:
    if "zup_perelman" in str(base_1c).lower():
        return ROBOT_PSWD_PERELMAN
    if ROBOT_PSWD:
        return ROBOT_PSWD
    return ROBOT_PSWD_DEFAULT


def get_creator_id(deal: dict) -> int:
    return int(os.getenv("BITRIX_TASK_CREATOR_ID") or deal.get("assignedById") or deal.get("createdBy") or 69)


def get_responsible_id(deal: dict, env_name: str) -> int:
    return int(os.getenv(env_name) or deal.get("assignedById") or deal.get("createdBy") or 69)


def make_task_description(deal: dict, fio: str, department_name: str, files_count: int) -> str:
    return (
        f"Смарт-процесс: {deal.get('title', '')}\n"
        f"Сотрудник: {fio}\n"
        f"Направление: {department_name}\n"
        f"ID смарт-процесса: {deal.get('id')}\n"
        f"Количество вложений: {files_count}"
    )


def append_generated_file(target_list: list[str], generated_result: str, department_name: str, document_name: str) -> None:
    if not generated_result:
        print(f"{department_name}: не сформирован документ '{document_name}'")
        return

    if os.path.isfile(generated_result):
        target_list.append(generated_result)
        return

    print(f"{department_name}: ошибка при формировании документа '{document_name}': {generated_result}")


def sanitize_value(value):
    if value is None:
        return ""
    if isinstance(value, dict):
        return {key: sanitize_value(inner_value) for key, inner_value in value.items()}
    if isinstance(value, list):
        return [sanitize_value(item) for item in value]
    return value


def attach_files_to_task(task_id: int, file_paths: list[str]) -> None:
    for file_path in file_paths:
        upload_status = upload_and_attach_file_to_task(
            task_id=task_id,
            file_path=file_path,
            folder_id=DEFAULT_UPLOAD_FOLDER_ID,
        )
        if isinstance(upload_status, str):
            print(upload_status)


def add_checklist_to_task(task_id: int, checklist: list[str]) -> None:
    for index, item in enumerate(checklist, start=1):
        checklist_status = add_checklist_item(task_id=task_id, title=item, sort_index=index * 10)
        if isinstance(checklist_status, str):
            print(f"Не удалось добавить чек-лист в задачу {task_id}: {checklist_status}")
            break


def create_department_task(deal: dict, department_key: str, fio: str, file_paths: list[str]) -> None:
    config = TASK_CONFIG[department_key]
    creator_id = get_creator_id(deal)
    responsible_id = get_responsible_id(deal, config["responsible_env"])
    description = make_task_description(
        deal=deal,
        fio=fio,
        department_name=config["title"],
        files_count=len(file_paths),
    )

    created_task = add_task(
        creator_id=creator_id,
        responsible_id=responsible_id,
        group_id=141,
        title=config["title"],
        description=description,
        company_name=deal.get("title", ""),
    )
    if isinstance(created_task, str):
        print(created_task)
        return

    task_data = created_task
    if "result" in task_data and isinstance(task_data["result"], dict):
        task_data = task_data["result"]
    if "task" in task_data and isinstance(task_data["task"], dict):
        task_data = task_data["task"]
    if "item" in task_data and isinstance(task_data["item"], dict):
        task_data = task_data["item"]

    task_id_value = task_data.get("id") or task_data.get("ID")
    if not task_id_value:
        print(f"Не удалось определить ID созданной задачи: {created_task}")
        return

    task_id = int(task_id_value)
    add_checklist_to_task(task_id=task_id, checklist=config["checklist"])
    attach_files_to_task(task_id=task_id, file_paths=file_paths)


def generate_hr_documents(parsed_data: dict, fio: str, base_1c: str) -> list[str]:
    hr_files: list[str] = []
    department_name = "Отдел кадров"

    append_generated_file(
        hr_files,
        mat_resp_sample_former(
            input_sample="mat_response/input_sample/Договор о полной индивидуальной материальной ответственности.docx",
            output_result="mat_response/output_result/output_Договор о полной индивидуальной материальной ответственности.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Договор о материальной ответственности",
    )
    append_generated_file(
        hr_files,
        nda_sample_former(
            input_sample="nda/input_sample/Обязательство о неразглашении коммерческой тайны.docx",
            output_result="nda/output_result/output_Обязательство о неразглашении коммерческой тайны.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Обязательство о неразглашении коммерческой тайны",
    )
    append_generated_file(
        hr_files,
        lna_sample_former(
            input_sample="lna/input_sample/Лист ознакомления с ЛНА.docx",
            output_result="lna/output_result/output_Лист ознакомления с ЛНА.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Лист ознакомления с ЛНА",
    )
    append_generated_file(
        hr_files,
        personal_secret_sample_former(
            input_sample="personal_secret/input_sample/Согласие на распространение персональных данных НОВОЕ.docx",
            output_result="personal_secret/output_result/output_Согласие на распространение персональных данных НОВОЕ.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Согласие на распространение персональных данных",
    )
    append_generated_file(
        hr_files,
        personal_secret_sample_former(
            input_sample="personal_secret/input_sample/Согласие на обработку персональных данных НОВОЕ.docx",
            output_result="personal_secret/output_result/output_Согласие на обработку персональных данных НОВОЕ.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Согласие на обработку персональных данных",
    )
    append_generated_file(
        hr_files,
        tkrf_sample_former(
            input_sample="tk_rf/input_sample/Трудовой договор.docx",
            output_result="tk_rf/output_result/output_Трудовой договор.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Трудовой договор",
    )
    append_generated_file(
        hr_files,
        zp_align_sample_former(
            input_sample="zp_alignment/input_sample/Заявление на перечисление заработной платы.docx",
            output_result="zp_alignment/output_result/output_Заявление на перечисление заработной платы.docx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Заявление на перечисление заработной платы",
    )
    append_generated_file(
        hr_files,
        formul_tkrf_sample_former(
            input_sample="formul_tkrf/input_sample/Формулировки для трудового договора.xlsx",
            output_result="formul_tkrf/output_result/output_Формулировки для трудового договора.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Формулировки для трудового договора",
    )

    onec_docs = form_order_and_sedo_docs_from_1c(
        fio=fio,
        base_1c=base_1c,
        username=ROBOT_LOGIN,
        pswd=get_1c_password(base_1c),
    )
    append_generated_file(
        hr_files,
        onec_docs["order"],
        department_name,
        "Приказ о приёме",
    )
    append_generated_file(
        hr_files,
        onec_docs["sedo"],
        department_name,
        "СЭДО",
    )

    job_title = parsed_data["job_info"]["Должность"]
    if job_title in EMPLOYEE_INSTRUCTIONS:
        append_generated_file(
            hr_files,
            employee_instr_sample_former(
                input_sample=f"employee_instructions/input_samples/{job_title}.docx",
                output_result=f"employee_instructions/output_result/output_{job_title}.docx",
                parsed_data_1c=parsed_data,
            ),
            department_name,
            "Должностная инструкция",
        )

    return hr_files


def generate_occup_safety_documents(parsed_data: dict) -> list[str]:
    occup_safety_files: list[str] = []
    append_generated_file(
        occup_safety_files,
        occup_safety_former(
            input_sample="occup_safety/input_sample/ЭДО Стандартные.docx",
            output_result="occup_safety/output_result/output_ЭДО Стандартные.docx",
            parsed_data_1c=parsed_data,
        ),
        "Служба охраны труда",
        "ЭДО Стандартные",
    )
    return occup_safety_files


def generate_military_documents(parsed_data: dict) -> list[str]:
    military_files: list[str] = []
    department_name = "Отдел воинского учета"

    append_generated_file(
        military_files,
        list_msg_sample_former(
            input_sample="military_registr/input_samples/Листок-сообщения.xlsx",
            output_result="military_registr/output_result/output_Листок-сообщения.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Листок сообщения",
    )
    append_generated_file(
        military_files,
        msg_13_sample_former(
            input_sample="military_registr/input_samples/Письмо 13.xlsx",
            output_result="military_registr/output_result/output_Письмо 13.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Письмо 13",
    )
    append_generated_file(
        military_files,
        application_sample_former(
            input_sample="military_registr/input_samples/Приложение 10.xlsx",
            output_result="military_registr/output_result/output_Приложение 10.xlsx",
            parsed_data_1c=parsed_data,
            sheet_name="Сведения прил. 10",
        ),
        department_name,
        "Приложение 10",
    )
    append_generated_file(
        military_files,
        application_sample_former(
            input_sample="military_registr/input_samples/Приложение 23.xlsx",
            output_result="military_registr/output_result/output_Приложение 23.xlsx",
            parsed_data_1c=parsed_data,
            sheet_name="Сверка",
        ),
        department_name,
        "Приложение 23",
    )
    append_generated_file(
        military_files,
        info_changes_sample_former(
            input_sample="military_registr/input_samples/Сведения об изменениях.xlsx",
            output_result="military_registr/output_result/output_Сведения об изменениях.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Сведения об изменениях",
    )
    append_generated_file(
        military_files,
        details_sample_former(
            input_sample="military_registr/input_samples/Сведения.xlsx",
            output_result="military_registr/output_result/output_Сведения.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Сведения о приеме на работу",
    )
    append_generated_file(
        military_files,
        notification_sample_former(
            input_sample="military_registr/input_samples/Уведомление 1.xlsx",
            output_result="military_registr/output_result/output_Уведомление 1.xlsx",
            parsed_data_1c=parsed_data,
            sheet_name="Уведомление 1",
        ),
        department_name,
        "Уведомление 1",
    )
    append_generated_file(
        military_files,
        notification_sample_former(
            input_sample="military_registr/input_samples/Уведомление 2.xlsx",
            output_result="military_registr/output_result/output_Уведомление 2.xlsx",
            parsed_data_1c=parsed_data,
            sheet_name="Уведомление 2",
        ),
        department_name,
        "Уведомление 2",
    )
    append_generated_file(
        military_files,
        f10_sample_former(
            input_sample="military_registr/input_samples/Ф10.xlsx",
            output_result="military_registr/output_result/output_Ф10.xlsx",
            parsed_data_1c=parsed_data,
        ),
        department_name,
        "Форма 10",
    )

    return military_files


def process_deal(deal: dict) -> None:
    fio = get_fio(deal)
    base_1c = deal.get(BASE_1C_FIELD)
    onec_password = get_1c_password(base_1c)

    if not fio or not base_1c:
        print(f"Сделка {deal.get('id')}: пропущена из-за отсутствия ФИО или строки подключения к 1С")
        return

    parsed_data = app_1c(fio=fio, base_1c=base_1c, username=ROBOT_LOGIN, pswd=onec_password)
    if isinstance(parsed_data, str):
        print(f"Сделка {deal.get('id')}: {parsed_data}")
        return

    parsed_data = sanitize_value(parsed_data)
    department_files: dict[str, list[str]] = {}

    if is_enabled(deal, HR_FLAG_FIELD):
        department_files["hr"] = generate_hr_documents(parsed_data, fio=fio, base_1c=base_1c)

    if is_enabled(deal, OCCUP_SAFETY_FLAG_FIELD):
        department_files["occup_safety"] = generate_occup_safety_documents(parsed_data)

    if is_enabled(deal, MILITARY_FLAG_FIELD):
        department_files["military"] = generate_military_documents(parsed_data)

    for department_key, file_paths in department_files.items():
        if file_paths:
            create_department_task(deal=deal, department_key=department_key, fio=fio, file_paths=file_paths)


def app() -> None:
    bitrix_deals = get_deals_bitrix(token=BITRIX_TOKEN, entityTypeId=1042, method="crm.item.list")

    for deal in bitrix_deals:
        process_deal(deal)


if __name__ == "__main__":
    app()
