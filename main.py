import os
import re
import shutil
import socket
import sys
import traceback
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv

from bitrix.bitrix import (
    add_checklist_item,
    add_task,
    build_task_crm_binding,
    build_task_deal_binding,
    company_info,
    download_file_from_url,
    ensure_bitrix_disk_subfolder,
    get_department_info,
    get_department_users,
    get_all_smart_process_items,
    get_company_userfield_code_by_title,
    get_deal_info,
    get_deals_bitrix,
    get_disk_attached_object_info,
    get_disk_file_info,
    get_inn_bitrix,
    get_smart_process_fields,
    get_smart_process_item,
    make_error_msg,
    resolve_bitrix_disk_folder_id,
    upload_and_attach_file_to_task,
    upload_file,
    update_smart_process_stage,
    update_smart_process_fields,
    update_task_crm_bindings,
)
from bitrix.template_sync import sync_templates_from_bitrix
from employee_instructions.form_empl_instr import employee_instr_sample_former
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
from parser_1с.bank_details import enrich_bank_details_from_1c
from parser_1с.print_forms import form_order_and_sedo_docs_from_1c
from parser_1с.personal_info import app_1c
from personal_secret.form_personal_secret import personal_secret_sample_former
from tk_rf.form_tkrf import tkrf_sample_former
from template_fallback import apply_common_placeholder_fallback
from zp_alignment.form_zp_alignment import zp_align_sample_former


load_dotenv()

BITRIX_TOKEN = os.getenv("BITRIX_TOKEN")
BITRIX_ERROR_CHAT_ID = os.getenv("BITRIX_ERROR_CHAT_ID", "chat12963")
ROBOT_LOGIN = os.getenv("ONE_C_DEFAULT_LOGIN")
ROBOT_PSWD = os.getenv("ONE_C_DEFAULT_PASSWORD")
ROBOT_PSWD_DEFAULT = os.getenv("ONE_C_DEFAULT_PASSWORD")
ROBOT_PSWD_PERELMAN = os.getenv("ONE_C_ZUP_PERELMAN_PASSWORD")
DEFAULT_UPLOAD_FOLDER_ID = int(os.getenv("BITRIX_TASK_UPLOAD_FOLDER_ID", "26188"))
SOURCE_SMART_PROCESS_ENTITY_TYPE_ID = int(os.getenv("BITRIX_SOURCE_SMART_PROCESS_ENTITY_TYPE_ID", "1042"))
SOURCE_SMART_PROCESS_STAGE_ID = os.getenv("BITRIX_SOURCE_SMART_PROCESS_STAGE_ID", "DT1042_5:UC_B608AW").strip()
SOURCE_SMART_PROCESS_NEXT_STAGE_ID = os.getenv("BITRIX_SOURCE_SMART_PROCESS_NEXT_STAGE_ID", "DT1042_5:UC_RUVYAU")
HR_COMPANY_RESPONSIBLE_FIELD = "ASSIGNED_BY_ID"
MILITARY_COMPANY_RESPONSIBLE_FIELD = "UF_CRM_1777389420"
OCCUP_SAFETY_COMPANY_RESPONSIBLE_FIELD = "UF_CRM_1777389391"
OCCUP_SAFETY_COMPANY_ACCOMPLICES_FIELD = "UF_CRM_1784720933"
MILITARY_COMPANY_ACCOMPLICES_FIELD = "UF_CRM_1784720890"
DEPARTMENT_STRUCTURE_IDS = {
    "hr": 4,
    "military": 18,
    "occup_safety": 20,
}
FIXED_AUDITOR_IDS = [5, 44]

HR_FLAG_FIELD = "ufCrm3_1770106591"
OCCUP_SAFETY_FLAG_FIELD = "ufCrm3_1770106632"
MILITARY_FLAG_FIELD = "ufCrm3_1770106622"
BASE_1C_FIELD = "ufCrm3_1770808935"
ARCHIVE_ENABLED_FIELD = "UF_CRM_3_1782457012"
ARCHIVE_DONE_FIELD = "UF_CRM_3_1782802037"
ARCHIVE_EXCLUDED_STAGE_ID = "DT1042_5:SUCCESS"
ARCHIVE_SOURCE_ITEM_FIELD = "UF_CRM_3_1782803163"
ARCHIVE_COMPANY_FOLDER_FIELD = "UF_CRM_1782128121750"
HR_BITRIX_COMPANY_FOLDER_FIELD = "UF_CRM_1782128963490"
ARCHIVE_SOURCE_ENTITY_TYPE_ID = 1058
ARCHIVE_SOURCE_FILE_FIELDS = [
    ("UF_CRM_7_1781789098", "Паспорт (главная страница, страница с пропиской, 13-я стр, печать о браке)"),
    ("UF_CRM_7_1781789108", "СНИЛС"),
    ("UF_CRM_7_1781789116", "ИНН"),
    ("UF_CRM_7_1782039237327", "Документ об образовании"),
    ("UF_CRM_7_1782039319076", "Военный билет (для военнообязанных) скан всех страниц"),
    ("UF_CRM_7_1782039355442", "Трудовая книжка (если есть бумажная) скан всех страниц"),
    ("UF_CRM_7_1782039370244", "Выписка из электронной трудовой книжки"),
    ("UF_CRM_7_1782039386160", "Банковские реквизиты"),
    ("UF_CRM_7_1782464853", "Паспорт (разворот с фото)"),
    ("UF_CRM_7_1782464894", "Нотариально заверенный перевод паспорта (все страницы)"),
    ("UF_CRM_7_1782464952", "Карточка дактилоскопии мигранта с 2-х сторон"),
    ("UF_CRM_7_1782464965", "Реквизиты банковской карточки"),
    ("UF_CRM_7_1782464983", "Полис ДМС"),
    ("UF_CRM_7_1782465011", "Медицинское заключение (скан всех листов с 2-х сторон)"),
    ("UF_CRM_7_1782465376", "Скриншот информации о месте пребывания в приложении АМИНА"),
    ("UF_CRM_7_1782465416", "Миграционная карта с 2-х сторон"),
    ("UF_CRM_7_1782465635", "Паспорт скан всех страниц"),
    ("UF_CRM_7_1782465661", "Регистрация по городу Москве"),
    ("UF_CRM_7_1782466026695", "Патент скан с 2-х сторон"),
    ("UF_CRM_7_1782466072428", "Уведомление о подтверждении трудовой деятельности с отметкой органа МВД"),
    ("UF_CRM_7_1782466130646", "Все чеки за оплату патента обязательно"),
    ("UF_CRM_7_1782466409727", "Вид на жительство с регистрацией все страницы"),
    ("UF_CRM_7_1782466832713", "РВП, выданное в регионе работы с регистрацией все страницы"),
    ("UF_CRM_7_1782467049563", "РВПО, выданное в регионе работы с актуальным сроком действия с регистрацией все страницы"),
    ("UF_CRM_7_1782467070012", "Справка из ВУЗа"),
]


def notify_bitrix_error(message: str) -> None:
    error_text = str(message or "").strip()
    if not error_text:
        return

    print(error_text)

    if not BITRIX_TOKEN:
        print("Не удалось отправить ошибку в чат Bitrix: отсутствует BITRIX_TOKEN")
        return

    try:
        send_result = make_error_msg(BITRIX_TOKEN, BITRIX_ERROR_CHAT_ID, error_text[:15000])
        if isinstance(send_result, str) and send_result.startswith("Ошибка"):
            print(f"Не удалось отправить ошибку в чат Bitrix: {send_result}")
    except Exception as notification_error:
        print(f"Не удалось отправить ошибку в чат Bitrix: {notification_error}")


def report_exception_to_bitrix(error: BaseException, *, context: str, deal: dict | None = None) -> None:
    deal_id = ""
    deal_title = ""
    if isinstance(deal, dict):
        deal_id = str(deal.get("id") or "")
        deal_title = str(deal.get("title") or "")

    traceback_text = traceback.format_exc()
    if not traceback_text or traceback_text == "NoneType: None\n":
        traceback_text = "".join(traceback.format_exception(type(error), error, error.__traceback__))

    host_name = socket.gethostname()
    context_lines = [
        f"Ошибка в AprioriNewEmployee",
        f"Контекст: {context}",
        f"Хост: {host_name}",
        f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
    ]
    if deal_id:
        context_lines.append(f"Сделка ID: {deal_id}")
    if deal_title:
        context_lines.append(f"Сделка: {deal_title}")

    context_lines.extend(
        [
            f"Тип ошибки: {type(error).__name__}",
            f"Сообщение: {error}",
            "Traceback:",
            traceback_text.strip(),
        ]
    )
    notify_bitrix_error("\n".join(context_lines))


def _global_exception_hook(exc_type, exc_value, exc_traceback) -> None:
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    traceback_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    notify_bitrix_error(
        "\n".join(
            [
                "Необработанная ошибка в AprioriNewEmployee",
                f"Хост: {socket.gethostname()}",
                f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
                f"Тип ошибки: {exc_type.__name__}",
                f"Сообщение: {exc_value}",
                "Traceback:",
                traceback_text.strip(),
            ]
        )
    )
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


sys.excepthook = _global_exception_hook

def get_employee_instruction_template_name(parsed_data: dict) -> str:
    job_title = str(parsed_data["job_info"].get("Должность", "") or "").strip()
    if job_title != "Повар":
        return job_title

    department_name = str(parsed_data["job_info"].get("Подразделение", "") or "").strip().lower()
    if "холод" in department_name:
        return "Повар ХЦ"
    if "заготов" in department_name:
        return "Повар ЗЦ"
    if "горяч" in department_name:
        return "Повар ГЦ"
    return "Повар"


def get_employee_instruction_template_path(parsed_data: dict) -> Path:
    job_title = get_employee_instruction_template_name(parsed_data)
    return Path("employee_instructions/input_samples") / f"{job_title}.docx"

HR_TEMPLATE_DIRECTORIES = {
    "mat_response": ("mat_response/input_sample", "mat_response/output_result", mat_resp_sample_former),
    "nda": ("nda/input_sample", "nda/output_result", nda_sample_former),
    "lna": ("lna/input_sample", "lna/output_result", lna_sample_former),
    "personal_secret": ("personal_secret/input_sample", "personal_secret/output_result", personal_secret_sample_former),
    "tk_rf": ("tk_rf/input_sample", "tk_rf/output_result", tkrf_sample_former),
    "zp_alignment": ("zp_alignment/input_sample", "zp_alignment/output_result", zp_align_sample_former),
}

MILITARY_TEMPLATE_HANDLERS = {
    "Листок-сообщения": lambda input_sample, output_result, parsed_data: list_msg_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
    ),
    "Письмо 13": lambda input_sample, output_result, parsed_data: msg_13_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
    ),
    "Приложение 10": lambda input_sample, output_result, parsed_data: application_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
        sheet_name="Сведения прил. 10",
    ),
    "Приложение 23": lambda input_sample, output_result, parsed_data: application_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
        sheet_name="Сверка",
    ),
    "Сведения об изменениях": lambda input_sample, output_result, parsed_data: info_changes_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
    ),
    "Сведения": lambda input_sample, output_result, parsed_data: details_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
    ),
    "Уведомление 1": lambda input_sample, output_result, parsed_data: notification_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
        sheet_name="Уведомление 1",
    ),
    "Уведомление 2": lambda input_sample, output_result, parsed_data: notification_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
        sheet_name="Уведомление 2",
    ),
    "Ф10": lambda input_sample, output_result, parsed_data: f10_sample_former(
        input_sample=input_sample,
        output_result=output_result,
        parsed_data_1c=parsed_data,
    ),
}


def iter_template_files(input_dir: Path, allowed_suffixes: tuple[str, ...] | None = None):
    if not input_dir.exists():
        return

    for template_path in sorted(input_dir.rglob("*")):
        if not template_path.is_file() or template_path.name.startswith("~$"):
            continue
        if allowed_suffixes and template_path.suffix.lower() not in allowed_suffixes:
            continue
        yield template_path


def build_output_path(output_dir: Path, template_path: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"output_{template_path.name}"


def form_generic_excel_template(input_sample: str, output_result: str, parsed_data: dict) -> str:
    """Copy a new Excel template so the common fallback can fill every worksheet."""
    try:
        output_path = Path(output_result)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_sample, output_path)
        return str(output_path)
    except Exception as error:
        return f"Ошибка во время формирования файла: {error}"


def process_template_directory(
    target_list: list[str],
    input_dir: Path,
    output_dir: Path,
    former,
    parsed_data: dict,
    department_name: str,
    allowed_suffixes: tuple[str, ...] | None = None,
) -> None:
    for template_path in iter_template_files(input_dir, allowed_suffixes=allowed_suffixes):
        output_path = build_output_path(output_dir, template_path)
        append_generated_file(
            target_list,
            finalize_generated_result(
                former(
                    input_sample=str(template_path),
                    output_result=str(output_path),
                    parsed_data_1c=parsed_data,
                ),
                parsed_data,
            ),
            department_name,
            template_path.stem,
        )


TASK_CONFIG = {
    "hr": {
        "flag_field": HR_FLAG_FIELD,
        "responsible_env": "BITRIX_HR_TASK_RESPONSIBLE_ID",
        "company_responsible_field": HR_COMPANY_RESPONSIBLE_FIELD,
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
        "company_responsible_field": OCCUP_SAFETY_COMPANY_RESPONSIBLE_FIELD,
        "fixed_responsible_id": 68,
        "company_accomplices_field": OCCUP_SAFETY_COMPANY_ACCOMPLICES_FIELD,
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
        "company_responsible_field": MILITARY_COMPANY_RESPONSIBLE_FIELD,
        "fixed_responsible_id": 68,
        "company_accomplices_field": MILITARY_COMPANY_ACCOMPLICES_FIELD,
        "company_auditors_mode": "same_field",
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


def _extract_user_ids(value) -> list[int]:
    if isinstance(value, str):
        normalized = value.strip()
        if not normalized:
            return []
        if any(separator in normalized for separator in (",", ";")):
            result: list[int] = []
            for part in re.split(r"[;,]", normalized):
                user_id = _extract_user_id(part)
                if user_id and user_id not in result:
                    result.append(user_id)
            return result

    if isinstance(value, list):
        result: list[int] = []
        for item in value:
            user_id = _extract_user_id(item)
            if user_id and user_id not in result:
                result.append(user_id)
        return result

    user_id = _extract_user_id(value)
    return [user_id] if user_id else []


def _extract_user_id(value) -> int | None:
    if isinstance(value, list):
        value = value[0] if value else None
    if isinstance(value, dict):
        value = value.get("ID") or value.get("id") or value.get("VALUE")
    if value in (None, ""):
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def get_parent_deal_id(deal: dict) -> int | None:
    parent_deal_id = deal.get("parentId2") or deal.get("dealId") or deal.get("parentDealId")
    if not parent_deal_id:
        return None
    try:
        return int(parent_deal_id)
    except (TypeError, ValueError):
        return None


def get_company_id_from_deal_context(deal: dict) -> int | None:
    company_candidates = (
        deal.get("companyId"),
        deal.get("COMPANY_ID"),
        deal.get("parentId4"),
        deal.get("parentCompanyId"),
    )
    for company_id in company_candidates:
        if not company_id:
            continue
        try:
            return int(company_id)
        except (TypeError, ValueError):
            continue

    parent_deal_id = get_parent_deal_id(deal)
    if not parent_deal_id:
        return None

    deal_info = get_deal_info(parent_deal_id)
    if not isinstance(deal_info, dict):
        return None

    company_id = deal_info.get("COMPANY_ID")
    if not company_id:
        return None

    try:
        return int(company_id)
    except (TypeError, ValueError):
        return None


def get_company_name_from_deal_context(deal: dict) -> str:
    company_id = get_company_id_from_deal_context(deal)
    if company_id:
        company = company_info(company_id)
        if isinstance(company, dict):
            company_name = str(
                company.get("TITLE")
                or company.get("title")
                or company.get("COMPANY_TITLE")
                or ""
            ).strip()
            if company_name:
                return company_name

    return str(deal.get("title") or "").strip()


def get_company_inn_from_deal_context(deal: dict) -> str:
    company_id = get_company_id_from_deal_context(deal)
    if not company_id:
        return ""

    company = company_info(company_id)
    if isinstance(company, dict):
        for field_name in ("INN", "RQ_INN", "UF_CRM_INN"):
            company_inn = str(company.get(field_name) or "").strip()
            if company_inn:
                return company_inn

    try:
        return str(get_inn_bitrix(company_id) or "").strip()
    except Exception:
        return ""


def get_company_task_responsible_id(deal: dict, department_key: str) -> int | None:
    user_ids = get_company_task_responsible_ids(deal, department_key)
    return user_ids[0] if user_ids else None


def get_company_task_responsible_ids(deal: dict, department_key: str) -> list[int]:
    company_id = get_company_id_from_deal_context(deal)
    if not company_id:
        return []

    company = company_info(company_id)
    if not isinstance(company, dict):
        return []

    config = TASK_CONFIG[department_key]
    field_code = config.get("company_responsible_field")
    if not field_code and config.get("company_responsible_title"):
        field_code = get_company_userfield_code_by_title(config["company_responsible_title"])

    if not field_code:
        return []

    return _extract_user_ids(company.get(field_code))


def get_company_task_user_ids_by_field(deal: dict, field_code: str | None) -> list[int]:
    if not field_code:
        return []

    company_id = get_company_id_from_deal_context(deal)
    if not company_id:
        return []

    company = company_info(company_id)
    if not isinstance(company, dict):
        return []

    return _extract_user_ids(company.get(field_code))


def prioritize_hr_files(file_paths: list[str]) -> list[str]:
    def sort_key(file_path: str) -> tuple[int, str]:
        file_name = Path(file_path).name.lower()
        if "сэдо" in file_name:
            return (2, file_name)
        if "приказ" in file_name:
            return (1, file_name)
        return (0, file_name)

    return sorted(file_paths, key=sort_key)


def get_company_task_responsible_debug(deal: dict, department_key: str) -> dict:
    debug_info = {
        "department_key": department_key,
        "company_id": None,
        "field_code": None,
        "field_value": None,
        "responsible_ids": [],
        "responsible_id": None,
    }

    company_id = get_company_id_from_deal_context(deal)
    debug_info["company_id"] = company_id
    if not company_id:
        return debug_info

    company = company_info(company_id)
    if not isinstance(company, dict):
        return debug_info

    config = TASK_CONFIG[department_key]
    field_code = config.get("company_responsible_field")
    if not field_code and config.get("company_responsible_title"):
        field_code = get_company_userfield_code_by_title(config["company_responsible_title"])

    debug_info["field_code"] = field_code
    if not field_code:
        return debug_info

    field_value = company.get(field_code)
    responsible_ids = _extract_user_ids(field_value)
    responsible_id = responsible_ids[0] if responsible_ids else None
    debug_info["field_value"] = field_value
    debug_info["responsible_ids"] = responsible_ids
    debug_info["responsible_id"] = responsible_id
    return debug_info


def get_company_task_auditors(deal: dict, department_key: str, responsible_id: int | None = None) -> list[int]:
    config = TASK_CONFIG[department_key]
    if config.get("company_auditors_mode") != "same_field":
        return get_department_auditors(department_key)

    responsible_ids = get_company_task_responsible_ids(deal, department_key)
    auditors: list[int] = []
    for user_id in responsible_ids:
        if not user_id or user_id == responsible_id or user_id in auditors:
            continue
        auditors.append(user_id)
    return auditors


def get_department_auditors(department_key: str) -> list[int]:
    department_id = DEPARTMENT_STRUCTURE_IDS.get(department_key)
    auditors: list[int] = []
    if not department_id:
        return FIXED_AUDITOR_IDS.copy()

    department_info = get_department_info(department_id)
    if isinstance(department_info, dict):
        head_id = _extract_user_id(
            department_info.get("UF_HEAD")
            or department_info.get("HEAD")
            or department_info.get("MANAGER")
        )
        if head_id and head_id not in auditors:
            auditors.append(head_id)

    users = get_department_users(department_id)
    if isinstance(users, list):
        deputy_id = None
        for user in users:
            position = str(user.get("WORK_POSITION") or user.get("UF_WORK_POSITION") or "").strip().lower()
            if "замест" in position:
                deputy_id = _extract_user_id(user.get("ID"))
                break
        if deputy_id and deputy_id not in auditors:
            auditors.append(deputy_id)

    for auditor_id in FIXED_AUDITOR_IDS:
        if auditor_id not in auditors:
            auditors.append(auditor_id)

    return auditors


def is_enabled(deal: dict, field_name: str) -> bool:
    return str(deal.get(field_name, "")).upper() == "Y"


def is_truthy_value(value) -> bool:
    return str(value or "").strip().lower() in {"y", "yes", "true", "1", "да"}


def is_falsy_value(value) -> bool:
    return str(value or "").strip().lower() in {"n", "no", "false", "0", "нет", ""}


def normalize_bitrix_field_code(field_code: str) -> str:
    return "".join(ch for ch in str(field_code or "").lower() if ch.isalnum())


def get_bitrix_field_value(entity: dict, field_code: str):
    normalized_target = normalize_bitrix_field_code(field_code)
    for key, value in entity.items():
        if normalize_bitrix_field_code(key) == normalized_target:
            return value
    return None


def smart_process_field_code_to_api(field_code: str) -> str:
    parts = str(field_code or "").split("_")
    if len(parts) < 4:
        return field_code
    return f"ufCrm{parts[2]}_{parts[3]}"


def build_archive_folder_name(fio: str) -> str:
    return f"{' '.join((fio or '').split())} {datetime.now().strftime('%d.%m')}".strip()


def build_personal_docs_folder_name(fio: str, citizenship_marker: str) -> str:
    normalized_fio = " ".join((fio or "").split()).strip()
    normalized_marker = str(citizenship_marker or "").strip()
    date_part = datetime.now().strftime("%d.%m.%Y")
    return f"{normalized_fio} {date_part} {normalized_marker}".strip()


def sanitize_filename_component(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\\\|?*]', "_", str(value or "").strip())
    sanitized = re.sub(r"\s+", " ", sanitized).strip(" .")
    return sanitized or "file"


def extract_numeric_bitrix_ids(value) -> list[int]:
    found_ids: list[int] = []

    def _collect(inner_value) -> None:
        if inner_value in (None, "", False):
            return
        if isinstance(inner_value, list):
            for item in inner_value:
                _collect(item)
            return
        if isinstance(inner_value, dict):
            for key in ("id", "ID", "value", "VALUE", "fileId", "FILE_ID"):
                if key in inner_value:
                    _collect(inner_value[key])
            return

        raw_value = str(inner_value).strip()
        if not raw_value:
            return
        if raw_value.isdigit():
            found_ids.append(int(raw_value))
            return

        for match in re.findall(r"\bn(\d+)\b", raw_value, flags=re.IGNORECASE):
            found_ids.append(int(match))
        for match in re.findall(r"\b(\d+)\b", raw_value):
            found_ids.append(int(match))

    _collect(value)
    return found_ids


def collect_bitrix_file_candidates(value) -> list[dict]:
    candidates: list[dict] = []

    def _collect(inner_value) -> None:
        if inner_value in (None, "", False):
            return

        if isinstance(inner_value, list):
            for item in inner_value:
                _collect(item)
            return

        if isinstance(inner_value, dict):
            download_url = (
                inner_value.get("urlMachine")
                or inner_value.get("URL_MACHINE")
                or
                inner_value.get("DOWNLOAD_URL")
                or inner_value.get("downloadUrl")
                or inner_value.get("url")
            )
            file_name = (
                inner_value.get("NAME")
                or inner_value.get("name")
                or inner_value.get("NAME_WITH_EXTENSION")
                or inner_value.get("fileName")
            )
            file_id = inner_value.get("id") or inner_value.get("ID") or inner_value.get("fileId") or inner_value.get("FILE_ID")

            if download_url:
                candidates.append(
                    {
                        "id": int(file_id) if str(file_id or "").isdigit() else None,
                        "download_url": str(download_url),
                        "name": str(file_name or ""),
                    }
                )

            for key in ("id", "ID", "value", "VALUE", "fileId", "FILE_ID", "fileData", "FILE_DATA"):
                if key in inner_value:
                    _collect(inner_value[key])
            return

        raw_value = str(inner_value).strip()
        if not raw_value:
            return
        if raw_value == "0":
            return
        if raw_value.isdigit():
            candidates.append({"id": int(raw_value)})
            return

        for match in re.findall(r"\bn(\d+)\b", raw_value, flags=re.IGNORECASE):
            candidates.append({"id": int(match)})

    _collect(value)

    unique_candidates: list[dict] = []
    seen_pairs: set[tuple] = set()
    for candidate in candidates:
        key = (candidate.get("id"), candidate.get("download_url"), candidate.get("name"))
        if key in seen_pairs:
            continue
        seen_pairs.add(key)
        unique_candidates.append(candidate)

    return unique_candidates


def summarize_bitrix_value(value, max_len: int = 200) -> str:
    text = repr(value)
    if len(text) > max_len:
        return f"{text[:max_len]}..."
    return text


def get_file_field_codes(entity_type_id: int) -> list[str]:
    fields = get_smart_process_fields(entity_type_id)
    if not isinstance(fields, dict):
        return []

    file_field_codes: list[str] = []
    for field_code, metadata in fields.items():
        if not isinstance(metadata, dict):
            continue
        haystack = " ".join(
            str(metadata.get(key, "") or "")
            for key in ("type", "userTypeId", "title", "formLabel")
        ).lower()
        if "file" in haystack or "disk" in haystack or "файл" in haystack:
            file_field_codes.append(field_code)

    return file_field_codes


def is_foreign_citizen_deal(deal: dict) -> bool:
    return is_truthy_value(get_bitrix_field_value(deal, "UF_CRM_3_1768399223"))


def resolve_disk_file_info(file_id: int) -> dict | None:
    file_info = get_disk_file_info(file_id)
    if isinstance(file_info, dict) and file_info:
        return file_info

    attached_info = get_disk_attached_object_info(file_id)
    if not isinstance(attached_info, dict) or not attached_info:
        return None

    if attached_info.get("DOWNLOAD_URL"):
        return attached_info

    object_id = attached_info.get("OBJECT_ID") or attached_info.get("OBJECTId") or attached_info.get("FILE_ID")
    if object_id:
        nested_file_info = get_disk_file_info(object_id)
        if isinstance(nested_file_info, dict) and nested_file_info:
            return nested_file_info
    return None


def mark_archive_export_done(deal_id: int | str) -> bool:
    field_key = smart_process_field_code_to_api(ARCHIVE_DONE_FIELD)
    candidate_values = ["Y", "1", 1, "Да", True]

    for candidate_value in candidate_values:
        update_result = update_smart_process_fields(
            entity_id=deal_id,
            entity_type_id=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
            fields={field_key: candidate_value},
        )
        if isinstance(update_result, str):
            continue

        refreshed_deal = get_smart_process_item(SOURCE_SMART_PROCESS_ENTITY_TYPE_ID, deal_id)
        if isinstance(refreshed_deal, dict) and is_truthy_value(get_bitrix_field_value(refreshed_deal, ARCHIVE_DONE_FIELD)):
            return True

    return False


def upload_hr_files_to_company_bitrix_folder(deal: dict, fio: str, hr_files: list[str]) -> None:
    if not hr_files:
        return

    hr_files = prioritize_hr_files(hr_files)

    company_id = get_company_id_from_deal_context(deal)
    if not company_id:
        print(f"Сделка {deal.get('id')}: не найдена компания для выгрузки кадровых документов в Bitrix")
        return

    company = company_info(company_id)
    if not isinstance(company, dict):
        print(f"Сделка {deal.get('id')}: не удалось получить компанию {company_id} для выгрузки кадровых документов в Bitrix")
        return

    folder_value = company.get(HR_BITRIX_COMPANY_FOLDER_FIELD)
    if not folder_value:
        return

    parent_folder_id = resolve_bitrix_disk_folder_id(folder_value)
    if not parent_folder_id:
        print(
            f"Сделка {deal.get('id')}: не удалось определить folderId из поля "
            f"{HR_BITRIX_COMPANY_FOLDER_FIELD}: {folder_value}"
        )
        return

    subfolder_name = build_archive_folder_name(fio)
    subfolder = ensure_bitrix_disk_subfolder(parent_folder_id, subfolder_name)
    if not isinstance(subfolder, dict):
        print(
            f"Сделка {deal.get('id')}: не удалось подготовить папку Bitrix "
            f"'{subfolder_name}' в {parent_folder_id}: {subfolder}"
        )
        return

    target_folder_id = subfolder.get("ID") or subfolder.get("id")
    if not target_folder_id:
        print(
            f"Сделка {deal.get('id')}: Bitrix не вернул ID созданной папки "
            f"'{subfolder_name}': {subfolder}"
        )
        return

    uploaded_count = 0
    for file_path in hr_files:
        if not os.path.isfile(file_path):
            continue

        upload_result = upload_file(
            file_path=file_path,
            filename=Path(file_path).name,
            folder_id=target_folder_id,
        )
        if upload_result:
            uploaded_count += 1
        else:
            print(
                f"Сделка {deal.get('id')}: не удалось загрузить кадровый документ "
                f"в Bitrix: {file_path}"
            )

    print(
        f"Сделка {deal.get('id')}: загружено кадровых документов в Bitrix: "
        f"{uploaded_count}/{len(hr_files)}"
    )


def download_files_from_source_item(source_item: dict, source_entity_type_id: int, destination_dir: Path) -> list[str]:
    destination_dir.mkdir(parents=True, exist_ok=True)
    downloaded_paths: list[str] = []
    seen_file_ids: set[int] = set()

    file_field_codes = get_file_field_codes(source_entity_type_id)
    candidate_values = [get_bitrix_field_value(source_item, field_code) for field_code in file_field_codes]

    # Fallback: some Bitrix file fields are returned in нестандартном формате или без ожидаемой меты.
    for key, value in source_item.items():
        if value in candidate_values:
            continue
        if isinstance(value, (list, dict)):
            candidate_values.append(value)
            continue
        text_value = str(value or "").strip()
        if re.search(r"\bn\d+\b", text_value, flags=re.IGNORECASE):
            candidate_values.append(value)

    for field_value in candidate_values:
        for file_id in extract_numeric_bitrix_ids(field_value):
            if file_id in seen_file_ids:
                continue
            seen_file_ids.add(file_id)

            file_info = resolve_disk_file_info(file_id)
            if not isinstance(file_info, dict):
                continue

            download_url = file_info.get("DOWNLOAD_URL") or file_info.get("DOWNLOADUrl")
            file_name = file_info.get("NAME") or file_info.get("NAME_WITH_EXTENSION") or f"{file_id}.bin"
            if not download_url or not file_name:
                continue

            destination_path = destination_dir / file_name
            if destination_path.exists():
                destination_path = destination_dir / f"{file_id}_{file_name}"

            error = download_file_from_url(download_url, str(destination_path))
            if isinstance(error, str):
                print(error)
                continue

            downloaded_paths.append(str(destination_path))

    return downloaded_paths


def download_named_files_from_source_item_fields(source_item: dict, destination_dir: Path) -> list[str]:
    destination_dir.mkdir(parents=True, exist_ok=True)
    downloaded_paths: list[str] = []
    seen_candidates: set[tuple] = set()

    for field_code, field_title in ARCHIVE_SOURCE_FILE_FIELDS:
        field_value = get_bitrix_field_value(source_item, field_code)
        file_candidates = collect_bitrix_file_candidates(field_value)
        if not file_candidates:
            if field_value not in (None, "", 0, "0", [], {}):
                print(
                    f"Поле {field_code} не распознано как файловое значение: "
                    f"{summarize_bitrix_value(field_value)}"
                )
            continue

        file_index = 0
        for file_candidate in file_candidates:
            candidate_key = (
                file_candidate.get("id"),
                file_candidate.get("download_url"),
                file_candidate.get("name"),
            )
            if candidate_key in seen_candidates:
                continue
            seen_candidates.add(candidate_key)

            file_id = file_candidate.get("id")
            download_url = file_candidate.get("download_url")
            original_name = str(file_candidate.get("name") or "")

            if not download_url and file_id is not None:
                file_info = resolve_disk_file_info(file_id)
                if not isinstance(file_info, dict):
                    print(f"Не удалось получить информацию о файле {file_id} для поля {field_code}")
                    continue

                download_url = file_info.get("DOWNLOAD_URL") or file_info.get("DOWNLOADUrl")
                original_name = str(
                    file_info.get("NAME")
                    or file_info.get("NAME_WITH_EXTENSION")
                    or original_name
                    or f"{file_id}.bin"
                )

            extension = Path(original_name).suffix or ".bin"
            if not download_url:
                print(
                    f"Для поля {field_code} не найден URL скачивания: "
                    f"{summarize_bitrix_value(file_candidate)}"
                )
                continue

            file_index += 1
            base_name = sanitize_filename_component(field_title)
            final_name = f"{base_name}{extension}" if file_index == 1 else f"{base_name} ({file_index}){extension}"
            destination_path = destination_dir / final_name

            error = download_file_from_url(download_url, str(destination_path))
            if isinstance(error, str):
                print(error)
                continue

            downloaded_paths.append(str(destination_path))

    return downloaded_paths


def process_archive_export_branch(deal: dict) -> bool:
    stage_id = str(
        deal.get("stageId") or deal.get("STAGE_ID") or deal.get("stage_id") or ""
    ).strip()
    if stage_id == ARCHIVE_EXCLUDED_STAGE_ID:
        return False

    enabled_value = get_bitrix_field_value(deal, ARCHIVE_ENABLED_FIELD)
    done_value = get_bitrix_field_value(deal, ARCHIVE_DONE_FIELD)
    if not is_truthy_value(enabled_value) or not is_falsy_value(done_value):
        return False

    company_id = get_company_id_from_deal_context(deal)
    if not company_id:
        print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u0430 \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u044f \u0434\u043b\u044f \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438")
        return False

    company = company_info(company_id)
    if not isinstance(company, dict):
        print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u043d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u044e {company_id}")
        return False

    target_root = str(get_bitrix_field_value(company, ARCHIVE_COMPANY_FOLDER_FIELD) or "").strip()
    if not target_root:
        print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u0443 \u043a\u043e\u043c\u043f\u0430\u043d\u0438\u0438 \u043d\u0435 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043e \u043f\u043e\u043b\u0435 {ARCHIVE_COMPANY_FOLDER_FIELD}")
        return False

    source_item_ids = extract_numeric_bitrix_ids(get_bitrix_field_value(deal, ARCHIVE_SOURCE_ITEM_FIELD))
    if not source_item_ids:
        print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d ID \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0430 \u0432 \u043f\u043e\u043b\u0435 {ARCHIVE_SOURCE_ITEM_FIELD}")
        return False

    source_item = get_smart_process_item(ARCHIVE_SOURCE_ENTITY_TYPE_ID, source_item_ids[0])
    if not isinstance(source_item, dict):
        print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u043d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c \u044d\u043b\u0435\u043c\u0435\u043d\u0442 {ARCHIVE_SOURCE_ENTITY_TYPE_ID}/{source_item_ids[0]}: {source_item}")
        return False

    citizenship_subfolder = "\u0418\u0413" if is_foreign_citizen_deal(deal) else "\u0420\u0424"
    destination_dir = Path(target_root) / build_personal_docs_folder_name(get_fio(deal), citizenship_subfolder)
    #Коммент
    downloaded_paths = download_named_files_from_source_item_fields(source_item, destination_dir)
    if not downloaded_paths:
        print(
            f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: "
            f"\u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u043e \u0444\u0430\u0439\u043b\u043e\u0432 "
            f"\u0432 \u043f\u043e\u043b\u044f\u0445 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u0430 "
            f"{ARCHIVE_SOURCE_ENTITY_TYPE_ID}/{source_item_ids[0]}"
        )
        return False

    if not mark_archive_export_done(deal.get("id")):
        print(
            f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: "
            f"\u043d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c "
            f"\u043e\u0442\u043c\u0435\u0442\u0438\u0442\u044c \u043f\u043e\u043b\u0435 "
            f"{ARCHIVE_DONE_FIELD} \u043f\u043e\u0441\u043b\u0435 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438"
        )
        return False

    print(f"\u0421\u0434\u0435\u043b\u043a\u0430 {deal.get('id')}: \u0432\u044b\u0433\u0440\u0443\u0436\u0435\u043d\u043e {len(downloaded_paths)} \u0444\u0430\u0439\u043b\u043e\u0432 \u0432 {destination_dir}")
    return True


def is_military_liable(parsed_data: dict) -> bool:
    bitrix_fields = parsed_data.get("bitrix_fields", {})
    return is_truthy_value(bitrix_fields.get("UF_CRM_3_1768399393", ""))


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


def get_task_deadline_for_today() -> str:
    local_deadline = datetime.now().replace(hour=18, minute=0, second=0, microsecond=0)
    bitrix_deadline = local_deadline - timedelta(hours=3)
    return bitrix_deadline.strftime("%d.%m.%Y %H:%M:%S")


def make_task_description(deal: dict, fio: str, department_name: str, files_count: int) -> str:
    return (
        f"Смарт-процесс: {deal.get('title', '')}\n"
        f"Сотрудник: {fio}\n"
        f"Направление: {department_name}\n"
        f"ID смарт-процесса: {deal.get('id')}\n"
        f"Количество вложений: {files_count}"
    )


def make_task_title(base_title: str, fio: str, company_name: str) -> str:
    prefix, separator, suffix = base_title.partition(":")
    fio = " ".join(str(fio or "").split())
    company_name = " ".join(str(company_name or "").split())
    details = ", ".join(part for part in (fio, company_name) if part)

    if not details:
        return base_title

    if separator:
        return f"{prefix}: {details}. {suffix.strip()}"
    return f"{base_title}: {details}"


def append_generated_file(target_list: list[str], generated_result: str, department_name: str, document_name: str) -> None:
    if not generated_result:
        print(f"{department_name}: не сформирован документ '{document_name}'")
        return

    if os.path.isfile(generated_result):
        target_list.append(generated_result)
        return

    print(f"{department_name}: ошибка при формировании документа '{document_name}': {generated_result}")


def finalize_generated_result(generated_result: str, parsed_data: dict) -> str:
    if generated_result and os.path.isfile(generated_result):
        try:
            apply_common_placeholder_fallback(generated_result, parsed_data)
        except Exception as error:
            print(f"Не удалось выполнить fallback-заполнение для {generated_result}: {error}")
    return generated_result


def sanitize_value(value):
    if value is None:
        return ""
    if isinstance(value, dict):
        return {key: sanitize_value(inner_value) for key, inner_value in value.items()}
    if isinstance(value, list):
        return [sanitize_value(item) for item in value]
    return value


def enrich_with_bitrix_fields(parsed_data: dict, deal: dict) -> dict:
    bitrix_fields = parsed_data.setdefault("bitrix_fields", {})
    bitrix_fields["UF_CRM_3_1768399223"] = deal.get("UF_CRM_3_1768399223") or deal.get("ufCrm3_1768399223") or ""
    bitrix_fields["UF_CRM_3_1768399393"] = deal.get("UF_CRM_3_1768399393") or deal.get("ufCrm3_1768399393") or ""
    return parsed_data


def has_hire_document_data(parsed_data: dict) -> bool:
    job_info = parsed_data.get("job_info", {})
    return bool(job_info.get("ДатаПриемаНаРаботу") and job_info.get("Должность"))


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


def get_crm_bindings_for_task(deal: dict) -> list[str]:
    crm_bindings: list[str] = []

    parent_deal_id = (
        deal.get("parentId2")
        or deal.get("dealId")
        or deal.get("parentDealId")
    )
    if parent_deal_id:
        crm_bindings.append(build_task_deal_binding(parent_deal_id))

    item_id = deal.get("id")
    if item_id:
        crm_bindings.append(build_task_crm_binding(SOURCE_SMART_PROCESS_ENTITY_TYPE_ID, item_id))

    return crm_bindings


def create_department_task(deal: dict, department_key: str, fio: str, file_paths: list[str]) -> bool:
    config = TASK_CONFIG[department_key]
    creator_id = get_creator_id(deal)
    responsible_debug = get_company_task_responsible_debug(deal, department_key)
    responsible_id = (
        config.get("fixed_responsible_id")
        or responsible_debug["responsible_id"]
        or get_responsible_id(deal, config["responsible_env"])
    )
    accomplices = get_company_task_user_ids_by_field(deal, config.get("company_accomplices_field"))
    accomplices = [user_id for user_id in accomplices if user_id and user_id != responsible_id]
    auditors = get_company_task_auditors(deal, department_key, responsible_id=responsible_id)
    if department_key == "hr":
        file_paths = prioritize_hr_files(file_paths)

    print(
        f"task_create_debug: department='{department_key}'; deal_id='{deal.get('id')}'; "
        f"responsible_id='{responsible_id}'; accomplices={accomplices}; auditors={auditors}; "
        f"files_count={len(file_paths)}"
    )
    task_title = make_task_title(
        base_title=config["title"],
        fio=fio,
        company_name=get_company_name_from_deal_context(deal),
    )

    description = make_task_description(
        deal=deal,
        fio=fio,
        department_name=config["title"],
        files_count=len(file_paths),
    )

    crm_bindings = get_crm_bindings_for_task(deal)

    created_task = add_task(
        creator_id=creator_id,
        responsible_id=responsible_id,
        title=task_title,
        description=description,
        company_name=deal.get("title", ""),
        deadline=get_task_deadline_for_today(),
        auditors=auditors,
        accomplices=accomplices,
    )
    if isinstance(created_task, str):
        print(created_task)
        return False

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
        return False

    task_id = int(task_id_value)
    if crm_bindings:
        crm_binding_result = update_task_crm_bindings(task_id=task_id, crm_bindings=crm_bindings)
        if isinstance(crm_binding_result, str):
            print(f"Не удалось связать задачу {task_id} с CRM-элементом: {crm_binding_result}")
    add_checklist_to_task(task_id=task_id, checklist=config["checklist"])
    attach_files_to_task(task_id=task_id, file_paths=file_paths)
    return True


def move_deal_to_next_stage(deal: dict) -> None:
    if not SOURCE_SMART_PROCESS_NEXT_STAGE_ID:
        return

    deal_id = deal.get("id")
    if not deal_id:
        print("Не удалось перевести сделку на следующий этап: отсутствует ID элемента")
        return

    update_result = update_smart_process_stage(
        entity_id=deal_id,
        entity_type_id=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
        stage_id=SOURCE_SMART_PROCESS_NEXT_STAGE_ID,
    )
    if isinstance(update_result, str):
        print(f"Не удалось перевести сделку {deal_id} на следующий этап: {update_result}")


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

    job_title = get_employee_instruction_template_name(parsed_data)
    job_template_path = get_employee_instruction_template_path(parsed_data)
    if job_template_path.exists():
        append_generated_file(
            hr_files,
            employee_instr_sample_former(
                input_sample=str(job_template_path),
                output_result=f"employee_instructions/output_result/output_{job_title}.docx",
                parsed_data_1c=parsed_data,
            ),
            department_name,
            "Должностная инструкция",
        )

    return hr_files


def generate_occup_safety_documents(parsed_data: dict) -> list[str]:
    occup_safety_files: list[str] = []

    input_dir = Path("occup_safety/input_sample")
    output_dir = Path("occup_safety/output_result")
    output_dir.mkdir(parents=True, exist_ok=True)

    for template_path in sorted(input_dir.iterdir()):
        if not template_path.is_file() or template_path.name.startswith("~$"):
            continue

        output_path = output_dir / f"output_{template_path.name}"
        append_generated_file(
            occup_safety_files,
            occup_safety_former(
                input_sample=str(template_path),
                output_result=str(output_path),
                parsed_data_1c=parsed_data,
            ),
            "Служба охраны труда",
            template_path.stem,
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


def generate_hr_documents_v2(parsed_data: dict, fio: str, base_1c: str) -> list[str]:
    hr_files: list[str] = []
    department_name = "Отдел кадров"

    for input_dir, output_dir, former in HR_TEMPLATE_DIRECTORIES.values():
        process_template_directory(
            target_list=hr_files,
            input_dir=Path(input_dir),
            output_dir=Path(output_dir),
            former=former,
            parsed_data=parsed_data,
            department_name=department_name,
            allowed_suffixes=(".docx",),
        )

    onec_docs = form_order_and_sedo_docs_from_1c(
        fio=fio,
        base_1c=base_1c,
        username=ROBOT_LOGIN,
        pswd=get_1c_password(base_1c),
    )
    append_generated_file(hr_files, onec_docs["order"], department_name, "Приказ о приёме")
    append_generated_file(hr_files, onec_docs["sedo"], department_name, "СЭДО")

    job_title = get_employee_instruction_template_name(parsed_data)
    job_template_path = get_employee_instruction_template_path(parsed_data)
    if job_template_path.exists():
        append_generated_file(
            hr_files,
            finalize_generated_result(
                employee_instr_sample_former(
                    input_sample=str(job_template_path),
                    output_result=f"employee_instructions/output_result/output_{job_title}.docx",
                    parsed_data_1c=parsed_data,
                ),
                parsed_data,
            ),
            department_name,
            "Должностная инструкция",
        )

    return hr_files


def generate_occup_safety_documents_v2(parsed_data: dict) -> list[str]:
    occup_safety_files: list[str] = []
    process_template_directory(
        target_list=occup_safety_files,
        input_dir=Path("occup_safety/input_sample"),
        output_dir=Path("occup_safety/output_result"),
        former=occup_safety_former,
        parsed_data=parsed_data,
        department_name="Служба охраны труда",
        allowed_suffixes=(".docx", ".xlsx"),
    )
    return occup_safety_files


def generate_military_documents_v2(parsed_data: dict) -> list[str]:
    military_files: list[str] = []
    department_name = "Отдел воинского учета"
    input_dir = Path("military_registr/input_samples")
    output_dir = Path("military_registr/output_result")

    for template_path in iter_template_files(input_dir, allowed_suffixes=(".xlsx",)):
        former = MILITARY_TEMPLATE_HANDLERS.get(template_path.stem)
        if former is None:
            former = form_generic_excel_template

        append_generated_file(
            military_files,
            finalize_generated_result(
                former(
                    str(template_path),
                    str(build_output_path(output_dir, template_path)),
                    parsed_data,
                ),
                parsed_data,
            ),
            department_name,
            template_path.stem,
        )

    if not military_files:
        print(
            "military_generate_warning: универсальная ветка не сформировала ни одного файла, "
            "используется fallback на прямую генерацию шаблонов ВУ"
        )
        return generate_military_documents(parsed_data)

    return military_files


def process_deal(deal: dict) -> None:
    try:
        fio = get_fio(deal)
        base_1c = deal.get(BASE_1C_FIELD)
        onec_password = get_1c_password(base_1c)

        if not fio or not base_1c:
            send_result = make_error_msg(BITRIX_TOKEN, BITRIX_ERROR_CHAT_ID,
                                         f"Ошибка в Роботе по приему на работу\nСделка {deal.get('id')}: пропущена из-за отсутствия ФИО или строки подключения к 1С")
            print(f"Сделка {deal.get('id')}: пропущена из-за отсутствия ФИО или строки подключения к 1С")
            return

        parsed_data = app_1c(
            fio=fio,
            base_1c=base_1c,
            username=ROBOT_LOGIN,
            pswd=onec_password,
            inn_org=get_company_inn_from_deal_context(deal),
        )
        if isinstance(parsed_data, str):
            notify_bitrix_error(
                f"Ошибка в Роботе по приему на работу\n"
                f"Контекст: process_deal / app_1c\n"
                f"Сделка ID: {deal.get('id')}\n"
                f"Сделка: {deal.get('title', '')}\n"
                f"ФИО: {fio}\n"
                f"Сообщение: {parsed_data}"
            )
            return

        parsed_data = enrich_with_bitrix_fields(parsed_data, deal)

        migrant_job_info = parsed_data.get("job_info", {})
        migrant_personal_info = parsed_data.get("personal_info", {})
        migrant_document_info = parsed_data.get("document_info", {})
        migrant_bitrix_fields = parsed_data.get("bitrix_fields", {})

        if not has_hire_document_data(parsed_data):
            print(f"Сделка {deal.get('id')}: пропущена, в 1С не найден документ ПриемНаРаботу для {fio}")
            send_result = make_error_msg(BITRIX_TOKEN, BITRIX_ERROR_CHAT_ID,
                                         f"Ошибка в Роботе по приему на работу\nСделка {deal.get('id')}: пропущена, в 1С не найден документ ПриемНаРаботу для {fio}")
            return

        try:
            parsed_data = enrich_bank_details_from_1c(
                parsed_data=parsed_data,
                fio=fio,
                base_1c=base_1c,
                username=ROBOT_LOGIN,
                pswd=onec_password,
            )
        except Exception as error:
            print(f"Сделка {deal.get('id')}: не удалось уточнить банковские реквизиты из 1С: {error}")
            report_exception_to_bitrix(error, context="process_deal / enrich_bank_details_from_1c", deal=deal)

        parsed_data = sanitize_value(parsed_data)
        department_files: dict[str, list[str]] = {}
        enabled_departments: list[str] = []

        if is_enabled(deal, HR_FLAG_FIELD):
            enabled_departments.append("hr")
            department_files["hr"] = generate_hr_documents_v2(parsed_data, fio=fio, base_1c=base_1c)
            upload_hr_files_to_company_bitrix_folder(deal=deal, fio=fio, hr_files=department_files["hr"])

        if is_enabled(deal, OCCUP_SAFETY_FLAG_FIELD):
            enabled_departments.append("occup_safety")
            department_files["occup_safety"] = generate_occup_safety_documents_v2(parsed_data)

        if is_enabled(deal, MILITARY_FLAG_FIELD) and is_military_liable(parsed_data):
            enabled_departments.append("military")
            department_files["military"] = generate_military_documents_v2(parsed_data)

        created_any_task = False
        for department_key in enabled_departments:
            file_paths = department_files.get(department_key, [])
            if not file_paths:
                print(
                    f"task_create_warning: department='{department_key}'; deal_id='{deal.get('id')}'; "
                    f"files_count=0; задача будет создана без вложений"
                )
            created_any_task = create_department_task(
                deal=deal,
                department_key=department_key,
                fio=fio,
                file_paths=file_paths,
            ) or created_any_task

        if created_any_task:
            move_deal_to_next_stage(deal)
    except Exception as error:
        report_exception_to_bitrix(error, context="process_deal", deal=deal)
        raise


def app() -> None:


    sync_templates_from_bitrix()
    archive_candidates = get_all_smart_process_items(
        entity_type_id=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
    )
    if isinstance(archive_candidates, str):
        print(f"Не удалось получить элементы для ветки выгрузки: {archive_candidates}")
        archive_candidates = []

    processed_archive_count = 0
    for deal in archive_candidates:
        if process_archive_export_branch(deal):
            processed_archive_count += 1

    print(f"Ветка выгрузки файлов: обработано элементов {processed_archive_count}")


    bitrix_deals = get_deals_bitrix(
        token=BITRIX_TOKEN,
        entityTypeId=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
        method="crm.item.list",
        stage_id=SOURCE_SMART_PROCESS_STAGE_ID,
    )

    print(
        f"В работу взято элементов: {len(bitrix_deals)} "
        f"(entityTypeId={SOURCE_SMART_PROCESS_ENTITY_TYPE_ID}, stageId={SOURCE_SMART_PROCESS_STAGE_ID})"
    )

    for deal in bitrix_deals:
        process_deal(deal)


def app() -> None:
    try:
        sync_templates_from_bitrix()
        archive_candidates = get_all_smart_process_items(
            entity_type_id=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
        )
        if isinstance(archive_candidates, str):
            notify_bitrix_error(
                f"Ошибка в AprioriNewEmployee\n"
                f"Контекст: app / get_all_smart_process_items\n"
                f"Сообщение: {archive_candidates}"
            )
            print(
                f"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c "
                f"\u044d\u043b\u0435\u043c\u0435\u043d\u0442\u044b \u0434\u043b\u044f \u0432\u0435\u0442\u043a\u0438 "
                f"\u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438: {archive_candidates}"
            )
            archive_candidates = []

        processed_archive_count = 0
        for deal in archive_candidates:
            if process_archive_export_branch(deal):
                processed_archive_count += 1

        print(
            f"\u0412\u0435\u0442\u043a\u0430 \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438 "
            f"\u0444\u0430\u0439\u043b\u043e\u0432: \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u0430\u043d\u043e "
            f"\u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 {processed_archive_count}"
        )

        bitrix_deals = get_deals_bitrix(
            token=BITRIX_TOKEN,
            entityTypeId=SOURCE_SMART_PROCESS_ENTITY_TYPE_ID,
            method="crm.item.list",
            stage_id=SOURCE_SMART_PROCESS_STAGE_ID,
        )

        print(
            f"В работу взято элементов: {len(bitrix_deals)} "
            f"(entityTypeId={SOURCE_SMART_PROCESS_ENTITY_TYPE_ID}, stageId={SOURCE_SMART_PROCESS_STAGE_ID})"
        )

        for deal in bitrix_deals:
            process_deal(deal)
    except Exception as error:
        send_result = make_error_msg(BITRIX_TOKEN, BITRIX_ERROR_CHAT_ID, "Прием на работу\nНе смог корректно отработать: подробная ошибка в следующем сообщении")
        make_error_msg(BITRIX_TOKEN, BITRIX_ERROR_CHAT_ID,
                       str(error))
        #report_exception_to_bitrix(error, context="app")
        raise


if __name__ == "__main__":
    app()
