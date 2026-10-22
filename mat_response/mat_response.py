from datetime import datetime

from docx import Document
from python_docx_replace import docx_get_keys, docx_replace

from template_utils import (
    expand_relative_month_keys,
    inflect_to_dative_phrase,
    inflect_to_genitive_phrase,
)


def _format_date_for_template(value) -> str:
    if value in (None, ""):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")

    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue

    return text


def mat_resp_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    org_info = parsed_data_1c["job_info"]["Организация"]
    personal_info = parsed_data_1c["personal_info"]
    document_info = parsed_data_1c["document_info"]
    job_info = parsed_data_1c["job_info"]

    mat_resp_input_words = {
        "Дата приказа о приеме": _format_date_for_template(job_info["ДатаПриемаНаРаботу"]),
        "Полное наименование": inflect_to_dative_phrase(
            org_info["НаименованиеПолное"], "Полное наименование"
        ),
        "Краткое наименование": org_info["НаименованиеСокращенное"],
        "Должность руководителя": org_info["ДолжностьРуководителя"],
        "Руководитель": org_info["Руководитель"],
        "Основание подписи": job_info.get("ОснованиеПодписи", ""),
        "Основания подписи": job_info.get("ОснованиеПодписи", ""),
        "Гражданство страны": inflect_to_genitive_phrase(
            personal_info["Страна"], "Гражданство страны"
        ),
        "Полное имя работника": personal_info["ФИО"],
        "Должность работника": inflect_to_genitive_phrase(
            job_info["Должность"], "Должность работника"
        ),
        "Юридический адрес": org_info["ЮридическийАдрес"],
        "Фактический адрес": org_info["ФактическийАдрес"],
        "ИНН организации": org_info["ИНН_Орг"],
        "КПП организации": org_info["КПП_Орг"],
        "ОГРН организации": org_info["ОГРН(ИП)"],
        "Расчетный счет организации": org_info["РасСчет"],
        "БИК банка организации": org_info["БикБанка"],
        "Адрес по прописке": personal_info["Адрес_прописка"],
        "Дата рождения": document_info["ДатаРождения"],
        "Место рождения": document_info["МестоРождения"],
        "Серия документа, удостоверяющего личность": document_info["Серия"],
        "Номер документа, удостоверяющего личность": document_info["Номер"],
        "Орган, выдавший документ, удостоверяющего личность": document_info["КемВыдан"],
        "Дата выдачи документа, удостоверяющего личность": document_info["ДатаВыдачи"],
        "СНИЛС": personal_info["СНИЛС"],
        "Корреспондентский счет организации": org_info["КорСчет"],
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern="<>")
        docx_replace(doc, "<>", **expand_relative_month_keys(mat_resp_input_words, keys))
        doc.save(output_result)
        return output_result
    except Exception as e:
        return f"Ошибка во время формирования файла: {e}"
