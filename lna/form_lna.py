from docx import Document
from python_docx_replace import docx_get_keys, docx_replace

from template_utils import expand_relative_month_keys


def lna_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    org_info = parsed_data_1c["job_info"]["Организация"]
    job_info = parsed_data_1c["job_info"]

    lna_input_words = {
        "Краткое наименование": org_info["НаименованиеСокращенное"],
        "Полное наименование": org_info["НаименованиеПолное"],
        "Дата приказа о приеме": job_info["ДатаПриемаНаРаботу"],
        "Должность работника": job_info["Должность"],
        "Должность руководителя": org_info.get("ДолжностьРуководителя", ""),
        "Руководитель": org_info.get("Руководитель", ""),
        "Основание подписи": job_info.get("ОснованиеПодписи", ""),
        "Основания подписи": job_info.get("ОснованиеПодписи", ""),
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern="<>")
        docx_replace(doc, "<>", **expand_relative_month_keys(lna_input_words, keys))
        doc.save(output_result)
        return output_result
    except Exception as e:
        return f"Ошибка во время формирования файла: {e}"
