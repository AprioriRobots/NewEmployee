import os

from docx import Document
from python_docx_replace import docx_get_keys, docx_replace

from name_utils import get_fio_part, get_name_patronymic_initials
from template_utils import expand_relative_month_keys, inflect_to_genitive_phrase


def _build_employee_instruction_words(parsed_data_1c: dict) -> dict:
    org_info = parsed_data_1c["job_info"]["\u041e\u0440\u0433\u0430\u043d\u0438\u0437\u0430\u0446\u0438\u044f"]
    fio = parsed_data_1c["personal_info"]["\u0424\u0418\u041e"]

    surname = get_fio_part(fio, 0)

    return {
        "\u0414\u0430\u0442\u0430 \u043f\u0440\u0438\u043a\u0430\u0437\u0430 \u043e \u043f\u0440\u0438\u0435\u043c\u0435": parsed_data_1c["job_info"]["\u0414\u0430\u0442\u0430\u041f\u0440\u0438\u0435\u043c\u0430\u041d\u0430\u0420\u0430\u0431\u043e\u0442\u0443"],
        "\u0414\u043e\u0433\u043e\u0432\u043e\u0440 \u2116": parsed_data_1c["job_info"]["\u0422\u043a\u041d\u043e\u043c\u0435\u0440"],
        "\u041f\u043e\u043b\u043d\u043e\u0435 \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435": inflect_to_genitive_phrase(
            org_info["\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435\u041f\u043e\u043b\u043d\u043e\u0435"],
            "\u041f\u043e\u043b\u043d\u043e\u0435 \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
        ),
        "\u041a\u0440\u0430\u0442\u043a\u043e\u0435 \u043d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435": org_info["\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435\u0421\u043e\u043a\u0440\u0430\u0449\u0435\u043d\u043d\u043e\u0435"],
        "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c \u0440\u0443\u043a\u043e\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044f": inflect_to_genitive_phrase(
            org_info["\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c\u0420\u0443\u043a\u043e\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044f"],
            "\u0414\u043e\u043b\u0436\u043d\u043e\u0441\u0442\u044c \u0440\u0443\u043a\u043e\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044f",
        ),
        "\u0424\u0430\u043c\u0438\u043b\u0438\u044f": surname,
        "\u0424\u0430\u043c\u0438\u043b\u0438\u044f \u0440\u0430\u0431\u043e\u0442\u043d\u0438\u043a\u0430": surname,
        "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b \u0438\u043c\u0435\u043d\u0438, \u043e\u0442\u0447\u0435\u0441\u0442\u0432\u0430": get_name_patronymic_initials(fio),
        "\u0418\u043d\u0438\u0446\u0438\u0430\u043b\u044b \u0440\u0430\u0431\u043e\u0442\u043d\u0438\u043a\u0430": get_name_patronymic_initials(fio),
        "\u0420\u0443\u043a\u043e\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044c": org_info["\u0420\u0443\u043a\u043e\u0432\u043e\u0434\u0438\u0442\u0435\u043b\u044c"],
    }


def employee_instr_sample_former(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    employee_instr_input_words = _build_employee_instruction_words(parsed_data_1c)

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern="<>")
        docx_replace(doc, "<>", **expand_relative_month_keys(employee_instr_input_words, keys))
        doc.save(output_result)
        return output_result
    except Exception as e:
        return f"\u041e\u0448\u0438\u0431\u043a\u0430 \u0432\u043e \u0432\u0440\u0435\u043c\u044f \u0444\u043e\u0440\u043c\u0438\u0440\u043e\u0432\u0430\u043d\u0438\u044f \u0444\u0430\u0439\u043b\u0430: {e}"


def employee_instr_sample_former2(input_sample: str, output_result: str, parsed_data_1c: dict) -> list[str]:
    employee_instr_input_words = _build_employee_instruction_words(parsed_data_1c)
    output_results = []

    for file in os.listdir(input_sample):
        try:
            doc = Document(os.path.join(input_sample, file))
            keys = docx_get_keys(doc, pattern="<>")
            docx_replace(doc, "<>", **expand_relative_month_keys(employee_instr_input_words, keys))

            output_path = os.path.join(output_result, f"output_{file}")
            doc.save(output_path)
            output_results.append(output_path)
        except Exception:
            continue

    return output_results
