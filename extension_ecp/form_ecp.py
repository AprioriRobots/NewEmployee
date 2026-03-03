from docx import Document
from python_docx_replace import docx_replace, docx_get_keys
from pywintypes import com_error

from bitrix.bitrix import get_all_companies, get_inn_bitrix
from parser_1с.personal_info import only_org_info


def form_ecp_act(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    # act_input_words: dict = {
    #     'Рабочая дата': parsed_data_1c['org_info']['РабочаяДата'],
    #     'Краткое наименование': parsed_data_1c['org_info']['НаименованиеСокращенное'],
    #     'ИНН организации': parsed_data_1c['org_info']['ИНН_Орг'],
    #     'КПП организации': parsed_data_1c['org_info']['КПП_Орг'],
    #     'Юридический адрес': parsed_data_1c['org_info']['ЮридическийАдрес'],
    #     'Расчетный счёт организации': parsed_data_1c['org_info']['РасСчет'],
    #     'БИК банка организации': parsed_data_1c['org_info']['БикБанка'],
    #     'Корреспондентский счет организации': parsed_data_1c['org_info']['КорСчет'],
    #     'Должность руководителя': parsed_data_1c['org_info']['ДолжностьРуководителя'],
    #     'Руководитель': parsed_data_1c['org_info']['Руководитель']
    # }
    act_input_words = {'БИК банка организации': '044525225', 'Должность руководителя': 'Генеральный директор', 'ИНН организации': '7709952203', 'КПП организации': '770901001', 'Корреспондентский счет организации': '30101810400000000225', 'Краткое наименование': 'ООО "ВИННИК"', 'Рабочая дата': '02.03.2026', 'Расчетный счёт организации': '40702810238000170440 в ПАО  СБЕРБАНК  г. Москва', 'Руководитель': 'Куликова Маргарита Геннадьевна', 'Юридический адрес': '101000, город Москва, ул. Покровка, дом 16, ПОМ II-КОМН 2'}

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **act_input_words)

        doc.save(output_result)
        return output_result

    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


def form_ecp_invoice(input_sample: str, output_result: str, parsed_data_1c: dict) -> str:
    invoice_input_words: dict = {
        'Рабочая дата': parsed_data_1c['org_info']['РабочаяДата'],
        'Краткое наименование': parsed_data_1c['org_info']['НаименованиеСокращенное'],
        'ИНН организации': parsed_data_1c['org_info']['ИНН_Орг'],
        'КПП организации': parsed_data_1c['org_info']['КПП_Орг'],
        'Юридический адрес': parsed_data_1c['org_info']['ЮридическийАдрес'],
        'Расчетный счёт организации': parsed_data_1c['org_info']['РасСчет'],
        'БИК банка организации': parsed_data_1c['org_info']['БикБанка'],
        'Корреспондентский счет организации': parsed_data_1c['org_info']['КорСчет']
    }

    try:
        doc = Document(input_sample)
        keys = docx_get_keys(doc, pattern='<>')
        print(keys)
        docx_replace(doc, '<>', **invoice_input_words)

        doc.save(output_result)
        return output_result

    except Exception as e:
        return f'Ошибка во время формирования файла: {e}'


if __name__ == '__main__':
    # act_ecp_result: str = form_ecp_act(input_sample='input_samples/Пример акта.docx',
    #                                    output_result='output_result/output_Пример акта.docx',
    #                                    parsed_data_1c={})

    companies: list = get_all_companies()

    for company in companies:

        if (company['TITLE'] == 'Тестовая компания' or company['UF_CRM_1771251320632'] is None
                or company['UF_CRM_1771251801387'] is None):
            continue

        inn_company: str | int = get_inn_bitrix(company_id=company['ID'])
        inn_company = '7709952203'
        base_1c: str = company['UF_CRM_1771251801387']
        base_1c = 'Srvr="srv2023.apriori.loc:1541";Ref="zup_perelman";'
        # print(inn_company, base_1c, sep='\t')

        org_info: dict | None = None

        try:
            if 'zup_perelman' in base_1c:
                org_info: dict = only_org_info(inn_org=inn_company, base_1c=base_1c,
                                                                   username='Робот', pswd='Q1234567890Q')
            else:
                org_info: dict = only_org_info(inn_org=inn_company, base_1c=base_1c,
                                                                   username='Робот', pswd='YUSliot3Cb')
        except com_error as e:
            print(f'Ошибка во время авторизации: {e}\nБаза: {base_1c}')
            continue

        if org_info is not None:
            act_ecp_result: str = form_ecp_act(input_sample='input_samples/Пример акта.docx',
                                           output_result='output_result/output_Пример акта.docx',
                                           parsed_data_1c=org_info)

            invoice_ecp_result: str = form_ecp_act(input_sample='input_samples/Пример счёта.docx',
                                           output_result='output_result/output_Пример счёта.docx',
                                           parsed_data_1c=org_info)

            break
