from bitrix.bitrix import get_inn_bitrix, get_token_report_file, get_all_companies
from parser_1с.personal_info import organisation_info

from extension_ecp.form_ecp import form_ecp_act, form_ecp_invoice




if __name__ == '__main__':
    # act_ecp_result: str = form_ecp_act(input_sample='input_samples/Пример акта.docx',
    #                                    output_result='output_result/output_Пример акта.docx',
    #                                    parsed_data_1c={})

    invoice_tokens_file: str | None = get_token_report_file()

    if invoice_tokens_file is None:
        print('Файл "ОТКЕНЫ ОТЧЕТНОСТЬ" не был загружен. Завершение программы')
        pass

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
            act_ecp_result: str = form_ecp_act(input_sample='extension_ecp/input_samples/Пример акта.docx',
                                           output_result='extension_ecp/output_result/output_Пример акта.docx',
                                           parsed_data_1c=org_info)

            invoice_ecp_result: str = form_ecp_act(input_sample='extension_ecp/input_samples/Пример счёта.docx',
                                           output_result='extension_ecp/output_result/output_Пример счёта.docx',
                                           parsed_data_1c=org_info)

            break