from bitrix.bitrix import (get_inn_bitrix, get_token_report_file, get_all_companies, add_task, upload_file,
                           upload_to_uploadurl, attach_file_to_task)
from parser_1с.personal_info import organisation_info, only_org_info

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
            file_subname = org_info['org_info']['НаименованиеСокращенное'].replace('"', '').replace('\'', '')

            act_ecp_result: str = form_ecp_act(input_sample='extension_ecp/input_samples/Пример акта.docx',
                                           output_result=f'extension_ecp/output_result/{file_subname}_Пример акта.docx',
                                           parsed_data_1c=org_info)

            invoice_ecp_result: str = form_ecp_invoice(input_sample='extension_ecp/input_samples/Пример счёта.docx',
                                                       output_result=f'extension_ecp/output_result/{file_subname}_Пример счёта.docx',
                                                       parsed_data_1c=org_info)

            created_task: dict | str = add_task(creator_id=69, responsible_id=69, number=5)
            task_id = created_task['id']


            # Загружаем акт
            uploaded_url_act = upload_file(file_path=act_ecp_result,
                                           filename=f'{file_subname}_Пример акта.pdf',
                                           folder_id=24342)
            uploaded_file_act = upload_to_uploadurl(upload_url=uploaded_url_act, field_name='file',
                                                    file_path=act_ecp_result)
            uploaded_file_id_act = uploaded_file_act['result']['ID']
            attached_file_act = attach_file_to_task(task_id=task_id, file_id=uploaded_file_id_act)


            # Загружаем счет
            uploaded_url_invoice = upload_file(file_path=invoice_ecp_result,
                                           filename=f'{file_subname}_Пример счёта.pdf',
                                           folder_id=24342)
            uploaded_file_invoice= upload_to_uploadurl(upload_url=uploaded_url_act, field_name='file',
                                                    file_path=invoice_ecp_result)
            uploaded_file_id_invoice = uploaded_file_invoice['result']['ID']
            attached_file_invoice = attach_file_to_task(task_id=task_id, file_id=uploaded_file_id_invoice)
            break