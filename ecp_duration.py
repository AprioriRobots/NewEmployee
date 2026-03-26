import datetime
import time

from bitrix.bitrix import (get_inn_bitrix, get_token_report_file, get_all_companies, add_task_ecp, upload_file,
                           upload_to_uploadurl, attach_file_to_task, get_buh_id, make_task_dependence, update_task,
                           get_task_by_id, get_contractors_reqs)
from parser_1с.personal_info import organisation_info, only_org_info

from extension_ecp.form_ecp import form_ecp_act, form_ecp_invoice
from extension_ecp.data_funcs import app_data_extracter, extract_row_if_red



def match_data(extracted_data: list[dict], inn_company: str | int) -> dict | None:
    for data in extracted_data:
        if str(data['ИНН']) == str(inn_company):
            founded_data: dict = {
                'ДатаОкончанияЭЦП': data['ДатаОкончанияЭЦП'],
                'ИНН': str(data['ИНН'])
            }

            return founded_data

    return None

def format_today_date() -> str:
  today = datetime.date.today()
  formatted_date = today.strftime("%d-%m-%Y")

  return formatted_date


if __name__ == '__main__':
    # act_ecp_result: str = form_ecp_act(input_sample='input_samples/Пример акта.docx',
    #                                    output_result='output_result/output_Пример акта.docx',
    #                                    parsed_data_1c={})

    invoice_tokens_file: str | None = get_token_report_file()
    contractors_req_data: dict = get_contractors_reqs()

    if invoice_tokens_file is None:
        print('Файл "ОТКЕНЫ ОТЧЕТНОСТЬ" не был загружен. Завершение программы')

        pass

    extracted_data: list[dict] = app_data_extracter(path_file='extension_ecp/report_tokens/ТОКЕНЫ ОТЧЕТНОСТЬ.xlsx')

    companies: list = get_all_companies()

    for company in companies:

        if (company['TITLE'] == 'Тестовая компания' or company['UF_CRM_1771251320632'] is None
                or company['UF_CRM_1771251801387'] is None):
            continue

        buh: str | None = get_buh_id(company_id=company['ID'])
        inn_company: str | int = get_inn_bitrix(company_id=company['ID'])
        # inn_company
        founded_data: dict | None = match_data(extracted_data=extracted_data, inn_company=inn_company)

        if founded_data is None or buh is None:
            # Значит нет в списке извлеченных данных, переходим к следующей карточке компании
            continue

        inn_company, ecp_date = founded_data['ИНН'], founded_data['ДатаОкончанияЭЦП']
        base_1c: str = company['UF_CRM_1771251801387']

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

            descr_for_main_task: str = (f"У {company['TITLE']} {ecp_date} истекает срок действия ЭЦП.\n\nПрошу связаться с"
                                        f" клиентом и согласовать продление ЭЦП, а также уточнить, будут ли они"
                                        f" продлевать через нас или самостоятельно.")
            title_for_main_task: str = f"{company['TITLE']}, продление ЭЦП"
            created_main_task: dict | str = add_task_ecp(creator_id=43, responsible_id=buh, group_id=236, title=title_for_main_task,
                                                     description=descr_for_main_task, company_name=company['TITLE'])
            main_task_id = created_main_task['task']['id']

            time.sleep(1)

            updated_task: dict | str = update_task(task_id=main_task_id, company_name=company['TITLE'])

            # Получаем номер документа
            current_num_doc: str | int = get_task_by_id(task_id=main_task_id)['ufAuto713946616750']
            org_info['org_info']['НомерДокумента'] = current_num_doc

            file_subname = org_info['org_info']['НаименованиеСокращенное'].replace('"', '').replace('\'', '')

            act_ecp_result: str = form_ecp_act(input_sample='extension_ecp/input_samples/Пример акта.docx',
                                               output_result=f'extension_ecp/output_result/{file_subname}_Пример акта.docx',
                                               parsed_data_1c=org_info, reqs_bank_main_org=contractors_req_data)

            invoice_ecp_result: str = form_ecp_invoice(input_sample='extension_ecp/input_samples/Пример счёта.docx',
                                                       output_result=f'extension_ecp/output_result/{file_subname}_Пример счёта.docx',
                                                       parsed_data_1c=org_info, reqs_bank_main_org=contractors_req_data)

            today: str  = format_today_date()

            # Загружаем акт
            uploaded_url_act = upload_file(file_path=act_ecp_result,
                                           filename=f'{file_subname}_акт_{today}.pdf',
                                           folder_id=24342)

            # uploaded_file_act = upload_to_uploadurl(upload_url=uploaded_url_act, field_name='file',
            #                                         file_path=act_ecp_result)

            uploaded_file_id_act = uploaded_url_act['ID']
            attached_file_act = attach_file_to_task(task_id=main_task_id, file_id=uploaded_file_id_act)

            # Загружаем счет
            uploaded_url_invoice = upload_file(file_path=invoice_ecp_result,
                                               filename=f'{file_subname}_счёт_{today}.pdf',
                                               folder_id=24342)

            # uploaded_file_invoice= upload_to_uploadurl(upload_url=uploaded_url_act, field_name='file',
            #                                         file_path=invoice_ecp_result)

            uploaded_file_id_invoice = uploaded_url_invoice['ID']
            attached_file_invoice = attach_file_to_task(task_id=main_task_id, file_id=uploaded_file_id_invoice)
            # break