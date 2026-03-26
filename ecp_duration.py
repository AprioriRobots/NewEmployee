import datetime
import time
import logging
import shutil
from pathlib import Path


from bitrix.bitrix import (get_inn_bitrix, get_token_report_file, get_all_companies, add_task_ecp, upload_file,
                           upload_to_uploadurl, attach_file_to_task, get_buh_id, make_task_dependence, update_task,
                           get_task_by_id, get_contractors_reqs)
from parser_1с.personal_info import organisation_info, only_org_info

from extension_ecp.form_ecp import form_ecp_act, form_ecp_invoice
from extension_ecp.data_funcs import app_data_extracter, extract_row_if_red


LOGGER_FORMAT = '%(asctime)s : %(name)s : %(levelname)s : %(message)s'
LOG_FILE = 'ecp_duration.log'
logger = logging.getLogger(__name__)


def setup_logging() -> None:
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    log_path = str(Path(LOG_FILE).resolve())
    has_handler = any(
        isinstance(handler, logging.FileHandler) and getattr(handler, 'baseFilename', None) == log_path
        for handler in root_logger.handlers
    )

    if has_handler:
        return

    file_handler = logging.FileHandler(filename=LOG_FILE, mode='a', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter(LOGGER_FORMAT))
    root_logger.addHandler(file_handler)


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
    formatted_date = today.strftime('%d-%m-%Y')

    return formatted_date


def clear_output_result_dir(output_dir: str | Path) -> None:
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    logger.info('Очистка директории результатов: %s', output_path)

    for item in output_path.iterdir():
        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
        except OSError as exc:
            logger.exception('Ошибка очистки элемента %s', item)
            print(f'Не удалось удалить {item}: {exc}')


if __name__ == '__main__':
    setup_logging()
    logger.info('Запуск ecp_duration начат')

    clear_output_result_dir('extension_ecp/output_result')

    invoice_tokens_file: str | None = get_token_report_file()
    contractors_req_data: dict = get_contractors_reqs()
    logger.info('Файл отчета по токенам получен: %s', bool(invoice_tokens_file))

    if invoice_tokens_file is None:
        logger.warning('Файл отчета по токенам не был загружен')
        print('Файл "ТОКЕНЫ ОТЧЕТНОСТЬ" не был загружен. Завершение программы')
        pass

    extracted_data: list[dict] = app_data_extracter(path_file='extension_ecp/report_tokens/ТОКЕНЫ ОТЧЕТНОСТЬ.xlsx')
    logger.info('Количество извлеченных строк из отчета: %d', len(extracted_data))

    companies: list = get_all_companies()
    logger.info('Количество компаний, загруженных из Битрикс: %d', len(companies) if companies else 0)

    for company in companies:

        if (company['TITLE'] == 'Тестовая компания' or company['UF_CRM_1771251320632'] is None
                or company['UF_CRM_1771251801387'] is None):
            logger.info('Компания %s пропущена из-за фильтров', company.get('ID'))
            continue

        buh: str | None = get_buh_id(company_id=company['ID'])
        inn_company: str | int = get_inn_bitrix(company_id=company['ID'])
        logger.info('Обработка компании id=%s', company.get('ID'))

        founded_data: dict | None = match_data(extracted_data=extracted_data, inn_company=inn_company)

        if founded_data is None or buh is None:
            logger.warning('Нет подходящих данных по токену или бухгалтера для компании id=%s', company.get('ID'))
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
            logger.exception('Ошибка авторизации для базы %s', base_1c)
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
            logger.info('Основная задача создана: task_id=%s company_id=%s', main_task_id, company.get('ID'))

            time.sleep(1)

            updated_task: dict | str = update_task(task_id=main_task_id, company_name=company['TITLE'])
            logger.info('Основная задача обновлена: task_id=%s', main_task_id)

            current_num_doc: str | int = get_task_by_id(task_id=main_task_id)['ufAuto713946616750']
            org_info['org_info']['НомерДокумента'] = current_num_doc
            logger.info('Номер документа получен для task_id=%s', main_task_id)

            file_subname = org_info['org_info']['НаименованиеСокращенное'].replace('"', '').replace("'", '')

            act_ecp_result: str = form_ecp_act(input_sample='extension_ecp/input_samples/Пример акта.docx',
                                               output_result=f'extension_ecp/output_result/{file_subname}_Пример акта.docx',
                                               parsed_data_1c=org_info, reqs_bank_main_org=contractors_req_data)
            logger.info('Акт сформирован: %s', act_ecp_result)

            invoice_ecp_result: str = form_ecp_invoice(input_sample='extension_ecp/input_samples/Пример счёта.docx',
                                                       output_result=f'extension_ecp/output_result/{file_subname}_Пример счёта.docx',
                                                       parsed_data_1c=org_info, reqs_bank_main_org=contractors_req_data)
            logger.info('Счет сформирован: %s', invoice_ecp_result)

            today: str = format_today_date()

            uploaded_url_act = upload_file(file_path=act_ecp_result,
                                           filename=f'{file_subname}_акт_{today}.pdf',
                                           folder_id=24342)

            uploaded_file_id_act = uploaded_url_act['ID']
            attached_file_act = attach_file_to_task(task_id=main_task_id, file_id=uploaded_file_id_act)
            logger.info('Акт загружен и прикреплен: task_id=%s file_id=%s', main_task_id, uploaded_file_id_act)

            uploaded_url_invoice = upload_file(file_path=invoice_ecp_result,
                                               filename=f'{file_subname}_счёт_{today}.pdf',
                                               folder_id=24342)

            uploaded_file_id_invoice = uploaded_url_invoice['ID']
            attached_file_invoice = attach_file_to_task(task_id=main_task_id, file_id=uploaded_file_id_invoice)
            logger.info('Счет загружен и прикреплен: task_id=%s file_id=%s', main_task_id, uploaded_file_id_invoice)

    logger.info('Выполнение ecp_duration завершено')
