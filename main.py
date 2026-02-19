import base64
import datetime
import os
import win32com.client

from dotenv import load_dotenv

from parser_1с.personal_info import app_1c
from bitrix.bitrix import get_deals_bitrix, add_file_to_smart_process

from mat_response.mat_response import mat_resp_sample_former
from nda.form_nda import nda_sample_former
from lna.form_lna import lna_sample_former
from personal_secret.form_personal_secret import personal_secret_sample_former
from tk_rf.form_tkrf import tkrf_sample_former
from zp_alignment.form_zp_alignment import zp_align_sample_former
from formul_tkrf.form_formul_tkrf import formul_tkrf_sample_former


load_dotenv()

BITRIX_TOKEN = os.getenv('BITRIX_TOKEN')


def encrypt_to_base64(file_lst: list, fio: str) -> list[list]:
    encrypted_files: list[list] = []

    for file in file_lst:
        with open(file, 'rb') as f:
            file_content = f.read()
        # кодируем в base64, получаем строку без префикса и переноса строк
        file_content_b64 = base64.b64encode(file_content).decode('utf-8')

        file_name: str = file.split('/')[-1].replace('output', fio)
        pair: list = [file_name, file_content_b64]
        encrypted_files.append(pair)

    return encrypted_files


def app():
    # Парсим из битрикса смарт-процессы с определенной стадией (Проверка документов)
    bitrix_deals: list[dict] | str = get_deals_bitrix(token=BITRIX_TOKEN, entityTypeId=1042,
                                                      method='crm.item.list')
    needed_deal = ''

    # Временный код, пока протестим для одного смарт-процесса
    for deal in bitrix_deals:
        if deal['title'] == 'Приём на работу #91':
            needed_deal = deal

    fio: str = f'{needed_deal['ufCrm3_1768399087']} {needed_deal['ufCrm3_1768400217']} {needed_deal['ufCrm3_1768399180']}'
    base_1c: str = needed_deal['ufCrm3_1770808935']

    # Изымаем нужные параметры (данные) из ЗУП базы 1с (инфа по базе берется из смарт-процесса)
    parsed_data: dict = app_1c(fio=fio, base_1c=base_1c, username='Робот',
                               pswd='YUSliot3Cb')

    # Создаем список документов. По ним мы потом будет кодировать в base64 содержимое файлов
    formed_files: list = []

    # Далее формируем кадровые документы по заданым параметрам
    mat_resp: str = mat_resp_sample_former(
        input_sample='mat_response/input_sample/Договор о полной индивидуальной материальной ответственности.docx',
        output_result='mat_response/output_result/output_Договор о полной индивидуальной материальной ответственности.docx',
        parsed_data_1c=parsed_data
    )
    formed_files.append(mat_resp)

    nda: str = nda_sample_former(
        input_sample='nda/input_sample/Обязательство о неразглашении коммерческой тайны.docx',
        output_result='nda/output_result/output_Обязательство о неразглашении коммерческой тайны.docx',
        parsed_data_1c=parsed_data
    )
    formed_files.append(nda)

    lna: str = lna_sample_former(input_sample='lna/input_sample/Лист ознакомления с ЛНА.docx',
                                 output_result='lna/output_result/output_Лист ознакомления с ЛНА.docx',
                                 parsed_data_1c=parsed_data)
    formed_files.append(lna)

    distribution_data: str = personal_secret_sample_former(
        input_sample='personal_secret/input_sample/Согласие на распространение персональных данных НОВОЕ.docx',
        output_result='personal_secret/output_result/output_Согласие на распространение персональных данных НОВОЕ.docx',
        parsed_data_1c=parsed_data
    )
    formed_files.append(distribution_data)

    processing_data: str = personal_secret_sample_former(
        input_sample='personal_secret/input_sample/Согласие на обработку персональных данных НОВОЕ.docx',
        output_result='personal_secret/output_result/output_Согласие на обработку персональных данных НОВОЕ.docx',
        parsed_data_1c=parsed_data
    )
    formed_files.append(processing_data)

    tkrf: str = tkrf_sample_former(input_sample='tk_rf/input_sample/Трудовой договор.docx',
                                   output_result='tk_rf/output_result/output_Трудовой договор.docx',
                                   parsed_data_1c=parsed_data)
    formed_files.append(tkrf)

    zp_align: str = zp_align_sample_former(
        input_sample='zp_alignment/input_sample/Заявление на перечисление заработной платы.docx',
        output_result='zp_alignment/output_result/output_Заявление на перечисление заработной платы.docx',
        parsed_data_1c=parsed_data
    )
    formed_files.append(zp_align)

    tkrf_formul: str = formul_tkrf_sample_former(
        input_sample='formul_tkrf/input_sample/Формулировки для трудового договора.xlsx',
        output_result='formul_tkrf/output_result/output_Формулировки для трудового договора.xlsx',
        parsed_data_1c=parsed_data
    )

    # Далее идет функция кодирования каждого файла из зписка formed_files. Вернется список списков, где каждый список lst[0] - название файла, lst[1] - содержимое в base64
    encrypted_form_files: list[list] = encrypt_to_base64(file_lst=formed_files, fio=fio)
    # encrypted_trk_rf = encrypted_form_files[5]

    status = add_file_to_smart_process(token=BITRIX_TOKEN, entity_id=needed_deal['id'], entity_type_id=1042,
                                       encrypted_files=encrypted_form_files)
    print(status)


if __name__ == '__main__':
    app()
    # # Робот прием
    # parsed_data: dict = app_1c(fio='Фам Тхи Тху Ха', base_1c='zup_bazilevskii', username='Робот',
    #                            pswd='YUSliot3Cb')
    #
    #
    # mat_resp: str = mat_resp_sample_former(
    #     input_sample='mat_response/input_sample/Договор о полной индивидуальной материальной ответственности.docx',
    #     output_result='mat_response/output_result/output_Договор о полной индивидуальной материальной ответственности.docx',
    #     parsed_data_1c=parsed_data
    # )
    #
    # nda: str = nda_sample_former(
    #     input_sample='nda/input_sample/Обязательство о неразглашении коммерческой тайны.docx',
    #     output_result='nda/output_result/output_Обязательство о неразглашении коммерческой тайны.docx',
    #     parsed_data_1c=parsed_data
    # )
    #
    # lna: str = lna_sample_former(input_sample='lna/input_sample/Лист ознакомления с ЛНА.docx',
    #                              output_result='lna/output_result/output_Лист ознакомления с ЛНА.docx',
    #                              parsed_data_1c=parsed_data)
    #
    #
    # distribution_data: str = personal_secret_sample_former(
    #     input_sample='personal_secret/input_sample/Согласие на распространение персональных данных НОВОЕ.docx',
    #     output_result='personal_secret/output_result/output_Согласие на распространение персональных данных НОВОЕ.docx',
    #     parsed_data_1c=parsed_data
    # )
    #
    # processing_data: str = personal_secret_sample_former(
    #     input_sample='personal_secret/input_sample/Согласие на обработку персональных данных НОВОЕ.docx',
    #     output_result='personal_secret/output_result/output_Согласие на обработку персональных данных НОВОЕ.docx',
    #     parsed_data_1c=parsed_data
    # )
    #
    #
    # tkrf: str = tkrf_sample_former(input_sample='tk_rf/input_sample/Трудовой договор.docx',
    #                                output_result='tk_rf/output_result/output_Трудовой договор.docx',
    #                                parsed_data_1c=parsed_data)
    #
    # zp_align: str = zp_align_sample_former(
    #     input_sample='zp_alignment/input_sample/Заявление на перечисление заработной платы.docx',
    #     output_result='zp_alignment/output_result/output_Заявление на перечисление заработной платы.docx',
    #     parsed_data_1c=parsed_data
    # )
