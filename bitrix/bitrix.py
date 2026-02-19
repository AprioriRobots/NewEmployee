import os
import requests
import json
import base64

from bitrix24 import Bitrix24, BitrixError
from dotenv import load_dotenv


load_dotenv()

BITRIX_TOKEN = os.getenv('BITRIX_TOKEN')


def make_error_msg(token: str, chat_id: str | int, msg: str) -> int | str:
    url: str = f'https://b24.apriori.rest/rest/69/{token}/'

    try:
        bx24 = Bitrix24(url)

        params = {
            "DIALOG_ID": f"{str(chat_id)}",
            "MESSAGE": msg
        }

        msg_response = bx24.callMethod('im.message.add', params=params)
        return msg_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def add_task(token: str, creator_id: int | str, responsible_id: int | str) -> dict | str:
    url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'

    try:
        bx24 = Bitrix24(url)

        params = {
            "fields": {
                'TITLE': 'Тестовая заявка на формирование ТК РФ',
                'DEADLINE': '2026-01-31T23:59:59+02:00',
                'CREATED_BY': creator_id,
                'RESPONSIBLE_ID': responsible_id,
            }
        }

        msg_response = bx24.callMethod('tasks.task.add', params=params)
        return msg_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def get_task_list(token: str) -> dict | str:
    # url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'
    #
    # try:
    #     bx24 = Bitrix24(url)
    #     msg_response = bx24.callMethod('tasks.task.list')
    #     return msg_response
    # except BitrixError as e:
    #     return f'Ошибка во время отправки запроса {e}'

    order = {
        'DEADLINE': 'asc',
        'PRIORITY': 'desc'
    }
    filter_ = {
        # '!STATUS': 6,
        # '>=DEADLINE': datetime.utcnow().strftime('%Y-%m-%d'),
        # 'RESPONSIBLE_ID': 547,
        '::SUBFILTER-PARAMS': {'FAVORITE': 'Y'}
    }
    select = [
        'ID', 'TITLE', 'DESCRIPTION', 'STATUS', 'subStatus',
        'DEADLINE', 'CREATED_DATE', 'RESPONSIBLE_ID',
        'ACCOMPLICES', 'AUDITORS', 'TAGS', 'COUNTERS',
        'PRIORITY', 'MARK'
    ]
    params = {
        'WITH_TIMER_INFO': True,
        'WITH_RESULT_INFO': True,
        'WITH_PARSED_DESCRIPTION': True
    }
    final_params = {
                'order': order,
                # 'filter': filter_,
                'select': select,
                'params': params
    }

    try:
        url: str = f'https://b24.apriori.rest/rest/69/{token}/tasks.task.list'

        # https://corp.tknovosib.ru/rest/717/ipvi7vmargk9d92f/crm.deal.list

        response = requests.post(url=url)
        items = response.text
        for entity in items:
            print('Entity:', entity)
    except Exception as e:
        print('Request failed:', e)


def get_task_by_id(token: str, task_id: str | int) -> dict | str:
    url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'

    try:
        bx24 = Bitrix24(url)

        params = {
            'id': task_id,
            'select': [
                'responsible.name',
                'responsible.email'
            ]
        }

        msg_response = bx24.callMethod('tasks.task.get', params=params)
        return msg_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def get_social_network_lst(token: str, social_network_id: str | int) -> dict | str:
    url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'

    try:
        bx24 = Bitrix24(url)
        params: dict = {
            'filter': {
                'ID': social_network_id
            },
            'select': ['ID', 'NAME']
        }
        msg_response = bx24.callMethod('socialnetwork.api.workgroup.list', params=params)
        return msg_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


# def get_crm_list(token: str) -> dict | str:
#
#
#     order = {
#         'DEADLINE': 'asc',
#         'PRIORITY': 'desc'
#     }
#     filter_ = {
#         # '!STATUS': 6,
#         # '>=DEADLINE': datetime.utcnow().strftime('%Y-%m-%d'),
#         # 'RESPONSIBLE_ID': 547,
#         '::SUBFILTER-PARAMS': {'FAVORITE': 'Y'}
#     }
#     select = [
#         'ID', 'TITLE', 'DESCRIPTION', 'STATUS', 'subStatus',
#         'DEADLINE', 'CREATED_DATE', 'RESPONSIBLE_ID',
#         'ACCOMPLICES', 'AUDITORS', 'TAGS', 'COUNTERS',
#         'PRIORITY', 'MARK'
#     ]
#     params = {
#         'WITH_TIMER_INFO': True,
#         'WITH_RESULT_INFO': True,
#         'WITH_PARSED_DESCRIPTION': True
#     }
#     final_params = {
#                 'order': order,
#                 # 'filter': filter_,
#                 # 'select': select,
#                 'params': params
#     }
#
#     try:
#         url: str = f'https://b24.apriori.rest/rest/69/{token}/crm.item.list'
#
#         response = requests.post(url=url)
#         items = response.text
#         for entity in items:
#             print('Entity:', entity)
#     except Exception as e:
#         print('Request failed:', e)



# def get_current_deal(toke: str, id_deal: str | int) -> dict | str:
#     url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'
#
#     try:
#         bx24 = Bitrix24(url)
#         params: dict = {
#             'filter': {
#                 'ID': social_network_id
#             },
#         }
#         msg_response = bx24.callMethod('crm.deal.get', params=params)
#         return msg_response
#     except BitrixError as e:
#         return f'Ошибка во время отправки запроса {e}'
#     pass


def crm_type_lst(token: str, method: str):
    # crm.type.list
    url = f"https://b24.apriori.rest/rest/69/{token}/{method}"
    response = requests.get(url=url)
    items = response.text
    return items



def get_deals_bitrix(token: str, entityTypeId, method) -> list:

    url = f"https://b24.apriori.rest/rest/69/{token}/{method}"

    params = {
        "entityTypeId": entityTypeId
    }
    # "filter": {
    #             "=stageId": "DT1042_5:UC_II89PK"
    #         }

    response = requests.post(url=url, params=params)
    items = response.json()

    filtered_items = []
    for item in items['result']['items']:
        if item['stageId'] == 'DT1042_5:UC_II89PK':
            filtered_items.append(item)
    return filtered_items


def upload_file(token, file_path, filename, folder_id: int | str):
    """
    Загрузка файла в Битрикс24
    """
    webhook_url = f'https://b24.apriori.rest/rest/69/{token}/'
    with open(file_path, 'rb') as file:
        files = {'file': (filename, file, 'application/octet-stream')}
        response = requests.post(
            webhook_url + "disk.folder.uploadfile",
            files=files,
            data={'id': folder_id}  # ID папки
        )
    if response.status_code == 200:
        result = response.json()
        if result.get('result'):
            upload_url = result['result']['uploadUrl']
            return upload_url
    return None


def upload_to_uploadurl(upload_url, field_name, file_path):
    """
    Отправить файл на UploadUrl с помощью POST-запроса multipart/form-data.
    """
    with open(file_path, 'rb') as f:
        files = {field_name: f}
        response = requests.post(upload_url, files=files)

    if response.status_code == 200:
        print("Файл успешно загружен")
        return response.json()
    else:
        print(f"Ошибка загрузки: {response.status_code}")
        return None


def update_bitrix_file(file_path: str, folder_id: int | str, filename: str) -> str:

    uploaded_url = upload_file(token=BITRIX_TOKEN, file_path=file_path, filename=filename, folder_id=folder_id)
    uploaded_file = upload_to_uploadurl(upload_url=uploaded_url, field_name='file', file_path=file_path)
    return f'Загрузили обновленную версию файла {filename} в битрикс24-диск. Статус ответа: \n{uploaded_file}!'


def add_comments_deal(token: str, entity_id: int | str):
    url: str = f'https://b24.apriori.rest/rest/api/69/{token}/'

    with open(r'C:\Users\apriori_ii2\PycharmProjects\AprioriNewEmployee\nda\output_result\output_Обязательство о неразглашении коммерческой тайны.docx', 'rb') as f:
        file_content = f.read()
    # кодируем в base64, получаем строку без префикса и переноса строк
    file_content_b64 = base64.b64encode(file_content).decode('utf-8')

    try:
        bx24 = Bitrix24(url)
        params: dict = {
            'fields': {
                'ENTITY_ID': entity_id,
                'ENTITY_TYPE': 'order',
                "COMMENT": "Тестовое добавление документов",
                # "FILES": [
				# 	[
				# 		"Test_file_nda.docx",
				# 		file_content_b64
				# 	]
                # ]
            }
        }
        msg_response = bx24.callMethod('crm.timeline.comment.add', params=params)
        return msg_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def add_file_to_smart_process(token: str, entity_id: int | str, entity_type_id: int | str, encrypted_files: list) -> str:
    url: str = f'https://b24.apriori.rest/rest/69/{token}/crm.item.update'

    # with open(
    #         r'C:\Users\apriori_ii2\PycharmProjects\AprioriNewEmployee\nda\output_result\output_Обязательство о неразглашении коммерческой тайны.docx',
    #         'rb') as f:
    #     file_content = f.read()
    # # кодируем в base64, получаем строку без префикса и переноса строк
    # file_content_b64 = base64.b64encode(file_content).decode('utf-8')

    try:
        params: dict = {
            'entityTypeId': int(entity_type_id),
            'id': int(entity_id),

            'fields': {
                'ufCrm3_1770989930': encrypted_files
            }
        }
        msg_response = requests.post(url, json=params)
        # msg_response = bx24.callMethod('crm.item.update', params=params)
        return msg_response
    except IndexError as e:
        return f'Ошибка во время отправки запроса {e}'



if __name__ == '__main__':
    get_type_lst = crm_type_lst(token=BITRIX_TOKEN, method='crm.type.list')

    bitrix_deals: list[dict] | str = get_deals_bitrix(token=BITRIX_TOKEN, entityTypeId=1042,
                                                      method='crm.item.list')
    needed_deal = ''

    for deal in bitrix_deals:
        if deal['title'] == 'Приём на работу #88':
            needed_deal = deal

    file_id = 16235
    add_file =  add_file_to_smart_process(token=BITRIX_TOKEN, entity_id=needed_deal['id'], entity_type_id=1042,
                                          file_id=file_id)
