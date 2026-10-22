import os
import requests
import json
import base64
import datetime
import mimetypes
import re
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse, unquote

from bitrix24 import Bitrix24, BitrixError
from dotenv import load_dotenv


load_dotenv()

BITRIX_TOKEN = os.getenv('BITRIX_TOKEN')
BITRIX_PORTAL_BASE_URL = 'https://b24.apriori.rest'


def make_error_msg(token: str, chat_id: str | int, msg: str) -> int | str:

    print('here')
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


def build_task_crm_binding(entity_type_id: int | str, entity_id: int | str) -> str:
    return f"T{int(entity_type_id):x}_{int(entity_id)}"


def build_task_deal_binding(deal_id: int | str) -> str:
    return f"D_{int(deal_id)}"


def add_task(creator_id: int | str, responsible_id: int | str, title: str, description: str, company_name: str,
             contains_results: bool = True, contains_gantt_links: bool = True,
             crm_bindings: list[str] | None = None, deadline: str | None = None,
             auditors: list[int] | None = None,
             accomplices: list[int] | None = None) -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/'

    try:
        bx24 = Bitrix24(url)

        params: dict = {
            "fields": {
                "TITLE": title,
                "RESPONSIBLE_ID": responsible_id,
                "CREATED_BY": creator_id,
                "GROUP_ID": 141,
                "DESCRIPTION": description,
                "DESCRIPTION_IN_BBCODE": "N",
                "TASK_CONTROL": "N",
                "ALLOW_TIME_TRACKING": "Y"
            }
        }

        if contains_results:
            params["fields"]["UF_TASK_WEBDAV_FILES"] = []
        if crm_bindings:
            params["fields"]["UF_CRM_TASK"] = crm_bindings
        if deadline:
            params["fields"]["DEADLINE"] = deadline
        if auditors:
            params["fields"]["AUDITORS"] = auditors
        if accomplices:
            params["fields"]["ACCOMPLICES"] = accomplices

        return bx24.callMethod('tasks.task.add', params=params)
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'

def add_checklist_item(task_id: int | str, title: str, sort_index: int | None = None,
                       is_complete: str = 'N') -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/'

    try:
        bx24 = Bitrix24(url)

        fields: dict = {
            "TITLE": title,
            "IS_COMPLETE": is_complete
        }
        if sort_index is not None:
            fields["SORT_INDEX"] = sort_index

        params: dict = {
            "taskId": int(task_id),
            "fields": fields
        }

        return bx24.callMethod('tasks.task.checklist.add', params=params)
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def update_task(task_id: int, company_name: str) -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/tasks.task.update'

    params = {
        "id": int(task_id),
        "fields": {
            "UF_AUTO_494863223278": company_name
        }
    }

    response = requests.post(url, json=params)

    if response.status_code == 200:
        data = response.json()
        if 'result' in data:
            return data['result']['task']
        else:
            return f"Ошибка в ответе API: {data}"
    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def update_task_crm_bindings(task_id: int | str, crm_bindings: list[str]) -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/tasks.task.update'

    params = {
        "id": int(task_id),
        "fields": {
            "UF_CRM_TASK": crm_bindings,
        }
    }

    response = requests.post(url, json=params)

    if response.status_code == 200:
        data = response.json()
        if 'result' in data:
            return data['result']
        return f"Ошибка в ответе API: {data}"

    return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"



def make_task_dependence(main_task_id: int | str, sub_task_id: int | str) -> dict | str:
    url = f"https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/task.dependence.add"

    params: dict = {
        "taskIdFrom": main_task_id,
        "taskIdTo": sub_task_id,
        "linkType": 1
    }

    response = requests.post(url, json=params)

    if response.status_code == 200:
        data = response.json()

        return data
    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


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
        items = response.json()
        for entity in items:
            print('Entity:', entity)
    except Exception as e:
        print('Request failed:', e)


def get_task_by_id(task_id: str | int) -> dict | str:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/tasks.task.get'


    params = {
        'taskId': task_id,
        'select': [
            'UF_AUTO_494863223278'                                                                          # 'UF_AUTO_494863223278'
        ]
    }
    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']['task']

    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def get_uf_tasks() -> dict | str:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/tasks.task.getFields'

    response = requests.post(webhook_url)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']

    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"




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


def crm_type_lst(token: str, method: str):
    # crm.type.list
    url = f"https://b24.apriori.rest/rest/69/{token}/{method}"
    response = requests.get(url=url)
    items = response.text
    return items



def get_deals_bitrix(token: str, entityTypeId, method, stage_id: str | None = None) -> list:
    url = f"https://b24.apriori.rest/rest/69/{token}/{method}"

    source_items: list[dict] = []
    start = 0
    page_index = 0
    last_response_payload = {}

    while True:
        page_index += 1
        params = {
            "entityTypeId": entityTypeId,
            "start": start,
        }

        response = requests.post(url=url, params=params)
        items = response.json()
        last_response_payload = items
        result = items.get('result') or {}
        page_items = result.get('items') or []

        if not page_items and start == 0:
            print(f"Bitrix: метод {method} не вернул элементов. Ответ: {items}")
            return []

        source_items.extend(page_items)

        next_start = items.get("next")
        if next_start in (None, ""):
            next_start = result.get("next")

        print(
            f"Bitrix pagination: page={page_index}, start={start}, "
            f"items_on_page={len(page_items)}, next={next_start}"
        )

        if next_start in (None, "", -1):
            break

        try:
            next_start = int(next_start)
        except (TypeError, ValueError):
            break

        if next_start <= start:
            break

        start = next_start

    if not source_items:
        print(
            f"Bitrix: метод {method} вернул пустой итоговый список после пагинации. "
            f"Ответ последней страницы: {last_response_payload}"
        )
        return []

    stage_counts: dict[str, int] = {}
    for item in source_items:
        item_stage_id = item.get('stageId') or item.get('stageID') or ""
        if item_stage_id:
            stage_counts[item_stage_id] = stage_counts.get(item_stage_id, 0) + 1

    if not stage_id:
        print(
            f"Bitrix: не задан BITRIX_SOURCE_SMART_PROCESS_STAGE_ID. "
            f"Всего элементов: {len(source_items)}. Распределение по этапам: {stage_counts}"
        )
        return []

    filtered_items = []
    available_stage_ids = set()
    for item in source_items:
        item_stage_id = item.get('stageId') or item.get('stageID') or ""
        if item_stage_id:
            available_stage_ids.add(item_stage_id)
        if item_stage_id == stage_id:
            filtered_items.append(item)

    print(
        f"Bitrix: получено элементов {len(source_items)}, "
        f"после фильтра stageId={stage_id} осталось {len(filtered_items)}. "
        f"Найденные этапы: {sorted(available_stage_ids)}. "
        f"Распределение по этапам: {stage_counts}"
    )
    return filtered_items


def upload_file(file_path, filename, folder_id: int | str):
    """
    Загрузка файла в Битрикс24
    """
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/disk.folder.uploadfile'
    with open(file_path, 'rb') as file:
        file_content_b64 = base64.b64encode(file.read()).decode('utf-8')

    payload = {
        'id': int(folder_id),
        'data': {
            'NAME': filename
        },
        'fileContent': [filename, file_content_b64]
    }
    response = requests.post(webhook_url, json=payload)
    if response.status_code == 200:
        result = response.json()
        if result.get('result'):
            return result['result']
    print(f"Ошибка загрузки в диск Битрикс: {response.status_code}, {response.text}")
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

    uploaded_url = upload_file(file_path=file_path, filename=filename, folder_id=folder_id)
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


def update_smart_process_stage(entity_id: int | str, entity_type_id: int | str, stage_id: str) -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.item.update'

    params: dict = {
        'entityTypeId': int(entity_type_id),
        'id': int(entity_id),
        'fields': {
            'stageId': stage_id
        }
    }

    response = requests.post(url, json=params)

    if response.status_code == 200:
        data = response.json()
        if 'result' in data:
            return data['result']
        return f"Ошибка в ответе API: {data}"

    return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def get_all_companies() -> list | None:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.company.list'

    # 'UF_CRM_1771251801387' - имя базы ЗУП, 'UF_CRM_1771251320632' - почты email_hr

    # select': ["UF_CRM_1771251801387", "UF_CRM_1771251320632", 'UF_CRM_1772109198124', 'ASSIGNED_BY_ID', 'TITLE']

    params = {
        'select': ["UF_CRM_1771251801387", "UF_CRM_1771251320632", 'UF_CRM_1772109198124', 'ASSIGNED_BY_ID', 'TITLE']
    }

    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']

    print('Error:', response.content)
    return None


def get_inn_bitrix(company_id: int) -> str:
    url = f"https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.requisite.list"

    payload = json.dumps({
        "filter": {
        "entityTypeId": 4,
        "ENTITY_ID": company_id
    }})
    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    return response.json()['result'][0]['RQ_INN']


def get_req_id(company_id: int) -> str:
    # Почти та же самая функция, что и get_inn_bitrix, только получаем id банковсокго реквизита
    url = f"https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.requisite.list"

    payload = json.dumps({
        "filter": {
        "entityTypeId": 4,
        "ENTITY_ID": company_id
    }})
    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    return response.json()['result'][0]['ID']

def get_req_link_list():
    url = f"https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.requisite.link.list"


    response = requests.request("POST", url)
    answer = response.json()
    return answer


# crm.company.get
def get_buh_id(company_id: int) -> str | None:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.company.get'

    # 'UF_CRM_1771251801387' - имя базы ЗУП, 'UF_CRM_1771251320632' - почты email_hr

    # select': ["UF_CRM_1771251801387", "UF_CRM_1771251320632", 'UF_CRM_1772109198124', 'ASSIGNED_BY_ID', 'TITLE']

    params = {
        'id': company_id
    }

    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']['UF_CRM_1772546685']

    print('Error:', response.content)
    return None



def get_folder_from_disk(token: str, folder_id: int, method: str) -> list[dict] | str:
    url: str = f'https://b24.apriori.rest/rest/69/{token}/'
    try:
        bx24 = Bitrix24(url)

        params = {
            "id": folder_id
        }

        disk_response = bx24.callMethod(method, params=params)
        return disk_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def get_file_from_disk(token: str, file_id: int) -> dict | str:
    url: str = f'https://b24.apriori.rest/rest/69/{token}/'
    try:
        bx24 = Bitrix24(url)

        params = {
            "id": int(file_id)
        }

        disk_response = bx24.callMethod('disk.folder.get', params=params)
        return disk_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def get_root_folders(token: str) -> list[dict] | str:
    # disk.storage.getlist

    url: str = f'https://b24.apriori.rest/rest/69/{token}/'
    try:
        bx24 = Bitrix24(url)


        disk_response = bx24.callMethod('disk.storage.getlist')

        for el in disk_response:
            if el['NAME'] == 'Априори':

                return el
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def save_file_from_url(url: str, path_to_save, filename: str) -> str:
    response = requests.get(url, stream=True)

    if response.status_code == 200:
        with open(f'{path_to_save}/{filename}', 'wb') as f:
            for chunk in response.iter_content(1024):
                f.write(chunk)

        return f'{path_to_save}/{filename}'
    else:
        return f'Ошибка во время скачивания файла: {response.status_code}'



def get_token_report_file() -> str | None:
    bitrix_folder: list[dict] | str = get_folder_from_disk(token=BITRIX_TOKEN, folder_id=193,       # 193 - папка "Общее"
                                                           method='disk.folder.getchildren')

    for item in bitrix_folder:
        if item['NAME'].lower() == 'токены отчетность.xlsx':
            url_file: str | int = item['DOWNLOAD_URL']
            downloaded_file: str = save_file_from_url(url=url_file,
                                                      path_to_save='extension_ecp/report_tokens', filename=item['NAME'])
            return downloaded_file

    return None


def attach_file_to_task(task_id: int | str, file_id: int | str) -> dict | str:
    url: str = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/'
    try:
        bx24 = Bitrix24(url)

        params: dict = {
            "taskId": task_id,
            "fileId": file_id
        }

        disk_response = bx24.callMethod('tasks.task.files.attach', params=params)
        return disk_response
    except BitrixError as e:
        return f'Ошибка во время отправки запроса {e}'


def upload_and_attach_file_to_task(task_id: int | str, file_path: str, folder_id: int | str) -> dict | str:
    original_filename = os.path.basename(file_path).replace('output_', '')
    filename = f"task_{task_id}_{original_filename}"
    uploaded_file = upload_file(file_path=file_path, filename=filename, folder_id=folder_id)

    if not uploaded_file or 'ID' not in uploaded_file:
        return f'Не удалось загрузить файл {filename} в диск Битрикс'

    return attach_file_to_task(task_id=task_id, file_id=uploaded_file['ID'])


def add_checklist_item(task_id: int | str, title: str, sort_index: int | None = None,
                       is_complete: str = 'N') -> dict | str:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/task.checklistitem.add.json'

    params: dict = {
        "TASKID": int(task_id),
        "FIELDS": {
            "TITLE": title
        }
    }
    if is_complete in ('Y', 'N'):
        params["FIELDS"]["IS_COMPLETE"] = is_complete
    if sort_index is not None:
        params["FIELDS"]["SORT_INDEX"] = sort_index

    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        data = response.json()
        if 'result' in data:
            return data['result']
        return f"Ошибка в ответе API: {data}"
    return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def get_uf_lst() -> dict | str:
    # user.userfield.list
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.deal.userfield.list'

    response = requests.post(webhook_url)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']

    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def get_bank_req(req_id: int | str) -> dict | str | None:
    # crm.requisite.bankdetail.list
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.requisite.bankdetail.list'

    response = requests.post(webhook_url)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            req_lst = result['result']

            for req in req_lst:
                if req['ID'] == int(req_id) or req['ID'] == str(req_id):
                    return req

         # return result['result']
        return None
    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def company_info(company_id: int):
    # crm.company.get
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.company.get'

    params: dict = {
        "id": company_id
    }

    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']

    else:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"


def get_deal_info(deal_id: int | str) -> dict | str:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.deal.get'

    params: dict = {
        "id": deal_id
    }

    response = requests.post(webhook_url, json=params)

    if response.status_code == 200:
        result = response.json()

        if 'result' in result:
            return result['result']

        return f"Ошибка в ответе API: {result}"

    return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"

def _bitrix_webhook_url(method: str) -> str:
    return f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/{method}'


def _bitrix_post(method: str, payload: dict | None = None, timeout: int = 60) -> dict | str:
    response = requests.post(_bitrix_webhook_url(method), json=payload or {}, timeout=timeout)

    if response.status_code != 200:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"

    data = response.json()
    if "error" in data:
        return f"Bitrix API error: {data['error']} {data.get('error_description', '')}".strip()

    return data


def get_all_smart_process_items(entity_type_id: int | str, select: list[str] | None = None) -> list[dict] | str:
    items: list[dict] = []
    start = 0

    while True:
        payload: dict = {
            "entityTypeId": int(entity_type_id),
            "start": start,
        }
        if select:
            payload["select"] = select

        response = _bitrix_post("crm.item.list", payload)
        if isinstance(response, str):
            return response

        result = response.get("result") or {}
        page_items = result.get("items") or []
        items.extend(page_items)

        next_start = response.get("next")
        if next_start in (None, ""):
            next_start = result.get("next")
        if next_start in (None, "", -1):
            break

        try:
            next_start = int(next_start)
        except (TypeError, ValueError):
            break

        if next_start <= start:
            break
        start = next_start

    return items


def get_smart_process_item(entity_type_id: int | str, item_id: int | str) -> dict | str:
    response = _bitrix_post(
        "crm.item.get",
        {
            "entityTypeId": int(entity_type_id),
            "id": int(item_id),
        },
    )
    if isinstance(response, str):
        return response

    result = response.get("result") or {}
    item = result.get("item")
    if isinstance(item, dict):
        return item
    return f"Ошибка в ответе API: {response}"


def get_smart_process_fields(entity_type_id: int | str) -> dict | str:
    response = _bitrix_post(
        "crm.item.fields",
        {
            "entityTypeId": int(entity_type_id),
        },
    )
    if isinstance(response, str):
        return response
    return response.get("result") or {}


def update_smart_process_fields(entity_id: int | str, entity_type_id: int | str, fields: dict) -> dict | str:
    response = _bitrix_post(
        "crm.item.update",
        {
            "entityTypeId": int(entity_type_id),
            "id": int(entity_id),
            "fields": fields,
        },
    )
    if isinstance(response, str):
        return response
    return response.get("result") or response


def get_disk_file_info(file_id: int | str) -> dict | str:
    response = _bitrix_post("disk.file.get", {"id": int(file_id)})
    if isinstance(response, str):
        return response
    return response.get("result") or {}


def get_disk_attached_object_info(attached_object_id: int | str) -> dict | str:
    response = _bitrix_post("disk.attachedObject.get", {"id": int(attached_object_id)})
    if isinstance(response, str):
        return response
    return response.get("result") or {}


def extract_bitrix_disk_folder_id(folder_value) -> int | None:
    raw_value = str(folder_value or "").strip()
    if not raw_value:
        return None

    if raw_value.isdigit():
        return int(raw_value)

    match = re.search(r"[?&]folderId=(\d+)\b", raw_value, flags=re.IGNORECASE)
    if match:
        return int(match.group(1))

    match = re.search(r"\bfolderId[=/](\d+)\b", raw_value, flags=re.IGNORECASE)
    if match:
        return int(match.group(1))

    return None


def get_disk_storages() -> list[dict] | str:
    storages: list[dict] = []
    start = 0

    while True:
        payload = {"start": start} if start else {}
        response = _bitrix_post("disk.storage.getlist", payload)
        if isinstance(response, str):
            return response

        result = response.get("result") or []
        if not isinstance(result, list):
            return []
        storages.extend(result)

        next_start = response.get("next")
        if next_start in (None, "", -1):
            break
        try:
            next_start = int(next_start)
        except (TypeError, ValueError):
            break
        if next_start <= start:
            break
        start = next_start

    return storages


def get_shared_disk_root_folder_id() -> int | None:
    storages = get_disk_storages()
    if not isinstance(storages, list):
        return None

    for storage in storages:
        if not isinstance(storage, dict):
            continue
        if storage.get("ENTITY_TYPE") == "common" and storage.get("ENTITY_ID") == "shared_files_s1":
            root_object_id = storage.get("ROOT_OBJECT_ID")
            if str(root_object_id or "").isdigit():
                return int(root_object_id)
        if storage.get("NAME") == "Общий диск":
            root_object_id = storage.get("ROOT_OBJECT_ID")
            if str(root_object_id or "").isdigit():
                return int(root_object_id)

    return None


def resolve_bitrix_shared_path_folder_id(folder_value) -> int | None:
    raw_value = str(folder_value or "").strip()
    if not raw_value:
        return None

    marker = "/docs/shared/path/"
    marker_index = raw_value.find(marker)
    if marker_index < 0:
        return None

    relative_path = raw_value[marker_index + len(marker):]
    relative_path = relative_path.split("#", 1)[0].strip()
    if relative_path.startswith("?"):
        relative_path = relative_path[1:]
    path_segments = [unquote(segment).strip() for segment in relative_path.split("/") if segment.strip()]
    if not path_segments:
        return None

    current_folder_id = get_shared_disk_root_folder_id()
    if not current_folder_id:
        return None

    for segment in path_segments:
        children = get_disk_folder_children(current_folder_id)
        if not isinstance(children, list):
            return None

        next_folder_id = None
        for child in children:
            if not isinstance(child, dict):
                continue
            child_name = str(child.get("NAME") or child.get("name") or "").strip()
            child_type = str(child.get("TYPE") or child.get("type") or "").lower()
            if child_name == segment and child_type in {"folder", "storage"}:
                child_id = child.get("ID") or child.get("id")
                if str(child_id or "").isdigit():
                    next_folder_id = int(child_id)
                    break

        if not next_folder_id:
            return None
        current_folder_id = next_folder_id

    return current_folder_id


def resolve_bitrix_disk_folder_id(folder_value) -> int | None:
    folder_id = extract_bitrix_disk_folder_id(folder_value)
    if folder_id:
        return folder_id
    return resolve_bitrix_shared_path_folder_id(folder_value)


def get_disk_folder_children(folder_id: int | str) -> list[dict] | str:
    response = _bitrix_post("disk.folder.getchildren", {"id": int(folder_id)})
    if isinstance(response, str):
        return response

    result = response.get("result")
    if isinstance(result, list):
        return result
    if isinstance(result, dict):
        items = result.get("items")
        if isinstance(items, list):
            return items
    return []


def create_bitrix_disk_subfolder(parent_folder_id: int | str, folder_name: str) -> dict | str:
    response = _bitrix_post(
        "disk.folder.addsubfolder",
        {
            "id": int(parent_folder_id),
            "data": {
                "NAME": str(folder_name or "").strip(),
            },
        },
    )
    if isinstance(response, str):
        return response

    result = response.get("result")
    if isinstance(result, dict):
        return result
    return f"Bitrix API error: unexpected addsubfolder response {response}"


def ensure_bitrix_disk_subfolder(parent_folder_id: int | str, folder_name: str) -> dict | str:
    normalized_name = str(folder_name or "").strip()
    if not normalized_name:
        return "Bitrix folder name is empty"

    children = get_disk_folder_children(parent_folder_id)
    if not isinstance(children, list):
        return children

    for child in children:
        if not isinstance(child, dict):
            continue
        child_name = str(child.get("NAME") or child.get("name") or "").strip()
        child_type = str(child.get("TYPE") or child.get("type") or "").lower()
        if child_name == normalized_name and child_type in {"folder", "storage"}:
            return child

    return create_bitrix_disk_subfolder(parent_folder_id, normalized_name)


def _make_authorized_download_url(download_url: str) -> str:
    raw_url = str(download_url or "").strip()
    if not raw_url:
        return raw_url

    absolute_url = raw_url if raw_url.startswith("http") else urljoin(BITRIX_PORTAL_BASE_URL, raw_url)
    parsed_url = urlparse(absolute_url)
    query_params = dict(parse_qsl(parsed_url.query, keep_blank_values=True))
    query_params.setdefault("auth", BITRIX_TOKEN or "")
    print(query_params)
    return urlunparse(
        (
            parsed_url.scheme,
            parsed_url.netloc,
            parsed_url.path,
            parsed_url.params,
            urlencode(query_params),
            parsed_url.fragment,
        )
    )


def _resolve_destination_path(destination_path: str, response: requests.Response) -> str:
    current_path = destination_path
    if os.path.splitext(current_path)[1].lower() != ".bin":
        return current_path

    content_disposition = response.headers.get("Content-Disposition", "")
    filename_match = re.search(r"filename\*=UTF-8''([^;]+)", content_disposition, flags=re.IGNORECASE)
    if filename_match:
        filename = unquote(filename_match.group(1)).strip().strip('"')
        extension = os.path.splitext(filename)[1]
        if extension:
            return f"{os.path.splitext(current_path)[0]}{extension}"

    filename_match = re.search(r'filename="?([^";]+)"?', content_disposition, flags=re.IGNORECASE)
    if filename_match:
        filename = filename_match.group(1).strip()
        extension = os.path.splitext(filename)[1]
        if extension:
            return f"{os.path.splitext(current_path)[0]}{extension}"

    content_type = (response.headers.get("Content-Type", "") or "").split(";")[0].strip().lower()
    guessed_extension = mimetypes.guess_extension(content_type) if content_type else None
    if guessed_extension:
        return f"{os.path.splitext(current_path)[0]}{guessed_extension}"

    return current_path


def download_file_from_url(download_url: str, destination_path: str) -> str | None:
    response = requests.get(download_url, timeout=120)
    if response.status_code != 200:
        return f"Ошибка скачивания файла: {response.status_code}, {response.text}"

    final_destination_path = _resolve_destination_path(destination_path, response)

    os.makedirs(os.path.dirname(final_destination_path), exist_ok=True)
    with open(final_destination_path, "wb") as file:
        file.write(response.content)
    return None


def download_file_from_url(download_url: str, destination_path: str) -> str | None:
    candidate_urls: list[str] = []
    raw_url = str(download_url or "").strip()
    if raw_url:
        candidate_urls.append(raw_url if raw_url.startswith("http") else urljoin(BITRIX_PORTAL_BASE_URL, raw_url))

    authorized_url = _make_authorized_download_url(download_url)
    if authorized_url and authorized_url not in candidate_urls:
        candidate_urls.append(authorized_url)

    response = None
    last_error: str | None = None

    for candidate_url in candidate_urls:
        response = requests.get(candidate_url, timeout=120)
        if response.status_code == 200:
            break
        last_error = f"Ошибка скачивания файла: {response.status_code}, {response.text}"
    else:
        return last_error or "Ошибка скачивания файла: не удалось выполнить запрос"

    final_destination_path = _resolve_destination_path(destination_path, response)
    if final_destination_path != destination_path and os.path.exists(destination_path):
        os.remove(destination_path)

    os.makedirs(os.path.dirname(final_destination_path), exist_ok=True)
    with open(final_destination_path, "wb") as file:
        file.write(response.content)
    return None


def get_company_userfield_code_by_title(field_title: str) -> str | None:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/crm.company.userfield.list'
    response = requests.post(webhook_url)

    if response.status_code != 200:
        return None

    result = response.json()
    userfields = result.get('result') or []
    normalized_title = str(field_title or '').strip().lower()

    partial_match: str | None = None

    for userfield in userfields:
        field_name = userfield.get('FIELD_NAME')
        labels = (
            userfield.get('EDIT_FORM_LABEL'),
            userfield.get('LIST_COLUMN_LABEL'),
            userfield.get('LIST_FILTER_LABEL'),
            userfield.get('MANDATORY_ERROR_MESSAGE'),
        )
        for label in labels:
            values = label.values() if isinstance(label, dict) else [label]
            for value in values:
                normalized_value = str(value or '').strip().lower()
                if normalized_value == normalized_title:
                    return field_name
                if normalized_title and normalized_title in normalized_value and not partial_match:
                    partial_match = field_name

    return partial_match


def get_department_info(department_id: int | str) -> dict | str | None:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/department.get'
    params = {
        "ID": int(department_id)
    }
    response = requests.post(webhook_url, json=params)

    if response.status_code != 200:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"

    result = response.json()
    items = result.get("result") or []
    if not items:
        return None
    return items[0]


def get_department_users(department_id: int | str) -> list[dict] | str:
    webhook_url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/user.get'
    params = {
        "FILTER": {
            "UF_DEPARTMENT": int(department_id),
            "ACTIVE": True,
        }
    }
    response = requests.post(webhook_url, json=params)

    if response.status_code != 200:
        return f"HTTP ошибка: {response.status_code}, сообщение: {response.text}"

    result = response.json()
    return result.get("result") or []



def get_contractors_reqs() -> dict:
    req_data: dict = {
        'ИНН_исполнителя': '',
        'Рас_счёт_исполнителя': '',
        'БИК_банка_исполнителя': '',
        'Корр_счет_исполнителя': ''
    }
    company_lst = get_all_companies()

    for company in company_lst:
        if company['ID'] == '21' or company['ID'] == 21:        # ID 21 - ИП ДУДКИНА ИННА НИКОЛАЕВНА
            inn = get_inn_bitrix(company_id=company['ID'])
            req_id = get_req_id(company_id=company['ID'])

            needed_req = get_bank_req(req_id=req_id)

            req_data['ИНН_исполнителя'] = inn
            req_data['Рас_счёт_исполнителя'] = needed_req['RQ_ACC_NUM']
            req_data['БИК_банка_исполнителя'] = needed_req['RQ_BIK']
            req_data['Корр_счет_исполнителя'] = needed_req['RQ_COR_ACC_NUM']

            return req_data


if __name__ == '__main__':
    company_lst = get_all_companies()

    for company in company_lst:
        if company['ID'] == '21' or company['ID'] == 21:
            inn = get_inn_bitrix(company_id=company['ID'])
            req_id = get_req_id(company_id=company['ID'])

            needed_req = get_bank_req(req_id=req_id)
            print(needed_req)
            pass

    # 'ufAuto713946616750'
    current_task = get_task_by_id(task_id=1418)


    pass


    # created_task: dict | str = add_task(creator_id=69, responsible_id=69,
    #                                     name_company='Тестовая компания', ecp_date='01.01.2026')
    # pass
    #
    #
    #
    # created_task: dict | str = add_task(creator_id=69, responsible_id=69)
    # task_id = created_task['id']
    # task_id = 1309
    #
    # uploaded_url_act = upload_file(file_path='../extension_ecp/output_result/output_Пример акта.docx', filename='output_Пример акта.docx', folder_id=24342)
    # uploaded_file_act = upload_to_uploadurl(upload_url=uploaded_url_act, field_name='file', file_path='../extension_ecp/output_result/output_Пример акта.docx')
    # uploaded_file_id_act = uploaded_file_act['result']['ID']
    #
    # attached_file_act = attach_file_to_task(task_id=task_id, file_id=uploaded_file_id_act)

    # root_folders = get_root_folders(token=BITRIX_TOKEN)

    # toke_rep_file = get_token_report_file()           24342 id папки "ПРОДЛЕНИЕ ЭЦП (тест)"

    # folder = get_folder_from_disk(token=BITRIX_TOKEN, folder_id=193,
    #                      method='disk.folder.getchildren')
    #
    #
    # get_type_lst = crm_type_lst(token=BITRIX_TOKEN, method='crm.type.list')
    #
    # bitrix_deals: list[dict] | str = get_deals_bitrix(token=BITRIX_TOKEN, entityTypeId=1042,
    #                                                   method='crm.item.list')
    # needed_deal = ''
    #
    # for deal in bitrix_deals:
    #     if deal['title'] == 'Приём на работу #88':
    #         needed_deal = deal
    #
    # file_id = 16235
    # add_file =  add_file_to_smart_process(token=BITRIX_TOKEN, entity_id=needed_deal['id'], entity_type_id=1042,
    #                                       file_id=file_id)
