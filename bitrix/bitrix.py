import os
import requests
import json
import base64
import datetime

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


def add_task(creator_id: int | str, responsible_id: int | str, group_id: int| str,
             title: str, description: str, company_name: str,
             contains_results: bool = True, contains_gantt_links: bool = True) -> dict | str:
    url = f'https://b24.apriori.rest/rest/69/{BITRIX_TOKEN}/'

    try:
        bx24 = Bitrix24(url)

        params: dict = {
            "fields": {
                "TITLE": title,
                "RESPONSIBLE_ID": responsible_id,
                "CREATED_BY": creator_id,
                "GROUP_ID": group_id,
                "DESCRIPTION": description,
                "DESCRIPTION_IN_BBCODE": "N",
                "TASK_CONTROL": "N"
            }
        }

        if contains_results:
            params["fields"]["UF_TASK_WEBDAV_FILES"] = []

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
        if item['stageId'] == 'DT1042_5:UC_B608AW':
            filtered_items.append(item)
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
