import os
import requests

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


if __name__ == '__main__':
    task_list = get_task_list(token=BITRIX_TOKEN)
    # created_task = add_task(token=BITRIX_TOKEN, creator_id=69, responsible_id=69)
    # network_lst = get_social_network_lst(token=BITRIX_TOKEN, social_network_id=141)
    pass
