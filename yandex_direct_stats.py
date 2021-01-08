import data_handler as dh
import email_handler as eh

import yaml
import os
import sys


script_path = os.path.dirname(sys.argv[0])

with open(os.path.join(script_path, r'config.yml'), 'r', encoding='utf-8') as stream:
    data_loaded = yaml.safe_load(stream)

# Yandex Direct settings
ACCESS_TOKEN = data_loaded['token']
YANDEX_LOGIN = data_loaded['yandex_login']
CLIENT_NAME = data_loaded['client_name']
FIELDS = data_loaded['fields']
DATE_RANGE_API = data_loaded['date_range_api']
DATE_RANGE_INT = data_loaded['date_range_int']
REPORT_TYPE = data_loaded['report_type']
ORDER_BY = data_loaded['order_by']


# email settings
EMAIL_PASSWORD = data_loaded['email_password']
EMAIL_USER = data_loaded['email_user']

# email body settings
RECEIVER = data_loaded['receiver']
body = data_loaded['email_body']
date = dh.date_range_exclude_today(date_range=DATE_RANGE_INT)
filename = dh.report_filename(client_name=CLIENT_NAME, date_range=DATE_RANGE_INT)
report_file_path = os.path.join(script_path, rf'{filename}')

if __name__ == '__main__':
    dh.report_wrapper(login=YANDEX_LOGIN,
                      token=ACCESS_TOKEN,
                      fieldnames=FIELDS,
                      client_name=CLIENT_NAME)
    eh.email_wrapper(user=EMAIL_USER,
                     password=EMAIL_PASSWORD,
                     receiver=RECEIVER,
                     body=body,
                     attachments=report_file_path)