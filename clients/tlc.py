import data_handler as dh
import gspread
import datetime as dt


DAYS_RANGE = 7
date_from = str(dt.date.today() - dt.timedelta(days=DAYS_RANGE))
date_to = str(dt.date.today() - dt.timedelta(days=1))


def main(date_from=date_from, date_to=date_to):
    # Yandex Direct settings
    ACCESS_TOKEN = 'AgAAAAAmQiFDAAahJf8XZZD3pE5SuY-PYLmbd_I'
    YANDEX_LOGIN = 'e-16244691'
    CLIENT_NAME = 'ТЛК'


    yandex_report = dh.get_report(login=YANDEX_LOGIN,
                                  token=ACCESS_TOKEN,
                                  fieldnames=['CampaignName',
                                              'Impressions',
                                              'Clicks',
                                              'Ctr',
                                              'AvgCpc',
                                              'Cost'],
                                  order_by='Cost',
                                  report_type='CAMPAIGN_PERFORMANCE_REPORT',
                                  report_name='TLC_CAMPAIGN_REPORT',
                                  date_from=date_from,
                                  date_to=date_to)

    gc = gspread.oauth()
    sh = gc.open("tlc-wood.ru Контекстная реклама")
    worksheet = sh.worksheet("Яндекс Директ")

    dh.add_report_date_to_google_sheet(worksheet=worksheet,
                                       days=DAYS_RANGE,
                                       start_column='A',
                                       end_column='F')

    headline_columns = ["Кампания",
                        "Показы",
                        "Клики",
                        "CTR, %",
                        "CPC, ₽",
                        "Стоимость, ₽"]
    dh.add_report_headline_to_google_sheets(worksheet,
                                            columns=headline_columns)

    worksheet.append_rows(yandex_report[1:])

    total_values = dh.values_for_total_row(yandex_report)
    total_row = [
        'ИТОГО',
        total_values['Impressions'],
        total_values['Clicks'],
        total_values['Ctr'],
        total_values['AvgCpc'],
        total_values['Cost']
    ]
    worksheet.append_row(total_row)
    dh.format_summary_row_in_google_sheets(worksheet,
                                           total_row,
                                           start_column='A',
                                           end_column='F')

    dh.format_last_added_report_in_google_sheets(worksheet,
                                                 prop='Показы')
