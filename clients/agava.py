import yandex_direct_stats.data_handler as dh
import gspread
import datetime as dt
import argparse

ACCESS_TOKEN = 'AgAAAAAmQiFDAAahJf8XZZD3pE5SuY-PYLmbd_I'
YANDEX_LOGIN = 'e-16304919'
SPREADSHEET = "agava-ltd.ru Контекстная реклама"
WORKSHEET = "Яндекс Директ"


def main(days_range):
    date_from = str(dt.date.today() - dt.timedelta(days=days_range))
    date_to = str(dt.date.today() - dt.timedelta(days=1))
    yandex_report = dh.get_report(login=YANDEX_LOGIN,
                                  token=ACCESS_TOKEN,
                                  fieldnames=['CampaignName',
                                              'Impressions',
                                              'Clicks',
                                              'Ctr',
                                              'AvgCpc',
                                              'Cost',
                                              'Conversions',
                                              'ConversionRate',
                                              'CostPerConversion'],
                                  order_by='Cost',
                                  report_type='CAMPAIGN_PERFORMANCE_REPORT',
                                  report_name='AGAVA_CAMPAIGN_REPORT',
                                  date_from=date_from,
                                  date_to=date_to)

    gc = gspread.oauth()
    sh = gc.open(SPREADSHEET)
    worksheet = sh.worksheet(WORKSHEET)

    dh.add_report_date_to_google_sheet(worksheet=worksheet,
                                       days=days_range,
                                       start_column='A',
                                       end_column='I')

    headline_columns = ["Кампания",
                        "Показы",
                        "Клики",
                        "CTR, %",
                        "CPC, ₽",
                        "Стоимость, ₽",
                        "Переход на страницу «Контакты»",
                        "Конверсия, %",
                        "Цена за переход на «Контакты», ₽"]
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
        total_values['Cost'],
        total_values['Conversions'],
        total_values['ConversionRate'],
        total_values['CostPerConversion']
    ]
    worksheet.append_row(total_row)
    dh.format_summary_row_in_google_sheets(worksheet,
                                           total_row,
                                           start_column='A',
                                           end_column='I')

    dh.format_last_added_report_in_google_sheets(worksheet,
                                                 prop='Показы')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--days', action='store_const')
    args = parser.parse_args()
    
    days_range = args.days

    main(days_range)