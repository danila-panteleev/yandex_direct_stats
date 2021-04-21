import data_handler as dh
import gspread
import datetime as dt

DAYS_RANGE = 7
date_from = str(dt.date.today() - dt.timedelta(days=DAYS_RANGE))
date_to = str(dt.date.today() - dt.timedelta(days=1))


def main(date_from=date_from, date_to=date_to):
    # Yandex Direct settings
    ACCESS_TOKEN = 'AgAAAAAmQiFDAAahJf8XZZD3pE5SuY-PYLmbd_I'
    YANDEX_LOGIN = 'mcurb4tova'
    CLIENT_NAME = 'Бронепоезд'


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
                                  order_by='CampaignName',
                                  report_type='CAMPAIGN_PERFORMANCE_REPORT',
                                  goals=['67653724', '168502726', '168502954', '168503038'],
                                  date_from=date_from,
                                  date_to=date_to)

    gc = gspread.oauth()
    sh = gc.open("b-poezd.ru Контекстная реклама")
    worksheet = sh.worksheet("Январь по кампаниям")
    dh.add_report_date_to_google_sheet(worksheet=worksheet,
                                       days=DAYS_RANGE,
                                       start_column='A',
                                       end_column='I')

    headline_columns = ["Кампания",
                        "Показы",
                        "Клики",
                        "CTR, %",
                        "CPC, ₽",
                        "Стоимость, ₽",
                        "Конверсии",
                        "Конверсия, %",
                        "Цена за конверсию, ₽"]
    dh.add_report_headline_to_google_sheets(worksheet, headline_columns)

    worksheet.append_rows(yandex_report[1:])

    impressions_sum = 0
    clicks_sum = 0
    summary_ctr = 0
    summary_cpc = 0
    summary_conversions = 0
    summary_conversion_rate = 0
    summary_conversion_cost = 0
    summary_cost = 0

    for i in range(1, len(yandex_report)):
        impressions_sum += int(yandex_report[i][1])
        clicks_sum += int(yandex_report[i][2])
        if yandex_report[i][5] != '--':
            summary_conversions += int(yandex_report[i][5])
        summary_cost += float(yandex_report[i][8])
    summary_ctr = clicks_sum / impressions_sum
    summary_cpc = summary_cost / clicks_sum
    summary_conversion_rate = summary_conversions / clicks_sum
    summary_conversion_cost = summary_cost / summary_conversions

    summary_row = ['ИТОГО',
                   str(impressions_sum),
                   str(clicks_sum),
                   f'{summary_ctr:.2f}',
                   f'{summary_cpc:.2f}',
                   str(summary_conversions),
                   f'{summary_conversion_rate:.2f}',
                   f'{summary_conversion_cost:.2f}',
                   f'{summary_cost:.2f}']

    worksheet.append_row(summary_row)

    report_headline_row_index = worksheet.findall('Кампания')[-1].row
    last_summary_row_index = worksheet.findall('ИТОГО')[-1].row
    worksheet.format(f'A{last_summary_row_index}:I{last_summary_row_index}',
                     {'textFormat': {'bold': True}})

    worksheet.format(f'B{report_headline_row_index}:I{last_summary_row_index}',
                     {'horizontalAlignment': 'CENTER'})
