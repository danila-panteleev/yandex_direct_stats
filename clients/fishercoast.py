import data_handler as dh
import gspread
import datetime as dt


def main():
    # Yandex Direct settings
    ACCESS_TOKEN = 'AgAAAAAmQiFDAAahJf8XZZD3pE5SuY-PYLmbd_I'
    YANDEX_LOGIN = 'e-16700996'
    CLIENT_NAME = 'Рыбачий берег'
    DAYS_RANGE = 7
    date_from = str(dt.date.today() - dt.timedelta(days=DAYS_RANGE))
    date_to = str(dt.date.today() - dt.timedelta(days=1))

    yandex_report_campaigns = dh.get_report(login=YANDEX_LOGIN,
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
                                            report_name='FISHERCOAST_CAMPAIGN_REPORT',
                                            date_from=date_from,
                                            date_to=date_to)
    yandex_report_adgroups = dh.get_report(login=YANDEX_LOGIN,
                                           token=ACCESS_TOKEN,
                                           fieldnames=['CampaignName',
                                                       'AdGroupName',
                                                       'Impressions',
                                                       'Clicks',
                                                       'Ctr',
                                                       'AvgCpc',
                                                       'Cost',
                                                       'Conversions',
                                                       'ConversionRate',
                                                       'CostPerConversion'],
                                           order_by='Date',
                                           sort_order='ASCENDING',
                                           report_type='ADGROUP_PERFORMANCE_REPORT',
                                           report_name='FISHERCOAST_ADGROUP_REPORT',
                                           date_from=date_from,
                                           date_to=date_to)

    gc = gspread.oauth()
    sh = gc.open("Оптметаллобаза / Контекстная реклама")
    worksheet_campaigns = sh.worksheet("Яндекс Директ по кампаниям")
    worksheet_daily = sh.worksheet("Яндекс Директ по группам объявлений")

    headline_campaigns = ["Кампания",
                          "Показы",
                          "Клики",
                          "CTR, %",
                          "CPC, ₽",
                          "Стоимость, ₽",
                          "Конверсии",
                          "Конверсия, %",
                          "Цена за конверсию, ₽"]

    headline_adgroup = ["Кампания",
                        "Группа объявлений",
                        "Показы",
                        "Клики",
                        "CTR, %",
                        "CPC, ₽",
                        "Стоимость, ₽",
                        "Конверсии",
                        "Конверсия, %",
                        "Цена за конверсию, ₽"]

    dh.add_report_date_to_google_sheet(worksheet_campaigns,
                                       DAYS_RANGE,
                                       'A',
                                       'I')

    headline_campaigns = ["Кампания",
                          "Показы",
                          "Клики",
                          "CTR, %",
                          "CPC, ₽",
                          "Стоимость, ₽",
                          "Конверсии",
                          "Конверсия, %",
                          "Цена за конверсию, ₽"]
    dh.add_report_headline_to_google_sheets(worksheet_campaigns,
                                            headline_campaigns)

    worksheet_campaigns.append_rows(yandex_report_campaigns[1:])

    total_values = dh.values_for_total_row(yandex_report_campaigns)
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
    worksheet_campaigns.append_row(total_row)
    dh.format_summary_row_in_google_sheets(worksheet_campaigns,
                                           total_row,
                                           'A',
                                           'I')

    dh.format_last_added_report_in_google_sheets(worksheet_campaigns, 'Показы')

    worksheet_daily.append_rows(yandex_report_daily[1:])
    dh.format_last_added_report_in_google_sheets(worksheet_daily, 'Показы')
