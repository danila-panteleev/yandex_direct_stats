from yandex_direct_stats.yandex_direct_stats import data_handler as dh
import datetime as dt
import gspread


def main():
    # Yandex Direct settings
    ACCESS_TOKEN = 'AgAAAAAmQiFDAAahJf8XZZD3pE5SuY-PYLmbd_I'
    YANDEX_LOGIN = 'e-16304928'
    DATE_RANGE_API = 'CUSTOM_DATE'
    DAYS_RANGE = 7

    date_from = str(dt.date.today() - dt.timedelta(days=DAYS_RANGE))
    date_to = str(dt.date.today() - dt.timedelta(days=1))

    report_date_range = [dh.date_range_exclude_today(DAYS_RANGE)]

    gc = gspread.oauth()
    sh = gc.open("auroracomp.ru Контекстная реклама")
    worksheet_summary = sh.worksheet("Общая статистика Яндекс")
    worksheet_campaigns = sh.worksheet("По кампаниям Я.Директ")
    worksheet_service_keywords = sh.worksheet("Я.Директ Услуги по фразам")
    services_total_report = dh.get_report(login=YANDEX_LOGIN,
                                          token=ACCESS_TOKEN,
                                          fieldnames=['Impressions',
                                                      'Clicks',
                                                      'Ctr',
                                                      'AvgCpc',
                                                      'Cost',
                                                      'Conversions',
                                                      'ConversionRate',
                                                      'CostPerConversion'],
                                          order_by='Impressions',
                                          report_type='CUSTOM_REPORT',
                                          report_name="SERVICES_TOTAL_REPORT",
                                          date_from=date_from,
                                          date_to=date_to,
                                          filter_item=[{
                                              "Field": "CampaignId",
                                              "Operator": "IN",
                                              "Values": ["57893145",
                                                         "57893158"]
                                          }])
    shop_total_report = dh.get_report(login=YANDEX_LOGIN,
                                      token=ACCESS_TOKEN,
                                      report_name="SHOP_TOTAL_REPORT",
                                      fieldnames=['Impressions',
                                                  'Clicks',
                                                  'Ctr',
                                                  'AvgCpc',
                                                  'Cost',
                                                  'Conversions',
                                                  'ConversionRate',
                                                  'CostPerConversion'],
                                      order_by='Impressions',
                                      report_type='CUSTOM_REPORT',
                                      date_from=date_from,
                                      date_to=date_to,
                                      filter_item=[{
                                          "Field": "CampaignId",
                                          "Operator": "IN",
                                          "Values": ["57923409",
                                                     "57893152",
                                                     "57893149",
                                                     "41292209"]
                                      }])
    campaigns_report = dh.get_report(login=YANDEX_LOGIN,
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
                                     report_name="CAMPAIGNS_REPORT",
                                     date_from=date_from,
                                     date_to=date_to
                                     )
    keywords_services_report = dh.get_report(login=YANDEX_LOGIN,
                                             token=ACCESS_TOKEN,
                                             fieldnames=['Criterion',
                                                         'Impressions',
                                                         'Clicks',
                                                         'Ctr',
                                                         'AvgCpc',
                                                         'Cost',
                                                         'Conversions',
                                                         'ConversionRate',
                                                         'CostPerConversion'],
                                             order_by='Cost',
                                             report_type='CUSTOM_REPORT',
                                             report_name="KEYWORDS_SERVICES_TOTAL_REPORT",
                                             date_from=date_from,
                                             date_to=date_to,
                                             filter_item=[{
                                                 "Field": "CampaignId",
                                                 "Operator": "IN",
                                                 "Values": ["57893145"]
                                             }])
    summary_report = report_date_range.copy()
    for i in range(len(services_total_report[1])):
        summary_report.append(services_total_report[1][i])
        try:
            summary_report.append(shop_total_report[1][i])
        except IndexError:
            ...
    worksheet_summary.append_row(summary_report)
    summary_report_row_index = worksheet_summary.findall(report_date_range[0])[-1].row
    worksheet_summary.format(f'A{summary_report_row_index}:Q{summary_report_row_index}',
                             {
                                 'horizontalAlignment': 'CENTER',
                                 "borders": {
                                     "bottom": {
                                         "style": "SOLID",
                                         "width": 1,
                                     },
                                     "right": {
                                         "style": "SOLID",
                                         "width": 1,
                                     }
                                 }
                             })

    dh.add_report_date_to_google_sheet(worksheet=worksheet_campaigns,
                                       days=DAYS_RANGE,
                                       start_column='A',
                                       end_column='I')
    campaigns_report_headline = ["Кампания",
                                 "Показы",
                                 "Клики",
                                 "CTR, %",
                                 "Цена за клик, ₽",
                                 "Стоимость",
                                 "Конверсии",
                                 "Конверсия",
                                 "Цена за конв., ₽"]
    dh.add_report_headline_to_google_sheets(worksheet=worksheet_campaigns,
                                            columns=campaigns_report_headline)
    worksheet_campaigns.append_rows(campaigns_report[1:])
    total_values = dh.values_for_total_row(campaigns_report)
    campaigns_report_total_row = [
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
    worksheet_campaigns.append_row(campaigns_report_total_row)
    dh.format_last_added_report_in_google_sheets(worksheet=worksheet_campaigns, prop=campaigns_report_headline[0])
    dh.format_summary_row_in_google_sheets(worksheet=worksheet_campaigns,
                                           total_row=campaigns_report_total_row)

    keywords_report_headline = ['Ключевая фраза'] + campaigns_report_headline[1:].copy()
    dh.add_report_date_to_google_sheet(worksheet=worksheet_service_keywords,
                                       days=DAYS_RANGE,
                                       start_column='A',
                                       end_column='I')
    dh.add_report_headline_to_google_sheets(worksheet=worksheet_service_keywords,
                                            columns=keywords_report_headline)
    worksheet_service_keywords.append_rows(keywords_services_report[1:])
    dh.format_last_added_report_in_google_sheets(worksheet=worksheet_service_keywords,
                                                 prop=keywords_report_headline[1])
