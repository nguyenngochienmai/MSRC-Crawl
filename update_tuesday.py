import json
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re


df = pd.read_csv('june.csv')

# Create a new DataFrame with the desired columns
new_columns = ['Product Family',
               'Impact Scope',
               'Impact',
               'Max Severity',
               'Details',
               'Base Server Score',
               'Impacted OS',
               'Recommendation',
               'Publicly Disclosed',
               'Exploited'
               ]
new_df = pd.DataFrame(columns=new_columns)

# Copy data from the original CSV to the new DataFrame
new_df['Product Family'] = df['Product Family']
new_df['Impact'] = df['Impact']
new_df['Max Severity'] = df['Max Severity']
new_df['Details'] = df['Details']
new_df['Base Server Score'] = df['Base Score']



def get_data (vuln):
    url1 = (
        "https://api.msrc.microsoft.com/sug/v2.0/en-US/affectedProduct?$orderBy=releaseDate%20desc&$filter=cveNumber%20eq%20%27"+ vuln +"%27"
    )
    data1 = requests.get(url1).json()
    unique_products = set()
    unique_kb_numbers = set()

    for item in data1["value"]:
        product = f"- {item['product']}"
        unique_products.add(product)

        kb_articles = item["kbArticles"]
        for article in kb_articles:
            article_name = article["articleName"]
            if article_name.isdigit():
                kb_number = f"KB{article_name}"
                unique_kb_numbers.add(kb_number)

    products_string = "\n".join(unique_products)
    kb_numbers_string = ",".join(unique_kb_numbers)

    url = (
        "https://api.msrc.microsoft.com/sug/v2.0/en-US/vulnerability/" + vuln
    )
    data = requests.get(url).json()
    publiclyDisclosed = f"{data['publiclyDisclosed']}."
    exploited = f"{data['exploited']}."

    return products_string, kb_numbers_string, publiclyDisclosed, exploited

with pd.ExcelWriter('june.xlsx', engine='xlsxwriter') as writer:
    new_df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Access the XlsxWriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add the desired format for the URLs
    url_format = workbook.add_format({'color': 'blue', 'underline': 1})
    bold_format = workbook.add_format({'bold': True})
    italic_format = workbook.add_format({'italic': True})

    for row_num, (id_cve, url) in enumerate(zip(df['Details'], df['Details.1']), start=1):
        hyperlink_formula = f'=HYPERLINK("{url}","{id_cve}")'
        worksheet.write_formula(row_num, new_columns.index('Details'), hyperlink_formula, url_format)
        pf = df['Product Family'][row_num-1]

        impacted_os, kb_article, publiclyDisclosed, exploited = get_data(id_cve)
        worksheet.write(row_num, new_columns.index('Impacted OS'), impacted_os)

        bold_start = len('Workaround:\nSolution:\n- Tải các bản cập nhật bảo mật mới nhất trong tháng 06/2024 dành cho ' + pf +', bao gồm:\n')
        bold_end = bold_start + len(kb_article)
        string = 'Workaround:\nSolution:\n- Tải các bản cập nhật bảo mật mới nhất trong tháng 06/2024 dành cho ' + pf +', bao gồm:\n' + kb_article +'.'
        format_text = (
                string[:bold_start], bold_format,
                string[bold_start: bold_end ], italic_format,
                string[bold_end:], italic_format
            )
        worksheet.write_rich_string(
                row_num,
                new_columns.index('Recommendation'),
                *format_text
            )
        worksheet.write(row_num, new_columns.index('Publicly Disclosed'), publiclyDisclosed)
        worksheet.write(row_num, new_columns.index('Exploited'), exploited)
