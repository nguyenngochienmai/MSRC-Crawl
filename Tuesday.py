import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import xlsxwriter
import re
import time


browser = webdriver.Firefox()

df = pd.read_csv('may.csv')

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


# Create a new Excel writer with XlsxWriter engine
with pd.ExcelWriter('may.xlsx', engine='xlsxwriter') as writer:
    new_df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Access the XlsxWriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # Add the desired format for the URLs
    url_format = workbook.add_format({'color': 'blue', 'underline': 1})
    bold_format = workbook.add_format({'bold': True})
    italic_format = workbook.add_format({'italic': True})


    # Apply the format to the 'Details' column
    for row_num, (id_cve, url) in enumerate(zip(df['Details'], df['Details.1']), start=1):
        hyperlink_formula = f'=HYPERLINK("{url}","{id_cve}")'
        worksheet.write_formula(row_num, new_columns.index('Details'), hyperlink_formula, url_format)

        browser.get(url)
        time.sleep(5)

        # Fill the "Impacted OS" column        
        get_products = browser.find_elements(By.XPATH, '//div[@data-automation-key="product"]')
        products = set(product.text for product in get_products)
        impacted_os = '\n'.join(f'- {product}.' for product in products)
        worksheet.write(row_num, new_columns.index('Impacted OS'), impacted_os)

        # Fill the "Recommendation" column
        get_articles = browser.find_elements(By.XPATH, '//div[@data-automation-key="kbArticles"]')
        articles = set(article.text for article in get_articles)
        kb_articles = []
        for article in articles:
            numbers = re.findall(r'\d+', article)
            kb_articles.extend(f'KB{number}' for number in numbers)
            kb_articles_text = ', '.join(kb_articles)
            bold_start = len('Workaround:\nSolution:\n- Tải các bản cập nhật bảo mật mới nhất trong tháng 05/2024 dành cho Windows, bao gồm:\n')
            bold_end = bold_start + len(kb_articles_text)
            string = 'Workaround:\nSolution:\n- Tải các bản cập nhật bảo mật mới nhất trong tháng 05/2024 dành cho Windows, bao gồm:\n' + kb_articles_text +'.'
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
        
        # Fill the "Publicly Disclosed" column
        get_publiclyDisclosed = browser.find_elements(By.XPATH, '//div[@data-automation-key="publiclyDisclosed"]')
        publiclyDisclosed = set(publiclyDisclosed.text for publiclyDisclosed in get_publiclyDisclosed)
        publiclyDisclosed = ''.join(publiclyDisclosed)
        worksheet.write(row_num, new_columns.index('Publicly Disclosed'), publiclyDisclosed)

        # Fill the "Exploited" column
        get_exploited = browser.find_elements(By.XPATH, '//div[@data-automation-key="exploited"]')
        exploited = set(exploited.text for exploited in get_exploited)
        exploited = ''.join(exploited)
        worksheet.write(row_num, new_columns.index('Exploited'), exploited)

    browser.quit()
