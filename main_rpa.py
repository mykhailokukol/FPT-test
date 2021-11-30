from RPA.Browser.Selenium import Selenium
from RPA.Excel.Application import Application
from datetime import timedelta
from time import sleep
import os
import pandas as pd

browser_lib = Selenium()
BASE_DIR = os.path.dirname(os.path.realpath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
writer = pd.ExcelWriter('output/table.xlsx')


def open_website(url):
    browser_lib.open_available_browser(url)


def download_pdf():
    try:
        browser_lib.click_element_when_visible('id:investments-table-object >> tag:tbody >> tag:tr >> tag:td >> tag:a')
        browser_lib.click_element_when_visible('id:business-case-pdf')
        sleep(10.0)
    except Exception:
        print('ERROR: Unable to download the PDF file because it\'s page does not exists.')


def open_view_page():
    browser_lib.click_element_when_visible('xpath://html//body//main//div[1]//div//div//div['
                                           '3]//div//div//div//div//div//div//div//div//div//div//div//div//div//div'
                                           '//div//div[''1]//div//div[2]//div//div//div//div[1]//div['
                                           '2]//div//div//div//div[2]//a')


def save_table(url):
    # IT IS DOES NOT WORKS !
    # ERROR: '<win32com.gen_py.Microsoft Excel 16.0 Object Library._Workbook instance at 0x92097552>'
    # object has no attribute '__len__'
    #
    # app = Application()
    #
    # app.open_application(visible=True)
    # app.open_workbook(os.path.join(OUTPUT_DIR, 'table.xlsx'))
    # app.set_active_worksheet(sheetname='Agencies')
    if url == 'index_page':
        amounts_list = []
        amounts = browser_lib.find_elements('css:span.h1.w900')
        for amount in amounts:
            amounts_list.append([amount.text])
        df = pd.DataFrame(amounts_list)
        df.to_excel(writer, sheet_name='Agencies', index=False, header=False)
        # for value, i in amounts_list, range(1, 11):
        #     app.write_to_cells(row=i, column=1, value=value)
    elif url == 'view_page':
        table = browser_lib.find_element('id:investments-table-object')
        table = table.get_attribute('outerHTML')
        df = pd.read_html(table)[0]
        df.to_excel(writer, sheet_name='Agency', index=False, header=False)
    writer.save()
    # app.save_excel()
    # app.quit_application()


def main():
    try:
        browser_lib.set_download_directory(OUTPUT_DIR)
        open_website('http://itdashboard.gov/')
        # Click 'Dive in' button
        browser_lib.click_link('css:.btn.btn-default.btn-lg-2x.trend_sans_oneregular')
        # Set browser waits because of 'lazy load'
        browser_lib.set_browser_implicit_wait(timedelta(seconds=10))
        # Save data from index page
        save_table(url='index_page')
        open_view_page()
        # Save data from view page
        save_table(url='view_page')
        writer.close()
        # Download PDF file in table
        download_pdf()
    finally:
        browser_lib.close_all_browsers()


if __name__ == '__main__':
    main()
