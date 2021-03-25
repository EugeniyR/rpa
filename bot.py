import os
import re
import time

from RPA.Browser.Selenium import Selenium
from RPA.Desktop import Desktop
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
#from RPA.HTTP import HTTP
from RPA.Tables import Tables


# ===== the below parameters can be changed =====
WORKBOOK_NAME = 'workbook.xlsx'
SHEET_GENERAL = 'Agencies'
BOT_TASK = os.path.join('src', 'config.txt')
OUTPUT_DIR = os.path.join('src','output')

# ===== Do not change the following below =======
BASE_URL = "https://itdashboard.gov/"
DIVE_IN = "css:a.btn-lg-2x"
DATA_FEEDS_URL = "css:div.tuck-5 > p > a[href='/drupal/data/datafeeds']"
AGENCY_TITLES = "css:#agency-tiles-2-widget > div > div > div > div > div > div > div:nth-child(2) > a > span.h4.w200"
AGENCY_SPENDINGS = "css:#agency-tiles-2-widget > div > div > div > div > div > div > div:nth-child(2) > a > span.h1.w900"


def get_extended_info(item):
    INVESTMENTS_TABLE = "css:#investments-table-object"
    TABLE_OBJECT_NEXT = "css:#investments-table-object_next.paginate_button.next.disabled"
    INVESTMENT_DROPDOWN = "css:select.form-control:nth-child(1)"
    INVESTMENT_SHOW_ALL = "css:select.form-control:nth-child(1) > option:last-child"
    TABLE_HEADERS = "css:.dataTables_scrollHeadInner > table > thead > tr:nth-child(2) > th"
    TABLE_ROWS = "css:#investments-table-object.datasource-table.usa-table-borderless.dataTable.no-footer tbody tr"
    INVESTMENT_UII = "css:tr > td > a"
    # DOWNLOAD_ELEMENT = "#business-case-pdf > a"
    DOWNLOAD_BUTTON = "link:Download Business Case PDF"
    GEN_TEXT = "Generating PDF"

    print('Getting extended info on ', item)
    browser.click_element('partial link:'+item)
    browser.wait_until_element_is_visible(INVESTMENTS_TABLE, timeout=10)
    browser.click_element(INVESTMENT_DROPDOWN)
    browser.click_element(INVESTMENT_SHOW_ALL)
    browser.wait_until_element_is_visible(TABLE_OBJECT_NEXT, timeout=15)

    column_list = []
    rows = browser.get_element_count(TABLE_ROWS)
    columns = browser.get_element_count(TABLE_HEADERS)
    for i in range(rows):
        rows_list = []
        for j in range(columns):
            get_cell = browser.get_table_cell(INVESTMENTS_TABLE, row=i+3, column=j+1)
            rows_list.append(get_cell)
        column_list.append(rows_list)

    acronym = ''
    for word in item.split():
        acronym += word[0]
    files.open_workbook(os.path.join(OUTPUT_DIR, WORKBOOK_NAME))
    update_excel(acronym, column_list, WORKBOOK_NAME)
    browser.set_download_directory(os.path.join(os.getcwd(), OUTPUT_DIR, acronym), download_pdf=True)

    uii_links = []
    uii_list = browser.get_webelements(INVESTMENT_UII)
    for i in range(len(uii_list)):
        href = browser.get_webelement("css:tr:nth-child("+str(i+1)+") > td > a").get_attribute('href')
        uii_links.append(href)

    for url in uii_links:
        browser.go_to(url)
        browser.wait_until_element_is_visible(DOWNLOAD_BUTTON)
        browser.click_element(DOWNLOAD_BUTTON)
        browser.wait_until_element_is_not_visible(GEN_TEXT)
        time.sleep(7)
        
    time.sleep(5)  # to finish downloading before closing the browser


def get_agencies_info():
    print("Getting the list of displayed Agencies")
    agencies_titles = browser.get_webelements(AGENCY_TITLES)
    agencies_spendings = browser.get_webelements(AGENCY_SPENDINGS)

    agencies_list = []
    for agency in agencies_titles:
        agencies_list.append(browser.get_text(agency)) 

    spendings_list = []
    for spending in agencies_spendings:
        spendings_list.append(browser.get_text(spending))

    agencies_data = list(zip(agencies_list, spendings_list))
    agencies_table = tables.create_table(data=agencies_data)
    files.create_workbook()
    update_excel(SHEET_GENERAL, agencies_table, WORKBOOK_NAME)


def update_excel(sheet_name, table_name, file_name):
    files.create_worksheet(sheet_name)
    files.set_active_worksheet(sheet_name)
    files.append_rows_to_worksheet(content=table_name)
    files.save_workbook(os.path.join(OUTPUT_DIR, file_name))


def main():
    if not os.path.exists(OUTPUT_DIR):
        print ("Output directory was not found and will be created")
        os.makedirs(OUTPUT_DIR)

    #browser.set_download_directory(os.path.join(os.getcwd(), OUTPUT_DIR), download_pdf=True)
    browser.open_available_browser(BASE_URL, maximized=True)
    browser.wait_until_element_is_visible(AGENCY_TITLES)
    # extract summary data on agecies and spending
    get_agencies_info()
    # extract the data on particular agency and add to the Excel file
    with open(BOT_TASK, "r") as f:
        for line in f:
            stripped_line = line.strip()
            if stripped_line:
                if not re.match(r'(^(\s+)?#)', stripped_line):
                    # browser.open_available_browser(BASE_URL, maximized=True)
                    browser.go_to(BASE_URL)
                    browser.wait_until_element_is_visible(AGENCY_TITLES)
                    get_extended_info(stripped_line)

    print('Done')


if __name__ == "__main__":
    
    prefs = {
        "download.default_directory": os.path.join(os.getcwd(), OUTPUT_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }

    browser = Selenium()
    desktop = Desktop()
    tables = Tables() 
    files = Files()
    #http = HTTP()
    main()
    