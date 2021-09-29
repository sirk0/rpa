from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

import time
from pathlib import Path

browser_lib = Selenium()
excel_lib = Files()
file_system_lib = FileSystem()
TIMEOUT = "10s"


def open_the_website(url):
    browser_lib.open_available_browser(url)


def get_departments() -> dict:
    locator_dive_in = "//a[contains(text(),'DIVE IN')]"
    browser_lib.wait_until_element_is_visible(locator_dive_in, timeout=TIMEOUT)
    browser_lib.click_link(locator_dive_in)
    locator_departments = "//div[@id='agency-tiles-widget']//a[contains(@href,'/drupal/summary/') and not(contains(@class, 'btn')) and not(img)]"
    browser_lib.wait_until_element_is_visible(locator_departments, timeout=TIMEOUT)
    elements = browser_lib.get_webelements(locator_departments)

    departments = {}
    for element in elements:
        department, _, total = element.text.split("\n")
        departments[department] = total
    return departments


def get_individual_investments(department):
    locator_department = "//div[@id='agency-tiles-widget']//*[text()='{}']".format(
        department
    )
    browser_lib.wait_until_element_is_visible(locator_department, timeout=TIMEOUT)
    browser_lib.click_element(locator_department)
    table = get_table()
    return table


def get_table():
    locator_paging = "//*[@name='investments-table-object_length']"
    browser_lib.wait_until_element_is_visible(locator_paging, timeout=TIMEOUT)
    locator_all = locator_paging + "/*[text()='All']"
    browser_lib.click_element(locator_all)
    locator_columns = "//div[@class='dataTables_scrollHead']//tr[@role='row']//th"
    browser_lib.wait_until_element_is_enabled(locator_columns, timeout=TIMEOUT)
    columns = browser_lib.find_elements(locator_columns)
    column_names = [column.text for column in columns]
    print(column_names)
    locator_rows = "//table[@id='investments-table-object']/tbody/tr"
    time.sleep(10)
    browser_lib.wait_until_element_is_visible(locator_rows, timeout=TIMEOUT)
    rows = browser_lib.get_webelements(locator_rows)
    print(len(rows))
    table = [column_names]
    for row in rows:
        tds = [td.text for td in row.find_elements_by_xpath(".//td")]
        table.append(tds)
    links = [
        uii.get_attribute("href") for uii in row.find_elements_by_xpath("//td[1]/a")
    ]
    for link in links:
        print(link)
        download_pdf(link)
    return table


def download_pdf(link):
    try:
        filename = link.split("/")[-1] + ".pdf"
        source = str(Path("~").expanduser().joinpath("Downloads").joinpath(filename))
        file_system_lib.remove_file(source)
        open_the_website(link)
        locator_download = "//a[contains(text(),'Download Business Case PDF')]"
        browser_lib.wait_until_element_is_visible(locator_download, timeout=TIMEOUT)
        browser_lib.find_element(
            "//a[contains(text(),'Download Business Case PDF')]"
        ).click()
        time.sleep(10)
        file_system_lib.copy_file(source, f"output/{filename}")

    finally:
        browser_lib.close_browser()


def create_excel_worksheet(name, content):
    excel_lib.create_worksheet(name, content=content)


def main():
    try:
        file_system_lib.create_directory("output")
        file_system_lib.empty_directory("output")
        open_the_website("https://itdashboard.gov/")
        departments = get_departments()
        print(departments)
        excel_lib.create_workbook()
        content = [["department", "total"], *departments.items()]
        create_excel_worksheet("Agencies", content)
        department = "National Science Foundation"
        investments = get_individual_investments(department)
        create_excel_worksheet("Investments", investments)
    finally:
        browser_lib.close_all_browsers()
        excel_lib.save_workbook("output/Workbook.xlsx")


if __name__ == "__main__":
    main()
