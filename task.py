"""Template robot with Python."""
import os
import re
import time
from datetime import timedelta

from RPA.Browser.Selenium import Selenium, webdriver
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF


class ItDashboard:

    def __init__(self):
        self.browser = Selenium()
        self.link = "https://itdashboard.gov/"
        self.file = Files()
        self.pdf = PDF()
        self.browser.set_download_directory(f'{os.getcwd()}/output/pdfs')
        self.sepnt_budgtet_sheet = 'Agencies'
        self.indivisial_investment_sheet = 'Indivisual Investement'
        self.pdf_links = []
        self.rows = []

    def get_agencies(self):
        self.browser.open_chrome_browser(self.link)
        time.sleep(10)
        self.browser.click_element_when_visible('//a[@aria-controls="home-dive-in"][@class="btn btn-default btn-lg-2x trend_sans_oneregular"]')
        agencies = self.browser.find_elements('//*[@id="agency-tiles-widget"]//div[@class="tuck-5"]//div[@class="col-sm-12"]//div[1]')
        agency_rows = [{"Department Name": self.split_name_spent(agency)[0], 'Budget': self.split_name_spent(agency)[1]} for agency in agencies]
        self.file.append_rows_to_worksheet(name=self.sepnt_budgtet_sheet, content=agency_rows, header=True)

    def it_dashboard_all(self):
        self.create_workbook_worsheets()
        self.get_agencies()
        self.browser.click_element_when_visible('//*[@id="agency-tiles-widget"]//a[contains(@href, "/drupal/summary/422") and @class="btn btn-default btn-sm"]')
        time.sleep(10)
        headers = self.browser.find_elements('//div[@class="dataTables_scrollHeadInner"]/table[@class="datasource-table usa-table-borderless dataTable no-footer"]/thead//th')
        headers_ls = [header.text for header in headers]
        self.browser.select_from_list_by_value('//*[@id="investments-table-object_length"]//select[@name="investments-table-object_length"]', '-1')
        time.sleep(15)
        rows_len = len(self.browser.find_elements('//*[@id="investments-table-object"]/tbody/tr'))
        rows = self.create_row(headers_ls, rows_len)
        self.file.append_rows_to_worksheet(rows, self.indivisial_investment_sheet, header=True)
        self.file.save_workbook()
        # self.browser.close_browser()

    def create_row(self, headers_ls, rows_len):
        print(rows_len)
        for row_num in range(1, rows_len + 1):
            try:
                self.browser.does_page_contain_link(
                    f'//*[@id="investments-table-object"]/tbody/tr[{row_num}]/td[1]/a')
                self.pdf_links.append({
                    'link': self.browser.get_element_attribute(
                             f'//*[@id="investments-table-object"]/tbody/tr[{row_num}]/td[1]/a', 'href'),
                    'name': self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{row_num}]/td[1]/a').text,
                    'investment_title': self.browser.get_table_cell('//*[@id="investments-table-object"]', row_num, 3)
                })
            except Exception as e:
                print(e)
            row = {
                headers_ls[col_num]: self.browser.get_table_cell('//*[@id="investments-table-object"]', row_num, col_num + 1)
                   for col_num in range(len(headers_ls))
                }
            self.rows.append(row)
        return self.rows

    def create_workbook_worsheets(self):
        workbook = self.file.create_workbook('output/agancies.xlsx')
        workbook.create_worksheet(self.sepnt_budgtet_sheet)
        workbook.create_worksheet(self.indivisial_investment_sheet)

    def download_uii_pdf(self):
        # self.browser.open_chrome_browser('https://itdashboard.gov/')
        for link in self.pdf_links:
            print(link)
            self.browser.go_to(link["link"])

            # self.browser.wait_until_page_contains_element('//*[@id="business-case-pdf"]/a', timedelta(seconds=5))
            self.browser.click_link('//*[@id="business-case-pdf"]/a')
            time.sleep(5)

    def read_pdf(self):
        # self.pdf_links = [1]
        for link in self.pdf_links:
            # self.pdf.open_pdf(f'output/pdfs/{name["name"]}.pdf')
            text = self.pdf.get_text_from_pdf(f'output/pdfs/{link["name"]}.pdf', [1], details=False, trim=True)
            print(text[1].find('Unique Investment Identifier (UII): '+link["name"]))
            print(text[1].find('Name of this Investment: ' +link["investment_title"]))

    @staticmethod
    def split_name_spent(agency):
        return agency.text.replace('\n', '').split('Total FY2021 Spending:')


obj = ItDashboard()
obj.it_dashboard_all()
obj.download_uii_pdf()
obj.read_pdf()
