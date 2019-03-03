import re
import time

from selenium import webdriver
from xlsx_writer import XlsxWrite


class EmailScraper:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.page_count = 10
        self.url = 'https://www.google.com/search?q=%s&start=%s&num=15'
        self.search_keyword = '"Plumber" "India" "@gmail.com"'.replace(" ", "+")
        self.xlsx_writer = XlsxWrite()

    def start_requests(self):
        for i in range(1, self.page_count+1):
            self.driver.get(url=self.url % (self.search_keyword, i*10))
            self.process_page(i)
        self.xlsx_writer.save_file()
        time.sleep(10)
        self.driver.close()

    def process_page(self, page):
        items = []
        results = self.driver.find_elements_by_css_selector('.st')
        for result in results:
            emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', result.text)
            items.extend(emails)
        print("Found %s no of emails in page %s" %(len(items), page))
        self.xlsx_writer.write_to_sheet(items)


if __name__ == "__main__":
    res = EmailScraper()
    res.start_requests()

