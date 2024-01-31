from selenium import webdriver as wb
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotInteractableException
from datetime import datetime
from collections import defaultdict
import time

RNIS_URL = 'https://reg-rnis.mos.ru/login#maintab.telematic'
routes = [
]


class AutoClicker:
    def __init__(self):
        self.routes = defaultdict(lambda: defaultdict(lambda: defaultdict(str)))
        self.browser = None
        self.state = 1

    def __call__(self, route, bus_numb, datetime_from, reset):
        try:
            if reset:
                self.browser.find_element(By.ID, 'button-1155-btnIconEl').click()
                time.sleep(0.5)
                self.browser.find_element(By.ID, 'button-1036-btnEl').click()
            link = self.routes[route][bus_numb][datetime_from]
            link.click()
        except ElementNotInteractableException:
            ActionChains(self.browser).move_to_element(link).click()

    def run_webdriver(self):
        chrome_options = wb.ChromeOptions()
        chrome_options.add_argument(r'user-data-dir=C:\Users\SharipovRR\AppData\Local\Google\Chrome\User Data')
        self.browser = wb.Chrome(options=chrome_options)

    def download_routes(self, dwn_window):
        if self.routes:
            self.routes.clear()
        table_rows = self.browser.find_element(By.ID, 'myData').find_elements(By.TAG_NAME, 'tr')[1:]
        step_value = len(table_rows)//10
        counter = step_value
        for row in table_rows:
            counter -= 1
            if not counter:
                dwn_window.event_generate('<<Updated>>', when='tail')
                time.sleep(0.3)
                counter = step_value
            values = [el.text for el in row.find_elements(By.TAG_NAME, 'td')]
            route = values[1]
            bus_numb = values[2].replace(' ', '')
            datetime_from = datetime.strptime(values[3], '%Y-%m-%d %H:%M')
            link_obj = row.find_element(By.TAG_NAME, 'a')
            self.routes[route][bus_numb][datetime_from] = link_obj

        dwn_window.end()

    def pause(self):
        self.state = 0

    def stop(self):
        self.browser.close()
        self.browser.quit()
