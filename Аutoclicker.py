from selenium import webdriver as wb
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementNotInteractableException
from datetime import datetime
from pprint import pprint
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
            ActionChains(self.browser).move_to_element(link)

    def run_webdriver(self):
        chrome_options = wb.ChromeOptions()
        chrome_options.add_argument(r'user-data-dir=C:\Users\SharipovRR\AppData\Local\Google\Chrome\User Data')
        self.browser = wb.Chrome(options=chrome_options)

    def upload_routes(self):
        if self.routes:
            self.routes.clear()
            print('Словарь очищен!')

        table_rows = self.browser.find_element(By.ID, 'myData').find_elements(By.TAG_NAME, 'tr')
        for i, row in enumerate(table_rows[1:]):
            values = [el.text for el in row.find_elements(By.TAG_NAME, 'td')]
            route = values[1]
            bus_numb = values[2].replace(' ', '')
            datetime_from = datetime.strptime(values[3], '%Y-%m-%d %H:%M')
            routes.append({
                'route': route,
                'bus_numb': bus_numb,
                'start_time': datetime_from
            })
            link_obj = row.find_element(By.TAG_NAME, 'a')
            self.routes[route][bus_numb][datetime_from] = link_obj
            print(f'Рейс{i} загружен')

    def pause(self):
        self.state = 0

    def stop(self):
        self.browser.close()
        self.browser.quit()
