from selenium import webdriver as wb
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (ElementNotInteractableException,
                                        ElementClickInterceptedException, NoSuchElementException)
from datetime import datetime
from collections import defaultdict
import time
from tkinter.messagebox import showerror

RNIS_URL = 'https://reg-rnis.mos.ru/login#maintab.telematic'


class AutoClicker:
    def __init__(self, profile_path):
        self.profile_path = profile_path
        self.routes = defaultdict(lambda: defaultdict(lambda: defaultdict(str)))
        self.browser = None
        self.state = 1

    def __bool__(self):
        if self.browser:
            return True
        return False

    def __call__(self, route, bus_numb, datetime_from, reset):
        try:
            if reset:
                self.browser.find_element(By.ID, 'button-1155-btnIconEl').click()
                time.sleep(0.5)
                self.browser.find_element(By.ID, 'button-1036-btnEl').click()
            link = self.routes[route][bus_numb][datetime_from]
            link.click()
        except (ElementNotInteractableException, ElementClickInterceptedException):
            ActionChains(self.browser).move_to_element(link).click()

    def run_webdriver(self):
        self.browser = wb.Chrome(options=self.get_options())

    def get_options(self):
        chrome_options = wb.ChromeOptions()
        if self.profile_path:
            chrome_options.add_argument(fr'user-data-dir={self.profile_path}')
        return chrome_options

    def download_routes(self, dwn_window):
        try:
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
                bus_numb = values[2].strip()
                datetime_from = datetime.strptime(values[3], '%Y-%m-%d %H:%M')
                link_obj = row.find_element(By.TAG_NAME, 'a')
                self.routes[route][bus_numb][datetime_from] = link_obj

        except Exception:
            showerror('Ошибка', 'Упс!Возникла ошибка при загрузке рейсов.')
            self.routes.clear()
        finally:
            if dwn_window:
                dwn_window.end()

    def pause(self):
        self.state = 0

    def stop(self):
        if self.browser:
            self.browser.quit()
            self.browser = None
