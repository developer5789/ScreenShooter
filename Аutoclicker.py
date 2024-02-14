from selenium import webdriver as wb
import time


class AutoClicker:
    def __init__(self, profile_path=None):
        self.profile_path = profile_path
        self.browser = None
        self.state = 1
        self.skip = False
        self.setInterface = 0
        self.setval = ''
        self.read_setval()

    def __bool__(self):
        if self.browser:
            return True
        return False

    def __call__(self, bus_numb, datetime_from, datetime_to, reset):
        if not self.setInterface:
            self.browser.execute_script("document.getElementById('combobox-1032-inputEl').value = 'Последний час';")
            self.setInterface = 1
            time.sleep(0.5)

        if reset:
            self.browser.execute_script("""document.getElementById('textfield-1123-inputEl').value = 'q';
                                            $("#button-1154-btnIconEl").click();""")
            time.sleep(0.5)

        setval_call = f'setval("{bus_numb}", "{datetime_from}", "{datetime_to}");'

        self.browser.execute_script(self.setval + setval_call)

    def run_webdriver(self):
        self.browser = wb.Chrome(options=self.get_options())
        self.browser.get('https://reg-rnis.mos.ru/')

    def get_options(self):
        chrome_options = wb.ChromeOptions()
        if self.profile_path:
            chrome_options.add_argument(fr'user-data-dir={self.profile_path}')
        return chrome_options

    def pause(self):
        self.state = 0

    def stop(self):
        if self.browser:
            self.browser.quit()
            self.browser = None

    def skip_route(self):
        self.skip = True

    def read_setval(self):
        with open('new_setval.js', encoding='utf-8') as f:
            self.setval = f.read()
