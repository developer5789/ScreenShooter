from selenium import webdriver as wb
from selenium.common import exceptions
import time
import json


class AutoClicker:
    """Cтроит треки рейсов в браузере."""
    def __init__(self, profile_path=None):
        """Инициализация аттрибутов

            Аргументы:
                profile_path(str): путь до папки профиля Chrome
        """
        self.profile_path = profile_path
        self.browser = None
        self.state = 0
        self.skip = False
        self.setInterface = 0
        self.setval = ''
        self.read_setval()

    def __bool__(self):
        """Есть ли обьект webdriver в памяти."""
        if self.browser:
            return True
        return False

    def __call__(self, bus_numb: str, datetime_from: str, datetime_to: str, reset: bool):
        """Строит трек в браузере, ставит на исполнение js-код

            Аргументы:
                bus_numb(str): гос.номер ТС,
                datetime_from(str): время начала рейса,
                datetime_to(str): время конца рейса,
                reset(bool): нужно ли делать сброс предыдущего трека.


        """
        if not self.setInterface:
            self.browser.execute_script("Ext.getCmp('combobox-1032').setValue('1h')")
            self.setInterface = 1

        if reset:
            self.browser.execute_script("""Ext.getCmp('textfield-1123').setValue('q');
                                            $("#button-1154-btnIconEl").click();""")
            time.sleep(0.5)

        setval_call = f'setval("{bus_numb}", "{datetime_from}", "{datetime_to}");'

        self.browser.execute_script(self.setval + setval_call)

    def run_webdriver(self):
        """Открывает браузер и сайт РНИС."""
        self.read_setval()
        self.browser = wb.Chrome(options=self.get_options())
        self.browser.get('https://reg-rnis.mos.ru/')
        self.browser.execute_cdp_cmd('Network.enable', {})
        self.setInterface = 0

    def get_options(self):
        """Создает и возвращает параметры обьекта chromedriver."""
        chrome_options = wb.ChromeOptions()
        chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
        chrome_options.add_experimental_option('detach', True)
        chrome_options.add_argument('--enable-logging')
        chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        if self.profile_path:
            chrome_options.add_argument(fr'user-data-dir={self.profile_path}')
        return chrome_options

    def pause(self):
        """Ставит на паузу автокликер."""
        self.state = 0

    def stop(self):
        """Закрывает браузер, очищает память."""
        if self.browser:
            self.browser.quit()
            self.browser = None

    def skip_route(self):
        """Задает self.skip значение True. Данный аттрибут определяет
         нужно ли пропускать рейс в таблице.
         """
        self.skip = True

    def read_setval(self):
        """Читает js-код из файла 'setval.txt'"""
        with open('setval.txt', encoding='utf-8') as f:
            self.setval = f.read()

    def check_track(self):
        for i in range(10):
            cords = self.find_cords()
            if cords:
                time.sleep(0.5)
                return True
            elif cords is not None:
                return False
            else:
                time.sleep(0.5)

    def find_cords(self):
        log_entries = self.browser.get_log("performance")
        for entry in log_entries:
            message_obj = json.loads(entry.get("message"))
            message = message_obj.get("message")
            method = message.get("method")

            if method == 'Network.responseReceived':
                response_url: str = message.get('params', {}).get('response', {}).get('url', '')
                if response_url.startswith('https://reg-rnis.mos.ru/service/geo/layerstracks'):
                    request_id = message.get('params', {}).get('requestId', '')
                    err_counter = 0
                    while True:
                        try:
                            response = self.browser.execute_cdp_cmd('Network.getResponseBody', {'requestId': request_id})
                            response_body = json.loads(response.get('body', ''))
                            cords = list(response_body['result'].values())[0]
                            return len(cords) > 30
                        except exceptions.WebDriverException:
                            if err_counter < 3:
                                err_counter += 1
                                time.sleep(0.5)
                                continue
                            break
                    print(f'Ошибок {err_counter}')

    def focus_on_track(self, table): # надо доработать
        previous_id = table.current_item - 1
        previous_route = table.item(str(previous_id))['values'][1] if previous_id >= 0 else None
        current_route = table.item(str(previous_id + 1))['values'][1]
        if current_route != previous_route or previous_route is None:
            try:
                self.browser.execute_script("""let points = $("#panel-1258-innerCt table");
                                            const m = (points.length - points.length  % 2) / 2;
                                            points[m].click();
                                            flag = document.querySelector('.ol-overlay-container');
                                            if (flag){
                                            flag.remove()
                                            };
                                            """)
                time.sleep(2)

            except Exception as err:
                pass

