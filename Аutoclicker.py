from selenium import webdriver as wb
import time


class AutoClicker:
    def __init__(self, profile_path=None):
        self.profile_path = profile_path
        self.browser = None
        self.state = 1
        self.skip = False
        self.js_st = 'myArr = [{{"number": "{number}", "datefrom": "{date_from}",' \
                     ' "dateto": "{date_to}"}}];setval(0);'

    def __bool__(self):
        if self.browser:
            return True
        return False

    def __call__(self, bus_numb, datetime_from, datetime_to, reset):
        if reset:
            self.browser.execute_script("""Ext.getCmp('textfield-1123').setValue('q');
                                            $("#button-1154-btnIconEl").click();""")
            time.sleep(0.5)

        js_script = self.js_st.format(number=bus_numb,
                                      date_from=datetime_from,
                                      date_to=datetime_to)
        self.browser.execute_script(js_script)

    def run_webdriver(self):
        self.browser = wb.Chrome(options=self.get_options())

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
