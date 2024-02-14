from datetime import datetime


class Loger:

    @staticmethod
    def enter_in_log(err):
        message = f'time: {datetime.now()}  error: {err}'
        Loger.write(message)

    @staticmethod
    def write(err_message):
        with open('errors_log.txt', 'a') as f:
            f.write(err_message + f'\n {len(err_message)*'-'} \n')

