from datetime import datetime


class Loger:
    """Логер для записи ошибок приложения."""
    @staticmethod
    def enter_in_log(err):
        """Внесение ошибки в лог
        
        Аргументы:
               err(Exception): исключение.
        """
        message = f'time: {datetime.now()}  error: {err}'
        Loger.write(message)

    @staticmethod
    def write(err_message):
        """Запись ошибки в текстовый файл errors_log.txt

            Аргументы:
               err_message(str): строка с текстом ошибки.
        """
        with open('errors_log.txt', 'a') as f:
            f.write(err_message + f"\n {len(err_message)*'-'} \n")

