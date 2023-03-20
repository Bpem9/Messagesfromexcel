
from typing import List
import numpy
import pandas as pd
import win32com.client as client

excel_data = pd.read_excel('Входные данные.xlsx', sheet_name='Лист1')


class Receiver:
    @staticmethod
    def get_logins(data: List) -> List:
        return list(set(self.data['login'].array))


class EnterSendman:
    @staticmethod
    def send_to_gmail_for_logins(logins: List) -> None:
        outlook = client.Dispatch('Outlook.Application')
        for login in logins:
            message = outlook.CreateItem(0)
            message.Display(0)
            message.To = str(login) + '@gmail.com'
            message.Subject = 'Test'
            message.Body = str(data[data.login == login].iloc[:, 1])
            # message.send()


excel_logins = Receiver().get_logins(excel_data)
EnterSendman.send_to_gmail_for_logins(excel_logins)
