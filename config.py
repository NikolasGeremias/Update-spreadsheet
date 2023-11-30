import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import hidden_data


class ConfigSelenium:
    def __init__(self, path=Path(os.getcwd() + '\\Exportacao')) -> None:
        self._folder = path
        self._date_range = self.default_date_range()
        self.service = Service(ChromeDriverManager().install())
        self.chrome_options = self._chrome_options()
        self.delete_data(self.folder)

    @property
    def date_range(self):
        return self._date_range

    @date_range.setter
    def date_range(self, days: int):
        current_date = datetime.now()
        input_date = current_date - timedelta(days=days)
        input_date = input_date.strftime('%d/%m/%Y')

        self._date_range = input_date

    @property
    def folder(self):
        return self._folder

    @folder.setter
    def folder(self, path):
        self._folder = path

    @staticmethod
    def default_date_range():
        current_date = datetime.now()
        input_date = current_date - timedelta(days=30)
        input_date = input_date.strftime('%d/%m/%Y')

        return input_date

    @staticmethod
    def delete_data(folder):
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))

    def _chrome_options(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless=new")
        prefs = {'download.default_directory': f'{self._folder}'}
        chrome_options.add_experimental_option('prefs', prefs)

        return chrome_options


class ConfigGoogleApi:
    def __init__(self, spreadsheet_id=hidden_data.spreadsheet_id, range=hidden_data.range) -> None:
        self._SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self._range = range
        self._spreadsheet_id = spreadsheet_id
        self.service = self.service()

    @property
    def range(self):
        return self._range

    @range.setter
    def range(self, range):
        self._range = range

    @property
    def spreadsheet_id(self):
        return self._spreadsheet_id

    @spreadsheet_id.setter
    def spreadsheet_id(self, spreadsheet_id):
        self._spreadsheet_id = spreadsheet_id

    def service(self, creds=None):
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self._SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:

                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self._SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            service = build('sheets', 'v4', credentials=creds)
        except Exception:
            DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
            service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)

        return service


def main():
    ...


if __name__ == '__main__':
    main()
