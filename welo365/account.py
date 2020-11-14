from __future__ import annotations

import logging
import os
import sys

from O365 import Account, Connection
from O365.drive import Folder
from O365.excel import WorkSheet
from pathlib import Path
from selenium import webdriver

logfile = Path.cwd() / 'output.log'
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
date_format = "%H:%M:%S"
formatter = logging.Formatter(log_format, date_format)
ch = logging.StreamHandler(sys.stderr)
ch.setFormatter(formatter)
ch.setLevel(logging.INFO)
logger.addHandler(ch)
fh = logging.FileHandler(logfile)
fh.setFormatter(formatter)
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)


def get_item(self, item_name: str):
    for item in self.get_items():
        if item_name.lower() in item.name.lower():
            return item


Folder.get_item = get_item


def protect(self):
    payload = {
        'options': {
            'allowFormatCells': False,
            'allowFormatColumns': False,
            'allowFormatRows': False,
            'allowInsertColumns': False,
            'allowInsertRows': False,
            'allowInsertHyperlinks': False,
            'allowDeleteColumns': False,
            'allowDeleteRows': False,
            'allowSort': True,
            'allowAutoFilter': True,
            'allowPivotTables': True
        }
    }
    return bool(self.session.post(json=payload))


def unprotect(self):
    bool(self.build_url('/protection/unprotect'))


WorkSheet.protect = protect
WorkSheet.unprotect = unprotect


class Folder(Folder):
    pass


class WorkSheet(WorkSheet):
    pass


class O365Account:
    def __init__(self, creds: tuple[str, str] = None, scopes: list[str] = None, scrape: bool = True):
        self.creds = creds or (os.environ.get('welo365_client_id'), os.environ.get('welo365_client_secret'))
        self.scopes = scopes or ['offline_access', 'Sites.Manage.All']
        self.account = Account(self.creds, auth_flow_type='authorization', scopes=scopes)
        self.con = self.account.con
        if scrape:
            self.con = self.scrape(self.con, scopes)
        if not self.account.is_authenticated:
            self.authenticate()

        self.storage = self.account.storage()
        self.drives = self.storage.get_drives()
        self.sharepoint = self.account.sharepoint
        self.my_drive = self.storage.get_default_drive()
        self.site = None
        self.root_folder = self.my_drive.get_root_folder()

    @staticmethod
    def scrape(con: Connection, scopes: list[str]):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        driver = webdriver.Chrome(options=chrome_options)
        auth_url, _ = con.get_authorization_url(requested_scopes=scopes)
        driver.get(auth_url)
        driver.implicitly_wait(5)
        email = driver.find_element_by_xpath('.//input[@type="email"]')
        email.send_keys(os.environ.get('okta_username'))
        submit = driver.find_element_by_xpath('.//input[@type="submit"]')
        submit.click()
        password = driver.find_element_by_xpath('.//input[@type="password"]')
        password.send_keys(os.environ.get('okta_password'))
        submit = driver.find_element_by_xpath('.//input[@value="Sign in"]')
        submit.click()
        driver.implicitly_wait(15)
        checkbox = driver.find_element_by_xpath('.//input[@type="checkbox"]')
        checkbox.click()
        submit = driver.find_element_by_xpath('.//input[@value="Yes"]')
        submit.click()
        driver.implicitly_wait(3)
        con.request_token(driver.current_url)
        driver.quit()
        return con

    def authenticate(self):
        result = self.account.authenticate()

    def get_drive(self):
        return self.my_drive

    def get_root_folder(self):
        return self.root_folder

    def get_folder_from_path(self, folder_path: str, site: str = None):
        folder_path = folder_path[1:] if folder_path[0] == '/' else folder_path

        if folder_path is None:
            return self.my_drive

        subfolders = folder_path.split('/')
        if len(subfolders) == 0:
            return self.my_drive

        if site:
            self.site = self.sharepoint().get_site('welocalize.sharepoint.com', site)

        drive = self.site.get_default_document_library() if self.site else self.my_drive

        items = drive.get_items()
        for subfolder in subfolders:
            try:
                subfolder_drive = list(filter(lambda x: subfolder in x.name, items))[0]
                items = subfolder_drive.get_items()
            except:
                raise ('Path {} not exist.'.format(folder_path))
        return subfolder_drive
