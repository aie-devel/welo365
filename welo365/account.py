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

DOMAIN = 'welocalize.sharepoint.com'


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


class O365Account(Account):
    def __init__(
            self,
            site: str = None,
            creds: tuple[str, str] = None,
            scopes: list[str] = None,
            auth_flow_type: str = 'authorization',
            scrape: bool = False
    ):
        creds = creds or (os.environ.get('welo365_client_id'), os.environ.get('welo365_client_secret'))
        scopes = scopes or ['offline_access', 'Sites.Manage.All']
        super().__init__(creds, scopes=scopes, auth_flow_type=auth_flow_type)
        if scrape:
            self.scrape(scopes)
        if not self.is_authenticated:
            self.authenticate()
        self.drives = self.storage().get_drives()
        self.site = self.get_site(site) if site else None
        self.drive = self.site.get_default_document_library() if self.site else self.storage().get_default_drive()
        self.root_folder = self.drive.get_root_folder()

    def scrape(self, scopes: list[str]):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        driver = webdriver.Chrome(options=chrome_options)
        auth_url, _ = self.con.get_authorization_url(requested_scopes=scopes)
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
        self.con.request_token(driver.current_url)
        driver.quit()

    def authenticate(self):
        result = self.authenticate()

    def get_site(self, site: str):
        return self.sharepoint().get_site(DOMAIN, f"/sites/{site}")

    def get_folder(self, *subfolders: str, site: str = None):
        if len(subfolders) == 0:
            return self.drive
        
        drive = self.get_site(site}.get_default_document_library() if site else self.drive

        if subfolders[0] != 'General':
            subfolders = ['General', *subfolders]

        items = drive.get_items()
        for subfolder in subfolders:
            try:
                subfolder_drive = list(filter(lambda x: subfolder in x.name, items))[0]
                items = subfolder_drive.get_items()
            except:
                raise ('Path {} not exist.'.format('/'.join(subfolders)))
        return subfolder_drive
