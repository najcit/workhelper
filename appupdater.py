# -*- coding: utf-8 -*-

import os
import tempfile
import time

import requests

from appresource import APP_EXE


class AppUpdater:
    def __init__(self, url='http://127.0.0.1:5000'):
        self.url = url
        self.app = None

    def close(self):
        pass

    def latest_version(self):
        response = requests.get(self.url+'/latest')
        return response.text

    def downloader(self, update_progress_func=None):
        if update_progress_func:
            update_progress_func(1)
        tmp_file_path = tempfile.mkdtemp()
        file_name = os.path.join(tmp_file_path, APP_EXE)
        self.app = file_name
        if update_progress_func:
            update_progress_func(10)
        print(file_name)
        response = requests.get(self.url + '/download', stream=True)
        if update_progress_func:
            update_progress_func(70)
        with open(file_name, 'wb') as f:
            f.write(response.content)
        if update_progress_func:
            update_progress_func(100)
