# -*- coding: utf-8 -*-

import asyncio
import os
import shutil
import tempfile
import time


class AppUpdater:
    def __init__(self):
        self.newest_app = None

    def close(self):
        pass

    def get_newest_version(self):
        return '1.0.1'

    def installer(self, update_progress_func=None):
        tmp_file_path = tempfile.mkdtemp()
        print(tmp_file_path)
        target = os.path.join(tmp_file_path, 'workhelper.exe')
        shutil.copy('server/workhelper_1.0.1.exe', target)
        self.newest_app = os.path.join(tmp_file_path, 'workhelper.exe')
        for i in range(0, 101):
            percent = i
            time.sleep(0.01)
            if update_progress_func:
                asyncio.run(update_progress_func(percent))
