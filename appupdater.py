import asyncio
import os
import shutil
import tempfile
import time


class AppUpdater:
    def __init__(self, version=None):
        self.current_version = version

    def is_new_version(self, version):
        return False

    def get_latest_version(self):
        pass

    def install(self, update_progress_func=None):
        tmp_file, tmp_file_path = tempfile.mkstemp()
        os.close(tmp_file)
        shutil.copy('server/workhelper_1.0.1.exe', tmp_file_path)

        for i in range(0, 101):
            percent = i
            time.sleep(0.1)
            if update_progress_func:
                asyncio.run(update_progress_func(percent))

    def download(self, version=None):
        pass
