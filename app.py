# -*- coding: utf-8 -*-
import os.path
import re
import shutil
import subprocess
import time

import win32con
import win32gui

from appcontroller import AppController
from applogger import AppLogger
from appmodel import AppModel
from appresource import APP_TITLE


class MyApp:
    found_window = False

    def __init__(self, **kwargs):
        self.model = None
        self.controller = None
        self.initialize(**kwargs)

    def __del__(self):
        self.destroy()

    def initialize(self, **kwargs):
        self.destroy()
        if not self.model and not self.controller:
            self.model = AppModel(**kwargs)
            self.controller = AppController(self.model)

    def destroy(self):
        if self.model:
            self.model.destroy()
        if self.controller:
            self.controller.destroy()

    def run(self):
        self.controller.run()

    @staticmethod
    def upgrade(app_path, install_path, timeout):
        AppLogger.info(f'{app_path}, {install_path}, {timeout}')
        time.sleep(timeout)
        shutil.copy(app_path, install_path)
        app_path = os.path.join(install_path,os.path.basename(app_path))
        cmd = f'{app_path}'
        subprocess.Popen(cmd)

    @staticmethod
    def show(**kwargs):
        def show_window(hwnd, wildcard):
            if re.match(wildcard, str(win32gui.GetWindowText(hwnd))):
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                MyApp.found_window = True

        win32gui.EnumWindows(show_window, APP_TITLE)
        if not MyApp.found_window:
            MyApp(**kwargs).run()
