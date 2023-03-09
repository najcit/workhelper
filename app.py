# -*- coding: utf-8 -*-

import re
import win32con
import win32gui

from appcontroller import AppController
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
