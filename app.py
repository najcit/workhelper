# -*- coding: utf-8 -*-
"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""
import click
import os
import re
import shutil
import subprocess
import time
import win32con
import win32gui
from appcontroller import AppController
from appmodel import AppModel
from appservice import AppService
from applogger import AppLogger
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
        app_path = os.path.join(install_path, os.path.basename(app_path))
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


@click.command()
@click.option('--root', default='', type=str, help='the root directory of the application')
@click.option('--theme', default='', type=str, help='the theme of the application')
@click.option('--enable_systray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--enable_env', is_flag=True, default=True, help='enable environment variable for the application')
@click.option('--app_path', default='', type=str, help='the newest application file')
@click.option('--install_path', default='', type=str, help='the install path of the application')
@click.option('--timeout', default=3, type=int, help='the timeout to install the application')
@click.option('--service', is_flag=True, default=False, help='run as the service')
def main(root, theme, enable_systray, enable_env, app_path, install_path, timeout, service):
    if os.path.isfile(app_path) and os.path.isdir(install_path):
        MyApp.upgrade(app_path, install_path, timeout)
    elif service:
        AppService().run()
    else:
        MyApp.show(root=root, theme=theme, enable_systray=enable_systray, enable_env=enable_env)


if __name__ == '__main__':
    main()
