# -*- coding: utf-8 -*-

import contextlib
import os
import subprocess
from PySimpleGUI import EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED, WIN_CLOSE_ATTEMPTED_EVENT, WIN_CLOSED

from applogger import AppLogger
from appupdater import AppUpdater
from appdatabase import AppDatabase
from appwindow import AppWindow
from appresource import *


class AppController(object):

    def __init__(self, model=None):
        self.model = None
        self.window = None
        self.updater = None
        self.database = None
        self.logger = None
        self.event_functions = self._init_event_functions()
        self.initialize(model)

    def __del__(self):
        self.destroy()

    def initialize(self, model):
        self.destroy()
        self.model = model
        if not self.window and self.model:
            self.window = AppWindow(self.model)
        if not self.updater:
            self.updater = AppUpdater()
        if not self.database:
            self.database = AppDatabase()
        if not self.logger:
            self.logger = AppLogger()

    def destroy(self):
        if self.window:
            self.window.close()
            self.window = None
        if self.updater:
            self.updater.close()
            self.updater = None
        if self.database:
            self.database.close()
            self.updater = None

    # noinspection PyArgumentList
    def run(self):
        with contextlib.suppress(Exception):
            while self.window:
                event, values = self.window.read()
                print(event, values, isinstance(event, tuple), type(event))
                if event in self.event_functions:
                    self.event_functions[event](event=event, values=values)
                else:
                    self.event_functions[E_DEFAULT](event=event, values=values)

    def refresh(self, **_kwargs):
        self.window.refresh()

    def attempt_exit(self, **_kwargs):
        self.window.attempt_exit()

    def exit(self, **_kwargs):
        self.window.exit()
        self.destroy()

    def show(self, **_kwargs):
        self.window.show()

    def search(self, **kwargs):
        values = kwargs['values']
        current_tag = values[E_ACTIVE_TAG]
        search_content = values['-CONTENT-']
        self.window.search(current_tag=current_tag, search_content=search_content)

    def back(self, **_kwargs):
        current_tag_list = self.model.current_tag.split('#')
        parent_tag = '#'.join(current_tag_list[:-1])
        self.window.back(parent_tag)

    def activate_tag(self, **kwargs):
        values = kwargs['values']
        current_tag = values[E_ACTIVE_TAG]
        self.window.activate_tag(current_tag)

    def select_app(self, **kwargs):
        event = kwargs['event']
        app_key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
        self.window.select_app(app_key)

    def import_app_info(self, **_kwargs):
        self.window.import_app_info()

    def export_app_info(self, **_kwargs):
        self.window.export_app_info()

    def show_view_icon(self, **_kwargs):
        self.window.show_view_icon(location=self.model.location)

    def show_view_list(self, **_kwargs):
        self.window.show_view_list(location=self.model.location)

    def sort_app(self, **_kwargs):
        self.window.sort_app(sort_order=self.model.sort_order)

    def select_root(self, **_kwargs):
        self.window.select_root()

    def select_theme(self, **_kwargs):
        self.window.select_theme()

    def select_font(self, **_kwargs):
        self.window.select_font()

    def set_window(self, **_kwargs):
        self.window.set_window()

    def show_local_app(self, **_kwargs):
        self.window.show_local_app()

    def hide_local_app(self, **_kwargs):
        self.window.hide_local_app()

    def show_shared_app(self, **_kwargs):
        self.window.show_shared_app()

    def hide_shared_app(self, **_kwargs):
        self.window.hide_shared_app()

    def show_browser_bookmark(self, **_kwargs):
        self.window.show_browser_bookmark()

    def hide_browser_bookmark(self, **_kwargs):
        self.window.hide_browser_bookmark()

    def check_update(self, **_kwargs):
        newest_version = self.updater.latest_version()
        downloader = self.updater.downloader
        self.window.check_update(newest_version, downloader)

    def finish_download(self, **_kwargs):
        if self.window.finish_download():
            cmd = f'{self.updater.app} --app_path {self.updater.app} --install_path {os.getcwd()}'
            subprocess.Popen(cmd)
            print(cmd)
            self.destroy()

    def about(self, **_kwargs):
        version = self.model.version
        self.window.about(version)

    def manage_tag(self, **_kwargs):
        self.window.manage_tag()

    def new_tag(self, **_kwargs):
        self.window.new_tag()

    def mod_tag(self, **kwargs):
        current_tag = kwargs['values'][E_ACTIVE_TAG]
        self.window.mod_tag(current_tag)

    def rmv_tag(self, **kwargs):
        current_tag = kwargs['values'][E_ACTIVE_TAG]
        self.window.rmv_tag(current_tag)

    def new_app(self, **kwargs):
        current_tag = kwargs['values'][E_ACTIVE_TAG]
        self.window.new_app(current_tag)

    def mod_app(self, **_kwargs):
        self.window.mod_app()

    def rmv_app(self, **_kwargs):
        self.window.rmv_app()

    def share_app(self, **_kwargs):
        self.window.share_app()

    def open_app_path(self, **_kwargs):
        self.window.open_app_path()

    def copy_app_path(self, **_kwargs):
        self.window.copy_app_path()

    def run_app(self, **kwargs):
        event = kwargs['event']
        self.window.run_app(event)
        # self.database.start_app(app_path)

    def default_action(self, **kwargs):
        event = kwargs['event']
        print('event', event)
        if isinstance(event, str) and (event.endswith('LeftClick') or event.endswith('RightClick')):
            self.select_app(**kwargs)
        elif isinstance(event, str) and event.endswith('DoubleClick'):
            self.run_app(**kwargs)
        elif isinstance(event, tuple) and event == (E_THREAD, E_DOWNLOAD_END):
            self.finish_download()

    def _init_event_functions(self):
        new_event_func = {}
        event_func = {
            EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED: self.show,
            WIN_CLOSE_ATTEMPTED_EVENT: self.attempt_exit,
            WIN_CLOSED: self.exit,
            E_SYSTRAY_SHOW: self.show,
            E_SYSTRAY_HIDE: self.attempt_exit,
            E_SYSTRAY_QUIT: self.exit,
            E_QUIT: self.exit,
            E_CHECK_UPDATE: self.check_update,
            E_ABOUT: self.about,
            E_IMPORT_APPS_INFO: self.import_app_info,
            E_EXPORT_APPS_INFO: self.export_app_info,
            E_SHOW_ICON_VIEW: self.show_view_icon,
            E_SHOW_LIST_VIEW: self.show_view_list,
            E_SELECT_ROOT: self.select_root,
            E_SELECT_THEME: self.select_theme,
            E_SELECT_FONT: self.select_font,
            E_SET_WINDOW: self.set_window,
            E_SHOW_LOCAL_APPS: self.show_local_app,
            E_HIDE_LOCAL_APPS: self.hide_local_app,
            E_SHOW_SHARED_APPS: self.show_shared_app,
            E_HIDE_SHARED_APPS: self.hide_shared_app,
            E_SHOW_BROWSER_BOOKMARKS: self.show_browser_bookmark,
            E_HIDE_BROWSER_BOOKMARKS: self.hide_browser_bookmark,
            E_SEARCH: self.search,
            E_RETURN: self.back,
            E_ACTIVE_TAG: self.activate_tag,
            E_REFRESH: self.refresh,
            E_MANAGE_TAG: self.manage_tag,
            E_NEW_TAG: self.new_tag,
            E_MOD_TAG: self.mod_tag,
            E_RMV_TAG: self.rmv_tag,
            E_SRT_APP: self.sort_app,
            E_NEW_APP: self.new_app,
            E_MOD_APP: self.mod_app,
            E_RMV_APP: self.rmv_app,
            E_SHARE_APP: self.share_app,
            E_RUN_APP: self.run_app,
            E_OPEN_APP_PATH: self.open_app_path,
            E_COPY_APP_PATH: self.copy_app_path,
            E_DEFAULT: self.default_action,
        }
        for key, value in event_func.items():
            if key and '&' in key:
                new_event_func[key.replace('&', '')] = value
            else:
                new_event_func[key] = value
        return new_event_func
