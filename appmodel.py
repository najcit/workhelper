# -*- coding: utf-8 -*-

import glob
import os

from appbookmark import AppBookmark
from applocalapp import AppLocalApp
from apputil import get_path_time
from appresource import *


class AppModel(object):
    def __init__(self, **kwargs):
        self.name = APP_TITLE
        self.version = APP_VERSION
        self.root = APP_DEFAULT_ROOT
        self.theme = APP_DEFAULT_THEME
        self.show_view = SHOW_ICON
        self.sort_order = SORT_DEFAULT
        self.current_tag = ALL_APPS
        self.search_content = ''
        self.enable_systray = False
        self.enable_env = False
        self.auto_update = False
        self.show_local = True
        self.show_browser = True
        self.location = None
        self.icons = None
        self.apps = None
        self.browsers = None
        self.event_functions = None
        self.event_hotkeys = None
        self.initialize(**kwargs)

    def __del__(self):
        self.destroy()

    def initialize(self, **kwargs):
        if 'root' in kwargs:
            self.root = kwargs['root'] if kwargs['root'] else APP_DEFAULT_ROOT
        if 'theme' in kwargs:
            self.theme = kwargs['theme'] if kwargs['theme'] else APP_DEFAULT_THEME
        if 'sort_order' in kwargs:
            self.sort_order = kwargs['sort_order']
        if 'show_view' in kwargs:
            self.show_view = kwargs['show_view']
        if 'current_tag' in kwargs:
            self.current_tag = kwargs['current_tag']
        if 'search_content' in kwargs:
            self.search_content = kwargs['search_content']
        if 'enable_systray' in kwargs:
            self.enable_systray = kwargs['enable_systray']
        if 'enable_env' in kwargs:
            self.enable_env = kwargs['enable_env']
        if 'auto_update' in kwargs:
            self.auto_update = kwargs['auto_update']
        if 'show_local' in kwargs:
            self.show_local = kwargs['show_local']
        if 'show_browser' in kwargs:
            self.show_browser = kwargs['show_browser']
        if 'location' in kwargs:
            self.location = kwargs['location']
        if not self.icons:
            self.icons = self.init_icons()
        if not self.apps:
            self.apps = self.init_apps()
        if not self.event_hotkeys:
            self.event_hotkeys = self.init_event_hotkeys()

    def destroy(self):
        pass

    @staticmethod
    def init_icons():
        return {
            'window': os.path.join(APP_ICON_PATH, 'window', APP_ICON),
            'app': glob.glob(os.path.join(APP_ICON_PATH, 'app', '*.png')),
            'return': os.path.join(APP_ICON_PATH, 'return', 'return.png'),
            'search': os.path.join(APP_ICON_PATH, 'search', 'search.png'),
            'bookmark': os.path.join(APP_ICON_PATH, 'bookmark', 'bookmark.png'),
            'category': os.path.join(APP_ICON_PATH, 'category', 'category.png'),
        }

    def init_apps(self):
        apps_info = {ALL_APPS: [], MY_APPS: [], LOCAL_APPS: [], ALL_BOOKMARKS: []}
        my_apps = []
        for category in os.listdir(self.root):
            path = os.path.join(self.root, category)
            apps = [{
                'name': app,
                'path': os.path.join(path, app),
                'icon': 'app',
                'type': 'app',
                'time': get_path_time(os.path.join(path, app)),
            } for app in os.listdir(path)]
            apps_info[ALL_APPS] += apps
            my_apps.append({
                'name': category,
                'path': path,
                'icon': 'category',
                'type': 'category',
                'time': get_path_time(path),
                'apps': apps
            })
        apps_info[MY_APPS] = my_apps
        local_app = AppLocalApp()
        local_apps_info = local_app.get_local_installed_app_list()
        apps_info[LOCAL_APPS] = local_apps_info
        apps_info[ALL_APPS] += apps_info[LOCAL_APPS]
        browsers = local_app.get_local_installed_browsers()
        for browser in browsers:
            bookmark = AppBookmark(browser)
            apps_info[bookmark.get_bookmark_name()] = bookmark.get_bookmarks()
        return apps_info

    @staticmethod
    def init_event_hotkeys():
        return {
            E_SYSTRAY_QUIT: ['<Control-q>', '<Control-Q>'],
            E_CHECK_UPDATE: ['<Control-u>', '<Control-U>'],
            E_ABOUT: ['<Control-a>', '<Control-A>'],
            E_IMPORT_APPS_INFO: ['<Control-i>', '<Control-I>'],
            E_EXPORT_APPS_INFO: ['<Control-x>', '<Control-X>'],
            E_SELECT_ROOT: ['<Control-o>', '<Control-O>'],
            E_SELECT_THEME: ['<Control-t>', '<Control-T>'],
            E_REFRESH: ['<F5>'],
            E_MANAGE_TAG: ['<Control-m>', '<Control-M>'],
            E_RUN_APP: ['<Control-r>', '<Control-R>'],
            E_OPEN_APP_PATH: ['<Control-p>', '<Control-P>'],
            E_COPY_APP_PATH: ['<Control-c>', '<Control-C>'],
        }
