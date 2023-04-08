# -*- coding: utf-8 -*-

import glob
import os

from appbookmark import AppBookmark
from applocalapp import AppLocalApp
from appmyapp import AppMyApp
from appresource import *
from appshareapp import AppSharedApp


class AppModel(object):

    def __init__(self, **kwargs):
        self.lang = kwargs['lang'] if 'lang' in kwargs else LANG_ENG
        self.name = APP_TITLE[self.lang]
        self.font = APP_FONT[self.lang]
        self.version = APP_VERSION
        self.root = APP_ROOT
        self.theme = APP_THEME
        self.enable_systray = kwargs['enable_systray'] if 'enable_systray' in kwargs else False
        self.enable_env = kwargs['enable_env'] if 'enable_env' in kwargs else False
        self.auto_update = kwargs['auto_update'] if 'auto_update' in kwargs else False
        self.show_view = SHOW_ICON
        self.sort_order = SORT_DEFAULT
        self.current_tag = ALL_APPS[self.lang]
        self.search_content = ''
        self.show_local = True
        self.show_shared = False
        self.show_browser = True
        self.location = None
        self.icons = None
        self.apps = None
        self.browsers = None
        self.event_hotkeys = None
        if self.enable_env:
            self.root = os.getenv('MY_ROOT', APP_ROOT)
            self.theme = os.getenv('MY_THEME', APP_THEME)
            font = os.getenv('MY_FONT', APP_FONT)
            if isinstance(font, str):
                font = font.split()
                self.font = (font[0], font[1])
            else:
                self.font = font
        self.init_event_hotkeys()
        self.initialize(**kwargs)

    def __del__(self):
        self.destroy()

    def initialize(self, **kwargs):
        if 'lang' in kwargs:
            self.lang = kwargs['lang'] or LANG_CHN
            self.name = APP_TITLE[self.lang]
            self.font = APP_FONT[self.lang]
            self.current_tag = ALL_APPS[self.lang]
            self.apps = self.init_apps()
        if 'root' in kwargs:
            self.root = kwargs['root'] or APP_ROOT
        if 'theme' in kwargs:
            self.theme = kwargs['theme'] or APP_THEME
        if 'font' in kwargs:
            self.font = kwargs['font'] or APP_FONT
        if 'sort_order' in kwargs:
            self.sort_order = kwargs['sort_order']
        if 'show_view' in kwargs:
            self.show_view = kwargs['show_view']
        if 'current_tag' in kwargs:
            self.current_tag = kwargs['current_tag']
        if 'search_content' in kwargs:
            self.search_content = kwargs['search_content']
        if 'show_local' in kwargs:
            self.show_local = kwargs['show_local']
        if 'show_shared' in kwargs:
            self.show_shared = kwargs['show_shared']
        if 'show_browser' in kwargs:
            self.show_browser = kwargs['show_browser']
        if 'location' in kwargs:
            self.location = kwargs['location']
        if not self.icons:
            self.icons = self.init_icons()
        if not self.apps:
            self.apps = self.init_apps()

    def destroy(self):
        pass

    @staticmethod
    def init_icons():
        return {
            'window': os.path.join(APP_ICON_PATH, 'window', APP_ICON),
            'default': glob.glob(os.path.join(APP_ICON_PATH, 'default', '*.png')),
            'return': os.path.join(APP_ICON_PATH, 'button', 'return.png'),
            'search': os.path.join(APP_ICON_PATH, 'button', 'search.png'),
            'bookmark': os.path.join(APP_ICON_PATH, 'folder', 'bookmark.png'),
            'category': os.path.join(APP_ICON_PATH, 'folder', 'category.png'),
        }

    def init_apps(self):
        apps_info = {ALL_APPS[self.lang]: [], MY_APPS[self.lang]: [],
                     LOCAL_APPS[self.lang]: [], SHARED_APPS[self.lang]: [],
                     ALL_BOOKMARKS[self.lang]: []}

        my_app = AppMyApp()
        my_apps_info = my_app.get_apps(self.root)
        apps_info[MY_APPS[self.lang]] = my_apps_info
        apps_info[ALL_APPS[self.lang]] += my_apps_info

        local_app = AppLocalApp()
        local_apps_info = local_app.get_apps()
        apps_info[LOCAL_APPS[self.lang]] = local_apps_info
        apps_info[ALL_APPS[self.lang]] += local_apps_info

        shared_app = AppSharedApp()
        shared_apps_info = shared_app.get_apps()
        apps_info[SHARED_APPS[self.lang]] = shared_apps_info
        apps_info[ALL_APPS[self.lang]] += shared_apps_info

        browsers = local_app.get_browsers()
        bookmark = AppBookmark(self.lang)
        for browser in browsers:
            apps_info[bookmark.name(browser)] = bookmark.get_bookmarks(browser)
            apps_info[ALL_BOOKMARKS[self.lang]] += apps_info[bookmark.name(browser)]

        return apps_info

    def init_event_hotkeys(self):
        self.event_hotkeys = {
            E_SYSTRAY_QUIT[self.lang]: ['<Control-q>', '<Control-Q>'],
            E_CHECK_UPDATE[self.lang]: ['<Control-u>', '<Control-U>'],
            E_ABOUT[self.lang]: ['<Control-a>', '<Control-A>'],
            E_IMPORT_APPS_INFO[self.lang]: ['<Control-i>', '<Control-I>'],
            E_EXPORT_APPS_INFO[self.lang]: ['<Control-x>', '<Control-X>'],
            E_SELECT_ROOT[self.lang]: ['<Control-o>', '<Control-O>'],
            E_SELECT_THEME[self.lang]: ['<Control-t>', '<Control-T>'],
            E_REFRESH[self.lang]: ['<F5>'],
            E_MANAGE_TAG[self.lang]: ['<Control-m>', '<Control-M>'],
            E_RUN_APP[self.lang]: ['<Control-r>', '<Control-R>'],
            E_OPEN_APP_PATH[self.lang]: ['<Control-p>', '<Control-P>'],
            E_COPY_APP_PATH[self.lang]: ['<Control-c>', '<Control-C>'],
        }
