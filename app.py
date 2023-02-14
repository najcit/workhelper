# -*- coding: utf-8 -*-

import contextlib
import copy
import glob
import os
import random
import re
import winreg

import psgtray as pt
import PySimpleGUI as sg
import shutil
import pyperclip
import win32con
import win32gui
import win32ui
import winshell
import yaml
from enum import Enum
from time import localtime, strftime

from PIL import Image

from appupdater import AppUpdater
from appdatabase import AppDatabase

E_FILE = '文件(&F)'
E_IMPORT_APPS_INFO = '导入应用信息(&I)'
E_EXPORT_APPS_INFO = '导出应用信息(&X)'
E_QUIT = '退出(&Q)'
E_EDIT = '编辑(&E)'
E_MANAGE_TAG = '管理标签(&M)'
E_ADD_TAG = '新建标签(&N)'
E_MOD_TAG = '编辑标签'
E_VIEW = '视图(&V)'
E_SHOW_ICON = '图标显示(&G)'
E_SHOW_LIST = '列表显示(&L)'
E_OPTION = '选项(&S)'
E_SELECT_ROOT = '选择根目录(&O)'
E_SELECT_THEME = '选择主题(&T)'
E_SHOW_LOCAL_APPS = '显示本地应用(&S)'
E_HIDE_LOCAL_APPS = '隐藏本地应用(&H)'
E_SELECT_FONT = '选择字体(&F)'
E_SET_WINDOW = '设置窗口(&W)'
E_HELP = '帮助(&H)'
E_CHECK_UPDATE = '检查更新(&U)'
E_ABOUT = '关于(&A)'
E_RMV_TAG = '删除标签(&D)'
E_SRT_APP = '排序应用(&S)'
E_ADD_APP = '添加应用(&A)'
E_REFRESH = '刷新(&R)'
E_MOD_APP = '修改应用(&M)'
E_RMV_APP = '删除应用(&D)'
E_RUN_APP = '运行应用(&R)'
E_OPEN_APP_PATH = '打开应用路径(&O)'
E_COPY_APP_PATH = '复制应用路径(&C)'
E_THREAD = '-THREAD-'
E_DOWNLOAD_END = '-DOWNLOAD-END-'
E_SYSTRAY_SHOW = '显示界面'
E_SYSTRAY_HIDE = '隐藏界面'
E_SYSTRAY_QUIT = '退出'

APP_TITLE_CHN = '工作助手'
APP_TITLE_ENG = 'Work Helper'
ALL_APP_CHN = '全部应用'
ALL_APP_ENG = 'All App'
LOCAL_APP_CHN = '本地应用'
LOCAL_APP_ENG = 'Local App'
E_SEARCH_CHN = '搜索'
E_SEARCH_ENG = 'Search'

APP_TITLE = APP_TITLE_CHN
ALL_APP = ALL_APP_CHN
LOCAL_APP = LOCAL_APP_CHN
E_SEARCH = E_SEARCH_CHN
APP_ICON = 'workhelper.ico'
APP_DB = 'workhelper.db'
APP_VERSION = '1.0.0'
APP_DEFAULT_ROOT = 'root'
APP_DEFAULT_THEME = 'DarkGreen7'


class SortType(Enum):
    DEFAULT = 1
    ASCEND = 2
    DESCEND = 3
    PREFERRED = 4


class ViewType(Enum):
    ICON = 1
    LIST = 2


class MyApp:
    found_window = False

    def __init__(self, root, theme, enable_systray, enable_env, auto_update, **kwargs):
        self.name = APP_TITLE
        self.version = APP_VERSION
        self.icon = APP_ICON
        self.app_db = APP_DB
        self.enable_systray = enable_systray
        self.enable_env = enable_env
        self.auto_update = auto_update
        self.root = root if root and root != '' else os.path.join(os.getcwd(), APP_DEFAULT_ROOT)
        self.theme = theme if theme and theme != '' else APP_DEFAULT_THEME
        self.icon_path = os.path.join(os.getcwd(), 'images')
        self.show_view = ViewType.ICON
        self.show_local_app_list = True
        self.filter_cond = None
        self.sort_type = None
        self.location = None
        self.app_list = None
        self.app_icon_path = os.path.join(os.getcwd(), 'images', '32x32')
        self.app_icon_list = None
        self.window = None
        self.systray = None
        self.updater = AppUpdater()
        self.database = AppDatabase(self.app_db)
        self.event_functions = self.init_event_functions()
        self.event_hotkeys = self.init_event_hotkeys()
        with contextlib.suppress(Exception):
            if self.enable_env:
                self.root = os.environ['MY_ROOT']
                self.theme = os.environ['MY_THEME']
        self.init(root, theme, **kwargs)

    def __del__(self):
        print('delete')

    def init(self, root=None, theme=None, **kwargs):
        self.root = root if root and root != '' and root != self.root else self.root
        self.theme = theme if theme and theme != '' and theme != self.theme else self.theme
        self.show_view = kwargs['show_view'] if 'show_view' in kwargs else ViewType.ICON
        self.filter_cond = kwargs['filter_cond'] if 'filter_cond' in kwargs else None
        self.sort_type = kwargs['sort_type'] if 'sort_type' in kwargs else None
        self.location = kwargs['location'] if 'location' in kwargs else None
        self.app_icon_list = self.init_app_icon_list()
        self.app_list = self.init_app_list()
        self.init_window()
        self.init_systray()
        return self

    # noinspection PyArgumentList
    def run(self):
        while self.window:
            event, values = self.window.read(timeout_key=sg.TIMEOUT_KEY)
            if self.systray and event == self.systray.key:
                event = values[event]
            print(event, values, isinstance(event, tuple), type(event))
            if event in self.event_functions:
                self.event_functions[event](event=event, values=values)
            else:
                self.event_functions['default'](event=event, values=values)

    def destroy(self):
        if self.window:
            self.window.close()
            self.window = None
        if self.systray:
            self.systray.close()
            self.systray = None

    def refresh(self, **kwargs):
        root = kwargs['root'] if 'root' in kwargs else None
        theme = kwargs['theme'] if 'theme' in kwargs else None
        self.destroy()
        self.init(root, theme, **kwargs)
        return self

    def attempt_exit(self, **_kwargs):
        if self.systray:
            self.window.hide()
            self.systray.show_icon()
        else:
            self.exit()

    def exit(self, **_kwargs):
        self.database.close()
        self.destroy()

    def show_top(self, **_kwargs):
        self.window.un_hide()
        self.window.bring_to_front()

    def select_root(self, **_kwargs):
        root = sg.popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if os.path.isdir(root):
            self.refresh(root=root)

    def select_theme(self, **_kwargs):
        cur_theme = self.theme
        layout = [[sg.Text('See how elements look under different themes by choosing a different theme here!')],
                  [sg.Listbox(values=sg.theme_list(), default_values=[cur_theme], size=(20, 12), key='THEME_LISTBOX',
                              enable_events=True)],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('设置主题', layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                cur_theme = values['THEME_LISTBOX'][0]
                break
        window.close()
        if cur_theme != self.theme:
            self.set_env('MY_THEME', cur_theme)
            self.refresh(theme=cur_theme)
        return self

    def check_update(self, **_kwargs):
        if self.updater.is_new_version(self.version):
            # 检查是否为最新版本
            text = '当前软件已是最新版本，版本号: {version}'.format(version=self.version)
            sg.popup_no_buttons(text, title=self.name, auto_close=True, auto_close_duration=10)
        else:
            # 下载最新版本, 更新下载进度
            async def update_progress(percentage):
                progressbar = self.window['-NOTIFICATION-']
                value = '最新版软件正在下载，进度为 {percentage}%。'.format(percentage=percentage)
                progressbar.update(value=value)
                if percentage >= 100:
                    value = '最新版软件下载已完成。'
                    progressbar.update(value=value)
                    self.window.write_event_value((E_THREAD, E_DOWNLOAD_END), percentage)

            self.window.start_thread(lambda: self.updater.install(update_progress_func=update_progress), ())

    def about(self, **_kwargs):
        text = '''\n当前已是最新版本: {version}\nCopyright © 2023–2026 by lidajun\n'''.format(version=self.version)
        sg.popup_no_buttons(text, title=self.name, auto_close=True, auto_close_duration=10)

    def import_apps_info(self, **_kwargs):
        result = False
        file = sg.popup_get_file(message='导入文件', title='选择导入文件', keep_on_top=True)
        if file:
            with open(str(file), "r", encoding='utf8') as file:
                app_cfg = yaml.load(file, yaml.FullLoader)
                for tab, apps in app_cfg.items():
                    tab_path = os.path.join(self.root, tab)
                    if not os.path.exists(tab_path):
                        os.mkdir(tab_path)
                    for app in apps:
                        app_path = os.path.join(tab_path, app)
                        if not os.path.exists(app_path):
                            os.mkdir(app_path)
            result = True
        if result:
            self.refresh()
        return result

    def export_apps_info(self, **_kwargs):
        result = False
        directory = sg.popup_get_folder(message='导出路径', title='选择导出路径', keep_on_top=True)
        if directory:
            file_path = os.path.join(directory, 'workhelper.yml')
            app_cfg = {}
            for tab, apps in self.app_list.items():
                if tab == '全部应用':
                    continue
                tab = os.path.basename(tab)
                apps = [app['name'] for app in apps]
                app_cfg[tab] = apps
            with open(file_path, 'w', encoding='utf-8') as file:
                yaml.dump(app_cfg, file, allow_unicode=True)
            result = True
        return result

    def show_icon(self, **_kwargs):
        self.refresh(location=self.location, show_view=ViewType.ICON)

    def show_list(self, **_kwargs):
        self.refresh(location=self.location, show_view=ViewType.LIST)

    def manage_tag(self, **_kwargs):
        values = [os.path.basename(tab) for tab, _ in self.app_list.items()]
        cur_tab = os.path.basename(self.window['-TAB-GROUP-'].get())
        col1 = sg.Column([[sg.Listbox(values, key='-TAB-LIST-', default_values=[cur_tab],
                                      select_mode=sg.LISTBOX_SELECT_MODE_SINGLE, enable_events=True, size=(30, 40))]],
                         size=(200, 400))
        col2 = sg.Column([[sg.Button('新建', key='-ADD-')],
                          [sg.Button('编辑', key='-EDIT-')],
                          [sg.Button('删除', key='-DELETE-')],
                          [sg.Button('导入', key='-IMPORT-')],
                          [sg.Button('导出', key='-EXPORT-')]],
                         size=(50, 350))
        col3 = [sg.Button('确定', key='-OK-'), sg.Button('取消', key='-CANCEL-')]
        layout = [[col1, col2], [col3]]
        window = sg.Window('管理标签', layout, keep_on_top=True, element_justification='right', finalize=True)

        def disable_buttons(tab):
            if tab == '全部应用':
                keys = ['-EDIT-', '-DELETE-']
            else:
                keys = []
            for key in ['-ADD-', '-EDIT-', '-DELETE-', '-IMPORT-', '-EXPORT-']:
                if key in keys:
                    window[key].update(disabled=True)
                else:
                    window[key].update(disabled=False)

        disable_buttons(cur_tab)
        while True:
            event, values = window.read()
            print(event, values)
            if event in (sg.WIN_CLOSED, '-CANCEL-'):
                break
            elif event == '-OK-':
                break
            elif event == '-ADD-':
                new_tab = self.add_tag()
                print(type(new_tab))
                if isinstance(new_tab, str):
                    values = window['-TAB-LIST-'].get_list_values()
                    values.append(new_tab)
                    index = values.index(new_tab)
                    window['-TAB-LIST-'].update(values, set_to_index=index)
            elif event == '-EDIT-':
                cur_tab = window['-TAB-LIST-'].get()[0]
                print('cur_tab', cur_tab)
                tab_key = os.path.join(self.root, cur_tab)
                new_tab = self.mod_tag(tab_key)
                if isinstance(new_tab, str):
                    values = window['-TAB-LIST-'].get_list_values()
                    index = values.index(cur_tab)
                    values[index] = new_tab
                    window['-TAB-LIST-'].update(values, set_to_index=index)
            elif event == '-DELETE-':
                cur_tab = window['-TAB-LIST-'].get()[0]
                tab_key = os.path.join(self.root, cur_tab)
                if self.rmv_tag(tab_key):
                    values = window['-TAB-LIST-'].get_list_values()
                    values.remove(cur_tab)
                    window['-TAB-LIST-'].update(values)
            elif event == '-IMPORT-':
                if self.import_apps_info():
                    values = [os.path.basename(tab) for tab, _ in self.app_list.items()]
                    window['-TAB-LIST-'].update(values)
            elif event == '-EXPORT-':
                self.export_apps_info()
            elif event == '-TAB-LIST-':
                tab = values[event][0]
                disable_buttons(tab)
            else:
                print(event, values)
        window.close()

    def add_tag(self, **_kwargs):
        result = False
        layout = [[sg.Text('新标签'), sg.Input(key='-NEW-TAB-')],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('新建标签', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                tab_name = window['-NEW-TAB-'].get()
                tab_path = os.path.join(self.root, tab_name)
                if not os.path.exists(tab_path):
                    os.mkdir(tab_path)
                result = tab_name
                sg.popup('成功添加' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def mod_tag(self, tab_key=None, **kwargs):
        if tab_key is None:
            tab_key = kwargs['values']['-TAB-GROUP-']
        result = False
        tab_name = os.path.basename(tab_key)
        tab_path = tab_key
        layout = [[sg.Text('当前标签'), sg.Input(key='CurTab', default_text=tab_name)],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('编辑标签', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                tab_name = window['CurTab'].get()
                if os.path.exists(tab_path):
                    os.renames(tab_path, os.path.join(self.root, tab_name))
                result = tab_name
                sg.popup('成功修改为' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def rmv_tag(self, tab_key=None, **kwargs):
        if tab_key is None:
            tab_key = kwargs['values']['-TAB-GROUP-']
        tab_name = os.path.basename(tab_key)
        tab_dir = tab_key
        result = sg.popup('是否删除' + tab_name + '?', keep_on_top=True, button_type=sg.POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(tab_dir)
            self.refresh()
        return result

    def sort_app(self, **_kwargs):
        if self.sort_type == SortType.DEFAULT:
            self.sort_type = SortType.ASCEND
        elif self.sort_type == SortType.ASCEND:
            self.sort_type = SortType.DESCEND
        elif self.sort_type == SortType.DESCEND:
            self.sort_type = SortType.DEFAULT
        self.refresh()

    def select_app(self, **kwargs):
        event = kwargs['event']
        key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
        self.window[key].set_focus()

    def add_app(self, tab_key=None, **kwargs):
        if tab_key is None:
            tab_key = kwargs['values']['-TAB-GROUP-']
        result = False
        tab_dir = tab_key
        layout = [[sg.Text('添加目标'), sg.Input(key='TargetPath'), sg.Button('浏览')],
                  [sg.Text('添加应用'), sg.Input(key='AppName'), sg.Button('图标')],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('添加应用', layout, element_justification='right')
        while True:
            event, values = window.read()
            if event == '浏览':
                file = sg.popup_get_file('选择目标文件', keep_on_top=True)
                if file:
                    window['TargetPath'].update(value=file)
                    app_name = os.path.basename(str(file)).capitalize().removesuffix('.exe')
                    window['AppName'].update(value=app_name)
            elif event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                target = window['TargetPath'].get()
                app_name = window['AppName'].get()
                app_dir = os.path.join(tab_dir, app_name)
                if not os.path.exists(app_dir):
                    os.mkdir(app_dir)
                link_path = os.path.join(app_dir, app_name + '.lnk')
                # print('target', target)
                # print('link_path', link_path)
                winshell.CreateShortcut(Path=link_path, Target=target, Icon=(target, 0))
                sg.popup('成功添加应用' + app_name + '!', no_titlebar=True, keep_on_top=True, auto_close=True,
                         auto_close_duration=2)
                result = True
                break
        window.close()
        if result:
            self.refresh()
        return result

    def mod_app(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        result = False
        app_dir = app_key.split('#')[1]
        app_name = os.path.basename(app_dir)
        layout = [[sg.Text('应用名称'), sg.Input(key='AppName', default_text=app_name)],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('编辑应用', layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                new_app_name = window['AppName'].get()
                os.renames(app_dir, os.path.join(os.path.dirname(app_dir), new_app_name))
                sg.popup(app_name + '成功修改成' + new_app_name + '!', no_titlebar=True, keep_on_top=True,
                         auto_close=True,
                         auto_close_duration=3)
                result = True
                break
        window.close()
        if result:
            self.refresh()
        return result

    def rmv_app(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        app_dir = app_key.split('#')[1]
        app_name = os.path.basename(app_dir)
        result = sg.popup('是否删除' + app_name + '?', no_titlebar=True, keep_on_top=True,
                          button_type=sg.POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(app_dir)
            self.refresh()
        return result

    def open_app_path(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        # print(cur_path)
        if os.path.exists(cur_path):
            os.startfile(cur_path)

    def run_app(self, **_kwargs):
        item = self.window.find_element_with_focus()
        if str(item).find('Button') > -1:
            app_path = item.metadata
            with contextlib.suppress(Exception):
                if os.path.isfile(app_path):
                    app = app_path
                elif os.path.isdir(app_path):
                    glob.glob('*.[lnk]')
                    for app_name in os.listdir(app_path):
                        if app_name.endswith((".lnk", ".exe")):
                            app = os.path.join(app_path, app_name)
                            break
                os.startfile(app)
                self.database.start_app(app_path)

    def copy_app_path(self, **_kwargs):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        pyperclip.copy(cur_path)

    @staticmethod
    def show_local_app_tag():
        print('show')

    def search(self, **kwargs):
        values = kwargs['values']
        tag = values['-TAB-GROUP-']
        content = values['-CONTENT-']
        filter_cond = {'tag': tag, 'content': content}
        if filter_cond != self.filter_cond:
            self.refresh(filter_cond=filter_cond)

    def default_action(self, **kwargs):
        event = kwargs['event']
        if isinstance(event, str) and event.endswith('Click'):
            self.select_app(event=event)
        elif isinstance(event, tuple) and event == (E_THREAD, E_DOWNLOAD_END):
            self.exit()
        else:
            self.run_app()

    def bind_hotkeys(self):
        for tag, apps in self.app_list.items():
            for app in apps:
                key = app['key']
                self.window[key].bind('<Button-1>', ' LeftClick')
                self.window[key].bind('<Button-2>', ' MiddleClick')
                self.window[key].bind('<Button-3>', ' RightClick')
        for event, hotkeys in self.event_hotkeys.items():
            for hotkey in hotkeys:
                self.window.bind(hotkey, event)

    def init_window(self):
        sg.theme(self.theme)
        show_icon = E_SHOW_ICON if self.show_view == ViewType.LIST else '!' + E_SHOW_ICON
        show_list = E_SHOW_LIST if self.show_view == ViewType.ICON else '!' + E_SHOW_LIST
        show_local_app_list = E_SHOW_LOCAL_APPS if not self.show_local_app_list else E_HIDE_LOCAL_APPS
        menu_def = [[E_FILE, [E_IMPORT_APPS_INFO, E_EXPORT_APPS_INFO, E_QUIT]],
                    [E_EDIT, [E_MANAGE_TAG, E_ADD_TAG]],
                    [E_VIEW, [show_icon, show_list, E_SRT_APP, E_REFRESH]],
                    [E_OPTION, [E_SELECT_ROOT, E_SELECT_THEME, show_local_app_list, E_SELECT_FONT, E_SET_WINDOW]],
                    [E_HELP, [E_CHECK_UPDATE, E_ABOUT]]]
        layout = [[sg.Menu(menu_def, key='-MENU-BAR-')]]
        col_layout = [[sg.Input(enable_events=False, key='-CONTENT-', expand_x=True, expand_y=True),
                       sg.Button(E_SEARCH, key=E_SEARCH, bind_return_key=True)]]
        layout += [[sg.Column(col_layout, expand_x=True)]]
        filter_tag = self.filter_cond['tag'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        filter_content = self.filter_cond['content'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        layout += self.init_content(filter_tag, filter_content)
        col_layout = [[sg.Text('欢迎使用' + self.name, key='-NOTIFICATION-'), sg.Sizegrip()]]
        layout += [[sg.Column(col_layout, expand_x=True)]]
        if self.location:
            self.window = sg.Window(self.name, layout, location=self.location, enable_close_attempted_event=True,
                                    resizable=True, finalize=True, icon=self.icon, font='Courier 12')
        else:
            self.window = sg.Window(self.name, layout, enable_close_attempted_event=True,
                                    resizable=True, finalize=True, icon=self.icon, font='Courier 12')
        self.window.set_min_size((300, 500))
        self.location = self.window.current_location()
        if filter_tag:
            self.window[filter_tag].select()
        if filter_content:
            self.window['-CONTENT-'].update(filter_content)
        self.bind_hotkeys()
        return self.window

    def init_systray(self):
        menu = [[], [E_SYSTRAY_SHOW, E_SYSTRAY_HIDE, E_SYSTRAY_QUIT]]
        if self.enable_systray and self.window:
            self.systray = pt.SystemTray(menu, icon=self.icon, window=self.window, tooltip=self.name,
                                         single_click_events=False)
        return self.systray

    def init_content(self, filter_tag, filter_content):
        tab_group = {}
        max_num_per_row = 1 if self.show_view == ViewType.LIST else 4
        app_right_click_menu = [[], [E_RUN_APP, E_MOD_APP, E_RMV_APP, E_OPEN_APP_PATH, E_COPY_APP_PATH]]
        tag_right_click_menu = [[], [E_ADD_TAG, E_RMV_TAG, E_SRT_APP, E_ADD_APP, E_REFRESH]]
        for tag, apps in self.app_list.items():
            tab_layout = []
            row_layout = []
            is_filter = True if tag == filter_tag else False
            if self.sort_type == SortType.ASCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper())
            elif self.sort_type == SortType.DESCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper(), reverse=True)
            for app in apps:
                col_layout = []
                # print(app)
                app_name, app_icon, app_path, app_key = app['name'], app['icon'], app['path'], app['key']
                if app_icon and os.path.isfile(app_icon):
                    icon = sg.Image(source=app_icon, tooltip=app_name, key='[icon]' + app_key, enable_events=True,
                                    size=(32, 32), background_color='white', right_click_menu=app_right_click_menu)
                    col_layout.append(icon)
                button = sg.Button(button_text=app_name[:20], tooltip=app_name, key=app_key, metadata=app_path,
                                   size=(20, 1), button_color=('black', 'white'), enable_events=False, border_width=0,
                                   right_click_menu=app_right_click_menu, expand_y=True)
                col_layout.append(button)
                if self.show_view == ViewType.LIST:
                    time = sg.Text(app['time'], background_color='white', expand_x=True)
                    col_layout.append(time)
                    path = sg.Text(text=app_path[:35], tooltip=app_path, size=(35, 1), justification='left',
                                   background_color='white', enable_events=True)
                    col_layout.append(path)
                    col = sg.Column([col_layout], background_color='white', expand_x=True, expand_y=True,
                                    right_click_menu=app_right_click_menu, grab=True)
                else:
                    col = sg.Column([col_layout], background_color='white')
                if (is_filter and re.search(filter_content, app_name, re.IGNORECASE)) or not is_filter:
                    row_layout.append(col)
                if len(row_layout) == max_num_per_row:
                    tab_layout.append(copy.deepcopy(row_layout))
                    row_layout.clear()
            else:
                if len(row_layout) > 0:
                    tab_layout.append(row_layout)
            tab_group[tag] = [[sg.Column(layout=tab_layout, scrollable=True, vertical_scroll_only=True,
                                         expand_x=True, expand_y=True)]]
        tab_group_layout = [[sg.Tab(title=os.path.basename(tab_key), layout=tab_layout, key=tab_key,
                                    right_click_menu=tag_right_click_menu)
                             for tab_key, tab_layout in tab_group.items()]]
        return [[sg.TabGroup(tab_group_layout, key='-TAB-GROUP-', expand_x=True, expand_y=True,
                             right_click_menu=tag_right_click_menu)]]

    def init_app_icon_list(self):
        return glob.glob(os.path.join(self.app_icon_path, '*.png'))

    def init_app_list(self):
        apps_info = {ALL_APP: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in apps_info:
                    apps_info[parent] = []
                apps = [{
                    'name': directory,
                    'key': parent + '#' + os.path.join(parent, directory),
                    'path': os.path.join(parent, directory),
                    'icon': random.choice(self.app_icon_list),
                    'time': self.timestamp_to_datetime(os.path.getmtime(os.path.join(parent, directory))),
                } for directory in directorys]
                apps_info[parent] += apps
                apps_info[ALL_APP] += apps
        if self.show_local_app_list:
            apps_info[LOCAL_APP] = self.init_local_installed_app_list()
        return apps_info

    def init_event_functions(self):
        new_event_func = {'default': self.default_action}
        event_func = {
            sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED: self.show_top,
            sg.WIN_CLOSE_ATTEMPTED_EVENT: self.attempt_exit,
            E_SYSTRAY_SHOW: self.show_top,
            E_SYSTRAY_HIDE: self.attempt_exit,
            E_SYSTRAY_QUIT: self.exit,
            E_QUIT: self.exit,
            E_CHECK_UPDATE: self.check_update,
            E_ABOUT: self.about,
            E_IMPORT_APPS_INFO: self.import_apps_info,
            E_EXPORT_APPS_INFO: self.export_apps_info,
            E_SHOW_ICON: self.show_icon,
            E_SHOW_LIST: self.show_list,
            E_SELECT_ROOT: self.select_root,
            E_SELECT_THEME: self.select_theme,
            E_SHOW_LOCAL_APPS: self.show_local_app_tag,
            E_SEARCH: self.search,
            E_REFRESH: self.refresh,
            E_MANAGE_TAG: self.manage_tag,
            E_ADD_TAG: self.add_tag,
            E_MOD_TAG: self.mod_tag,
            E_RMV_TAG: self.rmv_tag,
            E_SRT_APP: self.sort_app,
            E_ADD_APP: self.add_app,
            E_MOD_APP: self.mod_app,
            E_RMV_APP: self.rmv_app,
            E_RUN_APP: self.run_app,
            E_OPEN_APP_PATH: self.open_app_path,
            E_COPY_APP_PATH: self.copy_app_path,
        }
        for key, value in event_func.items():
            new_event_func[key] = value
            if '&' in key:
                new_event_func[key.replace('&', '')] = value
        return new_event_func

    @staticmethod
    def show(root, theme, enable_systray, enable_env, auto_update):
        def show_window(hwnd, wildcard):
            if re.match(wildcard, str(win32gui.GetWindowText(hwnd))):
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                MyApp.found_window = True

        win32gui.EnumWindows(show_window, APP_TITLE)
        if not MyApp.found_window:
            MyApp.start(root, theme, enable_systray, enable_env, auto_update)

    @staticmethod
    def start(root, theme, enable_systray, enable_env, auto_update, **kwargs):
        app = MyApp(root, theme, enable_systray, enable_env, auto_update, **kwargs)
        app.run()

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

    @staticmethod
    def set_env(key, value):
        if os.getenv(key) != value:
            command = f'start /b setx {key} {value}'.format(key=key, value=value)
            os.popen(command)

    @staticmethod
    def timestamp_to_datetime(time_stamp, format_string="%Y-%m-%d %H:%M:%S"):
        time_array = localtime(time_stamp)
        str_date = strftime(format_string, time_array)
        return str_date

    @staticmethod
    def make_png_from_app(icon_filepath, png_path):
        filename = os.path.basename(icon_filepath).split('.')[0] + '.png'
        png_filename = os.path.join(png_path, filename)
        if os.path.exists(png_filename):
            return png_filename
        if icon_filepath.find(".exe") != -1:
            large, _small = win32gui.ExtractIconEx(icon_filepath, 0)
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, 32, 32)
            hdc = hdc.CreateCompatibleDC()
            hdc.SelectObject(hbmp)
            hdc.DrawIcon((0, 0), large[0])
            bmp_bits = hbmp.GetBitmapBits(True)
            img = Image.frombuffer('RGBA', (32, 32), bmp_bits, 'raw', 'BGRA', 0, 1)
            img.save(png_filename)
        elif icon_filepath.find('.ico') != -1:
            img = Image.open(icon_filepath)
            img = img.resize((32, 32), Image.LANCZOS)
            img.save(png_filename)
        else:
            print(f'Could not support this file:{icon_filepath}')
            png_filename = ''
        return png_filename

    def init_local_installed_app_list(self):
        app_list = []
        key_name_list = [r'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                         r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall']
        for key_name in key_name_list:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_name, 0, winreg.KEY_READ)
            key_info = winreg.QueryInfoKey(key)[0]
            for i in range(key_info):
                app = {}
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    app['name'] = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    app['path'] = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                    app['key'] = app['path']
                    if not app['path']:
                        continue
                except EnvironmentError:
                    continue
                try:
                    app['icon'] = winreg.QueryValueEx(subkey, "DisplayIcon")[0].split(',')[0].replace('"', '')
                    if app['icon'].find('.exe') != -1:
                        app['path'] = app['icon']
                    app['icon'] = self.make_png_from_app(app['icon'], os.path.join(self.app_icon_path, 'local'))
                    if not app['icon']:
                        continue
                except EnvironmentError:
                    app['icon'] = ''
                try:
                    app['time'] = winreg.QueryValueEx(subkey, "InstallDate")[0]
                except EnvironmentError:
                    app['time'] = ''
                try:
                    app['version'] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                except EnvironmentError:
                    app['version'] = ''
                try:
                    app['publisher'] = winreg.QueryValueEx(subkey, "Publisher")[0]
                except EnvironmentError:
                    app['publisher'] = ''
                app_list.append(app)
        return app_list
