# -*- coding: utf-8 -*-

import contextlib
import copy
import glob
import os
import random
import re
import sys

import PySimpleGUI as sg
import psgtray as pt
import math
import shutil
import pyperclip
import winshell
import yaml
from time import localtime, strftime
from enum import Enum
from appupdater import AppUpdater
from appdatabase import AppDatabase


E_FILE = '文件(&F)'
E_NEW_WINDOW = '新建窗口(&N)'
E_CLOSE_WINDOW = '关闭窗口(&E)'
E_IMPORT_APP_INFO = '导入应用信息(&I)'
E_EXPORT_APP_INFO = '导出应用信息(&X)'
E_QUIT = '退出(&Q)'
E_EDIT = '编辑(&E)'
E_MANAGE_TAG = '管理标签(&M)'
E_ADD_TAG = '新建标签(&N)'
E_MOD_TAG = '编辑标签'
E_VIEW = '视图(&V)'
E_SHOW_ICON = '图标显示(&G)'
E_SHOW_LIST = '列表显示(&L)'
E_OPTION = '选项(&S)'
E_SELECT_ROOT = '选择目录(&O)'
E_SELECT_THEME = '选择主题(&T)'
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
E_SEARCH = '搜索'
E_THREAD = '-THREAD-'
E_DOWNLOAD_END = '-DOWNLOAD-END-'
E_SYSTRAY_SHOW = '显示界面'
E_SYSTRAY_HIDE = '隐藏界面'
E_SYSTRAY_QUIT = '退出'

APP_CHN_TITLE = '工作助手'
APP_ENG_TITLE = 'Work Helper'
APP_TITLE = APP_CHN_TITLE
APP_ICON = 'workhelper.ico'
APP_DB = 'workhelper.db'
APP_VERSION = '1.0.0'
ALL_TAG = '全部应用'


class SortType(Enum):
    DEFAULT = 1
    ASCEND = 2
    DESCEND = 3
    PREFERRED = 4


class MyApp:

    def __init__(self, root, theme, enable_systray, enable_env, location, name=None):
        self.name = name if name is not None else APP_TITLE
        self.version = APP_VERSION
        self.icon = APP_ICON
        self.app_db = APP_DB
        self.root = root
        self.theme = theme
        self.enable_systray = enable_systray
        self.enable_env = enable_env
        self.location = location
        self.app_icons = []
        self.list_view = False
        self.sort_type = SortType.DEFAULT
        self.filter_cond = None
        self.window = None
        self.systray = None
        self.updater = AppUpdater()
        self.database = AppDatabase(self.app_db)
        self.apps_info = self.init_apps_info()
        self.event_functions = self.init_event_functions()
        self.event_hotkeys = self.init_event_hotkeys()

    @staticmethod
    def show(root, theme, enable_systray, enable_env, location=None):
        app = MyApp(root, theme, enable_systray, enable_env, location)
        app.init()
        app.run()

    def init(self):
        if self.enable_env:
            with contextlib.suppress(Exception):
                self.root = os.environ['MY_ROOT']
                self.theme = os.environ['MY_THEME']
        if self.root == 'root':
            self.root = os.path.join(os.getcwd(), self.root)
        self.init_window()
        self.init_systray()
        return self

    def refresh(self, root=None, theme=None, filter_cond=None, location=None, **_kwargs):
        self.close()
        if root and root != self.root:
            self.root = root
            self.set_env('MY_ROOT', self.root)
        if theme and theme != self.theme:
            self.theme = theme
            self.set_env('MY_THEME', self.theme)
        self.filter_cond = filter_cond if filter_cond else None
        self.location = location if location else None
        self.init()
        return self

    def run(self):
        while True:
            event, values = self.read()
            print(event, values, isinstance(event, tuple), type(event))
            if event in self.event_functions and callable(self.event_functions[event]):
                # noinspection PyArgumentList
                self.event_functions[event](event=event, values=values)
            else:
                if isinstance(event, str) and event.endswith('Click'):
                    self.select_app(event=event)
                elif isinstance(event, tuple) and event == (E_THREAD, E_DOWNLOAD_END):
                    self.exit()
                else:
                    self.run_app()

    def read(self, timeout=None, timeout_key=sg.TIMEOUT_KEY, close=False):
        event, values = self.window.read(timeout, timeout_key, close)
        if self.systray and event == self.systray.key:
            event = values[event]
        return event, values

    def close(self):
        if self.window:
            self.window.close()
        if self.systray:
            self.systray.close()

    def show_top(self, **_kwargs):
        self.window.un_hide()
        self.window.bring_to_front()

    def attempt_exit(self, **_kwargs):
        if self.systray:
            self.window.hide()
            self.systray.show_icon()
        else:
            self.exit()

    def exit(self, **_kwargs):
        self.database.close()
        self.close()
        sys.exit(0)

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
            self.theme = cur_theme
            self.set_env('MY_THEME', self.theme)
            self.refresh()
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

    def new_window(self, **_kwargs):
        location = (self.location[0] + 20, self.location[1] + 20)
        MyApp.show(self.root, self.theme, self.enable_systray, self.enable_env, location)

    def manage_tag(self, **_kwargs):
        values = [os.path.basename(tab) for tab, _ in self.apps_info.items()]
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
                if self.import_app():
                    values = [os.path.basename(tab) for tab, _ in self.apps_info.items()]
                    window['-TAB-LIST-'].update(values)
            elif event == '-EXPORT-':
                self.export_app()
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

    def search(self, **kwargs):
        values = kwargs['values']
        tab = values['-TAB-GROUP-']
        content = values['-CONTENT-']
        filter_cond = {'tab': tab, 'content': content}
        if filter_cond != self.filter_cond:
            self.refresh(filter_cond=filter_cond)

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
            app_key = item.key
            with contextlib.suppress(Exception):
                app_dir = app_key.split('#')[1]
                for app_name in os.listdir(app_dir):
                    if app_name.endswith('.lnk'):
                        app = os.path.join(app_dir, app_name)
                        os.startfile(app)
                        break
                    elif app_name.endswith('.exe'):
                        app = os.path.join(app_dir, app_name)
                        os.popen(app)
                        break
            self.database.start_app(app_dir)

    def copy_app_path(self, **_kwargs):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        pyperclip.copy(cur_path)

    def import_app(self, **_kwargs):
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

    def export_app(self, **_kwargs):
        result = False
        directory = sg.popup_get_folder(message='导出路径', title='选择导出路径', keep_on_top=True)
        if directory:
            file_path = os.path.join(directory, 'workhelper.yml')
            app_cfg = {}
            for tab, apps in self.apps_info.items():
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
        self.list_view = False
        self.refresh(location=self.location)

    def show_list(self, **_kwargs):
        self.list_view = True
        self.refresh(location=self.location)

    def select_root(self, **_kwargs):
        root = sg.popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if os.path.isdir(root):
            self.refresh(root)

    def init_window(self):
        sg.theme(self.theme)
        show_icon = E_SHOW_ICON if self.list_view else '!' + E_SHOW_ICON
        show_list = E_SHOW_LIST if not self.list_view else '!' + E_SHOW_LIST
        menu_def = [[E_FILE, [E_NEW_WINDOW, E_CLOSE_WINDOW, E_IMPORT_APP_INFO, E_EXPORT_APP_INFO, E_QUIT]],
                    [E_EDIT, [E_MANAGE_TAG, E_ADD_TAG]],
                    [E_VIEW, [show_icon, show_list, E_SRT_APP, E_REFRESH]],
                    [E_OPTION, [E_SELECT_ROOT, E_SELECT_THEME, E_SELECT_FONT, E_SET_WINDOW]],
                    [E_HELP, [E_CHECK_UPDATE, E_ABOUT]]]
        layout = [[sg.Menu(menu_def, key='-MENU-BAR-')]]
        col_layout = [[sg.Input(enable_events=False, key='-CONTENT-', expand_x=True, expand_y=True),
                       sg.Button(E_SEARCH, key='-SEARCH-', bind_return_key=True)]]
        layout += [[sg.Column(col_layout, expand_x=True)]]
        filter_tag = self.filter_cond['tab'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        filter_content = self.filter_cond['content'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        layout += self.init_window_content(filter_tag, filter_content)
        col_layout = [[sg.Text('欢迎使用' + self.name, key='-NOTIFICATION-'), sg.Sizegrip()]]
        layout += [[sg.Column(col_layout, expand_x=True)]]
        if self.location:
            self.window = sg.Window(self.name, layout, location=self.location, enable_close_attempted_event=True,
                                    resizable=True, finalize=True, icon=self.icon, font='Courier 12',
                                    element_justification='top')
        else:
            self.window = sg.Window(self.name, layout, enable_close_attempted_event=True, resizable=True,
                                    finalize=True, icon=self.icon, font='Courier 12', element_justification='top')
        self.window.set_min_size((300, 500))
        self.location = self.window.current_location()
        if filter_tag:
            self.window[filter_tag].select()
        if filter_content:
            self.window['-CONTENT-'].update(filter_content)
        self.bind_hotkeys()
        return self.window

    def init_window_content(self, filter_tag, filter_content):
        tab_group = {}
        max_num_per_row = 1 if self.list_view else max(math.floor(len(self.apps_info) / 2), 4)
        app_right_click_menu = [[], [E_RUN_APP, E_MOD_APP, E_RMV_APP, E_OPEN_APP_PATH, E_COPY_APP_PATH]]
        tag_right_click_menu = [[], [E_ADD_TAG, E_RMV_TAG, E_SRT_APP, E_ADD_APP, E_REFRESH]]
        for tag, apps in self.apps_info.items():
            tab_layout = []
            row_layout = []
            is_filter = True if tag == filter_tag else False
            if self.sort_type == SortType.ASCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper())
            elif self.sort_type == SortType.DESCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper(), reverse=True)
            for app in apps:
                col_layout = []
                app_name = app['name']
                app_icon = app['icon']
                app_path = app['path']
                app_key = app['key']
                if os.path.isfile(app_icon):
                    icon = sg.Image(source=app_icon, tooltip=app_name, key='[icon]' + app_key, enable_events=True,
                                    background_color='white', right_click_menu=app_right_click_menu)
                    col_layout.append(icon)
                button = sg.Button(button_text=app_name[:20], tooltip=app_name, key=app_key, size=(20, 1),
                                   button_color=('black', 'white'), enable_events=False, border_width=0,
                                   right_click_menu=app_right_click_menu, expand_y=True)
                col_layout.append(button)
                if self.list_view:
                    path = sg.Text(text=app_path[:35], tooltip=app_path, size=(35, 1), justification='left',
                                   background_color='white', enable_events=True)
                    col_layout.append(path)
                    time = sg.Text(app['time'], background_color='white')
                    col_layout.append(time)
                    col = sg.Column([col_layout], background_color='white', expand_y=True, expand_x=True,
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

    def init_systray(self):
        menu = [[], [E_SYSTRAY_SHOW, E_SYSTRAY_HIDE, E_SYSTRAY_QUIT]]
        if self.enable_systray and self.window:
            self.systray = pt.SystemTray(menu, icon=self.icon, window=self.window, tooltip=self.name,
                                         single_click_events=False)
        return self.systray

    def init_apps_info(self):
        if not self.app_icons:
            self.app_icons = glob.glob(os.path.join(os.getcwd(), 'images', '*.png'))
        apps_info = {ALL_TAG: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in apps_info:
                    apps_info[parent] = []
                apps = [{
                    'name': directory,
                    'key': parent + '#' + os.path.join(parent, directory),
                    'path': os.path.join(parent, directory),
                    'icon': random.choice(self.app_icons),
                    'time': self.timestamp_to_datetime(os.path.getmtime(os.path.join(parent, directory))),
                } for directory in directorys]
                apps_info[parent] += apps
                apps_info[ALL_TAG] += apps
        return apps_info

    def init_event_functions(self):
        event_func = {
            sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED: self.show_top,
            sg.WIN_CLOSE_ATTEMPTED_EVENT: self.attempt_exit,
            E_SYSTRAY_SHOW: self.show_top,
            E_SYSTRAY_HIDE: self.attempt_exit,
            E_SYSTRAY_QUIT: self.exit,
            E_CLOSE_WINDOW: self.attempt_exit,
            E_QUIT: self.exit,
            E_CHECK_UPDATE: self.check_update,
            E_ABOUT: self.about,
            E_NEW_WINDOW: self.new_window,
            E_IMPORT_APP_INFO: self.import_app,
            E_EXPORT_APP_INFO: self.export_app,
            E_SHOW_ICON: self.show_icon,
            E_SHOW_LIST: self.show_list,
            E_SELECT_ROOT: self.select_root,
            E_SELECT_THEME: self.select_theme,
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
        new_event_func = {}
        for key, value in event_func.items():
            new_event_func[key] = value
            if '&' in key:
                new_event_func[key.replace('&', '')] = value
        return new_event_func

    @staticmethod
    def init_event_hotkeys():
        return {
            E_SYSTRAY_QUIT: ['<Control-q>', '<Control-Q>'],
            E_CLOSE_WINDOW: ['<Control-e>', '<Control-E>'],
            E_CHECK_UPDATE: ['<Control-u>', '<Control-U>'],
            E_ABOUT: ['<Control-a>', '<Control-A>'],
            E_NEW_WINDOW: ['<Control-n>', '<Control-N>'],
            E_IMPORT_APP_INFO: ['<Control-i>', '<Control-I>'],
            E_EXPORT_APP_INFO: ['<Control-x>', '<Control-X>'],
            E_SELECT_ROOT: ['<Control-o>', '<Control-O>'],
            E_SELECT_THEME: ['<Control-t>', '<Control-T>'],
            E_REFRESH: ['<F5>'],
            E_MANAGE_TAG: ['<Control-m>', '<Control-M>'],
            E_RUN_APP: ['<Control-r>', '<Control-R>'],
            E_OPEN_APP_PATH: ['<Control-p>', '<Control-P>'],
            E_COPY_APP_PATH: ['<Control-c>', '<Control-C>'],
        }

    def bind_hotkeys(self):
        for tag, apps in self.apps_info.items():
            for app in apps:
                key = app['key']
                self.window[key].bind('<Button-1>', ' LeftClick')
                self.window[key].bind('<Button-2>', ' MiddleClick')
                self.window[key].bind('<Button-3>', ' RightClick')
        for event, hotkeys in self.event_hotkeys.items():
            for hotkey in hotkeys:
                self.window.bind(hotkey, event)

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
