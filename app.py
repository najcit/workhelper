# -*- coding: utf-8 -*-

import contextlib
import copy
import os
import random
import re
import time

import PySimpleGUI as sg
import psgtray as pt
import math
import shutil
import pyperclip
import winshell
import yaml
from appupdater import AppUpdater

ALL_TAG = '全部应用'
EVENT_THREAD_KEY = '-THREAD-'
EVENT_DL_END_KEY = '-DOWNLOAD-END-'
EVENT_SHOW = '显示界面'
EVENT_HIDE = '隐藏界面'
EVENT_QUIT = '退出'
EVENT_QUIT_WITH_HOTKEY = '退出(&Q)'
EVENT_FILE = '文件(&F)'
EVENT_NEW_WINDOW = '新建窗口(&N)'
EVENT_CLOSE_WINDOW = '关闭窗口(&E)'
EVENT_IMPORT_APP_INFO = '导入应用信息(&I)'
EVENT_EXPORT_APP_INFO = '导出应用信息(&X)'
EVENT_EDIT = '编辑(&E)'
EVENT_MANAGE_TAG = '管理标签(&M)'
EVENT_ADD_TAG = '新建标签(&N)'
EVENT_MOD_TAG = '编辑标签'
EVENT_VIEW = '视图(&V)'
EVENT_SHOW_ICON = '图标(&G)'
EVENT_SHOW_LIST = '列表(&L)'
EVENT_OPTION = '选项(&S)'
EVENT_SELECT_ROOT = '选择目录(&O)'
EVENT_SELECT_THEME = '选择主题(&T)'
EVENT_SELECT_FONT = '选择字体(&F)'
EVENT_SET_WINDOW = '设置窗口(&W)'
EVENT_HELP = '帮助(&H)'
EVENT_CHECK_UPDATE = '检查更新(&U)'
EVENT_ABOUT = '关于(&A)'
EVENT_RMV_TAG = '删除标签'
EVENT_REFRESH = '刷新'
EVENT_ADD_APP = '添加应用'
EVENT_MOD_APP = '修改应用'
EVENT_RMV_APP = '删除应用'
EVENT_RUN_APP = '运行应用'
EVENT_OPEN_APP_PATH = '打开应用路径'
EVENT_COPY_APP_PATH = '复制应用路径'
EVENT_SEARCH = '搜索'


class MyApp:
    name = None
    version = None
    root = None
    theme = None
    enable_systray = False
    enable_env = False
    window = None
    systray = None
    location = None
    filter_cond = None
    apps_info = None
    event_functions = None
    event_hotkeys = None

    def __init__(self, root, theme, enable_systray, enable_env, location, name=None):
        self.name = name if name is not None else '工作助手'
        self.version = '1.0.0'
        self.icon = 'workhelper.ico'
        self.root = root
        self.theme = theme
        self.enable_systray = enable_systray
        self.enable_env = enable_env
        self.location = location
        self.updater = AppUpdater()
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

    def refresh(self, root=None, theme=None, filter_cond=None, **_kwargs):
        if root and root != self.root:
            self.root = root
            self.set_env('MY_ROOT', self.root)
        if theme and theme != self.theme:
            self.theme = theme
            self.set_env('MY_THEME', self.theme)
        self.filter_cond = filter_cond if filter_cond else {}
        self.close()
        self.init()
        return self

    def run(self):
        while True:
            event, values = self.read()
            print(event, values, isinstance(event, tuple), type(event), len(event) > 1)
            if event in self.event_functions and callable(self.event_functions[event]):
                # noinspection PyArgumentList
                self.event_functions[event](event=event, values=values)
            else:
                print(event, values, isinstance(event, tuple), type(event), len(event) > 1)
                if isinstance(event, str) and event.endswith('Click'):
                    self.select_app(event=event)
                elif isinstance(event, tuple) and event == (EVENT_THREAD_KEY, EVENT_DL_END_KEY):
                    print(event, values)
                    self.exit()
                else:
                    self.run_app()

    def init_window(self):
        sg.theme(self.theme)
        menu_def = [[EVENT_FILE, [EVENT_NEW_WINDOW, EVENT_CLOSE_WINDOW, EVENT_IMPORT_APP_INFO,
                                  EVENT_EXPORT_APP_INFO, EVENT_QUIT_WITH_HOTKEY]],
                    [EVENT_EDIT, [EVENT_MANAGE_TAG, EVENT_ADD_TAG]],
                    [EVENT_VIEW, [EVENT_SHOW_ICON, EVENT_SHOW_LIST]],
                    [EVENT_OPTION, [EVENT_SELECT_ROOT, EVENT_SELECT_THEME, EVENT_SELECT_FONT, EVENT_SET_WINDOW]],
                    [EVENT_HELP, [EVENT_CHECK_UPDATE, EVENT_ABOUT]]]
        layout = [[sg.Menu(menu_def, key='-MENUBAR-')]]
        layout += [sg.Input(enable_events=False, key='-CONTENT-', expand_x=True),
                   sg.Button(EVENT_SEARCH, bind_return_key=True)],
        filter_tab = self.filter_cond['tab'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        filter_content = self.filter_cond['content'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        tabs = []
        keys = []
        max_num_per_row = max(math.floor(len(self.apps_info) / 2), 4)
        app_right_click_menu = [[], [EVENT_RUN_APP, EVENT_MOD_APP, EVENT_RMV_APP, EVENT_OPEN_APP_PATH,
                                     EVENT_COPY_APP_PATH]]
        tab_right_click_menu = [[], [EVENT_ADD_TAG, EVENT_RMV_TAG, EVENT_ADD_APP, EVENT_REFRESH]]
        for tab, apps in self.apps_info.items():
            tab_name = os.path.basename(tab).ljust(4, ' ')
            tab_key = tab
            tab_layout = []
            row_layout = []
            is_filter = True if tab_key == filter_tab else False
            for app in apps:
                app_name = app['name']
                app_icon = app['icon']
                app_key = tab + '#' + app['path']
                button = sg.Button(button_text=app_name, tooltip=app_name, key=app_key, size=(20, 1),
                                   button_color=('black', 'white'), enable_events=True, border_width=0,
                                   right_click_menu=app_right_click_menu, expand_x=True, expand_y=True)
                if os.path.isfile(app_icon):
                    icon = sg.Image(source=app_icon, background_color='white')
                    col = sg.Column([[icon, button]], background_color='white')
                else:
                    col = sg.Column([[button]], background_color='white')
                if (is_filter and re.search(filter_content, app_name, re.IGNORECASE)) or not is_filter:
                    row_layout.append(col)
                    keys.append(app_key)
                if len(row_layout) == max_num_per_row:
                    tab_layout.append(copy.deepcopy(row_layout))
                    row_layout.clear()
            else:
                if len(row_layout) > 0:
                    tab_layout.append(row_layout)
            tabs.append(sg.Tab(tab_name, tab_layout, key=tab_key, expand_x=True, expand_y=True,
                               right_click_menu=tab_right_click_menu))
        layout += [[sg.TabGroup([tabs], key='-TAB-GROUP-', expand_x=True, expand_y=True,
                                right_click_menu=tab_right_click_menu)]]
        layout += [[sg.Text('欢迎使用' + self.name, key='-NOTIFICATION-'), sg.Sizegrip()]]
        if self.location:
            self.window = sg.Window(self.name, layout, location=self.location, enable_close_attempted_event=True,
                                    resizable=True, finalize=True, icon=self.icon, font='Courier 12')
        else:
            self.window = sg.Window(self.name, layout, enable_close_attempted_event=True, resizable=True,
                                    finalize=True, icon=self.icon, font='Courier 12')
        self.window.set_min_size(self.window.size)
        if filter_tab:
            self.window[filter_tab].select()
        if filter_content:
            self.window['-CONTENT-'].update(filter_content)
        for key in keys:
            self.window[key].bind('<Button-1>', ' LeftClick')
            self.window[key].bind('<Button-3>', ' RightClick')
        self.bind_hotkeys()
        return self.window

    def bind_hotkeys(self):
        for event, hotkeys in self.event_hotkeys.items():
            for hotkey in hotkeys:
                self.window.bind(hotkey, event)

    def init_systray(self):
        menu = [[], [EVENT_SHOW, EVENT_HIDE, EVENT_QUIT]]
        if self.enable_systray and self.window:
            self.systray = pt.SystemTray(menu, icon=self.icon, window=self.window, tooltip=self.name,
                                         single_click_events=False)
        return self.systray

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
            self.close()
            exit(0)

    def exit(self, **_kwargs):
        self.close()
        exit(0)

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
                    self.window.write_event_value((EVENT_THREAD_KEY, EVENT_DL_END_KEY), percentage)

            self.window.start_thread(lambda: self.updater.install(update_progress_func=update_progress), ())

    def about(self, **_kwargs):
        text = '''\n当前已是最新版本: {version}\nCopyright © 2023–2026 by lidajun\n'''.format(version=self.version)
        sg.popup_no_buttons(text, title=self.name, auto_close=True, auto_close_duration=10)

    def new_window(self, **_kwargs):
        location_x, location_y = self.window.current_location()
        location = (location_x + 20, location_y + 20)
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

    def copy_app_path(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        pyperclip.copy(cur_path)

    def import_app(self):
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

    def export_app(self):
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

    def select_root(self):
        root = sg.popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if os.path.isdir(root):
            self.refresh(root)

    def init_apps_info(self):
        apps_info = {ALL_TAG: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in apps_info:
                    apps_info[parent] = []

                def timestamp_to_datetime(time_stamp, format_string="%Y-%m-%d %H:%M:%S"):
                    time_array = time.localtime(time_stamp)
                    str_date = time.strftime(format_string, time_array)
                    return str_date

                apps = [{
                    'name': directory,
                    'path': os.path.join(parent, directory),
                    'icon': os.path.join(os.getcwd(), 'images', 'robot_' + str(random.randrange(1, 6)) + '.png'),
                    'time': timestamp_to_datetime(os.path.getmtime(os.path.join(parent, directory))),
                } for directory in directorys]
                apps_info[parent] += apps
                apps_info[ALL_TAG] += apps
        return apps_info

    def init_event_functions(self):
        event_func = {
            sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED: self.show_top,
            sg.WIN_CLOSE_ATTEMPTED_EVENT: self.attempt_exit,
            EVENT_SHOW: self.show_top,
            EVENT_HIDE: self.attempt_exit,
            EVENT_QUIT: self.exit,
            EVENT_CLOSE_WINDOW: self.attempt_exit,
            EVENT_QUIT_WITH_HOTKEY: self.exit,
            EVENT_CHECK_UPDATE: self.check_update,
            EVENT_ABOUT: self.about,
            EVENT_NEW_WINDOW: self.new_window,
            EVENT_IMPORT_APP_INFO: self.import_app,
            EVENT_EXPORT_APP_INFO: self.export_app,
            EVENT_SELECT_ROOT: self.select_root,
            EVENT_SELECT_THEME: self.select_theme,
            EVENT_SEARCH: self.search,
            EVENT_REFRESH: self.refresh,
            EVENT_MANAGE_TAG: self.manage_tag,
            EVENT_ADD_TAG: self.add_tag,
            EVENT_MOD_TAG: self.mod_tag,
            EVENT_RMV_TAG: self.rmv_tag,
            EVENT_ADD_APP: self.add_app,
            EVENT_MOD_APP: self.mod_app,
            EVENT_RMV_APP: self.rmv_app,
            EVENT_RUN_APP: self.run_app,
            EVENT_OPEN_APP_PATH: self.open_app_path,
            EVENT_COPY_APP_PATH: self.copy_app_path,
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
            EVENT_QUIT: ['<Control-q>', '<Control-Q>'],
            EVENT_CLOSE_WINDOW: ['<Control-e>', '<Control-E>'],
            EVENT_CHECK_UPDATE: ['<Control-u>', '<Control-U>'],
            EVENT_ABOUT: ['<Control-a>', '<Control-A>'],
            EVENT_NEW_WINDOW: ['<Control-n>', '<Control-N>'],
            EVENT_IMPORT_APP_INFO: ['<Control-i>', '<Control-I>'],
            EVENT_EXPORT_APP_INFO: ['<Control-x>', '<Control-X>'],
            EVENT_SELECT_ROOT: ['<Control-o>', '<Control-O>'],
            EVENT_SELECT_THEME: ['<Control-t>', '<Control-T>'],
            EVENT_REFRESH: ['<F5>'],
            EVENT_MANAGE_TAG: ['<Control-m>', '<Control-M>'],
            EVENT_RUN_APP: ['<Control-r>', '<Control-R>'],
            EVENT_OPEN_APP_PATH: ['<Control-p>', '<Control-P>'],
            EVENT_COPY_APP_PATH: ['<Control-c>', '<Control-C>'],
        }

    @staticmethod
    def set_env(key, value):
        if os.getenv(key) != value:
            command = f'start /b setx {key} {value}'.format(key=key, value=value)
            os.popen(command)
