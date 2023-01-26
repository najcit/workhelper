# -*- coding: utf-8 -*-

import contextlib
import copy
import os
import random
import re
import PySimpleGUI as sg
import psgtray as pt
import math
import shutil
import pyperclip
import winshell
import yaml
from appupdater import AppUpdater

THREAD_KEY = '-THREAD-'
DL_END_KEY = '-DOWNLOAD-END-'


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
        self.event_functions = {
            sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED: self.show_top,
            sg.WIN_CLOSE_ATTEMPTED_EVENT: self.attempt_exit,
            '显示界面': self.show_top,
            '隐藏界面': self.attempt_exit,
            '关闭窗口': self.attempt_exit,
            '退出': self.exit,
            '检查更新': self.check_update,
            '关于': self.about,
            '新建窗口': self.new_window,
            '导出应用信息': self.export_app,
            '打开目录': self.select_app_root,
            '选择主题': self.set_theme,
            '管理标签': self.manage_tab,
            '新建标签': self.add_tab,
            '修改标签': self.mod_tab,
            '删除标签': self.rmv_tab,
            '刷新': self.refresh,
            '新建应用': self.add_app,
            '修改应用': self.mod_app,
            '删除应用': self.rmv_app,
            '打开应用路径': self.open_app,
            '复制应用路径': self.copy_app_path,
            '搜索': self.search_app,
        }

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

    def refresh(self, root=None, theme=None, filter_cond=None):
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
            if event in self.event_functions and callable(self.event_functions[event]):
                # noinspection PyArgumentList
                self.event_functions[event](event=event, values=values)
            else:
                print(event, values, isinstance(event, tuple), type(event), len(event) > 1)
                if isinstance(event, str) and event.endswith('Click'):
                    self.select_app(event=event)
                elif isinstance(event, tuple) and event == (THREAD_KEY, DL_END_KEY):
                    print(event, values)
                    self.exit()
                else:
                    self.run_app()

    def init_window(self):
        sg.theme(self.theme)
        menu = [['文件', ['新建窗口', '关闭窗口', '导入应用信息', '导出应用信息', '退出']],
                ['编辑', ['管理标签', '新建标签']],
                ['设置', ['打开目录', '选择主题']],
                ['帮助', ['检查更新', '关于']]]
        layout = [[sg.MenubarCustom(menu)]]
        layout += [sg.Input(enable_events=True, key='-CONTENT-', expand_x=True),
                   sg.Button('搜索', bind_return_key=True)],
        whole = '全部应用'
        apps_info = {whole: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in apps_info:
                    apps_info[parent] = []
                apps = [{
                    'name': directory,
                    'path': os.path.join(parent, directory),
                    'icon': os.path.join(os.getcwd(), 'images', 'robot_' + str(random.randrange(1, 6)) + '.png'),
                } for directory in directorys]
                apps_info[parent] += apps
                apps_info[whole] += apps
        self.apps_info = apps_info
        filter_tab = self.filter_cond['tab'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        filter_content = self.filter_cond['content'] if self.filter_cond and len(self.filter_cond) > 0 else ''
        tabs = []
        keys = []
        max_num_per_row = max(math.floor(len(apps_info) / 2), 4)
        app_right_click_menu = [['右击菜单'], ['执行应用', '修改应用', '删除应用', '打开应用路径', '复制应用路径']]
        tab_right_click_menu = [['右击菜单'], ['修改标签', '删除标签', '新建应用', '刷新']]
        for tab, apps in apps_info.items():
            tab_name = os.path.basename(tab).ljust(4, ' ')
            tab_key = tab
            tab_layout = []
            row_layout = []
            is_filter = True if tab_key == filter_tab else False
            for app in apps:
                app_name = app['name']
                app_icon = app['icon']
                app_key = tab + '#' + app['path']
                button = sg.Button(button_text=app_name, tooltip=app_name, key=app_key, size=(18, 1), font=12,
                                   button_color=('black', 'white'), enable_events=True, border_width=0,
                                   right_click_menu=app_right_click_menu, expand_x=True, expand_y=True)
                if os.path.isfile(app_icon):
                    icon = sg.Image(source=app_icon, size=(64, 64), background_color='white')
                    col = sg.Column([[icon, button]], background_color='white')
                else:
                    col = sg.Column([[button]], background_color='white')
                if is_filter:
                    if re.search(filter_content, app_name, re.IGNORECASE):
                        row_layout.append(col)
                        keys.append(app_key)
                else:
                    row_layout.append(col)
                    keys.append(app_key)

                if len(row_layout) == max_num_per_row:
                    tab_layout.append(copy.deepcopy(row_layout))
                    row_layout.clear()
            else:
                if len(row_layout) > 0:
                    tab_layout.append(row_layout)
            tabs.append(sg.Tab(tab_name[:4], tab_layout, key=tab_key, right_click_menu=tab_right_click_menu))
        layout += [[sg.TabGroup([tabs], key='-TAB-GROUP-', expand_x=True, expand_y=True,
                                right_click_menu=tab_right_click_menu)]]
        layout += [[sg.Text('欢迎使用' + self.name, key='-NOTIFICATION-'), sg.Sizegrip()]]
        if self.location:
            self.window = sg.Window(self.name, layout, location=self.location, enable_close_attempted_event=True,
                                    resizable=True, finalize=True, icon=self.icon)
        else:
            self.window = sg.Window(self.name, layout, enable_close_attempted_event=True, resizable=True,
                                    finalize=True, icon=self.icon)
        self.window.set_min_size(self.window.size)
        if filter_tab:
            self.window[filter_tab].select()
        if filter_content:
            self.window['-CONTENT-'].update(filter_content)
        for key in keys:
            self.window[key].bind('<Button-1>', ' LeftClick')
            self.window[key].bind('<Button-3>', ' RightClick')
        return self.window

    def init_systray(self):
        menu = ['后台菜单', ['显示界面', '隐藏界面', '退出']]
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

    def set_theme(self, **_kwargs):
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
                    self.window.write_event_value((THREAD_KEY, DL_END_KEY), percentage)

            self.window.start_thread(lambda: self.updater.install(update_progress_func=update_progress), ())

    def about(self, **_kwargs):
        text = '''\n当前已是最新版本: {version}\nCopyright © 2023–2026 by lidajun\n'''.format(version=self.version)
        sg.popup_no_buttons(text, title=self.name, auto_close=True, auto_close_duration=10)

    def new_window(self, **_kwargs):
        location_x, location_y = self.window.current_location()
        location = (location_x + 20, location_y + 20)
        MyApp.show(self.root, self.theme, self.enable_systray, self.enable_env, location)

    def manage_tab(self):
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
                new_tab = self.add_tab()
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
                new_tab = self.mod_tab(tab_key)
                if isinstance(new_tab, str):
                    values = window['-TAB-LIST-'].get_list_values()
                    index = values.index(cur_tab)
                    values[index] = new_tab
                    window['-TAB-LIST-'].update(values, set_to_index=index)
            elif event == '-DELETE-':
                cur_tab = window['-TAB-LIST-'].get()[0]
                tab_key = os.path.join(self.root, cur_tab)
                if self.rmv_tab(tab_key):
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

    def add_tab(self):
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

    def mod_tab(self, tab_key=None, **kwargs):
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

    def rmv_tab(self, tab_key=None, **kwargs):
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

    def search_app(self, **kwargs):
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

    def open_app(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        # print(cur_path)
        if os.path.exists(cur_path):
            os.startfile(cur_path)

    def run_app(self):
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

    def select_app_root(self):
        root = sg.popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if os.path.isdir(root):
            self.refresh(root)

    @staticmethod
    def set_env(key, value):
        if os.getenv(key) != value:
            command = f'start /b setx {key} {value}'.format(key=key, value=value)
            os.popen(command)
