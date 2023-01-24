import copy
import os
import re
from contextlib import suppress
import PySimpleGUI as sg
import psgtray as pt
import math


class MyApp:
    root = None
    theme = None
    enable_tray = False
    enable_env = False
    window = None
    systray = None
    filter_condition = {}

    def __init__(self, root, theme, enable_tray, enable_env):
        self.root = root
        self.theme = theme
        self.enable_tray = enable_tray
        self.enable_env = enable_env

    def init(self):
        if self.enable_env:
            with suppress(Exception):
                self.root = os.environ['MY_ROOT']
                self.theme = os.environ['MY_THEME']
        if self.root == 'root':
            self.root = os.path.join(os.getcwd(), self.root)
        self.init_window()
        self.init_systray()
        return self

    def update(self, filter_condition=None):
        self.filter_condition = filter_condition if filter_condition else {}
        self.close()
        self.init()
        return self

    def init_window(self):
        sg.theme(self.theme)
        menu = [['编辑', ['管理标签', '添加标签']],
                ['设置', ['打开目录', '选择主题', '自动更新']],
                ['帮助', ['关于']]]
        layout = [[sg.MenubarCustom(menu)]]
        layout += [sg.Input(enable_events=True, key='content', expand_x=True), sg.Button('搜索')],
        whole = '全部应用'
        items_info = {whole: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in items_info:
                    items_info[parent] = []
                items_info[parent] += [os.path.join(parent, directory) for directory in directorys]
                items_info[whole] += [os.path.join(parent, directory) for directory in directorys]
        filter_tab = self.filter_condition['tab'] if len(self.filter_condition) > 0 else ''
        filter_content = self.filter_condition['content'] if len(self.filter_condition) > 0 else ''
        tabs = []
        keys = []
        max_num_per_row = max(math.ceil(len(items_info) / 2), 4)
        app_right_click_menu = [['右击菜单'], ['修改应用', '删除应用', '打开所在位置']]
        tab_right_click_menu = [['右击菜单'], ['修改标签', '删除标签', '添加应用', '刷新']]
        for tab, apps in items_info.items():
            tab_name = os.path.basename(tab).ljust(4, ' ')
            tab_key = tab
            tab_layout = []
            row_layout = []
            is_filter = True if tab_key == filter_tab else False
            for app in apps:
                app_name = os.path.basename(app)
                app_key = tab + '#' + app
                button = sg.Button(app_name, tooltip=app_name, key=app_key, size=(12, 2), mouseover_colors='blue',
                                   enable_events=True, right_click_menu=app_right_click_menu)
                if is_filter:
                    if re.search(filter_content, app_name, re.IGNORECASE):
                    # if app_name.find(filter_content) > -1:
                    #     re.match()
                        row_layout.append(button)
                        keys.append(app_key)
                else:
                    row_layout.append(button)
                    keys.append(app_key)

                if len(row_layout) == max_num_per_row:
                    tab_layout.append(copy.deepcopy(row_layout))
                    row_layout.clear()
            else:
                if len(row_layout) > 0:
                    tab_layout.append(row_layout)
            tabs.append(sg.Tab(tab_name[:4], tab_layout, key=tab_key, right_click_menu=tab_right_click_menu))
        layout += [[sg.TabGroup([tabs], expand_x=True, expand_y=True, right_click_menu=tab_right_click_menu)]]
        layout += [[sg.Sizegrip()]]
        self.window = sg.Window('工作助手', layout, finalize=True, enable_close_attempted_event=True, resizable=True)

        print('filter_tab', filter_tab)
        if filter_tab:
            self.window[filter_tab].select()
        if filter_content:
            self.window['content'].update(filter_content)

        for key in keys:
            self.window[key].bind('<Button-1>', ' LeftClick')
            self.window[key].bind('<Button-3>', ' RightClick')
        return self.window

    def init_systray(self):
        menu = ['后台菜单', ['显示界面', '隐藏界面', '退出']]
        if self.enable_tray and self.window:
            self.systray = pt.SystemTray(menu, single_click_events=False, window=self.window, tooltip='工作助手')
        return self.systray

    @staticmethod
    def set_env(key, value):
        if os.getenv(key) != value:
            command = f'start /b setx {key} {value}'.format(key=key, value=value)
            os.popen(command)

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
