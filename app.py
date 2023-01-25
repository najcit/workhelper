import contextlib
import copy
import os
import re
import PySimpleGUI as sg
import psgtray as pt
import math
import shutil
import pyperclip
import winshell
import yaml


class MyApp:
    root = None
    theme = None
    enable_tray = False
    enable_env = False
    window = None
    systray = None
    location = None
    filter_cond = {}
    app_info = None

    def __init__(self, root, theme, enable_tray, enable_env, location):
        self.root = root
        self.theme = theme
        self.enable_tray = enable_tray
        self.enable_env = enable_env
        self.location = location

    @staticmethod
    def show(root, theme, enable_tray, enable_env, location=None):
        app = MyApp(root, theme, enable_tray, enable_env, location)
        app.init()
        app.run()
        app.close()

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

    def update(self, root=None, theme=None, filter_cond=None):
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
            if event in (sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED, '显示界面'):
                self.window.un_hide()
                self.window.bring_to_front()
            elif event in (sg.WIN_CLOSE_ATTEMPTED_EVENT, '隐藏界面', '关闭窗口'):
                if self.systray:
                    self.window.hide()
                    self.systray.show_icon()
                else:
                    break
            elif event in (sg.WIN_CLOSED, 'Exit', '退出'):
                break
            elif event == '关于':
                sg.popup('工作助手', 'Copyright © 2010–2023 by lidajun')
            elif event == '新建窗口':
                location_x, location_y = self.window.current_location()
                location = (location_x + 20, location_y + 20)
                MyApp.show(self.root, self.theme, self.enable_tray, self.enable_env, location)
            elif event == '导入应用信息':
                self.import_app()
            elif event == '导出应用信息':
                self.export_app()
            elif event == '管理标签':
                self.manage_tab()
            elif event == '打开目录':
                self.select_app_root()
            elif event == '选择主题':
                self.set_theme()
            elif event == '添加标签':
                self.add_tab()
            elif event == '修改标签':
                cur_tab = values[0]
                self.mod_tab(cur_tab)
            elif event == '删除标签':
                cur_tab = values[0]
                self.rmv_tab(cur_tab)
            elif event == '刷新':
                self.update()
            elif event == '添加应用':
                cur_tab = values[0]
                self.add_app(cur_tab)
            elif event == '修改应用':
                item = self.window.find_element_with_focus()
                if self.mod_app(item.key):
                    self.update()
            elif event == '修改应用':
                item = self.window.find_element_with_focus()
                if str(item).find('Button') > -1:
                    self.run_app(item.key)
            elif event == '删除应用':
                item = self.window.find_element_with_focus()
                if self.rmv_app(item.key):
                    self.update()
            elif event == '打开应用路径':
                item = self.window.find_element_with_focus()
                self.open_app(item.key)
            elif event == '复制应用路径':
                item = self.window.find_element_with_focus()
                self.copy_app_path(item.key)
            elif type(event) == str and event.endswith('Click'):
                key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
                self.window[key].set_focus()
            elif event == '搜索':
                tab = values[0]
                content = values['content']
                filter_condition = {'tab': tab, 'content': content}
                if filter_condition != self.filter_cond:
                    self.update(filter_condition)
            else:
                # print(event, values)
                item = self.window.find_element_with_focus()
                if str(item).find('Button') > -1:
                    self.run_app(item.key)

    def init_window(self):
        sg.theme(self.theme)
        menu = [['文件', ['新建窗口', '关闭窗口', '导入应用信息', '导出应用信息', '退出']],
                ['编辑', ['管理标签', '添加标签']],
                ['设置', ['打开目录', '选择主题', '自动更新']],
                ['帮助', ['关于']]]
        layout = [[sg.MenubarCustom(menu)]]
        layout += [sg.Input(enable_events=True, key='content', expand_x=True), sg.Button('搜索', bind_return_key=True)],
        whole = '全部应用'
        app_info = {whole: []}
        for parent, directorys, _files in os.walk(self.root):
            sub_root = parent.replace(self.root + os.path.sep, '')
            if parent != self.root and sub_root.find(os.path.sep) == -1:
                if parent not in app_info:
                    app_info[parent] = []
                app_info[parent] += [os.path.join(parent, directory) for directory in directorys]
                app_info[whole] += [os.path.join(parent, directory) for directory in directorys]
        self.app_info = app_info
        filter_tab = self.filter_cond['tab'] if len(self.filter_cond) > 0 else ''
        filter_content = self.filter_cond['content'] if len(self.filter_cond) > 0 else ''
        tabs = []
        keys = []
        max_num_per_row = max(math.ceil(len(app_info) / 2), 4)
        app_right_click_menu = [['右击菜单'], ['执行应用', '修改应用', '删除应用', '打开应用路径', '复制应用路径']]
        tab_right_click_menu = [['右击菜单'], ['修改标签', '删除标签', '添加应用', '刷新']]
        for tab, apps in app_info.items():
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
        if self.location:
            self.window = sg.Window('工作助手', layout, location=self.location, enable_close_attempted_event=True,
                                    resizable=True, finalize=True)
        else:
            self.window = sg.Window('工作助手', layout, enable_close_attempted_event=True, resizable=True, finalize=True)
        self.window.set_min_size((600, 400))
        if filter_tab:
            self.window[filter_tab].select()
        if filter_content:
            self.window[filter_content].update(filter_content)
        for key in keys:
            self.window[key].bind('<Button-1>', ' LeftClick')
            self.window[key].bind('<Button-3>', ' RightClick')
        return self.window

    def init_systray(self):
        menu = ['后台菜单', ['显示界面', '隐藏界面', '退出']]
        if self.enable_tray and self.window:
            self.systray = pt.SystemTray(menu, single_click_events=False, window=self.window, tooltip='工作助手')
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

    def set_theme(self):
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
            self.update()
        return self

    def add_tab(self):
        result = False
        layout = [[sg.Text('新标签'), sg.Input(key='NewTab')],
                  [sg.Button('确定'), sg.Button('取消')]]
        tab_window = sg.Window('添加标签', layout, element_justification='right')
        while True:
            event, values = tab_window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                new_tab = tab_window['NewTab'].get()
                print(new_tab)
                os.mkdir(os.path.join(self.root, new_tab))
                sg.popup('You success to new a tab!', auto_close=True, auto_close_duration=1, keep_on_top=True)
                result = True
                break
        tab_window.close()
        return result

    def mod_tab(self, cur_tab):
        result = False
        layout = [[sg.Text('当前标签'), sg.Input(key='CurTab', default_text=cur_tab)],
                  [sg.Button('确定'), sg.Button('取消')]]
        tab_window = sg.Window('添加标签', layout)
        while True:
            event, values = tab_window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                new_tab = tab_window['CurTab'].get()
                os.renames(os.path.join(self.root, cur_tab), os.path.join(self.root, new_tab))
                sg.popup('You success to new a tab!', auto_close=True, auto_close_duration=1, keep_on_top=True)
                result = True
                break
        tab_window.close()
        return result

    @staticmethod
    def set_env(key, value):
        if os.getenv(key) != value:
            command = f'start /b setx {key} {value}'.format(key=key, value=value)
            os.popen(command)

    @staticmethod
    def rmv_tab(tab_key):
        tab_name = os.path.basename(tab_key)
        tab_dir = tab_key
        result = sg.popup('是否删除' + tab_name + '?', no_titlebar=True, keep_on_top=True,
                          button_type=sg.POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(tab_dir)
        return result

    @staticmethod
    def add_app(tab_key):
        result = False
        tab_dir = tab_key
        layout = [[sg.Text('添加目标'), sg.Input(key='TargetPath'), sg.Button('浏览')],
                  [sg.Text('添加应用'), sg.Input(key='AppName'), sg.Button('图标')],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('添加应用', layout, element_justification='right')
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
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
            elif event == '浏览':
                file = sg.popup_get_file('选择目标文件', keep_on_top=True)
                if file:
                    window['TargetPath'].update(value=file)
                    app_name = os.path.basename(str(file)).capitalize().removesuffix('.exe')
                    window['AppName'].update(value=app_name)
        window.close()
        return result

    @staticmethod
    def mod_app(app_key):
        result = False
        app_dir = app_key.split('#')[1]
        app_name = os.path.basename(app_dir)
        layout = [[sg.Text('应用名称'), sg.Input(key='AppName', default_text=app_name)],
                  [sg.Button('确定'), sg.Button('取消')]]
        window = sg.Window('添加应用', layout)
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                new_app_name = window['AppName'].get()
                os.renames(app_dir, os.path.join(os.path.dirname(app_dir), new_app_name))
                sg.popup(app_name + '成功修改成' + new_app_name + '!', no_titlebar=True, keep_on_top=True, auto_close=True,
                         auto_close_duration=3)
                result = True
                break
        window.close()
        return result

    @staticmethod
    def rmv_app(app_key):
        app_dir = app_key.split('#')[1]
        app_name = os.path.basename(app_dir)
        result = sg.popup('是否删除' + app_name + '?', no_titlebar=True, keep_on_top=True,
                          button_type=sg.POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(app_dir)
        return result

    @staticmethod
    def open_app(app_key):
        cur_path = app_key.split('#')[1]
        # print(cur_path)
        if os.path.exists(cur_path):
            os.startfile(cur_path)

    @staticmethod
    def run_app(app_key):
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

    @staticmethod
    def copy_app_path(app_key):
        cur_path = app_key.split('#')[1]
        pyperclip.copy(cur_path)

    def import_app(self):
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
            self.update()

    def export_app(self):
        directory = sg.popup_get_folder(message='导出路径', title='选择导出路径', keep_on_top=True)
        if directory:
            file_path = os.path.join(directory, 'workhelper.yml')
            app_cfg = {}
            for tab, apps in self.app_info.items():
                if tab == '全部应用':
                    continue
                tab = os.path.basename(tab)
                apps = [os.path.basename(app) for app in apps]
                app_cfg[tab] = apps
            with open(file_path, 'w', encoding='utf-8') as file:
                yaml.dump(app_cfg, file, allow_unicode=True)

    @staticmethod
    def manage_app(self):
        pass

    def select_app_root(self):
        root = sg.popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if os.path.isdir(root):
            self.update(root)
