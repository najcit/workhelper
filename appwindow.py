# -*- coding: utf-8 -*-

import copy
import os
import random
import re
import shutil
import time
import webbrowser
import tkinter
import pyperclip
import validators
import winshell
import yaml
from PySimpleGUI import theme as set_theme
from PySimpleGUI import TIMEOUT_KEY, WIN_CLOSED, LISTBOX_SELECT_MODE_SINGLE, POPUP_BUTTONS_YES_NO
from PySimpleGUI import popup_get_folder, set_options, popup_no_buttons, popup_get_file, theme_list, popup
from PySimpleGUI import Menu, Sizegrip, Tab, TabGroup, Image, Listbox, Combo, Text, Button, Window, Column, Input
from psgtray import SystemTray
from appresource import *
from apputil import find_or_make_icon, set_env


class AppWindow(object):

    def __init__(self, model):
        self.model = model
        self.window = None
        self.systray = None
        self.initialize()

    def __del__(self):
        self.close()

    def initialize(self, **kwargs):
        self.model.initialize(**kwargs)
        set_theme(self.model.theme)
        set_options(font=self.model.font)
        show_icon = E_SHOW_ICON_VIEW if self.model.show_view == SHOW_LIST else '!' + E_SHOW_ICON_VIEW
        show_list = E_SHOW_LIST_VIEW if self.model.show_view == SHOW_ICON else '!' + E_SHOW_LIST_VIEW
        show_local = E_SHOW_LOCAL_APPS if not self.model.show_local else E_HIDE_LOCAL_APPS
        show_browser = E_SHOW_BROWSER_BOOKMARKS if not self.model.show_browser else E_HIDE_BROWSER_BOOKMARKS
        menu_def = [[E_FILE, [E_IMPORT_APPS_INFO, E_EXPORT_APPS_INFO, E_QUIT]],
                    [E_EDIT, [E_MANAGE_TAG, E_ADD_TAG]],
                    [E_VIEW, [show_icon, show_list, E_SRT_APP, E_REFRESH]],
                    [E_OPTION, [E_SELECT_ROOT, E_SELECT_THEME, E_SELECT_FONT, E_SET_WINDOW, show_local, show_browser]],
                    [E_HELP, [E_CHECK_UPDATE, E_ABOUT]]]
        layout = [[Menu(menu_def, key='-MENU-BAR-')]]
        disable_button = True if self.model.current_tag.find("#") == -1 else False
        col_layout = [[Button(image_filename=self.model.icons['return'], key=E_RETURN, disabled=disable_button,
                              tooltip=E_RETURN),
                       Input(default_text=self.model.current_tag, key='-CURRENT-TAG-', size=(15, 1), disabled=True,
                             expand_x=True, expand_y=True),
                       Input(enable_events=False, key='-CONTENT-', expand_x=True, expand_y=True),
                       Button(image_filename=self.model.icons['search'], key=E_SEARCH, bind_return_key=True,
                              tooltip=E_SEARCH)
                       ]]
        layout += [[Column(col_layout, expand_x=True)]]
        layout += self._init_content()
        col_layout = [[Text('欢迎使用' + self.model.name, key='-NOTIFICATION-'), Sizegrip()]]
        layout += [[Column(col_layout, expand_x=True)]]
        if self.model.location:
            self.window = Window(self.model.name, layout, location=self.model.location,
                                 enable_close_attempted_event=True,
                                 resizable=True, finalize=True, icon=self.model.icons['window'])
        else:
            self.window = Window(self.model.name, layout, enable_close_attempted_event=True,
                                 resizable=True, finalize=True, icon=self.model.icons['window'])
        self.window.set_min_size((800, 400))
        self.model.location = self.window.current_location()
        print('current_tag', self.model.current_tag, 'search_content', self.model.search_content,
              'current_tag', self.model.current_tag)
        self.window[self.model.current_tag].select()
        self.window['-CURRENT-TAG-'].update(value=self.model.current_tag)
        self.window['-CONTENT-'].update(self.model.search_content)
        self._bind_hotkeys()

        menu = [[], [E_SYSTRAY_SHOW, E_SYSTRAY_HIDE, E_SYSTRAY_QUIT]]
        if self.model.enable_systray and self.window:
            self.systray = SystemTray(menu, icon=self.model.icon, window=self.window, tooltip=self.model.name,
                                      single_click_events=False)

    def close(self):
        if self.window:
            self.window.close()
            self.window = None
        if self.systray:
            self.systray.close()
            self.systray = None

    def read(self):
        event, values = WIN_CLOSED, None
        if self.window:
            event, values = self.window.read(timeout_key=TIMEOUT_KEY)
        if self.systray and event == self.systray.key:
            event = values[event]
        return event, values

    def refresh(self, **kwargs):
        self.close()
        self.initialize(**kwargs)

    def exit(self, **_kwargs):
        self.close()

    def attempt_exit(self, **_kwargs):
        if self.systray:
            self.hide()
        else:
            self.exit()

    def hide(self, **_kwargs):
        self.window.hide()

    def show(self):
        self.window.un_hide()
        self.window.bring_to_front()

    def search(self, current_tag, search_content):
        if current_tag != self.model.current_tag or search_content != self.model.search_content:
            self.refresh(current_tag=current_tag, search_content=search_content)

    def back(self, parent_tag):
        self.refresh(current_tag=parent_tag)

    def activate_tag(self, current_tag):
        print(self.model.current_tag)
        if self.model.current_tag.find(current_tag) == -1:
            self.model.current_tag = current_tag
            self.window['-CURRENT-TAG-'].update(value=current_tag)
        disabled = True if current_tag.find('#') == -1 else False
        self.window[E_RETURN].update(disabled=disabled)

    def select_app(self, app_key):
        self.window[app_key].set_focus()

    def import_app_info(self):
        result = False
        file = popup_get_file(message='导入文件', title='选择导入文件', keep_on_top=True)
        if file:
            with open(str(file), "r", encoding='utf8') as file:
                app_cfg = yaml.load(file, yaml.FullLoader)
                for tab, apps in app_cfg.items():
                    tab_path = os.path.join(self.model.root, tab)
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

    def export_app_info(self):
        result = False
        directory = popup_get_folder(message='导出路径', title='选择导出路径', keep_on_top=True)
        if directory:
            file_path = os.path.join(directory, 'workhelper.yml')
            app_cfg = {}
            for tab, apps in self.model.apps.items():
                if tab == '全部应用':
                    continue
                tab = os.path.basename(tab)
                apps = [app['name'] for app in apps]
                app_cfg[tab] = apps
            with open(file_path, 'w', encoding='utf-8') as file:
                yaml.dump(app_cfg, file, allow_unicode=True)
            result = True
        return result

    def show_view_icon(self, location):
        self.refresh(location=location, show_view=SHOW_ICON)

    def show_view_list(self, location):
        self.refresh(location=location, show_view=SHOW_LIST)

    def sort_app(self, sort_order):
        if sort_order == SORT_DEFAULT:
            sort_order = SORT_ASCEND
        elif sort_order == SORT_ASCEND:
            sort_order = SORT_DESCEND
        elif sort_order == SORT_DESCEND:
            sort_order = SORT_DEFAULT
        else:
            sort_order = SORT_DEFAULT
        self.refresh(sort_order=sort_order)

    def select_root(self):
        root = popup_get_folder(message='应用根目录', title='选择应用根目录', keep_on_top=True)
        if root and os.path.isdir(root):
            self.refresh(root=root)

    def select_theme(self):
        theme = self.model.theme
        layout = [[Text('主题'), Combo(theme_list(), default_value=theme, key='-THEME-COMBO-')],
                  [Button('确定'), Button('取消')]]
        window = Window('设置主题', layout, element_justification='right', keep_on_top=True)
        while True:
            event, values = window.read()
            if event in [WIN_CLOSED, '取消']:
                break
            elif event == '确定':
                theme = values['-THEME-COMBO-']
                break
        window.close()
        if theme != self.model.theme:
            set_env('MY_THEME', theme)
            self.refresh(theme=theme)

    def select_font(self):
        font_families = tkinter.font.families()
        size_families = [i for i in range(1, 26)]
        font = self.model.font
        layout = [[Text('字体'), Combo(font_families, default_value=font[0], size=(20, 1), key='-STYLE-COMBO-')],
                  [Text('大小'), Combo(size_families, default_value=font[1], size=(20, 1), key='-SIZE-COMBO-')],
                  [Button('确定'), Button('取消')]]
        window = Window('设置主题', layout, element_justification='right', keep_on_top=True)
        while True:
            event, values = window.read()
            if event in [WIN_CLOSED, '取消']:
                break
            elif event == '确定':
                font = (values['-STYLE-COMBO-'], values['-SIZE-COMBO-'])
                break
        window.close()
        if font != self.model.font:
            set_env('MY_FONT', font[0] + ' ' + str(font[1]))
            self.refresh(font=font)

    def set_window(self):
        cur_size = self.window.size
        layout = [[Text('窗口宽度'), Input(cur_size[0], key='-WINDOWS-WIDTH-')],
                  [Text('窗口高度'), Input(cur_size[1], key='-WINDOWS-HEIGHT-')],
                  [Button('确定'), Button('取消')]]
        window = Window('设置窗口', layout, element_justification='right', keep_on_top=True)
        while window:
            event, values = window.read()
            if event in [WIN_CLOSED, '取消']:
                break
            elif event == '确定':
                cur_size = (values['-WINDOWS-WIDTH-'], values['-WINDOWS-HEIGHT-'])
                break
        window.close()
        if cur_size != self.window.size:
            self.window.size = cur_size

    def show_local_app(self):
        self.refresh(show_local=True)

    def hide_local_app(self):
        self.refresh(show_local=False)

    def show_browser_bookmark(self):
        self.refresh(show_browser=True)

    def hide_browser_bookmark(self):
        self.refresh(show_browser=False)

    def check_update(self, is_new_version, installer):
        if is_new_version:
            text = '当前软件已是最新版本，版本号: {version}'.format(version=self.model.version)
            popup_no_buttons(text, title=self.model.name, auto_close=True, auto_close_duration=10)
        else:
            async def update_progress(percentage):
                progressbar = self.window['-NOTIFICATION-']
                value = '最新版软件正在下载，进度为 {percentage}%。'.format(percentage=percentage)
                progressbar.update(value=value)
                if percentage >= 100:
                    value = '最新版软件下载已完成。'
                    progressbar.update(value=value)
                    self.window.write_event_value((E_THREAD, E_DOWNLOAD_END), percentage)

            self.window.start_thread(lambda: installer(update_progress_func=update_progress), ())

    def about(self, version):
        text = f'\n当前已是最新版本: {version}\nCopyright © 2023–2026 by lidajun\n'
        popup_no_buttons(text, title=self.model.name, auto_close=True, auto_close_duration=10)

    def manage_tag(self):
        values = [os.path.basename(tab) for tab, _ in self.model.apps.items()]
        cur_tab = os.path.basename(self.window[E_ACTIVE_TAG].get())
        col1 = Column([[Listbox(values, key='-TAB-LIST-', default_values=[cur_tab],
                                select_mode=LISTBOX_SELECT_MODE_SINGLE, enable_events=True, size=(30, 40))]],
                      size=(200, 400))
        col2 = Column([[Button('新建', key='-ADD-')],
                       [Button('编辑', key='-EDIT-')],
                       [Button('删除', key='-DELETE-')],
                       [Button('导入', key='-IMPORT-')],
                       [Button('导出', key='-EXPORT-')]],
                      size=(50, 350))
        col3 = [Button('确定', key='-OK-'), Button('取消', key='-CANCEL-')]
        layout = [[col1, col2], [col3]]
        window = Window('管理标签', layout, keep_on_top=True, element_justification='right', finalize=True)

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
            if event in (WIN_CLOSED, '-CANCEL-'):
                break
            elif event == '-OK-':
                break
            elif event == '-ADD-':
                new_tab = self.new_tag()
                print(type(new_tab))
                if isinstance(new_tab, str):
                    values = window['-TAB-LIST-'].get_list_values()
                    values.append(new_tab)
                    index = values.index(new_tab)
                    window['-TAB-LIST-'].update(values, set_to_index=index)
            elif event == '-EDIT-':
                cur_tab = window['-TAB-LIST-'].get()[0]
                print('cur_tab', cur_tab)
                tab_key = os.path.join(self.model.root, cur_tab)
                new_tab = self.mod_tag(tab_key)
                if isinstance(new_tab, str):
                    values = window['-TAB-LIST-'].get_list_values()
                    index = values.index(cur_tab)
                    values[index] = new_tab
                    window['-TAB-LIST-'].update(values, set_to_index=index)
            elif event == '-DELETE-':
                cur_tab = window['-TAB-LIST-'].get()[0]
                tab_key = os.path.join(self.model.root, cur_tab)
                if self.rmv_tag(tab_key):
                    values = window['-TAB-LIST-'].get_list_values()
                    values.remove(cur_tab)
                    window['-TAB-LIST-'].update(values)
            elif event == '-IMPORT-':
                if self.import_app_info():
                    values = [os.path.basename(tab) for tab, _ in self.model.apps.items()]
                    window['-TAB-LIST-'].update(values)
            elif event == '-EXPORT-':
                self.export_app_info()
            elif event == '-TAB-LIST-':
                tab = values[event][0]
                disable_buttons(tab)
            else:
                print(event, values)
        window.close()

    def new_tag(self):
        result = False
        layout = [[Text('新标签'), Input(key='-NEW-TAB-')],
                  [Button('确定'), Button('取消')]]
        window = Window('新建标签', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                tab_name = window['-NEW-TAB-'].get()
                tab_path = os.path.join(self.model.root, tab_name)
                if not os.path.exists(tab_path):
                    os.mkdir(tab_path)
                result = tab_name
                popup_no_buttons('成功添加' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def mod_tag(self, current_tag):
        result = False
        tab_name = os.path.basename(current_tag)
        tab_path = current_tag
        layout = [[Text('当前标签'), Input(key='CurTab', default_text=tab_name)],
                  [Button('确定'), Button('取消')]]
        window = Window('编辑标签', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                tab_name = window['CurTab'].get()
                if os.path.exists(tab_path):
                    os.renames(tab_path, os.path.join(self.model.root, tab_name))
                result = tab_name
                popup_no_buttons('成功修改为' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def rmv_tag(self, current_tag):
        tab_name = os.path.basename(current_tag)
        tab_dir = current_tag
        result = popup('是否删除' + tab_name + '?', keep_on_top=True, button_type=POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(tab_dir)
            self.refresh()
        return result

    def add_app(self, current_tag):
        result = False
        layout = [[Text('添加目标'), Input(key='TargetPath'), Button('浏览')],
                  [Text('添加应用'), Input(key='AppName'), Button('图标')],
                  [Button('确定'), Button('取消')]]
        window = Window('添加应用', layout, element_justification='right')
        while True:
            event, values = window.read()
            if event == '浏览':
                file = popup_get_file('选择目标文件', keep_on_top=True)
                if file:
                    window['TargetPath'].update(value=file)
                    app_name = os.path.basename(str(file)).capitalize().removesuffix('.exe')
                    window['AppName'].update(value=app_name)
            elif event == WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                target = window['TargetPath'].get()
                app_name = window['AppName'].get()
                app_dir = os.path.join(current_tag, app_name)
                if not os.path.exists(app_dir):
                    os.mkdir(app_dir)
                link_path = os.path.join(app_dir, app_name + '.lnk')
                # print('target', target)
                # print('link_path', link_path)
                winshell.CreateShortcut(Path=link_path, Target=target, Icon=(target, 0))
                popup('成功添加应用' + app_name + '!', no_titlebar=True, keep_on_top=True, auto_close=True,
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
        layout = [[Text('应用名称'), Input(key='AppName', default_text=app_name)],
                  [Button('确定'), Button('取消')]]
        window = Window('编辑应用', layout)
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '取消':
                break
            elif event == '确定':
                new_app_name = window['AppName'].get()
                os.renames(app_dir, os.path.join(os.path.dirname(app_dir), new_app_name))
                popup(app_name + '成功修改成' + new_app_name + '!', no_titlebar=True, keep_on_top=True,
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
        result = popup('是否删除' + app_name + '?', no_titlebar=True, keep_on_top=True,
                       button_type=POPUP_BUTTONS_YES_NO)
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

    def copy_app_path(self):
        item = self.window.find_element_with_focus()
        app_key = item.key
        cur_path = app_key.split('#')[1]
        pyperclip.copy(cur_path)

    def run_app(self, event):
        key = event.replace(' DoubleClick', '')
        element = self.window[key]
        if not element:
            element = self.window.find_element_with_focus()
        app = element.metadata
        print(app)
        if app['type'] in ['category', 'folder']:
            self.refresh(current_tag=key)
        else:
            app_path = app['path']
            if os.path.isfile(app_path):
                os.startfile(app_path)
            elif os.path.isdir(app_path):
                for app_name in os.listdir(app_path):
                    if app_name.endswith((".lnk", ".exe")):
                        app = os.path.join(app_path, app_name)
                        os.startfile(app)
                        break
            elif validators.url(app_path):
                print('url', app_path)
                webbrowser.open(app_path)
            else:
                print('invalid path', app_path)

    def _get_apps_of_current_tag(self, tag, apps):
        current_tag_list = self.model.current_tag.split('#')
        if len(current_tag_list) == 1:
            return tag, apps
        else:
            if tag == current_tag_list[0]:
                for i in range(len(current_tag_list[1:])):
                    for app in apps:
                        if app['name'] == current_tag_list[i]:
                            tag += "#" + current_tag_list[i]
                            apps = app['apps']
                            break
        return tag, apps

    def _init_content(self):
        print(time.perf_counter())
        tab_group = {}
        max_num_per_row = 1 if self.model.show_view == SHOW_LIST else 4
        app_right_click_menu = [[], [E_RUN_APP, E_MOD_APP, E_RMV_APP, E_OPEN_APP_PATH, E_COPY_APP_PATH]]
        tag_right_click_menu = [[], [E_ADD_TAG, E_RMV_TAG, E_SRT_APP, E_ADD_APP, E_REFRESH]]
        for tag, apps in self.model.apps.items():
            # print('tag', tag, 'show_browser', self.model.show_browser)
            if not self.model.show_local and tag == LOCAL_APPS:
                continue
            if not self.model.show_browser and tag in [CHROME_BOOKMARKS, EDGE_BOOKMARKS, FIREFOX_BOOKMARKS]:
                continue
            tag, apps = self._get_apps_of_current_tag(tag, apps)
            tab_layout = []
            row_layout = []
            if self.model.sort_order == SORT_ASCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper())
            elif self.model.sort_order == SORT_DESCEND:
                apps = sorted(apps, key=lambda x: x['name'].upper(), reverse=True)
            for app in apps:
                col_layout = []
                app_name, app_type, app_icon, app_path = app['name'], app['type'], app['icon'], app['path']
                app_key = tag + '#' + app_name
                if not app_icon.endswith('.png'):
                    if app_type in ['category', 'bookmark', 'folder']:
                        app_icon = self.model.icons['bookmark']
                    elif app_type in ['app', 'website', 'url']:
                        app_icon = random.choice(self.model.icons['default'])
                    elif app_type in ['local']:
                        app_icon = find_or_make_icon(app_icon, os.path.join(APP_ICON_PATH, app_type))
                    else:
                        app_icon = ''
                icon = Image(source=app_icon, tooltip=app_name, key='[icon]' + app_key, enable_events=True,
                             size=(32, 32), background_color='white', right_click_menu=app_right_click_menu)
                col_layout.append(icon)
                button = Text(text=app_name, tooltip=app_name, key=app_key, metadata=app,
                              size=(30, 1), background_color='white', enable_events=False, border_width=0,
                              right_click_menu=app_right_click_menu, expand_x=True)
                col_layout.append(button)
                if self.model.show_view == SHOW_LIST:
                    date_time = Text(app['time'], background_color='white', size=(20, 1), expand_x=True)
                    col_layout.append(date_time)
                    path = Text(text=app_path, tooltip=app_path, size=(200, 1), justification='left',
                                background_color='white', enable_events=True)
                    col_layout.append(path)
                    col = Column([col_layout], background_color='white', expand_x=True, expand_y=True,
                                 right_click_menu=app_right_click_menu)
                else:
                    col = Column([col_layout], background_color='white', right_click_menu=app_right_click_menu)

                if self.model.search_content == '' or re.search(self.model.search_content, app_name, re.IGNORECASE):
                    row_layout.append(col)
                if len(row_layout) == max_num_per_row:
                    tab_layout.append(copy.deepcopy(row_layout))
                    row_layout.clear()
            else:
                if len(row_layout) > 0:
                    tab_layout.append(row_layout)
            tab_group[tag] = [[Column(layout=tab_layout, scrollable=True, vertical_scroll_only=True,
                                      expand_x=True, expand_y=True)]]
        tab_group_layout = [[Tab(title=os.path.basename(tab_key), layout=tab_layout, key=tab_key,
                                 right_click_menu=tag_right_click_menu)
                             for tab_key, tab_layout in tab_group.items()]]
        print(time.perf_counter())
        return [[TabGroup(tab_group_layout, key=E_ACTIVE_TAG, expand_x=True, expand_y=True,
                          enable_events=True, right_click_menu=tag_right_click_menu)]]

    def _bind_hotkeys(self):
        for tag, apps in self.model.apps.items():
            if not self.model.show_local and tag == LOCAL_APPS:
                continue
            if not self.model.show_browser and tag in [FIREFOX_BOOKMARKS, CHROME_BOOKMARKS, EDGE_BOOKMARKS]:
                continue
            tag, apps = self._get_apps_of_current_tag(tag, apps)
            for app in apps:
                key = tag + '#' + app['name']
                if self.model.search_content == '' or re.search(self.model.search_content, app['name'], re.IGNORECASE):
                    if self.window[key]:
                        self.window[key].bind('<Double-Button-1>', ' DoubleClick')
                        self.window[key].bind('<Button-1>', ' LeftClick')
                        self.window[key].bind('<Button-2>', ' MiddleClick')
                        self.window[key].bind('<Button-3>', ' RightClick')
        for event, hotkeys in self.model.event_hotkeys.items():
            for hotkey in hotkeys:
                self.window.bind(hotkey, event)
