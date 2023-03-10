# -*- coding: utf-8 -*-

import copy
import os
import random
import re
import shutil
import webbrowser

import pyperclip
import validators
import winshell
import yaml
from PySimpleGUI import TIMEOUT_KEY, theme, Menu, Sizegrip, Tab, TabGroup, Image, Listbox, theme_list
from PySimpleGUI import WIN_CLOSED, \
    popup_get_folder, Text, Button, Window, popup_no_buttons, popup_get_file, Column, \
    LISTBOX_SELECT_MODE_SINGLE, Input, popup, POPUP_BUTTONS_YES_NO

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
        theme(self.model.theme)
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
        col_layout = [[Text('????????????' + self.model.name, key='-NOTIFICATION-'), Sizegrip()]]
        layout += [[Column(col_layout, expand_x=True)]]
        if self.model.location:
            self.window = Window(self.model.name, layout, location=self.model.location,
                                 enable_close_attempted_event=True,
                                 resizable=True, finalize=True, icon=self.model.icons['window'], font='Courier 12')
        else:
            self.window = Window(self.model.name, layout, enable_close_attempted_event=True,
                                 resizable=True, finalize=True, icon=self.model.icons['window'], font='Courier 12')
        self.window.set_min_size((300, 500))
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

    def search(self, **kwargs):
        values = kwargs['values']
        current_tag = values[E_ACTIVE_TAG]
        filter_content = values['-CONTENT-']
        if current_tag != self.model.current_tag or filter_content != self.model.filter_content:
            self.refresh(current_tag=current_tag, filter_content=filter_content)

    def back(self, **_kwargs):
        current_tag_list = self.model.current_tag.split('#')
        current_tag = current_tag_list[0]
        if current_tag != self.model.current_tag:
            self.refresh(current_tag=current_tag)

    def activate_tag(self, current_tag):
        print(self.model.current_tag)
        if self.model.current_tag.find(current_tag) == -1:
            self.model.current_tag = current_tag
            self.window['-CURRENT-TAG-'].update(value=current_tag)

    def select_app(self, app_key):
        self.window[app_key].set_focus()

    def import_app_info(self):
        result = False
        file = popup_get_file(message='????????????', title='??????????????????', keep_on_top=True)
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
        directory = popup_get_folder(message='????????????', title='??????????????????', keep_on_top=True)
        if directory:
            file_path = os.path.join(directory, 'workhelper.yml')
            app_cfg = {}
            for tab, apps in self.model.apps.items():
                if tab == '????????????':
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
        root = popup_get_folder(message='???????????????', title='?????????????????????', keep_on_top=True)
        if root and os.path.isdir(root):
            self.refresh(root=root)

    def select_theme(self):
        cur_theme = self.model.theme
        layout = [[Text('See how elements look under different themes by choosing a different theme here!')],
                  [Listbox(values=theme_list(), default_values=[cur_theme], size=(20, 12), key='THEME_LISTBOX',
                           enable_events=True)],
                  [Button('??????'), Button('??????')]]
        window = Window('????????????', layout)
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '??????':
                break
            elif event == '??????':
                cur_theme = values['THEME_LISTBOX'][0]
                break
        window.close()
        if cur_theme != self.model.theme:
            set_env('MY_THEME', cur_theme)
            self.refresh(theme=cur_theme)
        return self

    def select_font(self):
        print('Selecting font')

    def set_window(self):
        print('Setting window')

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
            text = '??????????????????????????????????????????: {version}'.format(version=self.model.version)
            popup_no_buttons(text, title=self.model.name, auto_close=True, auto_close_duration=10)
        else:
            async def update_progress(percentage):
                progressbar = self.window['-NOTIFICATION-']
                value = '??????????????????????????????????????? {percentage}%???'.format(percentage=percentage)
                progressbar.update(value=value)
                if percentage >= 100:
                    value = '?????????????????????????????????'
                    progressbar.update(value=value)
                    self.window.write_event_value((E_THREAD, E_DOWNLOAD_END), percentage)

            self.window.start_thread(lambda: installer(update_progress_func=update_progress), ())

    def about(self, version):
        text = f'\n????????????????????????: {version}\nCopyright ?? 2023???2026 by lidajun\n'
        popup_no_buttons(text, title=self.model.name, auto_close=True, auto_close_duration=10)

    def manage_tag(self):
        values = [os.path.basename(tab) for tab, _ in self.model.apps.items()]
        cur_tab = os.path.basename(self.window[E_ACTIVE_TAG].get())
        col1 = Column([[Listbox(values, key='-TAB-LIST-', default_values=[cur_tab],
                                select_mode=LISTBOX_SELECT_MODE_SINGLE, enable_events=True, size=(30, 40))]],
                      size=(200, 400))
        col2 = Column([[Button('??????', key='-ADD-')],
                       [Button('??????', key='-EDIT-')],
                       [Button('??????', key='-DELETE-')],
                       [Button('??????', key='-IMPORT-')],
                       [Button('??????', key='-EXPORT-')]],
                      size=(50, 350))
        col3 = [Button('??????', key='-OK-'), Button('??????', key='-CANCEL-')]
        layout = [[col1, col2], [col3]]
        window = Window('????????????', layout, keep_on_top=True, element_justification='right', finalize=True)

        def disable_buttons(tab):
            if tab == '????????????':
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
        layout = [[Text('?????????'), Input(key='-NEW-TAB-')],
                  [Button('??????'), Button('??????')]]
        window = Window('????????????', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '??????':
                break
            elif event == '??????':
                tab_name = window['-NEW-TAB-'].get()
                tab_path = os.path.join(self.model.root, tab_name)
                if not os.path.exists(tab_path):
                    os.mkdir(tab_path)
                result = tab_name
                popup('????????????' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def mod_tag(self, current_tag):
        result = False
        tab_name = os.path.basename(current_tag)
        tab_path = current_tag
        layout = [[Text('????????????'), Input(key='CurTab', default_text=tab_name)],
                  [Button('??????'), Button('??????')]]
        window = Window('????????????', layout, keep_on_top=True, element_justification='right')
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '??????':
                break
            elif event == '??????':
                tab_name = window['CurTab'].get()
                if os.path.exists(tab_path):
                    os.renames(tab_path, os.path.join(self.model.root, tab_name))
                result = tab_name
                popup('???????????????' + tab_name, auto_close=True, auto_close_duration=3, keep_on_top=True)
                break
        window.close()
        if result:
            self.refresh()
        return result

    def rmv_tag(self, current_tag):
        tab_name = os.path.basename(current_tag)
        tab_dir = current_tag
        result = popup('????????????' + tab_name + '?', keep_on_top=True, button_type=POPUP_BUTTONS_YES_NO)
        if result:
            shutil.rmtree(tab_dir)
            self.refresh()
        return result

    def add_app(self, current_tag):
        result = False
        layout = [[Text('????????????'), Input(key='TargetPath'), Button('??????')],
                  [Text('????????????'), Input(key='AppName'), Button('??????')],
                  [Button('??????'), Button('??????')]]
        window = Window('????????????', layout, element_justification='right')
        while True:
            event, values = window.read()
            if event == '??????':
                file = popup_get_file('??????????????????', keep_on_top=True)
                if file:
                    window['TargetPath'].update(value=file)
                    app_name = os.path.basename(str(file)).capitalize().removesuffix('.exe')
                    window['AppName'].update(value=app_name)
            elif event == WIN_CLOSED or event == '??????':
                break
            elif event == '??????':
                target = window['TargetPath'].get()
                app_name = window['AppName'].get()
                app_dir = os.path.join(current_tag, app_name)
                if not os.path.exists(app_dir):
                    os.mkdir(app_dir)
                link_path = os.path.join(app_dir, app_name + '.lnk')
                # print('target', target)
                # print('link_path', link_path)
                winshell.CreateShortcut(Path=link_path, Target=target, Icon=(target, 0))
                popup('??????????????????' + app_name + '!', no_titlebar=True, keep_on_top=True, auto_close=True,
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
        layout = [[Text('????????????'), Input(key='AppName', default_text=app_name)],
                  [Button('??????'), Button('??????')]]
        window = Window('????????????', layout)
        while True:
            event, values = window.read()
            if event == WIN_CLOSED or event == '??????':
                break
            elif event == '??????':
                new_app_name = window['AppName'].get()
                os.renames(app_dir, os.path.join(os.path.dirname(app_dir), new_app_name))
                popup(app_name + '???????????????' + new_app_name + '!', no_titlebar=True, keep_on_top=True,
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
        result = popup('????????????' + app_name + '?', no_titlebar=True, keep_on_top=True,
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
        if app['type'] == 'category':
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

    def _init_content(self):
        tab_group = {}
        max_num_per_row = 1 if self.model.show_view == SHOW_LIST else 4
        app_right_click_menu = [[], [E_RUN_APP, E_MOD_APP, E_RMV_APP, E_OPEN_APP_PATH, E_COPY_APP_PATH]]
        tag_right_click_menu = [[], [E_ADD_TAG, E_RMV_TAG, E_SRT_APP, E_ADD_APP, E_REFRESH]]

        current_tag_list = self.model.current_tag.split('#')
        print('current_tag_list', current_tag_list)
        for tag, apps in self.model.apps.items():
            print('tag', tag, 'show_browser', self.model.show_browser)
            if not self.model.show_local and tag == LOCAL_APP:
                continue
            if not self.model.show_browser and tag in [FIREFOX_BOOKMARK, CHROME_BOOKMARK, IE_BOOKMARK, EDGE_BOOKMARK]:
                continue
            if len(current_tag_list) > 1 and tag == current_tag_list[0]:
                for app in apps:
                    if app['name'] == current_tag_list[1]:
                        tag = self.model.current_tag
                        apps = app['apps']
                        break
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
                        app_icon = random.choice(self.model.icons['app'])
                    elif app_type in ['local']:
                        app_icon = find_or_make_icon(app_icon, os.path.join(APP_ICON_PATH, app_type))
                    else:
                        app_icon = ''
                icon = Image(source=app_icon, tooltip=app_name, key='[icon]' + app_key, enable_events=True,
                             size=(32, 32), background_color='white', right_click_menu=app_right_click_menu)
                col_layout.append(icon)
                button = Text(text=app_name[:20], tooltip=app_name, key=app_key, metadata=app,
                              size=(20, 1), background_color='white', enable_events=False, border_width=0,
                              right_click_menu=app_right_click_menu, expand_x=True)
                col_layout.append(button)
                if self.model.show_view == SHOW_LIST:
                    time = Text(app['time'], background_color='white', expand_x=True)
                    col_layout.append(time)
                    path = Text(text=app_path[:35], tooltip=app_path, size=(35, 1), justification='left',
                                background_color='white', enable_events=True)
                    col_layout.append(path)
                    col = Column([col_layout], background_color='white', expand_x=True, expand_y=True,
                                 right_click_menu=app_right_click_menu, grab=True)
                else:
                    col = Column([col_layout], background_color='white')

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
        return [[TabGroup(tab_group_layout, key=E_ACTIVE_TAG, expand_x=True, expand_y=True,
                          enable_events=True, right_click_menu=tag_right_click_menu)]]

    def _bind_hotkeys(self):
        current_tag_list = self.model.current_tag.split('#')
        # print('current_tag_list', current_tag_list)
        for tag, apps in self.model.apps.items():
            if not self.model.show_local and tag == LOCAL_APP:
                continue
            if not self.model.show_browser and tag in [FIREFOX_BOOKMARK, CHROME_BOOKMARK, IE_BOOKMARK, EDGE_BOOKMARK]:
                continue
            if len(current_tag_list) > 1 and tag == current_tag_list[0]:
                for app in apps:
                    if app['name'] == current_tag_list[1]:
                        tag = self.model.current_tag
                        apps = app['apps']
                        break
            for app in apps:
                key = tag + '#' + app['name']
                if self.window[key]:
                    self.window[key].bind('<Double-Button-1>', ' DoubleClick')
                    self.window[key].bind('<Button-1>', ' LeftClick')
                    self.window[key].bind('<Button-2>', ' MiddleClick')
                    self.window[key].bind('<Button-3>', ' RightClick')
        for event, hotkeys in self.model.event_hotkeys.items():
            for hotkey in hotkeys:
                self.window.bind(hotkey, event)
