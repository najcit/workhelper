"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""

import os
import copy
import shutil
from math import ceil

import click
import winshell
import PySimpleGUI as sg
from psgtray import SystemTray
from contextlib import suppress


def set_env(key, value):
    if os.getenv(key) != value:
        command = f'start /b setx {key} {value}'.format(key=key, value=value)
        os.popen(command)


def init_window(root, theme, tray):
    set_env('MY_ROOT', root)
    set_env('MY_THEME', theme)

    sg.theme(theme)
    menu = [['编辑', ['管理标签', '添加标签']],
            ['设置', ['打开目录', '选择主题', '自动更新']],
            ['帮助', ['关于']]]
    layout = [[sg.MenubarCustom(menu)]]
    layout += [sg.Input(enable_events=True, key='content', expand_x=True), sg.Button('搜索')],
    whole = '全部应用'
    items_info = {whole: []}
    for parent, directorys, _files in os.walk(root):
        sub_root = parent.replace(root + os.path.sep, '')
        if parent != root and sub_root.find(os.path.sep) == -1:
            if parent not in items_info:
                items_info[parent] = []
            items_info[parent] += [os.path.join(parent, directory) for directory in directorys]
            items_info[whole] += [os.path.join(parent, directory) for directory in directorys]

    tabs = []
    keys = []
    max_num_per_row = max(ceil(len(items_info)/2), 4)
    # print(max_num_per_row)
    app_right_click_menu = [['右击菜单'], ['修改应用', '删除应用', '打开所在位置']]
    tab_right_click_menu = [['右击菜单'], ['修改标签', '删除标签', '添加应用', '刷新']]
    for tab, apps in items_info.items():
        tab_name = os.path.basename(tab).ljust(4, ' ')
        tab_key = tab
        tab_layout = []
        row_layout = []
        for app in apps:
            app_name = os.path.basename(app)
            app_key = tab+'#'+app
            button = sg.Button(app_name, tooltip=app_name, key=app_key, size=(12, 2), mouseover_colors='blue',
                               enable_events=True, right_click_menu=app_right_click_menu)
            row_layout.append(button)
            if len(row_layout) == max_num_per_row:
                tab_layout.append(copy.deepcopy(row_layout))
                row_layout.clear()
            keys.append(app_key)
        else:
            if len(row_layout) > 0:
                tab_layout.append(row_layout)
        tabs.append(sg.Tab(tab_name[:4], tab_layout, key=tab_key, right_click_menu=tab_right_click_menu))

    layout += [[sg.TabGroup([tabs], expand_x=True, expand_y=True, right_click_menu=tab_right_click_menu)]]
    window = sg.Window('工作助手', layout, finalize=True, enable_close_attempted_event=True)
    for key in keys:
        window[key].bind('<Button-1>', ' LeftClick')
        window[key].bind('<Button-3>', ' RightClick')
    menu = ['后台菜单', ['显示界面', '隐藏界面', '退出']]
    sys_tray = SystemTray(menu, single_click_events=False, window=window, tooltip='工作助手') if tray else None
    return window, sys_tray


def update_window(window, tray, root, theme):
    if tray:
        tray.close()
    window.close()
    return init_window(root, theme, tray)


def set_theme(theme):
    cur_theme = theme
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
            print('[LOG] User Chose Theme: ' + str(cur_theme))
            break
    window.close()
    return cur_theme


def add_tab(root):
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
            os.mkdir(os.path.join(root, new_tab))
            sg.popup('You success to new a tab!', auto_close=True, auto_close_duration=1, keep_on_top=True)
            result = True
            break
    tab_window.close()
    return result


def mod_tab(root, current_tab):
    result = False
    layout = [[sg.Text('当前标签'), sg.Input(key='CurTab', default_text=current_tab)],
              [sg.Button('确定'), sg.Button('取消')]]
    tab_window = sg.Window('添加标签', layout)
    while True:
        event, values = tab_window.read()
        if event == sg.WIN_CLOSED or event == '取消':
            break
        elif event == '确定':
            new_tab = tab_window['CurTab'].get()
            os.renames(os.path.join(root, current_tab), os.path.join(root, new_tab))
            sg.popup('You success to new a tab!', auto_close=True, auto_close_duration=1, keep_on_top=True)
            result = True
            break
    tab_window.close()
    return result


def rmv_tab(tab_key):
    tab_name = os.path.basename(tab_key)
    tab_dir = tab_key
    result = sg.popup('是否删除' + tab_name + '?', no_titlebar=True, keep_on_top=True, button_type=sg.POPUP_BUTTONS_YES_NO)
    if result:
        shutil.rmtree(tab_dir)
    return result


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


def rmv_app(app_key):
    app_dir = app_key.split('#')[1]
    app_name = os.path.basename(app_dir)
    result = sg.popup('是否删除' + app_name + '?', no_titlebar=True, keep_on_top=True, button_type=sg.POPUP_BUTTONS_YES_NO)
    if result:
        shutil.rmtree(app_dir)
    return result


def run_app(app_key):
    with suppress(Exception):
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


@click.command()
@click.option('--root', default='root', type=str, help='the root directory of the application')
@click.option('--theme', default='DarkGreen7', type=str, help='the theme of the application')
@click.option('--tray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--disable_env', is_flag=True, default=False, help='disable environment variable for the application')
def main(root, theme, tray, disable_env):
    if not disable_env:
        with suppress(Exception):
            root = os.environ['MY_ROOT']
            theme = os.environ['MY_THEME']
    if root == 'root':
        root = os.path.join(os.getcwd(), root)

    window, sys_tray = init_window(root, theme, tray)
    while True:
        event, values = window.read()
        if sys_tray and event == sys_tray.key:
            event = values[event]

        if event == '__TIMEOUT__':
            pass
        elif event in ('显示界面', sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED):
            window.un_hide()
            window.bring_to_front()
        elif event in ('隐藏界面', sg.WIN_CLOSE_ATTEMPTED_EVENT):
            if sys_tray:
                window.hide()
                sys_tray.show_icon()
            else:
                break
        elif event in (sg.WIN_CLOSED, 'Exit', '退出'):
            break
        elif event == '关于':
            sg.popup('工作助手', 'Copyright © 2010–2023 lidajun')
        elif event == '打开目录':
            folder = sg.popup_get_folder('Choose your folder', keep_on_top=True)
            if folder:
                root = str(folder)
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '选择主题':
            new_theme = set_theme(theme)
            if theme != new_theme:
                theme = new_theme
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '添加标签':
            if add_tab(root):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '修改标签':
            if mod_tab(root, values[0]):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '删除标签':
            if rmv_tab(values[0]):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '刷新':
            window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '添加应用':
            if add_app(values[0]):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '修改应用':
            item = window.find_element_with_focus()
            if mod_app(item.key):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '删除应用':
            item = window.find_element_with_focus()
            if rmv_app(item.key):
                window, sys_tray = update_window(window, sys_tray, root, theme)
        elif event == '打开所在位置':
            item = window.find_element_with_focus()
            cur_path = str(os.path.join(root, item.key))
            os.startfile(cur_path)
        elif event.endswith('Click'):
            key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
            window[key].set_focus()
        else:
            item = window.find_element_with_focus()
            if str(item).find('Button') > -1:
                run_app(item.key)
    if sys_tray:
        sys_tray.close()
    window.close()


if __name__ == '__main__':
    main()
