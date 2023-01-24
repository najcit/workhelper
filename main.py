"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""

import os
import shutil
from contextlib import suppress

import click
import winshell
import PySimpleGUI as sg
from app import MyApp


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


def mod_tab(root, cur_tab):
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
            os.renames(os.path.join(root, cur_tab), os.path.join(root, new_tab))
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
@click.option('--enable_tray', is_flag=True, default=False, help='enable system tray for the application')
@click.option('--enable_env', is_flag=True, default=True, help='enable environment variable for the application')
def main(root, theme, enable_tray, enable_env):
    app = MyApp(root, theme, enable_tray, enable_env)
    app.init()
    while True:
        event, values = app.read()
        if event in ('显示界面', sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED):
            app.window.un_hide()
            app.window.bring_to_front()
        elif event in ('隐藏界面', sg.WIN_CLOSE_ATTEMPTED_EVENT):
            if app.systray:
                app.window.hide()
                app.systray.show_icon()
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
                app.update()
        elif event == '选择主题':
            new_theme = set_theme(theme)
            if theme != new_theme:
                theme = new_theme
                app.update()
        elif event == '添加标签':
            if add_tab(root):
                app.update()
        elif event == '修改标签':
            if mod_tab(root, values[0]):
                app.update()
        elif event == '删除标签':
            if rmv_tab(values[0]):
                app.update()
        elif event == '刷新':
            app.update()
        elif event == '添加应用':
            if add_app(values[0]):
                app.update()
        elif event == '修改应用':
            item = app.window.find_element_with_focus()
            if mod_app(item.key):
                app.update()
        elif event == '删除应用':
            item = app.window.find_element_with_focus()
            if rmv_app(item.key):
                app.update()
        elif event == '打开所在位置':
            item = app.window.find_element_with_focus()
            cur_path = str(os.path.join(root, item.key))
            if os.path.isfile(cur_path):
                os.startfile(cur_path)
        elif type(event) == str and event.endswith('Click'):
            key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
            app.window[key].set_focus()
        elif event == '搜索':
            tab = values[0]
            content = values['content']
            filter_condition = {'tab': tab, 'content': content}
            if filter_condition != app.filter_cond:
                app.update(filter_condition)
        else:
            item = app.window.find_element_with_focus()
            if str(item).find('Button') > -1:
                run_app(item.key)
    app.close()


if __name__ == '__main__':
    main()
