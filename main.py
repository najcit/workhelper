"""
    A simple workstation that can include some self-customization tools
    can help developers work.
"""

import os
import copy
import winshell
import subprocess
import PySimpleGUI as sg
from psgtray import SystemTray


def set_env(key, value):
    if os.getenv(key) != value:
        command = f"start /b setx {key} {value}".format(key=key, value=value)
        print(command)
        os.system(command)


def make_window(root, theme):
    print(root, theme)
    set_env('MY_ROOT', root)
    set_env('MY_THEME', theme)
    sg.theme(theme)
    menu = [['编辑', ['管理标签', '添加标签']],
            ['设置', ['打开目录', '选择主题']],
            ['帮助', ['关于']]]
    layout = [[sg.MenubarCustom(menu)]]
    tabs = []
    keys = []
    right_click_menu = [['右击菜单'], ['修改项', '删除项', '打开所在位置']]
    if not os.path.exists(root):
        root = os.getcwd()
    for sub_root in os.listdir(root):
        if os.path.isfile(os.path.join(root, sub_root)):
            continue
        # print(sub_root)
        sub_layout = []
        row_layout = []
        items = os.listdir(os.path.join(root, sub_root))
        for i in range(len(items)):
            if os.path.isfile(os.path.join(root, sub_root, items[i])):
                continue
            # print(items[i])
            key = os.path.join(sub_root, items[i])
            keys.append(key)
            row_layout.append(sg.Button(items[i], key=key, size=(12, 2), mouseover_colors='blue',
                                        enable_events=True, right_click_menu=right_click_menu))
            if len(row_layout) == 4:
                sub_layout.append(copy.deepcopy(row_layout))
                row_layout.clear()
        else:
            if len(row_layout) > 0:
                sub_layout.append(copy.deepcopy(row_layout))
        tabs.append(sg.Tab(sub_root, sub_layout, key=sub_root))

    right_click_menu = [['右击菜单'], ['修改标签', '删除标签', '添加项', '刷新']]
    layout += [[sg.TabGroup([tabs], size=(460, 560), expand_x=True, expand_y=True, right_click_menu=right_click_menu)]]
    window = sg.Window('工作助手', layout, finalize=True, enable_close_attempted_event=True)
    for key in keys:
        # print(key)
        window[key].bind("<Button-3>", ' RightClick')
        window[key].bind("<Button-1>", ' LeftClick')

    menu = ['后台菜单', ['显示界面', '隐藏界面', '退出']]
    tray = SystemTray(menu, single_click_events=False, window=window, tooltip='工作助手')
    return window, tray


def upate_window(window, tray, root, theme):
    window.close()
    tray.close()
    return make_window(root, theme)


def set_theme(theme):
    cur_theme = theme
    layout = [[sg.Text("See how elements look under different themes by choosing a different theme here!")],
              [sg.Listbox(values=sg.theme_list(), default_values=[cur_theme], size=(20, 12), key='THEME_LISTBOX',
                          enable_events=True)],
              [sg.Button("确定"), sg.Button('取消')]]
    window = sg.Window('设置主题', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '取消':
            break
        elif event == '确定':
            cur_theme = values['THEME_LISTBOX'][0]
            print("[LOG] User Chose Theme: " + str(cur_theme))
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
            sg.popup("You success to new a tab!", auto_close=True, auto_close_duration=1, keep_on_top=True)
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
            sg.popup("You success to new a tab!", auto_close=True, auto_close_duration=1, keep_on_top=True)
            result = True
            break
    tab_window.close()
    return result


def rmv_tab(root, tab):
    result = sg.popup("确认是否删除" + tab + "?", button_type=sg.POPUP_BUTTONS_YES_NO)
    if result:
        os.rmdir(os.path.join(root, tab))
    return result


def add_item(root, tab):
    result = False
    layout = [[sg.Text('添加目标'), sg.Input(key='NewTarget'), sg.Button('浏览')],
              [sg.Text('新添加项'), sg.Input(key='NewItem'), sg.Button('图标')],
              [sg.Button('确定'), sg.Button('取消')]]
    window = sg.Window('添加项', layout, element_justification='right')
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '取消':
            break
        elif event == '确定':
            target = window['NewTarget'].get()
            # print('target', target)
            item = window['NewItem'].get()
            # print('item', item)
            item_path = os.path.join(os.getcwd(), root, tab, item)
            # print('item_path', item_path)
            if not os.path.exists(item_path):
                os.mkdir(item_path)
            link_path = os.path.join(item_path, item + '.lnk')
            # print('link_path', link_path)
            winshell.CreateShortcut(Path=link_path, Target=target, Icon=(target, 0))
            sg.popup("成功添加一个项!", auto_close=True, auto_close_duration=1, keep_on_top=True)
            result = True
            break
        elif event == '浏览':
            file = sg.popup_get_file('选择目标文件', keep_on_top=True)
            if file:
                print(str(file))
                window['NewTarget'].update(value=file)
                item = os.path.basename(str(file)).capitalize().removesuffix('.exe')
                window['NewItem'].update(value=item)

    window.close()
    return result


def mod_item(root, item):
    result = False
    tab = item.split(os.path.sep)[0]
    default_text = item.split(os.path.sep)[1]
    layout = [[sg.Text('当前项'), sg.Input(key='CurItem', default_text=default_text)],
              [sg.Button('确定'), sg.Button('取消')]]
    window = sg.Window('添加项', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '取消':
            break
        elif event == '确定':
            new_item = window['CurItem'].get()
            os.renames(os.path.join(root, item), os.path.join(root, tab, new_item))
            sg.popup("You success to new a item!", auto_close=True, auto_close_duration=1, keep_on_top=True)
            result = True
            break
    window.close()
    return result


def rmv_item(root, item):
    result = sg.popup("确认是否删除" + item + "?", button_type=sg.POPUP_BUTTONS_YES_NO)
    if result:
        os.remove(os.path.join(root, item))
    return result


def run_item(root, item):
    app_dir = os.path.join(root, item)
    items = os.listdir(app_dir)
    try:
        for item in items:
            if item.endswith('.lnk'):
                app = os.path.join(app_dir, item)
                os.startfile(app)
                break
            elif item.endswith('.exe'):
                app = os.path.join(app_dir, item)
                subprocess.Popen(app)
                break
    except:
        pass


if __name__ == '__main__':
    try:
        root = os.environ['MY_ROOT']
    except:
        root = os.path.join(os.getcwd(), 'root')

    try:
        theme = os.environ['MY_THEME']
    except:
        theme = 'DarkGreen7'

    window, tray = make_window(root, theme)

    while True:
        event, values = window.read()
        if event == tray.key:
            event = values[event]

        if event == '__TIMEOUT__':
            pass
        elif event in ('显示界面', sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED):
            window.un_hide()
            window.bring_to_front()
        elif event in ('隐藏界面', sg.WIN_CLOSE_ATTEMPTED_EVENT):
            window.hide()
            tray.show_icon()
        elif event in (sg.WIN_CLOSED, 'Exit', '退出'):
            break
        elif event == '关于':
            sg.popup('工作助手', 'Copyright © 2010–2023 lidajun')
        elif event == '打开目录':
            folder = sg.popup_get_folder('Choose your folder', keep_on_top=True)
            if folder:
                root = str(folder)
                window, tray = upate_window(window, tray, root, theme)
        elif event == '选择主题':
            new_theme = set_theme(theme)
            if theme != new_theme:
                theme = new_theme
                window, tray = upate_window(window, tray, root, theme)
        elif event == '添加标签':
            if add_tab(root):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '修改标签':
            if mod_tab(root, values[0]):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '删除标签':
            if rmv_tab(root, values[0]):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '刷新':
            window, tray = upate_window(window, tray, root, theme)
        elif event == '添加项':
            print(event, values)
            if add_item(root, values[0]):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '修改项':
            item = window.find_element_with_focus()
            if mod_item(root, item.key):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '删除项':
            item = window.find_element_with_focus()
            if rmv_item(root, item.key):
                window, tray = upate_window(window, tray, root, theme)
        elif event == '打开所在位置':
            item = window.find_element_with_focus()
            cur_path = str(os.path.join(root, item.key))
            os.startfile(cur_path)
        elif event.endswith('Click'):
            key = event.replace('Click', '').replace('Right', '').replace('Left', '').strip()
            window[key].set_focus()
        else:
            item = window.find_element_with_focus()
            if str(type(item)).find('Button') != -1:
                run_item(root, item.key)
    tray.close()
    window.close()
