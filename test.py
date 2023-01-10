import os
from turtle import title

import psutil
from pywinauto import Application, taskbar, mouse
from uiautomation import uiautomation


def main():
    # 检查 app 是否启动
    app = 'workhelper.exe'
    pid = -1
    for p in psutil.process_iter():
        if p.name().find(app) > -1:
            pid = p.pid
            print(p.pid, p.name(), p.cmdline())
            break

    # 启动 app
    if pid == -1:
        print('app start running')
        cmd = os.path.join(os.getcwd(), app)
        app = Application().start(cmd_line=cmd)
    else:
        print('app already run')
        app = Application().connect(process=pid)

    def __get_element_postion(element):
        element_position = element.rectangle()
        center_position = (int((element_position.left + element_position.right) / 2),
                           int((element_position.top + element_position.bottom) / 2))
        print(center_position)
        return center_position

    # 点击菜单
    # ev_frame = uiautomation.WindowControl(searchDepth=1, Name="工作助手")
    # ev_frame.SetTopmost(True)
    # ev_frame.TextControl(Name="场景编辑").Click()
    # ev_frame.Element
    # print(ev_frame.ButtonControl(searchDepth=7))
    dlg = app.window(title=u'工作助手', class_name='TkTopLevel')
    edit = dlg.child_window(title='UISpy', control_type="Button")
    print(dlg.print_ctrl_ids())
    mouse.click(button='left', coords=__get_element_postion(edit))
    # click on Outlook's icon
    # taskbar.ClickSystemTrayIcon("工作助手")
    # dlg = app.top_window()
    # dlg.print_control_identifiers()
    # dlg['TkChild'].
    # dlg.menu_select('编辑')
    # Select an item in the popup menu
    # print(app.PopupMenu.Menu())
    # app.PopupMenu.Menu().get_menu_path("Cancel Server Request")[0].click()
    # 检查项


if __name__ == '__main__':
    main()
