import configparser
import os
import re
import time

import aircv
import cv2
import numpy as np
import psutil
import pyautogui
import win32con
import win32gui


def main():
    app_name = 'workhelper'
    app = app_name + '.exe'
    app_window = u'工作助手'

    # 检查 app 是否启动
    pid = -1
    for p in psutil.process_iter():
        if p.name().find(app) > -1:
            pid = p.pid
            print(p.pid, p.name(), p.cmdline())
            break
    if pid == -1:
        print('app start running')
        app_path = os.path.join(os.getcwd(), app)
        os.popen(app_path)
        time.sleep(4)
    else:
        print('app already run')

    # 显示到屏幕最前端
    pyautogui.hotkey('winleft', 'd')

    def show_window(hwnd, wildcard):
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    win32gui.EnumWindows(show_window, ".*%s.*" % app_window)

    # 依次点击截图标记的位置
    # 获取截图文件
    img_list = []
    hsv_file = ''
    cfg_path = os.path.join(os.getcwd(), app_name)
    for file in os.listdir(cfg_path):
        file_path = os.path.join(cfg_path, file)
        if file.find('.ini') != -1:
            hsv_file = file_path
        elif file_path.find('.png') != -1:
            img_list.append(file_path)
    img_list = sorted(img_list, key=lambda v: int(os.path.basename(v).split('.')[0]))
    # print(img_list)

    # 获取截图标记信息
    cfg = configparser.ConfigParser()
    cfg.read(hsv_file)
    min_hsv = [int(cfg.get('min_hsv', 'h')), int(cfg.get('min_hsv', 's')), int(cfg.get('min_hsv', 'v'))]
    max_hsv = [int(cfg.get('max_hsv', 'h')), int(cfg.get('max_hsv', 's')), int(cfg.get('max_hsv', 'v'))]
    print('min_hsv', min_hsv)
    print('max_hsv', max_hsv)
    min_size = int(cfg.get('hsv_area', 'min_size'))
    max_size = int(cfg.get('hsv_area', 'max_size'))
    print('min_size', min_size)
    print('max_size', max_size)

    # 获取应用程序在屏幕位置
    screen_img = os.path.join(os.getcwd(), 'screen.png')
    pyautogui.screenshot(screen_img)
    app_img = img_list[0]
    src_image = aircv.imread(screen_img)
    dst_image = aircv.imread(app_img)
    result = aircv.find_template(src_image, dst_image)
    os.remove(screen_img)
    if not result or result.get('confidence') < 0.85:
        print('not found target')
        exit(0)
    window = result.get('rectangle')
    print('window', window, window[0])

    # 依次点击截图位置
    for img_file in img_list[1:]:
        img = cv2.imread(img_file)
        img_hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        thresh1 = np.array(min_hsv)
        thresh2 = np.array(max_hsv)
        img_range = cv2.inRange(img_hsv, thresh1, thresh2)
        img_morph = img_range.copy()
        cv2.erode(img_morph, (3, 3), img_morph, iterations=3)
        cv2.dilate(img_morph, (3, 3), img_morph, iterations=3)
        # cv2.imshow('img_morph', img_morph)
        # cv2.waitKey()
        contours, _hierarchy = cv2.findContours(img_morph, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        # print("len(contours)", len(contours))
        if len(contours) > 0:
            sort_contours = sorted(contours, key=cv2.contourArea, reverse=True)
            contour = sort_contours[0]
            # print('contourArea(contour)', cv2.contourArea(contour))
            if min_size <= cv2.contourArea(contour) <= max_size:
                x, y, w, h = cv2.boundingRect(contour)
                # print(x, y, w, h)
                # print(window[0][0]+x+int(w/2), window[0][1]+y+int(h/2))
                pyautogui.click(window[0][0]+x+int(w/2), window[0][1]+y+int(h/2))
                time.sleep(0.5)


if __name__ == '__main__':
    main()
