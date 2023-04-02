# -*- coding: utf-8 -*-
import ctypes
import datetime
import os
import win32gui
import win32ui
import winreg
from PIL import Image
from time import localtime, strftime


def set_env(key, value):
    if os.getenv(key) != value:
        command = f'start /b setx {key} "{value}"'.format(key=key, value=value)
        print(command)
        os.popen(command)


def date_str_to_datetime(date_str):
    formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]} 00:00:00"
    return formatted_date


def timestamp_to_datetime(time_stamp, format_string="%Y-%m-%d %H:%M:%S"):
    time_array = localtime(time_stamp)
    str_date = strftime(format_string, time_array)
    return str_date


def find_or_make_icon(icon_filepath, png_path):
    filename = os.path.basename(icon_filepath).split('.')[0] + '.png'
    png_filename = os.path.join(png_path, filename)
    if os.path.exists(png_filename):
        return png_filename
    if icon_filepath.find(".exe") != -1:
        large, _small = win32gui.ExtractIconEx(icon_filepath, 0)
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, 32, 32)
        hdc = hdc.CreateCompatibleDC()
        hdc.SelectObject(hbmp)
        hdc.DrawIcon((0, 0), large[0])
        bmp_bits = hbmp.GetBitmapBits(True)
        img = Image.frombuffer('RGBA', (32, 32), bmp_bits, 'raw', 'BGRA', 0, 1)
        img.save(png_filename)
    elif icon_filepath.find('.ico') != -1:
        img = Image.open(icon_filepath)
        img = img.resize((32, 32), Image.LANCZOS)
        img.save(png_filename)
    else:
        png_filename = ''
    return png_filename


def get_path_time(path):
    return timestamp_to_datetime(os.path.getmtime(path))


def formate_time(time):
    return datetime.datetime.fromtimestamp(int(time) / 1000000).strftime('%Y-%m-%d %H:%M:%S')


class Registry(object):

    def __init__(self, key_location, key_path):
        self.reg_key = winreg.OpenKey(key_location, key_path, 0, winreg.KEY_ALL_ACCESS)

    def set_key(self, name, value):
        try:
            _, reg_type = winreg.QueryValueEx(self.reg_key, name)
        except WindowsError:
            # If the value does not exist yet, we (guess) use a string as the reg_type
            reg_type = winreg.REG_SZ
        winreg.SetValueEx(self.reg_key, name, 0, reg_type, value)

    def delete_key(self, name):
        try:
            winreg.DeleteValue(self.reg_key, name)
        except WindowsError:
            # Ignores if the key value doesn't exist
            pass


class EnvVar(Registry):
    """
    Configures the HTTP_PROXY environment variable
    """

    def __init__(self):
        super(EnvVar, self).__init__(
            winreg.HKEY_CURRENT_USER,
            'Environment')

    def on(self, key, value):
        self.set_key(key, value)
        self.refresh()

    def off(self, key):
        self.delete_key(key)
        self.refresh()

    @staticmethod
    def refresh():
        hwnd_broadcast = 0xFFFF
        wm_settingchange = 0x1A
        smto_abortifhung = 0x0002
        result = ctypes.c_long()
        ctypes.windll.user32.SendMessageTimeoutW(hwnd_broadcast, wm_settingchange, 0, u'Environment',
                                                 smto_abortifhung, 5000, ctypes.byref(result))


if __name__ == '__main__':
    ev = EnvVar()
    ev.on('My', 'Test')
    ev.off('My')
