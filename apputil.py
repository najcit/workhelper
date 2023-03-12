# -*- coding: utf-8 -*-
import datetime
import os
from time import localtime, strftime

import win32gui
import win32ui
from PIL import Image


def set_env(key, value):
    if os.getenv(key) != value:
        command = f'start /b setx {key} "{value}"'.format(key=key, value=value)
        print(command)
        os.popen(command)


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
