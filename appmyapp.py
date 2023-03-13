# -*- coding: utf-8 -*-
import os

from apputil import get_path_time


class AppMyApp(object):
    def __init__(self):
        pass

    def __del__(self):
        pass

    @staticmethod
    def get_apps(root):
        my_apps = []
        for item in os.listdir(root):
            path = os.path.join(root, item)
            type_ = 'app' if os.path.isfile(path) else 'category'
            name = os.path.splitext(item)[0]
            apps = AppMyApp.get_apps(path) if type_ == 'category' else []
            my_apps.append({
                'name': name,
                'path': path,
                'icon': 'default',
                'type': type_,
                'time': get_path_time(path),
                'apps': apps
            })
        return my_apps

    @staticmethod
    def get_my_apps(root):
        return AppMyApp.get_apps(root)
