# -*- coding: utf-8 -*-
import os

from apputil import get_path_time


class AppMyApp(object):

    @staticmethod
    def get_apps(root):
        my_apps = []
        for item in os.listdir(root):
            path = os.path.join(root, item)
            type_ = 'app' if os.path.isfile(path) else 'category'
            name = os.path.splitext(item)[0]
            app = {'name': name, 'path': path, 'type': type_, 'parent': root,
                   'icon': 'default', 'time': get_path_time(path)}
            if type_ == 'category':
                app['apps'] = AppMyApp.get_apps(path)
            my_apps.append(app)
        return my_apps
