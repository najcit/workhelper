import contextlib

import requests


class AppSharedApp(object):

    @staticmethod
    def get_apps(url='http://127.0.0.1:5000'):
        with contextlib.suppress(Exception):
            response = requests.get(url + '/query/apps')
            return response.json()
        return []
