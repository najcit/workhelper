# -*- coding: utf-8 -*-

import configparser
import datetime
import json
import os
import sqlite3
from pprint import pprint

from appresource import FIREFOX_BOOKMARK, CHROME_BOOKMARK, EDGE_BOOKMARK, IE_BOOKMARK, UNKNOWN_BOOKMARK


class AppBookmark:

    def __init__(self, name):
        self.browser_name = name
        self.bookmark_name = {
            'firefox': FIREFOX_BOOKMARK,
            'chrome': CHROME_BOOKMARK,
            'edge': EDGE_BOOKMARK,
            'ie': IE_BOOKMARK,
        }
        self.bookmarks = {
            'firefox': self.get_firefox_bookmarks,
            'chrome': self.get_chrome_bookmarks,
            'edge': self.get_edge_bookmarks,
            'ie': self.get_ie_bookmarks,
        }

    def get_bookmark_name(self):
        if self.browser_name in self.bookmark_name:
            return self.bookmark_name[self.browser_name]
        return UNKNOWN_BOOKMARK

    def get_bookmarks(self):
        if self.browser_name in self.bookmarks:
            return self.bookmarks[self.browser_name]()
        return []

    @staticmethod
    def get_firefox_bookmarks():
        # 获取配置文件夹 places.sqlite
        mozilla_profile = os.path.join(os.getenv('APPDATA'), r'Mozilla\Firefox')
        mozilla_profile_ini = os.path.join(mozilla_profile, r'profiles.ini')
        profile = configparser.ConfigParser()
        profile.read(mozilla_profile_ini)
        mozilla_data_path = os.path.normpath(os.path.join(mozilla_profile, profile.get('Profile0', 'Path')))
        mozilla_bookmark = os.path.join(mozilla_data_path, r'places.sqlite')
        # print(mozilla_profile_ini, mozilla_data_path, mozilla_bookmark)
        # 读取 places.sqlite
        bookmarks = []
        if os.path.exists(mozilla_bookmark):
            firefox_connection = sqlite3.connect(mozilla_bookmark)
            cursor = firefox_connection.cursor()
            query = "select id,title,dateAdded from moz_bookmarks where type=2 and parent=2 order by position asc;"
            cursor.execute(query)
            for row in cursor:
                menu = {'id': row[0],
                        'name': row[1],
                        'path': os.path.join(FIREFOX_BOOKMARK, row[1]),
                        'icon': 'default',
                        'type': 'category',
                        'time': datetime.datetime.fromtimestamp(row[2] / 1000000).strftime('%Y-%m-%d %H:%M:%S')}
                bookmarks.append(menu)
            for menu in bookmarks:
                parent = menu['id']
                query = f"""select moz_bookmarks.title, moz_places.url, moz_bookmarks.dateAdded
                            from moz_bookmarks,moz_places 
                            where type=1 and parent={parent} and fk=moz_places.id order by position asc;"""
                cursor.execute(query)
                apps = []
                for row in cursor:
                    app = {'name': row[0],
                           'path': row[1],
                           'icon': 'default',
                           'type': 'app',
                           'time': datetime.datetime.fromtimestamp(row[2]/1000000).strftime('%Y-%m-%d %H:%M:%S')}
                    apps.append(app)
                menu['apps'] = apps
        return bookmarks

    @staticmethod
    def get_chrome_bookmarks():
        bookmarks = []
        bookmarks_file = os.path.join(os.getenv('APPDATA'), r'..\Local\Google\Chrome\User Data\Default\Bookmarks')
        if os.path.isfile(bookmarks_file):
            with open(bookmarks_file, encoding='utf-8') as file:
                data = json.load(file)['roots']
                for item in data:
                    id_ = data[item]['id']
                    time = int(data[item]['date_added'])
                    children = data[item]['children']
                    apps = []
                    for child in children:
                        pprint(child)
                        type_ = child['type']
                        time = int(child['date_added'])
                        app = {'name': child['name'],
                               'path': child['url'] if type_ == 'url' else '',
                               'icon': 'default',
                               'type': type_,
                               'time': datetime.datetime.fromtimestamp(time/1000000).strftime('%Y-%m-%d %H:%M:%S')}
                        apps.append(app)
                    menu = {'id': id_,
                            'name': item,
                            'path': os.path.join(CHROME_BOOKMARK, item),
                            'icon': 'default',
                            'type': 'category',
                            'time': datetime.datetime.fromtimestamp(time/1000000).strftime('%Y-%m-%d %H:%M:%S'),
                            'apps': apps}
                    bookmarks.append(menu)
        return bookmarks

    @staticmethod
    def get_edge_bookmarks():
        print('Getting edge bookmark_name')
        return []

    @staticmethod
    def get_ie_bookmarks():
        print('Getting ie bookmark_name')
        return []
