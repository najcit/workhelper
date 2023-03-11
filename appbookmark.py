# -*- coding: utf-8 -*-

import configparser
import json
import os
import sqlite3

from appresource import FIREFOX_BOOKMARKS, CHROME_BOOKMARKS, EDGE_BOOKMARKS, UNKNOWN_BOOKMARKS
from apputil import formate_time


class AppBookmark:

    def __init__(self, name):
        self.browser_name = name
        self.bookmark_name = {
            'firefox': FIREFOX_BOOKMARKS,
            'chrome': CHROME_BOOKMARKS,
            'edge': EDGE_BOOKMARKS,
        }
        self.bookmarks = {
            'firefox': self.get_firefox_bookmarks,
            'chrome': self.get_chrome_bookmarks,
            'edge': self.get_edge_bookmarks,
        }

    def get_bookmark_name(self):
        if self.browser_name in self.bookmark_name:
            return self.bookmark_name[self.browser_name]
        return UNKNOWN_BOOKMARKS

    def get_bookmarks(self):
        if self.browser_name in self.bookmarks:
            return self.bookmarks[self.browser_name]()
        return []

    @staticmethod
    def get_firefox_bookmarks():
        mozilla_profile = os.path.join(os.getenv('APPDATA'), r'Mozilla\Firefox')
        mozilla_profile_ini = os.path.join(mozilla_profile, r'profiles.ini')
        profile = configparser.ConfigParser()
        profile.read(mozilla_profile_ini)
        mozilla_data_path = os.path.normpath(os.path.join(mozilla_profile, profile.get('Profile0', 'Path')))
        mozilla_bookmark = os.path.join(mozilla_data_path, r'places.sqlite')
        # print(mozilla_profile_ini, mozilla_data_path, mozilla_bookmark)
        bookmarks = []
        if os.path.exists(mozilla_bookmark):
            firefox_connection = sqlite3.connect(mozilla_bookmark)
            cursor = firefox_connection.cursor()
            query = "select id,title,dateAdded from moz_bookmarks where type=2 and parent=2 order by position asc;"
            cursor.execute(query)
            for row in cursor:
                menu = {'id': row[0],
                        'name': row[1],
                        'path': os.path.join(FIREFOX_BOOKMARKS, row[1]),
                        'icon': 'default',
                        'type': 'category',
                        'time': formate_time(row[2])}
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
                           'time': formate_time(row[2])}
                    apps.append(app)
                menu['apps'] = apps
        return bookmarks

    @staticmethod
    def get_children(root):
        bookmarks = []
        if isinstance(root, dict):
            for key, value in root.items():
                id_ = value['id']
                name = value['name']
                type_ = value['type']
                path = value['url'] if type_ == 'url' else ''
                time = formate_time(value['date_added'])
                apps = AppBookmark.get_children(value['children']) if type_ == 'folder' else []
                if type_ == 'folder' and apps == []:
                    continue
                bookmark = {'id': id_,
                            'name': name,
                            'path': path,
                            'icon': 'default',
                            'type': type_,
                            'time': time,
                            'apps': apps}
                bookmarks.append(bookmark)
        elif isinstance(root, list):
            for item in root:
                id_ = item['id']
                name = item['name']
                type_ = item['type']
                path = item['url'] if type_ == 'url' else ''
                time = formate_time(item['date_added'])
                apps = AppBookmark.get_children(item['children']) if type_ == 'folder' else []
                if type_ == 'folder' and apps == []:
                    continue
                bookmark = {'id': id_,
                            'name': name,
                            'path': path,
                            'icon': 'default',
                            'type': type_,
                            'time': time,
                            'apps': apps}
                bookmarks.append(bookmark)
        return bookmarks

    @staticmethod
    def get_chrome_bookmarks():
        bookmarks = []
        bookmarks_file = os.path.join(os.getenv('APPDATA'), r'..\Local\Google\Chrome\User Data\Default\Bookmarks')
        if os.path.isfile(bookmarks_file):
            with open(bookmarks_file, encoding='utf-8') as file:
                root = json.load(file)['roots']
                bookmarks = AppBookmark.get_children(root['bookmark_bar']['children'])
        return bookmarks

    @staticmethod
    def get_edge_bookmarks():
        bookmarks = []
        bookmarks_file = os.path.join(os.getenv('APPDATA'), r'..\Local\Microsoft\Edge\User Data\Default\Bookmarks')
        if os.path.isfile(bookmarks_file):
            with open(bookmarks_file, encoding='utf-8') as file:
                root = json.load(file)['roots']
                bookmarks = AppBookmark.get_children(root['bookmark_bar']['children'])
        return bookmarks
