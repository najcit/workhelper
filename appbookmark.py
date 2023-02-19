import configparser
import datetime
import os
import sqlite3


class AppBookmark:

    def __init__(self, name):
        self.name = name

    def get_bookmarks(self):
        if self.name == 'firefox':
            return FireFoxBookmark.get_bookmarks()
        return []


class FireFoxBookmark:

    @staticmethod
    def get_bookmarks():
        # 获取配置文件夹 places.sqlite
        mozilla_profile = os.path.join(os.getenv('APPDATA'), r'Mozilla\Firefox')
        mozilla_profile_ini = os.path.join(mozilla_profile, r'profiles.ini')
        profile = configparser.ConfigParser()
        profile.read(mozilla_profile_ini)
        mozilla_data_path = os.path.normpath(os.path.join(mozilla_profile, profile.get('Profile0', 'Path')))
        mozilla_bookmark = os.path.join(mozilla_data_path, r'places.sqlite')
        # print(mozilla_profile_ini, mozilla_data_path, mozilla_bookmark)
        # 读取 places.sqlite
        menu_list = []
        if os.path.exists(mozilla_bookmark):
            firefox_connection = sqlite3.connect(mozilla_bookmark)
            cursor = firefox_connection.cursor()
            query = "select id,title,dateAdded from moz_bookmarks where type=2 and parent=2 order by position asc;"
            # x = 'left join moz_places on moz_bookmarks.fk=moz_places.id where type=1;'
            cursor.execute(query)
            for row in cursor:
                menu = {'id': row[0],
                        'name': row[1],
                        'path': 'https://www.baidu.com/',
                        'icon': 'category',
                        'type': 'category',
                        'time': datetime.datetime.fromtimestamp(row[2] / 1000000).strftime('%Y-%m-%d %H:%M:%S')}
                menu_list.append(menu)
            for menu in menu_list:
                query = """select id, title, dateAdded from moz_bookmarks 
                        where type = 1 and parent = {} order by position asc;""".format(menu['id'])
                cursor.execute(query)
                apps = []
                for row in cursor:
                    app = {'name': row[1],
                           'path': str(row[0]),
                           'icon': 'website',
                           'type': 'app',
                           'time': datetime.datetime.fromtimestamp(row[2] / 1000000).strftime('%Y-%m-%d %H:%M:%S')}
                    apps.append(app)
                menu['apps'] = apps
        return menu_list


class GoogleBookmark:

    def __init__(self):
        pass
